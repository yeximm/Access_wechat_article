from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field
import inspect
from threading import Event, Lock, Thread, current_thread
import time
from typing import Any, Callable, Iterable, Mapping
from uuid import uuid4

from .main_flow_models import (
    HomeArticleTarget,
    MainFlowContext,
    SingleArticleOptions,
    SingleArticleReceipt,
)


_FINAL_STATUSES = {"success", "skipped_collected", "failed", "cancelled"}


@dataclass(frozen=True, slots=True)
class SingleArticleUpdate:
    """单篇任务的可序列化阶段更新。"""

    task_id: str
    target_fingerprint: str
    phase: str = "finished"
    status: str = "failed"
    foreground_done: bool = False
    progress_counted: bool = False
    tab_closed: bool = False
    article_saved: bool = False
    detail_status: str = "pending"
    comments_status: str = "not_requested"
    offline_status: str = "not_requested"
    archive_dir: str | None = None
    article_title: str = ""
    message: str = ""
    error_stage: str | None = None
    error_detail: str | None = None
    duration_seconds: float = 0.0
    timestamp: float = field(default_factory=time.time)

    @classmethod
    def from_mapping(
        cls,
        value: Mapping[str, Any],
        *,
        task_id: str,
        target_fingerprint: str,
    ) -> "SingleArticleUpdate":
        status = _normalize_status(value.get("status"))
        phase = str(value.get("phase") or "finished")
        foreground_done = _as_bool(value, "foreground_done", "foregroundDone")
        if phase == "foreground" and _as_bool(value, "tab_closed", "tabClosed"):
            foreground_done = True
        progress_counted = _as_bool(value, "progress_counted", "progressCounted")
        return cls(
            task_id=str(value.get("task_id") or value.get("taskId") or task_id),
            target_fingerprint=str(
                value.get("target_fingerprint")
                or value.get("targetFingerprint")
                or target_fingerprint
            ),
            phase=phase,
            status=status,
            foreground_done=foreground_done,
            progress_counted=progress_counted,
            tab_closed=_as_bool(value, "tab_closed", "tabClosed"),
            article_saved=_as_bool(value, "article_saved", "articleSaved"),
            detail_status=str(
                value.get("detail_status") or value.get("detailStatus") or "pending"
            ),
            comments_status=str(
                value.get("comments_status")
                or value.get("commentsStatus")
                or "not_requested"
            ),
            offline_status=str(
                value.get("offline_status")
                or value.get("offlineStatus")
                or "not_requested"
            ),
            archive_dir=value.get("archive_dir") or value.get("archiveDir"),
            article_title=str(
                value.get("article_title") or value.get("articleTitle") or ""
            ),
            message=str(value.get("message") or ""),
            error_stage=value.get("error_stage") or value.get("errorStage"),
            error_detail=value.get("error_detail") or value.get("errorDetail"),
            duration_seconds=float(
                value.get("duration_seconds")
                or value.get("durationSeconds")
                or value.get("totalSeconds")
                or 0
            ),
        )

    def to_receipt(self) -> SingleArticleReceipt:
        return SingleArticleReceipt(
            task_id=self.task_id,
            target_fingerprint=self.target_fingerprint,
            status=self.status,  # type: ignore[arg-type]
            phase=self.phase,  # type: ignore[arg-type]
            foreground_done=self.foreground_done,
            progress_counted=self.progress_counted,
            tab_closed=self.tab_closed,
            article_saved=self.article_saved,
            detail_status=self.detail_status,
            comments_status=self.comments_status,
            offline_status=self.offline_status,
            archive_dir=self.archive_dir,
            article_title=self.article_title,
            message=self.message,
            error_stage=self.error_stage,
            error_detail=self.error_detail,
            duration_seconds=self.duration_seconds,
        )


@dataclass(slots=True)
class SingleArticleHandle:
    handle_id: str
    target: HomeArticleTarget
    options: SingleArticleOptions
    context: MainFlowContext | Any
    cancel_event: Event = field(default_factory=Event)
    foreground_event: Event = field(default_factory=Event)
    finished_event: Event = field(default_factory=Event)
    thread: Thread | None = None
    latest_receipt: SingleArticleReceipt | None = None
    foreground_receipt: SingleArticleReceipt | None = None
    finished_receipt: SingleArticleReceipt | None = None
    detail_job_id: str | None = None
    updates: deque[SingleArticleUpdate] = field(default_factory=deque)
    post_job_ids: dict[str, str] = field(default_factory=dict)
    post_statuses: dict[str, str] = field(default_factory=dict)
    process_ids: dict[str, str] = field(default_factory=dict)
    foreground_job_id: str | None = None
    started_process_types: set[str] = field(default_factory=set)
    finished_process_types: set[str] = field(default_factory=set)
    cancelled_job_keys: set[tuple[str, str]] = field(default_factory=set)
    post_services_started: bool = False
    lock: Lock = field(default_factory=Lock, repr=False)


class ArticleDispatcher:
    """将单篇任务统一适配为可等待前台回执的执行器。

    生产环境可传入现有 Huey 服务；测试和后续重构可传入 runner。主页窗口
    始终由主流程串行控制，评论和离线归档只通过阶段更新回传结果。
    """

    def __init__(
        self,
        *,
        service: Any | None = None,
        foreground_service: Any | None = None,
        post_services: Mapping[str, Any] | None = None,
        runner: Callable[..., Any] | None = None,
        poll_interval_seconds: float = 0.05,
        cancel_settle_timeout_seconds: float = 2.0,
        id_factory: Callable[[], str] | None = None,
    ) -> None:
        if service is None and foreground_service is None and runner is None:
            raise ValueError("ArticleDispatcher 必须提供 service 或 runner")
        if service is not None and runner is not None:
            raise ValueError("ArticleDispatcher 不能同时提供 service 和 runner")
        if foreground_service is not None and runner is not None:
            raise ValueError("ArticleDispatcher 不能在 runner 模式下提供 foreground_service")
        self._service = service
        self._foreground_service = foreground_service
        self._post_services = {
            str(key): value
            for key, value in (post_services or {}).items()
            if value is not None
        }
        self._runner = runner
        self._poll_interval_seconds = max(0.01, float(poll_interval_seconds))
        self._cancel_settle_timeout_seconds = max(
            self._poll_interval_seconds,
            float(cancel_settle_timeout_seconds),
        )
        self._id_factory = id_factory or (lambda: uuid4().hex[:12])
        self._handles: dict[str, SingleArticleHandle] = {}
        self._lock = Lock()
        self._closed = False

    def submit(
        self,
        target: HomeArticleTarget,
        options: SingleArticleOptions,
        context: MainFlowContext | Any,
    ) -> SingleArticleHandle:
        with self._lock:
            if self._closed:
                raise RuntimeError("单篇任务分发器已经关闭")
            handle = SingleArticleHandle(
                handle_id=f"article-{self._id_factory()}",
                target=target,
                options=options,
                context=context,
            )
            self._handles[handle.handle_id] = handle
            thread = Thread(
                target=self._run_handle,
                args=(handle,),
                name=f"awa-single-article-{handle.handle_id}",
                daemon=True,
            )
            handle.thread = thread
            thread.start()
            self._emit_context_event(
                handle,
                {
                    "type": "article_log",
                    "phase": "precheck",
                    "status": "running",
                    "message": "已接收单篇任务",
                },
            )
            return handle

    def wait_foreground_done(
        self,
        handle: SingleArticleHandle,
        *,
        timeout_seconds: float | None = None,
    ) -> SingleArticleReceipt:
        if not handle.foreground_event.wait(timeout_seconds):
            raise TimeoutError(f"单篇任务未在规定时间返回前台回执：{handle.handle_id}")
        with handle.lock:
            receipt = handle.foreground_receipt or handle.latest_receipt
        if receipt is None:
            raise RuntimeError(f"单篇任务前台回执为空：{handle.handle_id}")
        return receipt

    def wait_finished(
        self,
        handle: SingleArticleHandle,
        *,
        timeout_seconds: float | None = None,
    ) -> SingleArticleReceipt:
        if not handle.finished_event.wait(timeout_seconds):
            raise TimeoutError(f"单篇任务未完成：{handle.handle_id}")
        with handle.lock:
            receipt = handle.finished_receipt or handle.latest_receipt
        if receipt is None:
            raise RuntimeError(f"单篇任务最终回执为空：{handle.handle_id}")
        return receipt

    def poll_updates(self, handle: SingleArticleHandle) -> list[SingleArticleUpdate]:
        with handle.lock:
            result = list(handle.updates)
            handle.updates.clear()
            return result

    def cancel(self, handle: SingleArticleHandle) -> None:
        handle.cancel_event.set()
        self._cancel_jobs(handle)

    def shutdown(self) -> None:
        with self._lock:
            if self._closed:
                return
            self._closed = True
            handles = list(self._handles.values())
        for handle in handles:
            self.cancel(handle)
        for handle in handles:
            thread = handle.thread
            if thread is not None and thread is not current_thread():
                thread.join(timeout=1.0)
        services = []
        if self._service is not None:
            services.append(self._service)
        if self._foreground_service is not None:
            services.append(self._foreground_service)
        services.extend(self._post_services.values())
        seen: set[int] = set()
        for service in services:
            marker = id(service)
            if marker in seen:
                continue
            seen.add(marker)
            close = getattr(service, "shutdown", None)
            if callable(close):
                try:
                    close()
                except Exception:
                    continue

    def _run_handle(self, handle: SingleArticleHandle) -> None:
        foreground_process_id = f"{handle.handle_id}:foreground"
        self._start_process_event(
            handle,
            "foreground",
            foreground_process_id,
        )
        try:
            if self._runner is not None:
                result = self._runner(
                    target=handle.target,
                    options=handle.options,
                    context=handle.context,
                    emit=lambda value: self._emit(handle, value),
                    cancel_event=handle.cancel_event,
                )
                self._consume_runner_result(handle, result)
            else:
                self._run_service(handle)
        except Exception as exc:
            self._emit(
                handle,
                {
                    "phase": "finished",
                    "status": "failed",
                    "message": f"单篇任务执行失败：{type(exc).__name__}: {exc}",
                    "error_stage": "dispatcher",
                    "error_detail": str(exc),
                },
            )
        finally:
            # 无论单篇服务返回什么结果，都不能遗留阶段进程计数。
            # 正常路径会在各阶段最终回执到达时提前结束；这里是异常和取消路径的兜底。
            with handle.lock:
                process_items = tuple(handle.process_ids.items())
            process_items = process_items or (
                ("foreground", foreground_process_id),
                ("detail", f"{handle.handle_id}:detail"),
            )
            for process_type, process_id in process_items:
                self._finish_process_event(handle, process_type, process_id)
            with handle.lock:
                latest = handle.latest_receipt
            if latest is None:
                self._emit(
                    handle,
                    {
                        "phase": "finished",
                        "status": "cancelled"
                        if handle.cancel_event.is_set()
                        else "failed",
                        "message": "单篇任务未返回结果",
                    },
                )
            elif not handle.finished_event.is_set() and latest.status in _FINAL_STATUSES:
                self._emit(
                    handle,
                    {
                        **latest.to_dict(),
                        "phase": "finished",
                    },
                )

    def _consume_runner_result(self, handle: SingleArticleHandle, result: Any) -> None:
        if result is None:
            return
        if isinstance(result, SingleArticleReceipt):
            self._emit(handle, {**result.to_dict(), "phase": result.phase})
            return
        if isinstance(result, Mapping):
            updates = result.get("updates")
            if isinstance(updates, Iterable) and not isinstance(updates, (str, bytes, Mapping)):
                for update in updates:
                    if isinstance(update, Mapping):
                        self._emit(handle, update)
                remainder = {key: value for key, value in result.items() if key != "updates"}
                if any(key in remainder for key in ("status", "phase", "foreground_done", "foregroundDone")):
                    self._emit(handle, remainder)
            else:
                self._emit(handle, result)
            return
        if isinstance(result, Iterable) and not isinstance(result, (str, bytes)):
            for update in result:
                if isinstance(update, Mapping):
                    self._emit(handle, update)

    def _run_service(self, handle: SingleArticleHandle) -> None:
        try:
            if self._uses_split_services():
                self._run_split_service_body(handle)
            else:
                self._run_service_body(handle)
        finally:
            # 阶段结束由 _emit 根据详情回执尽早结束；这里仅作为异常兜底。
            self._finish_process_event(
                handle,
                "detail",
                handle.process_ids.get("detail", f"{handle.handle_id}:detail"),
            )

    def _uses_split_services(self) -> bool:
        foreground = self._foreground_service
        detail = self._service
        return bool(
            foreground is not None
            and detail is not None
            and callable(getattr(foreground, "start_foreground", None))
            and callable(getattr(foreground, "get", None))
            and callable(getattr(detail, "start_detail", None))
            and callable(getattr(detail, "get", None))
        )

    def _run_split_service_body(self, handle: SingleArticleHandle) -> None:
        """前台捕获和详情保存使用两个独立 Huey 任务。"""

        foreground = self._foreground_service
        detail = self._service
        if foreground is None or detail is None:
            raise RuntimeError("前台和详情服务未完整配置")

        payload = _card_payload(handle.target)
        foreground_kwargs = {
            "card_index": handle.target.sequence,
            "account_name": handle.target.account_name,
            "card": payload,
            "skip_collected_records": handle.options.skip_collected_records,
        }
        foreground_result = _call_supported(
            foreground.start_foreground,
            foreground_kwargs,
        )
        foreground_job_id = _job_id_from_result(
            foreground_result,
            "前台 Huey 服务未返回 jobId",
        )
        with handle.lock:
            handle.foreground_job_id = foreground_job_id
        if handle.cancel_event.is_set():
            self._cancel_jobs(handle)

        foreground_finished = False
        detail_started = False
        detail_finished = False
        cancel_deadline: float | None = None

        while True:
            cancelled = handle.cancel_event.is_set()
            if cancelled and cancel_deadline is None:
                cancel_deadline = time.monotonic() + self._cancel_settle_timeout_seconds
                self._cancel_jobs(handle)

            if not foreground_finished:
                try:
                    foreground_payload = foreground.get(foreground_job_id)
                    normalized = _service_update(
                        foreground_payload,
                        default_phase="foreground",
                    )
                    self._emit(handle, normalized)
                    if _is_foreground_finished(normalized):
                        foreground_finished = True
                        if (
                            not cancelled
                            and normalized.get("status") == "success"
                            and bool(normalized.get("foreground_done"))
                        ):
                            detail_job_id = self._start_detail_task(
                                handle,
                                detail,
                                payload,
                                normalized,
                            )
                            with handle.lock:
                                handle.detail_job_id = detail_job_id
                            detail_started = True
                        elif normalized.get("status") == "skipped_collected":
                            self._emit_finished(handle)
                            return
                        elif normalized.get("status") in {"failed", "cancelled"}:
                            self._emit_finished(handle)
                            return
                except Exception as exc:
                    foreground_finished = True
                    self._emit(
                        handle,
                        {
                            "phase": "foreground",
                            "status": "cancelled" if cancelled else "failed",
                            "error_stage": "foreground_poll",
                            "error_detail": str(exc),
                            "foregroundDone": False,
                        },
                    )
                    self._emit_cancelled_finished(
                        handle,
                        "前台任务已取消" if cancelled else "前台任务查询失败",
                    )
                    return

            if detail_started and not detail_finished:
                try:
                    detail_payload = detail.get(handle.detail_job_id or "")
                    normalized = _service_update(
                        detail_payload,
                        default_phase="detail",
                    )
                    self._emit(handle, normalized)
                    if _is_detail_finished(normalized):
                        detail_finished = True
                        if not cancelled and _detail_saved(normalized):
                            self._start_post_tasks(handle, normalized)
                except Exception as exc:
                    detail_finished = True
                    self._emit(
                        handle,
                        {
                            "phase": "detail",
                            "status": "cancelled" if cancelled else "failed",
                            "detailStatus": "cancelled" if cancelled else "failed",
                            "error_stage": "detail_poll",
                            "error_detail": str(exc),
                        },
                    )

            self._poll_post_tasks(handle)
            if detail_finished and self._post_tasks_finished(handle):
                if cancelled:
                    self._emit_cancelled_finished(handle, "单篇任务已取消")
                else:
                    self._emit_finished(handle)
                return

            if cancelled and cancel_deadline is not None:
                if time.monotonic() >= cancel_deadline:
                    self._mark_unsettled_posts_cancelled(handle)
                    self._emit_cancelled_finished(handle, "单篇任务取消后置任务未及时收敛")
                    return
            time.sleep(self._poll_interval_seconds)

    def _start_detail_task(
        self,
        handle: SingleArticleHandle,
        service: Any,
        card: Mapping[str, Any],
        foreground_result: Mapping[str, Any],
    ) -> str:
        result = _call_supported(
            service.start_detail,
            {
                "card_index": handle.target.sequence,
                "account_name": handle.target.account_name,
                "card": dict(card),
                "capture_result": (
                    foreground_result.get("captureResult")
                    or foreground_result.get("capture_result")
                ),
                "foreground_result": dict(foreground_result),
                "store_article_detail": handle.options.collect_article_detail,
            },
        )
        return _job_id_from_result(result, "详情 Huey 服务未返回 jobId")

    def _run_service_body(self, handle: SingleArticleHandle) -> None:
        start = getattr(self._service, "start", None)
        get = getattr(self._service, "get", None)
        if not callable(start) or not callable(get):
            raise TypeError("单篇 Huey 服务必须提供 start() 和 get()")
        payload = _card_payload(handle.target)
        kwargs = {
            "card_index": handle.target.sequence,
            "account_name": handle.target.account_name,
            "card": payload,
            "skip_collected_records": handle.options.skip_collected_records,
            "store_article_detail": handle.options.collect_article_detail,
            "collect_comments": handle.options.collect_comments,
            "store_comment_info": handle.options.collect_comments,
            "archive_offline_content": handle.options.archive_offline,
            "stateful_offline_cache": handle.options.offline_archive_mode == "beta",
        }
        result = _call_supported(start, kwargs)
        job_id = str(result.get("jobId") or result.get("job_id") or "")
        if not job_id:
            raise RuntimeError("Huey 单篇服务未返回 jobId")
        with handle.lock:
            handle.detail_job_id = job_id
        if handle.cancel_event.is_set():
            cancel = getattr(self._service, "cancel", None)
            if callable(cancel):
                try:
                    cancel(job_id)
                except Exception:
                    pass
        detail_finished = False
        cancel_deadline: float | None = None
        while True:
            cancelled = handle.cancel_event.is_set()
            if cancelled and cancel_deadline is None:
                cancel_deadline = time.monotonic() + self._cancel_settle_timeout_seconds
                self._cancel_jobs(handle)

            # 取消后仍读取详情任务的最终状态，避免 Huey job 和详情 attempt
            # 在主流程已经返回后继续无主运行。
            if not detail_finished:
                try:
                    payload_result = get(job_id)
                    normalized = _service_update(payload_result, default_phase="detail")
                    self._emit(handle, normalized)
                    if _is_detail_finished(normalized):
                        detail_finished = True
                        if not cancelled and _detail_saved(normalized):
                            self._start_post_tasks(handle, normalized)
                except Exception as exc:
                    detail_finished = True
                    self._emit(
                        handle,
                        {
                            "phase": "detail",
                            "status": "failed" if not cancelled else "cancelled",
                            "detailStatus": "failed" if not cancelled else "cancelled",
                            "error_stage": "detail_poll",
                            "error_detail": str(exc),
                        },
                    )

            self._poll_post_tasks(handle)
            if detail_finished and self._post_tasks_finished(handle):
                if cancelled:
                    self._emit_cancelled_finished(handle, "单篇任务已取消")
                else:
                    self._emit_finished(handle)
                return

            if cancelled and cancel_deadline is not None:
                if time.monotonic() >= cancel_deadline:
                    self._mark_unsettled_posts_cancelled(handle)
                    self._emit_cancelled_finished(handle, "单篇任务取消后置任务未及时收敛")
                    return
            time.sleep(self._poll_interval_seconds)

    def _start_post_tasks(
        self,
        handle: SingleArticleHandle,
        detail_payload: Mapping[str, Any],
    ) -> None:
        with handle.lock:
            if handle.post_services_started:
                return
            handle.post_services_started = True

        article_payload = _article_post_payload(detail_payload, handle.target)
        for name, service in self._post_services.items():
            enabled = (
                name == "comments" and handle.options.collect_comments
            ) or (
                name == "offline" and handle.options.archive_offline
            )
            if not enabled:
                continue
            start_post = getattr(service, "start_post", None)
            if not callable(start_post):
                self._emit(
                    handle,
                    {
                        "phase": name,
                        "status": "failed",
                        f"{name}_status": "failed",
                        "error_stage": name,
                        "error_detail": "后置服务未提供 start_post 接口",
                    },
                )
                continue
            try:
                post_article_payload = article_payload
                if name == "offline":
                    stateful_offline_cache = handle.options.offline_archive_mode == "beta"
                    post_article_payload = {
                        **article_payload,
                        # 主服务页“离线归档 beta”最终要落到离线缓存子进程的带状态开关。
                        "offlineArchiveMode": handle.options.offline_archive_mode,
                        "statefulOfflineCache": stateful_offline_cache,
                        "stateful_offline_cache": stateful_offline_cache,
                    }
                result = _call_supported(
                    start_post,
                    {
                        "article": post_article_payload,
                        "parent_task_id": handle.handle_id,
                    },
                )
                post_job_id = str(
                    result.get("jobId") or result.get("job_id") or ""
                ) if isinstance(result, Mapping) else ""
                if not post_job_id:
                    raise RuntimeError(f"{name} 后置服务未返回 jobId")
                handle.post_job_ids[name] = post_job_id
                process_id = f"{handle.handle_id}:{name}"
                handle.process_ids[name] = process_id
                self._emit_process_event(handle, name, process_id, "started")
            except Exception as exc:
                handle.post_statuses[name] = "failed"
                self._emit(
                    handle,
                    {
                        "phase": name,
                        "status": "failed",
                        f"{name}_status": "failed",
                        "error_stage": name,
                        "error_detail": str(exc),
                    },
                )

    def _poll_post_tasks(self, handle: SingleArticleHandle) -> None:
        for name, job_id in tuple(handle.post_job_ids.items()):
            service = self._post_services.get(name)
            get = getattr(service, "get", None)
            if not callable(get):
                continue
            try:
                payload = get(job_id)
            except Exception as exc:
                handle.post_statuses[name] = "failed"
                self._emit(
                    handle,
                    {
                        "phase": name,
                        "status": "failed",
                        f"{name}_status": "failed",
                        "error_stage": name,
                        "error_detail": str(exc),
                        "message": f"{name} 后置任务查询失败：{exc}",
                    },
                )
                self._finish_process_event(
                    handle,
                    name,
                    handle.process_ids.get(name, f"{handle.handle_id}:{name}"),
                )
                continue
            normalized = _service_update(payload, default_phase=name)
            self._emit(handle, normalized)
            status = _normalize_status(normalized.get("status"))
            handle.post_statuses[name] = status
            if status in _FINAL_STATUSES:
                self._finish_process_event(
                    handle,
                    name,
                    handle.process_ids.get(name, f"{handle.handle_id}:{name}"),
                )

    def _post_tasks_finished(self, handle: SingleArticleHandle) -> bool:
        for name, job_id in handle.post_job_ids.items():
            del job_id
            if handle.post_statuses.get(name) not in _FINAL_STATUSES:
                return False
        return True

    def _emit_finished(self, handle: SingleArticleHandle) -> None:
        with handle.lock:
            previous = handle.latest_receipt
        if previous is None:
            return
        failed = previous.status in {"failed", "cancelled"} or any(
            value == "failed"
            for value in (previous.comments_status, previous.offline_status)
        )
        if previous.status == "cancelled":
            final_status = "cancelled"
        elif previous.status == "skipped_collected":
            final_status = "skipped_collected"
        else:
            final_status = "failed" if failed else "success"
        self._emit(
            handle,
            {
                **previous.to_dict(),
                "phase": "finished",
                "status": final_status,
            },
        )

    def _emit_cancelled_finished(self, handle: SingleArticleHandle, message: str) -> None:
        """在所有可等待阶段收敛后发送一次不可被成功结果覆盖的取消回执。"""

        self._emit(
            handle,
            {
                "phase": "finished",
                "status": "cancelled",
                "message": message,
            },
        )

    def _cancel_jobs(self, handle: SingleArticleHandle) -> None:
        """按各服务真实 job ID 发出取消；重复调用必须安全。"""

        with handle.lock:
            foreground_job_id = handle.foreground_job_id
            detail_job_id = handle.detail_job_id
            post_jobs = tuple(handle.post_job_ids.items())
        foreground_service = self._foreground_service or self._service
        cancel_foreground = getattr(foreground_service, "cancel", None)
        if callable(cancel_foreground) and foreground_job_id:
            if self._claim_cancel_job(handle, "foreground", foreground_job_id):
                try:
                    cancel_foreground(foreground_job_id)
                except Exception:
                    pass
        cancel = getattr(self._service, "cancel", None)
        if callable(cancel) and detail_job_id:
            if self._claim_cancel_job(handle, "detail", detail_job_id):
                try:
                    cancel(detail_job_id)
                except Exception:
                    pass
        for name, job_id in post_jobs:
            service = self._post_services.get(name)
            post_cancel = getattr(service, "cancel", None)
            if callable(post_cancel) and job_id and self._claim_cancel_job(handle, name, job_id):
                try:
                    post_cancel(job_id)
                except Exception:
                    pass

    @staticmethod
    def _claim_cancel_job(
        handle: SingleArticleHandle,
        service_name: str,
        job_id: str,
    ) -> bool:
        key = (str(service_name), str(job_id))
        with handle.lock:
            if key in handle.cancelled_job_keys:
                return False
            handle.cancelled_job_keys.add(key)
            return True

    def _mark_unsettled_posts_cancelled(self, handle: SingleArticleHandle) -> None:
        """安全超时后把仍未返回的后置阶段标记为取消，并关闭进程计数。"""

        with handle.lock:
            pending = tuple(
                name
                for name in handle.post_job_ids
                if handle.post_statuses.get(name) not in _FINAL_STATUSES
            )
        for name in pending:
            handle.post_statuses[name] = "cancelled"
            self._emit(
                handle,
                {
                    "phase": name,
                    "status": "cancelled",
                    f"{name}_status": "cancelled",
                    "message": "取消等待超时，已停止等待该后置任务",
                },
            )
            self._finish_process_event(
                handle,
                name,
                handle.process_ids.get(name, f"{handle.handle_id}:{name}"),
            )

    def _emit(self, handle: SingleArticleHandle, value: Mapping[str, Any]) -> None:
        update = SingleArticleUpdate.from_mapping(
            value,
            task_id=handle.handle_id,
            target_fingerprint=handle.target.fingerprint,
        )
        with handle.lock:
            previous = handle.latest_receipt
            receipt = _merge_receipt(previous, update.to_receipt())
            handle.latest_receipt = receipt
            handle.updates.append(update)
            if update.foreground_done or update.status == "skipped_collected":
                handle.foreground_receipt = receipt
                handle.foreground_event.set()
            if update.phase == "finished":
                handle.finished_receipt = receipt
                handle.finished_event.set()

        # 前台安全回执到达后，详情阶段才开始计时；详情最终结果到达后立即结束，
        # 不把评论/离线归档的等待时间算进 detail 活跃进程。
        if update.foreground_done or update.status == "skipped_collected":
            self._finish_process_event(
                handle,
                "foreground",
                handle.process_ids.get("foreground", f"{handle.handle_id}:foreground"),
            )
            if update.status not in {"skipped_collected", "failed", "cancelled"}:
                self._start_process_event(
                    handle,
                    "detail",
                    handle.process_ids.get("detail", f"{handle.handle_id}:detail"),
                )

        if _is_detail_result(update):
            self._finish_process_event(
                handle,
                "detail",
                handle.process_ids.get("detail", f"{handle.handle_id}:detail"),
            )
        self._emit_context_event(
            handle,
            {
                "type": "article",
                **update.to_receipt().to_dict(),
            },
        )

    def _emit_process_event(
        self,
        handle: SingleArticleHandle,
        process_type: str,
        process_id: str,
        action: str,
    ) -> None:
        self._emit_context_event(
            handle,
            {
                "type": "process",
                "eventId": f"{process_id}:{action}",
                "action": action,
                "processId": process_id,
                "processType": process_type,
            },
        )

    def _finish_process_event(
        self,
        handle: SingleArticleHandle,
        process_type: str,
        process_id: str,
    ) -> None:
        with handle.lock:
            if process_type in handle.finished_process_types:
                return
            handle.finished_process_types.add(process_type)
        self._emit_process_event(handle, process_type, process_id, "finished")

    def _start_process_event(
        self,
        handle: SingleArticleHandle,
        process_type: str,
        process_id: str,
    ) -> None:
        process_kind = str(process_type or "article").strip() or "article"
        process_key = str(process_id or "").strip()
        if not process_key:
            return
        with handle.lock:
            if process_kind in handle.started_process_types:
                return
            handle.started_process_types.add(process_kind)
            handle.process_ids[process_kind] = process_key
        self._emit_process_event(handle, process_kind, process_key, "started")

    def _emit_context_event(
        self,
        handle: SingleArticleHandle,
        event: Mapping[str, Any],
    ) -> None:
        context = handle.context
        payload = {
            "taskId": str(getattr(context, "task_id", "") or ""),
            "articleTaskId": handle.handle_id,
            # 这些字段只用于日志和状态展示，不参与窗口坐标重新计算。
            "articleIndex": handle.target.sequence,
            "articleTitle": handle.target.title_raw or handle.target.title_display,
            "articleDate": handle.target.article_date,
            **dict(event),
        }
        sink = getattr(context, "event_sink", None)
        if not callable(sink):
            state = getattr(context, "state", None)
            sink = getattr(state, "handle_event", None)
        if not callable(sink):
            return
        try:
            sink(payload)
        except Exception:
            # 事件通道是辅助状态，不能反向让单篇采集失败。
            return


def _merge_receipt(
    previous: SingleArticleReceipt | None,
    current: SingleArticleReceipt,
) -> SingleArticleReceipt:
    if previous is None:
        return current
    values = previous.to_dict()
    for key, value in current.to_dict().items():
        if key in {"phase", "status"} or value not in {
            None,
            "",
            False,
            0,
            "pending",
            "not_requested",
        }:
            values[key] = value
    return SingleArticleReceipt.from_dict(values)


def _service_update(
    payload: Mapping[str, Any],
    *,
    default_phase: str | None = None,
) -> dict[str, Any]:
    raw_status = str(payload.get("status") or "").strip().lower()
    status = _normalize_status(payload.get("status"))
    phase = str(payload.get("phase") or default_phase or "")
    if not phase:
        if status == "skipped_collected":
            phase = "foreground"
        elif default_phase:
            phase = default_phase
        elif status in _FINAL_STATUSES:
            phase = "finished"
        else:
            phase = "precheck"
    article_saved = bool(
        payload.get("articleSaved", payload.get("article_saved", False))
    )
    if not article_saved and raw_status in {"completed", "success"}:
        article_saved = str(
            payload.get("detailStatus") or payload.get("detail_status") or ""
        ).lower() == "success" or bool(payload.get("articleId"))
    return {
        **dict(payload),
        "phase": phase,
        "status": status,
        "foreground_done": bool(payload.get("foregroundDone", False)),
        "progress_counted": bool(payload.get("progressCounted", status != "skipped_collected" and phase == "precheck")),
        "article_saved": article_saved,
    }


def _job_id_from_result(result: Any, message: str) -> str:
    if not isinstance(result, Mapping):
        raise RuntimeError(message)
    job_id = str(result.get("jobId") or result.get("job_id") or "").strip()
    if not job_id:
        raise RuntimeError(message)
    return job_id


def _is_foreground_finished(payload: Mapping[str, Any]) -> bool:
    status = _normalize_status(payload.get("status"))
    phase = str(payload.get("phase") or "")
    return (
        status == "skipped_collected"
        or status in {"failed", "cancelled"}
        or (status == "success" and phase == "foreground" and bool(payload.get("foreground_done")))
    )


def _is_detail_finished(payload: Mapping[str, Any]) -> bool:
    status = _normalize_status(payload.get("status"))
    phase = str(payload.get("phase") or "")
    return (
        status in _FINAL_STATUSES
        and phase in {"detail", "finished"}
    ) or status == "skipped_collected"


def _detail_saved(payload: Mapping[str, Any]) -> bool:
    """兼容详情服务的 snake_case/camelCase 成功字段。"""

    return bool(
        payload.get("articleSaved")
        or payload.get("article_saved")
        or str(payload.get("detailStatus") or payload.get("detail_status") or "").lower()
        == "success"
        or payload.get("articleId")
    )


def _is_detail_result(update: SingleArticleUpdate) -> bool:
    """判断更新是否已经给出详情阶段的最终结果。"""

    detail_status = str(update.detail_status or "").strip().lower()
    return (
        update.status in _FINAL_STATUSES
        and update.phase in {"detail", "finished"}
    ) or detail_status in {"success", "failed", "skipped"}


def _article_post_payload(
    payload: Mapping[str, Any],
    target: HomeArticleTarget,
) -> dict[str, Any]:
    result = dict(payload)
    result.setdefault("accountName", target.account_name)
    result.setdefault("articleTitle", target.title_raw or target.title_display)
    result.setdefault("publishedDate", target.article_date)
    result.setdefault("rawTitle", target.title_raw)
    result.setdefault("records", [{
        "index": target.sequence,
        "accountName": target.account_name,
        "publishedDate": target.article_date,
        "rawTitle": target.title_raw,
        "title": target.title_display,
    }])
    return result


def _card_payload(target: HomeArticleTarget) -> dict[str, Any]:
    return {
        "index": target.sequence,
        "accountName": target.account_name,
        "publishedDate": target.article_date,
        "dateText": target.article_date,
        "title": target.title_display,
        "rawTitle": target.title_raw,
        "clickPoint": [int(target.click_point[0]), int(target.click_point[1])],
        "visibleRect": list(target.visible_rect),
        "cardRect": list(target.card_rect),
        "homeWindowHandle": int(target.home_window_id or 0),
    }


def _call_supported(function: Callable[..., Any], kwargs: Mapping[str, Any]) -> Any:
    try:
        signature = inspect.signature(function)
    except (TypeError, ValueError):
        return function(**dict(kwargs))
    accepts_kwargs = any(
        parameter.kind is inspect.Parameter.VAR_KEYWORD
        for parameter in signature.parameters.values()
    )
    if accepts_kwargs:
        return function(**dict(kwargs))
    supported = {
        key: value
        for key, value in kwargs.items()
        if key in signature.parameters
    }
    return function(**supported)


def _normalize_status(value: Any) -> str:
    normalized = str(value or "failed").strip().lower()
    if normalized in {"running", "pending", "precheck"}:
        return normalized
    if normalized in {"completed", "ready-to-continue", "success"}:
        return "success" if normalized == "completed" or normalized == "success" else "success"
    if normalized in {"skipped", "skipped-collected", "skipped_collected"}:
        return "skipped_collected"
    if normalized in {"cancelled", "canceled"}:
        return "cancelled"
    if normalized in {"failed", "save-failed", "error"}:
        return "failed"
    return normalized if normalized in _FINAL_STATUSES else "failed"


def _as_bool(value: Mapping[str, Any], snake: str, camel: str) -> bool:
    return bool(value.get(snake, value.get(camel, False)))


__all__ = [
    "ArticleDispatcher",
    "SingleArticleHandle",
    "SingleArticleUpdate",
]
