from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from threading import Lock, RLock
import time
from types import SimpleNamespace
from typing import Any, Callable, Mapping
from uuid import uuid4

from huey import SqliteHuey
from huey.consumer import Consumer

from src.domain.enums import CaptureType, ResourceType, TaskStatus
from src.domain.models import ArticleTarget, MitmCaptureResult, ResourceManifest, TaskContext
from src.services.capture.collected_article_lookup_service import (
    CollectedArticleLookupService,
)
from src.services.capture.html_parse_save_service import HtmlParseSaveService
from src.services.capture.single_article_capture_service import SingleCaptureSettings
from src.services.runtime.database_write_coordinator import DatabaseWriteCoordinator
from src.services.task.huey_runtime_queue import (
    ensure_huey_storage_ready,
    runtime_huey_queue_dir,
)


_SAFE_ID_PATTERN = re.compile(r"[^A-Za-z0-9._-]+")
_ACTIVE_STATUSES = {"running"}
# 系统代理和 MITM 属于进程级全局资源，所有详情服务实例默认共享同一把锁。
_PROCESS_FOREGROUND_LOCK = Lock()


@dataclass(frozen=True, slots=True)
class InitialContentStorageTaskOptions:
    """初始内容存储单篇任务输入，由后端同步读取主页卡片后传入。"""

    card_index: int = 1
    account_name: str | None = None
    card: dict[str, Any] | None = None
    skip_collected_records: bool = False
    store_article_detail: bool = True
    stateful_offline_cache: bool = False


class _EmbeddedThreadConsumer(Consumer):
    """桌面应用内嵌 Huey worker，避免接管主进程信号处理。"""

    def _set_signal_handlers(self) -> None:
        return None


class _ForegroundLease:
    """保护主页点击和全局代理接管区间的互斥租约。"""

    def __init__(self, lock: Lock) -> None:
        self._lock = lock
        self._held = False

    def __enter__(self) -> "_ForegroundLease":
        self._lock.acquire()
        self._held = True
        return self

    def __exit__(self, exc_type: Any, exc_value: Any, traceback: Any) -> None:
        del exc_type, exc_value, traceback
        if self._held:
            self._held = False
            self._lock.release()


class InitialContentStorageHueyService:
    """使用会话级 SqliteHuey 执行单篇“详情捕获 -> 初始内容保存”任务。"""

    def __init__(
        self,
        *,
        temp_root: str | Path,
        config: Any,
        window_factory: Any,
        capture_factory: Any,
        database_path: str | Path,
        html_save: Any | None = None,
        write_coordinator: DatabaseWriteCoordinator | None = None,
        runner: Callable[..., dict[str, Any]] | None = None,
        lookup_service: CollectedArticleLookupService | None = None,
        session_id: str | None = None,
        job_id_factory: Callable[[], str] | None = None,
        now: Callable[[], datetime] = datetime.now,
        action: str = "initial-content-storage",
        title: str = "初始内容存储结果",
        flow_label: str = "初始内容存储测试",
        job_prefix: str = "initial-storage",
        queue_name: str = "initial-content-storage",
        task_name: str = "InitialContentStorageTask",
        wait_message_with_card: str = "已读取首篇文章卡片，正在等待Huey执行初始内容存储任务...",
        wait_message_without_card: str = "正在等待Huey执行初始内容存储任务...",
        extra_public_options: dict[str, Any] | None = None,
        worker_count: int = 1,
        post_phase: str = "post",
        foreground_lock: Lock | None = None,
    ) -> None:
        self._temp_root = Path(temp_root).resolve()
        self._config = config
        self._window_factory = window_factory
        self._capture_factory = capture_factory
        self._database_path = Path(database_path)
        self._write_coordinator = write_coordinator or DatabaseWriteCoordinator()
        self._html_save = html_save or HtmlParseSaveService(
            write_coordinator=self._write_coordinator,
            now=now,
        )
        self._runner = runner
        self._lookup_service = lookup_service or CollectedArticleLookupService()
        self._job_id_factory = job_id_factory or (lambda: uuid4().hex[:12])
        self._now = now
        self._action = action
        self._title = title
        self._flow_label = flow_label
        self._job_prefix = job_prefix
        self._queue_name = queue_name
        self._task_name = task_name
        self._wait_message_with_card = wait_message_with_card
        self._wait_message_without_card = wait_message_without_card
        self._extra_public_options = dict(extra_public_options or {})
        self._worker_count = max(1, int(worker_count))
        self._post_phase = str(post_phase or "post")
        self._jobs: dict[str, dict[str, Any]] = {}
        self._cancelled_jobs: set[str] = set()
        self._active_attempts: dict[str, Any] = {}
        self._lock = RLock()
        self._foreground_lock = (
            foreground_lock
            if foreground_lock is not None
            else _PROCESS_FOREGROUND_LOCK
        )
        self._consumer_started = False
        self._closed = False

        normalized_session_id = _safe_identifier(
            session_id or uuid4().hex[:12],
            fallback="session",
        )
        queue_dir = runtime_huey_queue_dir(self._temp_root)
        queue_dir.mkdir(parents=True, exist_ok=True)
        self._queue_database_path = (
            queue_dir / f"{_safe_identifier(queue_name)}-{normalized_session_id}.sqlite3"
        )
        self._huey = SqliteHuey(
            queue_name,
            filename=str(self._queue_database_path),
            results=True,
            store_none=False,
        )
        self._task_wrapper = self._huey.task(
            retries=0,
            name=task_name,
        )(self._execute_task)
        self._post_task_wrapper = self._huey.task(
            retries=0,
            name=f"{task_name}Post",
        )(self._execute_post_task)
        self._foreground_task_wrapper = self._huey.task(
            retries=0,
            name=f"{task_name}Foreground",
        )(self._execute_foreground_task)
        self._detail_task_wrapper = self._huey.task(
            retries=0,
            name=f"{task_name}Detail",
        )(self._execute_detail_task)
        self._consumer = _EmbeddedThreadConsumer(
            self._huey,
            workers=self._worker_count,
            worker_type="thread",
            periodic=False,
            check_worker_health=False,
        )

    def _foreground_lease(self) -> _ForegroundLease:
        return _ForegroundLease(self._foreground_lock)

    @property
    def queue_database_path(self) -> Path:
        return self._queue_database_path

    def start(
        self,
        *,
        card_index: int = 1,
        account_name: str | None = None,
        card: dict[str, Any] | None = None,
        skip_collected_records: bool = False,
        store_article_detail: bool = True,
        stateful_offline_cache: bool = False,
    ) -> dict[str, Any]:
        # 该任务当前只支持“存储文章详情”开启；前端禁用只是 UI 保护，后端仍强制锁定。
        options = InitialContentStorageTaskOptions(
            card_index=max(1, int(card_index)),
            account_name=_optional_text(account_name),
            card=_safe_card_payload(card),
            skip_collected_records=bool(skip_collected_records),
            store_article_detail=True,
            stateful_offline_cache=bool(stateful_offline_cache),
        )
        with self._lock:
            if self._closed:
                raise RuntimeError("Huey初始内容存储服务已经关闭")
            job_id = f"{self._job_prefix}-{_safe_identifier(self._job_id_factory())}"
            task = self._task_wrapper.s(job_id, _task_payload(options))
            initial = {
                "ok": False,
                "status": "running",
                "jobId": job_id,
                "hueyTaskId": task.id,
                "action": self._action,
                "title": self._title,
                "message": (
                    self._wait_message_with_card
                    if options.card is not None
                    else self._wait_message_without_card
                ),
                "tone": "info",
                "items": [
                    {"label": "流程", "value": self._flow_label},
                    {
                        "label": "跳过已采集记录",
                        "value": "开启" if options.skip_collected_records else "关闭",
                    },
                    {"label": "存储文章详情", "value": "开启（锁定）"},
                    {"label": "状态", "value": "等待执行"},
                ]
                + _initial_card_items(options),
                "records": [options.card] if options.card is not None else [],
                "accountName": options.account_name or "",
                "captureType": CaptureType.NONE.value,
                "startedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": self._public_options_payload(options),
            }
            self._jobs[job_id] = initial
            self._trim_jobs()
            self._enqueue_task(task)
            return dict(initial)

    def start_foreground(
        self,
        *,
        card_index: int = 1,
        account_name: str | None = None,
        card: dict[str, Any] | None = None,
        skip_collected_records: bool = False,
        stateful_offline_cache: bool = False,
    ) -> dict[str, Any]:
        """提交单篇前台阶段：跳过校验、代理捕获、点击、关闭标签和释放代理。"""

        options = InitialContentStorageTaskOptions(
            card_index=max(1, int(card_index)),
            account_name=_optional_text(account_name),
            card=_safe_card_payload(card),
            skip_collected_records=bool(skip_collected_records),
            store_article_detail=True,
            stateful_offline_cache=bool(stateful_offline_cache),
        )
        with self._lock:
            if self._closed:
                raise RuntimeError("Huey前台任务服务已经关闭")
            job_id = f"{self._job_prefix}-foreground-{_safe_identifier(self._job_id_factory())}"
            task = self._foreground_task_wrapper.s(job_id, _task_payload(options))
            initial = {
                "ok": False,
                "status": "running",
                "phase": "foreground",
                "jobId": job_id,
                "hueyTaskId": task.id,
                "action": self._action,
                "title": self._title,
                "message": "正在等待Huey执行单篇前台捕获任务...",
                "tone": "info",
                "items": [
                    {"label": "流程", "value": f"{self._flow_label} - 前台捕获"},
                    {
                        "label": "跳过已采集记录",
                        "value": "开启" if options.skip_collected_records else "关闭",
                    },
                    {"label": "状态", "value": "等待执行"},
                ]
                + _initial_card_items(options),
                "records": [options.card] if options.card is not None else [],
                "accountName": options.account_name or "",
                "captureType": CaptureType.NONE.value,
                "foregroundDone": False,
                "progressCounted": False,
                "tabClosed": False,
                "startedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": self._public_options_payload(options),
            }
            self._jobs[job_id] = initial
            self._trim_jobs()
            self._enqueue_task(task)
            return dict(initial)

    def start_detail(
        self,
        *,
        card_index: int = 1,
        account_name: str | None = None,
        card: dict[str, Any] | None = None,
        capture_result: Mapping[str, Any] | None = None,
        foreground_result: Mapping[str, Any] | None = None,
        store_article_detail: bool = True,
    ) -> dict[str, Any]:
        """提交单篇详情阶段：只消费前台捕获证据并保存文章详情。"""

        options = InitialContentStorageTaskOptions(
            card_index=max(1, int(card_index)),
            account_name=_optional_text(account_name),
            card=_safe_card_payload(card),
            skip_collected_records=False,
            store_article_detail=bool(store_article_detail),
            stateful_offline_cache=False,
        )
        capture_payload = dict(capture_result or {})
        foreground_payload = dict(foreground_result or {})
        task_payload = {
            "options": _task_payload(options),
            "captureResult": capture_payload,
            "foregroundResult": foreground_payload,
            "storeArticleDetail": bool(store_article_detail),
        }
        with self._lock:
            if self._closed:
                raise RuntimeError("Huey详情任务服务已经关闭")
            job_id = f"{self._job_prefix}-detail-{_safe_identifier(self._job_id_factory())}"
            task = self._detail_task_wrapper.s(job_id, task_payload)
            records = (
                [options.card]
                if options.card is not None
                else list(foreground_payload.get("records") or [])
            )
            initial = {
                "ok": False,
                "status": "running",
                "phase": "detail",
                "jobId": job_id,
                "hueyTaskId": task.id,
                "action": self._action,
                "title": self._title,
                "message": "正在等待Huey执行单篇详情保存任务...",
                "tone": "info",
                "items": [
                    {"label": "流程", "value": f"{self._flow_label} - 详情保存"},
                    {"label": "状态", "value": "等待执行"},
                ],
                "records": records,
                "accountName": options.account_name or str(foreground_payload.get("accountName") or ""),
                "captureType": str(capture_payload.get("capture_type") or CaptureType.NONE.value),
                "foregroundDone": bool(foreground_payload.get("foregroundDone")),
                "progressCounted": bool(foreground_payload.get("progressCounted", True)),
                "tabClosed": bool(foreground_payload.get("tabClosed")),
                "detailStatus": "running",
                "startedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": self._public_options_payload(options),
            }
            self._jobs[job_id] = initial
            self._trim_jobs()
            self._enqueue_task(task)
            return dict(initial)

    def get(self, job_id: str) -> dict[str, Any]:
        with self._lock:
            job = self._jobs.get(str(job_id))
            if job is None:
                raise KeyError(job_id)
            return dict(job)

    def start_post(
        self,
        *,
        article: Mapping[str, Any],
        parent_task_id: str = "",
    ) -> dict[str, Any]:
        """提交已完成详情的后置任务，不重新执行主页点击和详情捕获。"""

        article_payload = dict(article)
        article_payload["parentTaskId"] = str(parent_task_id or "")
        with self._lock:
            if self._closed:
                raise RuntimeError("Huey 后置任务服务已经关闭")
            job_id = f"{self._job_prefix}-post-{_safe_identifier(self._job_id_factory())}"
            task = self._post_task_wrapper.s(job_id, article_payload)
            initial = {
                "ok": False,
                "status": "running",
                "phase": self._post_phase,
                "jobId": job_id,
                "hueyTaskId": task.id,
                "action": self._action,
                "title": self._title,
                "message": f"正在等待{self._flow_label}后置任务执行...",
                "tone": "info",
                "parentTaskId": str(parent_task_id or ""),
                "articleId": article_payload.get("articleId"),
                "articleTitle": article_payload.get("articleTitle") or article_payload.get("rawTitle") or "",
                "options": {**self._extra_public_options, "postOnly": True},
            }
            self._jobs[job_id] = initial
            self._trim_jobs()
            self._enqueue_task(task)
            return dict(initial)

    def is_active(self) -> bool:
        with self._lock:
            return any(
                str(job.get("status") or "") in _ACTIVE_STATUSES
                for job in self._jobs.values()
            )

    def cancel(self, job_id: str) -> bool:
        """请求取消一个 Huey 任务；真正的窗口清理仍在任务 finally 中完成。"""

        key = str(job_id or "").strip()
        if not key:
            return False
        with self._lock:
            current = self._jobs.get(key)
            if current is None:
                return False
            if str(current.get("status") or "") not in _ACTIVE_STATUSES:
                return False
            self._cancelled_jobs.add(key)
            self._jobs[key] = {
                **current,
                "ok": False,
                "status": "cancelled",
                "phase": "finished",
                "message": "任务已请求取消，等待安全清理。",
                "tone": "warning",
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            attempt = self._active_attempts.get(key)
        if attempt is not None:
            try:
                attempt.cancel()
            except Exception:
                pass
        return True

    def _register_active_attempt(self, job_id: str, attempt: Any) -> None:
        """登记当前任务持有的子进程 attempt，并处理取消竞态。"""

        key = str(job_id or "").strip()
        if not key or attempt is None:
            return
        with self._lock:
            self._active_attempts[key] = attempt
            cancelled = key in self._cancelled_jobs
        if cancelled:
            try:
                attempt.cancel()
            except Exception:
                pass

    def _unregister_active_attempt(self, job_id: str, attempt: Any | None = None) -> None:
        """移除 attempt；避免迟到的 finally 清掉新 attempt。"""

        key = str(job_id or "").strip()
        if not key:
            return
        active_attempts = getattr(self, "_active_attempts", None)
        if not isinstance(active_attempts, dict):
            return
        lock = getattr(self, "_lock", None)
        if lock is None:
            current = active_attempts.get(key)
            if attempt is None or current is attempt:
                active_attempts.pop(key, None)
            return
        with lock:
            current = active_attempts.get(key)
            if attempt is None or current is attempt:
                active_attempts.pop(key, None)

    def shutdown(self) -> None:
        with self._lock:
            if self._closed:
                return
            self._closed = True
            consumer_started = self._consumer_started
        if consumer_started:
            self._consumer.stop(graceful=True)
        self._huey.storage.close()

    def _start_consumer(self) -> None:
        if self._consumer_started:
            return
        self._consumer.start()
        self._consumer_started = True

    def _enqueue_task(self, task: Any) -> None:
        ensure_huey_storage_ready(self._huey, self._queue_database_path)
        self._start_consumer()
        self._huey.enqueue(task)

    def _execute_task(self, job_id: str, options_payload: dict[str, Any]) -> None:
        options = _options_from_payload(options_payload)

        def update(payload: dict[str, Any]) -> None:
            merged = {
                "jobId": job_id,
                "action": self._action,
                "title": self._title,
                **payload,
                "options": self._public_options_payload(options),
            }
            with self._lock:
                if job_id in self._cancelled_jobs:
                    return
                current = self._jobs.get(job_id, {})
                self._jobs[job_id] = {**current, **merged}

        try:
            if self._is_cancelled(job_id):
                return
            should_run_capture = False
            if self._runner is not None:
                result = self._runner(options=options, on_update=update)
            elif options.card is not None:
                result = _manual_card_result(options)
                should_run_capture = True
            else:
                result = _missing_task_input_result(options)
            if not isinstance(result, dict):
                raise RuntimeError("初始内容存储任务返回了无法识别的结果")
            final = self._apply_collected_lookup(result, options)
            if final.get("status") == "skipped-collected":
                update(
                    {
                        **final,
                        "phase": "foreground",
                        "status": "skipped_collected",
                        "foregroundDone": True,
                        "progressCounted": False,
                        "tabClosed": False,
                        "detailStatus": "not_requested",
                    }
                )
            elif final.get("status") == "ready-to-continue":
                update(
                    {
                        **final,
                        "phase": "precheck",
                        "status": "running",
                        "progressCounted": True,
                    }
                )
            if should_run_capture and final.get("status") == "ready-to-continue":
                final = self._run_capture_and_save_flow(
                    job_id=job_id,
                    result=final,
                    options=options,
                    update=update,
                )
            final = {
                "jobId": job_id,
                "action": self._action,
                "title": self._title,
                **final,
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": self._public_options_payload(options),
            }
            final.setdefault("captureType", CaptureType.NONE.value)
            final.setdefault("phase", "finished")
            update(final)
        except Exception as exc:
            update(
                {
                    "ok": False,
                    "status": "failed",
                    "phase": "finished",
                    "message": f"初始内容存储任务失败：{exc}",
                    "tone": "error",
                    "items": [{"label": "失败原因", "value": str(exc)}],
                    "records": [],
                    "captureType": CaptureType.NONE.value,
                    "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    def _execute_foreground_task(self, job_id: str, options_payload: dict[str, Any]) -> None:
        options = _options_from_payload(options_payload)

        def update(payload: dict[str, Any]) -> None:
            merged = {
                "jobId": job_id,
                "action": self._action,
                "title": self._title,
                **payload,
                "phase": payload.get("phase") or "foreground",
                "options": self._public_options_payload(options),
            }
            with self._lock:
                if job_id in self._cancelled_jobs:
                    return
                current = self._jobs.get(job_id, {})
                self._jobs[job_id] = {**current, **merged}

        try:
            if self._is_cancelled(job_id):
                return
            if self._runner is not None:
                result = self._runner(options=options, on_update=update)
                if not isinstance(result, dict):
                    raise RuntimeError("前台任务返回了无法识别的结果")
            elif options.card is not None:
                result = _manual_card_result(options)
            else:
                result = _missing_task_input_result(options)

            final = self._apply_collected_lookup(dict(result), options)
            if final.get("status") == "skipped-collected":
                update(
                    {
                        **final,
                        "ok": True,
                        "phase": "foreground",
                        "status": "skipped_collected",
                        "foregroundDone": True,
                        "progressCounted": False,
                        "tabClosed": False,
                        "detailStatus": "not_requested",
                        "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                return
            if final.get("status") != "ready-to-continue":
                update(
                    {
                        **final,
                        "ok": False,
                        "phase": "foreground",
                        "status": "failed",
                        "foregroundDone": False,
                        "progressCounted": False,
                        "tabClosed": False,
                        "detailStatus": "failed",
                        "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                return

            update(
                {
                    **final,
                    "ok": False,
                    "phase": "foreground",
                    "status": "running",
                    "message": "前台校验通过，正在执行代理捕获和文章点击...",
                    "progressCounted": True,
                }
            )
            final = self._run_foreground_capture_flow(
                job_id=job_id,
                result=final,
                options=options,
                update=update,
            )
            final = {
                "jobId": job_id,
                "action": self._action,
                "title": self._title,
                **final,
                "phase": "foreground",
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": self._public_options_payload(options),
            }
            final.setdefault("captureType", CaptureType.NONE.value)
            update(final)
        except Exception as exc:
            update(
                {
                    "ok": False,
                    "status": "failed",
                    "phase": "foreground",
                    "foregroundDone": False,
                    "progressCounted": True,
                    "tabClosed": False,
                    "detailStatus": "failed",
                    "message": f"前台捕获任务失败：{exc}",
                    "tone": "error",
                    "items": [{"label": "失败原因", "value": str(exc)}],
                    "records": [],
                    "captureType": CaptureType.NONE.value,
                    "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    def _execute_detail_task(self, job_id: str, payload: dict[str, Any]) -> None:
        options = _options_from_payload(dict(payload.get("options") or {}))
        capture_payload = dict(payload.get("captureResult") or {})
        foreground_result = dict(payload.get("foregroundResult") or {})
        store_article_detail = bool(payload.get("storeArticleDetail", True))

        def update(value: dict[str, Any]) -> None:
            merged = {
                "jobId": job_id,
                "action": self._action,
                "title": self._title,
                **value,
                "phase": value.get("phase") or "detail",
                "options": self._public_options_payload(options),
            }
            with self._lock:
                if job_id in self._cancelled_jobs:
                    return
                current = self._jobs.get(job_id, {})
                self._jobs[job_id] = {**current, **merged}

        try:
            if self._is_cancelled(job_id):
                return
            records = [
                dict(record)
                for record in (
                    foreground_result.get("records")
                    or ([options.card] if options.card is not None else [])
                )
                if isinstance(record, Mapping)
            ]
            base_result = {
                **_manual_card_result(options),
                **foreground_result,
                "records": records,
                "accountName": (
                    _optional_text(options.account_name)
                    or _optional_text(foreground_result.get("accountName"))
                    or ""
                ),
            }
            items = list(foreground_result.get("items") or base_result.get("items") or [])
            account_name = str(base_result.get("accountName") or "")

            if not store_article_detail:
                update(
                    {
                        **base_result,
                        "ok": True,
                        "status": "completed",
                        "phase": "detail",
                        "detailStatus": "not_requested",
                        "articleSaved": False,
                        "message": "详情保存未开启，详情阶段已跳过。",
                        "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                return
            if not capture_payload:
                raise RuntimeError("详情阶段缺少前台捕获结果")

            capture_result = MitmCaptureResult.from_dict(capture_payload)
            capture_type = _capture_type_value(capture_result)
            records_for_target = records or ([options.card] if options.card is not None else [])
            if not records_for_target:
                raise RuntimeError("详情阶段缺少文章卡片信息")
            target_card = _detail_card_payload(records_for_target[0])
            target = _target_from_card(options, target_card)
            started_at = time.monotonic()
            capture_started_at = self._now()
            context = TaskContext(
                task_id=job_id,
                proxy_lease_id=f"{job_id}-storage",
                db_path=self._database_path,
                storage_root=_config_path(
                    self._config,
                    "storage",
                    "article_storage_root",
                    self._temp_root,
                ),
                temp_dir=_config_path(
                    self._config,
                    "storage",
                    "temp_dir",
                    self._temp_root,
                )
                / job_id,
                started_at=capture_started_at,
            )
            update(
                {
                    **base_result,
                    "ok": False,
                    "status": "running",
                    "phase": "detail",
                    "detailStatus": "running",
                    "message": "正在解析 HTML 并保存文章详情...",
                    "captureType": capture_type,
                    "accountName": account_name,
                    "items": items,
                    "records": records,
                }
            )
            request_timeout = float(
                getattr(getattr(self._config, "request", None), "request_timeout_seconds", 30.0)
            )
            save = self._html_save.save(
                context=context,
                target=target,
                capture_result=capture_result,
                attempt_started_at=capture_started_at,
                duration_seconds=float(
                    foreground_result.get("totalSeconds")
                    or foreground_result.get("durationSeconds")
                    or 0.0
                ),
                request_timeout_seconds=request_timeout,
            )
            if save.status is TaskStatus.SUCCESS and save.data is not None:
                final = self._build_save_success_result(
                    job_id=job_id,
                    base_result=base_result,
                    options=options,
                    update=update,
                    items=items,
                    records=records,
                    account_name=account_name,
                    capture_type=capture_type,
                    started_at=started_at,
                    context=context,
                    saved=save.data,
                )
                final["finishedAt"] = self._now().strftime("%Y-%m-%d %H:%M:%S")
                update(final)
                return

            error_message = str(getattr(save, "message", "") or "详情解析或保存失败")
            update(
                {
                    **base_result,
                    "ok": False,
                    "status": "save-failed",
                    "phase": "detail",
                    "detailStatus": "failed",
                    "articleSaved": False,
                    "foregroundDone": True,
                    "progressCounted": True,
                    "tabClosed": bool(foreground_result.get("tabClosed", True)),
                    "message": error_message,
                    "tone": "error",
                    "captureType": capture_type,
                    "items": items + [{"label": "失败原因", "value": error_message}],
                    "records": records,
                    "accountName": account_name,
                    "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )
        except Exception as exc:
            update(
                {
                    "ok": False,
                    "status": "failed",
                    "phase": "detail",
                    "detailStatus": "failed",
                    "articleSaved": False,
                    "foregroundDone": True,
                    "progressCounted": True,
                    "tabClosed": bool(foreground_result.get("tabClosed", True)),
                    "message": f"详情保存任务失败：{exc}",
                    "tone": "error",
                    "errorStage": "detail",
                    "errorDetail": str(exc),
                    "captureType": str(capture_payload.get("capture_type") or CaptureType.NONE.value),
                    "items": [{"label": "失败原因", "value": str(exc)}],
                    "records": foreground_result.get("records") or [],
                    "accountName": (
                        _optional_text(options.account_name)
                        or _optional_text(foreground_result.get("accountName"))
                        or ""
                    ),
                    "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    def _execute_post_task(
        self,
        job_id: str,
        article_payload: dict[str, Any],
    ) -> None:
        def update(payload: dict[str, Any]) -> None:
            merged = {
                "jobId": job_id,
                "action": self._action,
                "title": self._title,
                "phase": self._post_phase,
                **payload,
                "options": {**self._extra_public_options, "postOnly": True},
            }
            with self._lock:
                if job_id in self._cancelled_jobs:
                    return
                current = self._jobs.get(job_id, {})
                self._jobs[job_id] = {**current, **merged}

        try:
            if self._is_cancelled(job_id):
                return
            result = self._run_post_task(
                job_id=job_id,
                article=article_payload,
                update=update,
            )
            if not isinstance(result, dict):
                raise RuntimeError("后置任务返回了无法识别的结果")
            final = {
                "jobId": job_id,
                "phase": self._post_phase,
                **result,
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            update(final)
        except Exception as exc:
            update(
                {
                    "ok": False,
                    "status": "failed",
                    "phase": self._post_phase,
                    "message": f"后置任务失败：{exc}",
                    "tone": "error",
                    "errorStage": self._post_phase,
                    "errorDetail": str(exc),
                    "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    def _run_post_task(
        self,
        *,
        job_id: str,
        article: Mapping[str, Any],
        update: Callable[[dict[str, Any]], None],
    ) -> dict[str, Any]:
        del job_id, article, update
        raise RuntimeError("当前服务未实现后置任务")

    def _run_foreground_capture_flow(
        self,
        *,
        job_id: str,
        result: dict[str, Any],
        options: InitialContentStorageTaskOptions,
        update: Callable[[dict[str, Any]], None],
    ) -> dict[str, Any]:
        started_at = time.monotonic()
        items = list(result.get("items") or [])
        records = [dict(record) for record in list(result.get("records") or [])]
        account_name = _optional_text(result.get("accountName")) or _optional_text(
            options.account_name
        ) or ""
        target = _target_from_card(options, records[0] if records else {})
        settings = SingleCaptureSettings.from_app_config(self._config)
        tabs = self._window_factory.create_tab_service()
        clicker = self._window_factory.create_clicker()
        control = self._capture_factory.create_in_process_control()
        attempt: Any | None = None
        article_tab: Any | None = None
        attempt_finished = False
        tab_closed = False

        def publish(message: str) -> None:
            update(
                {
                    **result,
                    "ok": False,
                    "status": "running",
                    "phase": "foreground",
                    "message": message,
                    "tone": "info",
                    "items": list(items),
                    "records": records,
                    "accountName": account_name,
                    "captureType": CaptureType.NONE.value,
                    "foregroundDone": False,
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
            )

        def append_step(
            label: str,
            value: str,
            *,
            duration_seconds: float | None = None,
        ) -> None:
            item: dict[str, Any] = {"label": label, "value": value}
            if duration_seconds is not None:
                item["durationSeconds"] = round(max(0.0, duration_seconds), 3)
            items.append(item)

        with self._foreground_lease():
            try:
                publish("正在启动 MITM 捕获...")
                step_started = time.monotonic()
                baseline = tabs.capture_baseline()
                append_step(
                    "标签基线",
                    "已记录点击前浏览器标签状态",
                    duration_seconds=time.monotonic() - step_started,
                )

                step_started = time.monotonic()
                attempt_id = uuid4().hex[:12]
                attempt = control.start_attempt(
                    task_id=job_id,
                    attempt_id=attempt_id,
                    proxy_lease_id=f"{job_id}-{attempt_id}",
                    proxy_address=settings.proxy_address,
                    capture_config=settings.capture_config,
                )
                self._register_active_attempt(job_id, attempt)
                if self._is_cancelled(job_id):
                    try:
                        attempt.cancel()
                    except Exception:
                        pass
                    return {
                        **result,
                        "ok": False,
                        "status": "cancelled",
                        "phase": "foreground",
                        "foregroundDone": False,
                        "progressCounted": False,
                        "tabClosed": False,
                        "detailStatus": "cancelled",
                        "message": "单篇前台任务已取消。",
                    }
                attempt.wait_ready(timeout_seconds=settings.ready_timeout_seconds)
                append_step(
                    "MITM 捕获",
                    "已启动 MITM 并开启系统代理",
                    duration_seconds=time.monotonic() - step_started,
                )

                publish("正在点击文章卡片...")
                step_started = time.monotonic()
                clicker.click(target)
                append_step(
                    "点击文章",
                    "已发送文章卡片点击",
                    duration_seconds=time.monotonic() - step_started,
                )

                publish("正在确认文章标签打开...")
                step_started = time.monotonic()
                article_tab = tabs.wait_for_opened_article_tab(
                    baseline=baseline,
                    timeout_seconds=settings.title_timeout_seconds,
                    stable_delay_seconds=settings.title_stable_delay_seconds,
                    poll_initial_interval_seconds=(
                        settings.title_poll_initial_interval_seconds
                    ),
                    poll_max_interval_seconds=settings.title_poll_max_interval_seconds,
                )
                append_step(
                    "文章标签",
                    "已检测到文章标签打开",
                    duration_seconds=time.monotonic() - step_started,
                )

                publish("正在关闭文章标签...")
                step_started = time.monotonic()
                tabs.close_article_tab(article_tab, home_window_handle=target.home_window_handle)
                tab_closed = True
                append_step(
                    "关闭标签",
                    "已发送 Ctrl+W 关闭文章标签",
                    duration_seconds=time.monotonic() - step_started,
                )

                publish("正在停止 MITM 捕获...")
                step_started = time.monotonic()
                capture_result = attempt.stop_capture(
                    timeout_seconds=settings.result_timeout_seconds
                )
                attempt_finished = True
                append_step(
                    "关闭代理",
                    "已恢复系统代理并停止 MITM",
                    duration_seconds=time.monotonic() - step_started,
                )
                _append_capture_event_items(items, capture_result)
                capture_type = _capture_type_value(capture_result)
                capture_ok = (
                    getattr(capture_result, "status", None) is TaskStatus.SUCCESS
                    and capture_type != CaptureType.NONE.value
                )
                if not capture_ok:
                    error_message = str(
                        getattr(capture_result, "error_message", "")
                        or "未捕获到 HTML 或 reference"
                    )
                    append_step("捕获结果", error_message)
                    return {
                        **result,
                        "ok": False,
                        "status": "failed",
                        "phase": "foreground",
                        "message": error_message,
                        "tone": "error",
                        "items": items,
                        "records": records,
                        "accountName": account_name,
                        "captureType": capture_type,
                        "foregroundDone": bool(tab_closed),
                        "progressCounted": True,
                        "tabClosed": bool(tab_closed),
                        "detailStatus": "failed",
                        "totalSeconds": round(time.monotonic() - started_at, 3),
                    }

                append_step("捕获结果", f"已得到 {capture_type}")
                return {
                    **result,
                    "ok": True,
                    "status": "success",
                    "phase": "foreground",
                    "message": "前台采集完成，主页可以继续下一篇。",
                    "tone": "success",
                    "items": items,
                    "records": records,
                    "accountName": account_name,
                    "captureType": capture_type,
                    "captureResult": capture_result.to_dict(),
                    "foregroundDone": True,
                    "progressCounted": True,
                    "tabClosed": bool(tab_closed),
                    "articleSaved": False,
                    "detailStatus": "pending",
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
            except Exception as exc:
                append_step("失败原因", str(exc))
                return {
                    **result,
                    "ok": False,
                    "status": "failed",
                    "phase": "foreground",
                    "message": f"前台捕获失败：{exc}",
                    "tone": "error",
                    "items": items,
                    "records": records,
                    "accountName": account_name,
                    "captureType": CaptureType.NONE.value,
                    "foregroundDone": bool(tab_closed),
                    "progressCounted": True,
                    "tabClosed": bool(tab_closed),
                    "detailStatus": "failed",
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
            finally:
                self._unregister_active_attempt(job_id, attempt)
                if article_tab is not None and not tab_closed:
                    try:
                        tabs.close_article_tab(
                            article_tab,
                            home_window_handle=target.home_window_handle,
                        )
                    except Exception:
                        pass
                if attempt is not None and not attempt_finished:
                    try:
                        attempt.cancel()
                    except Exception:
                        pass

    def _run_capture_and_save_flow(
        self,
        *,
        job_id: str,
        result: dict[str, Any],
        options: InitialContentStorageTaskOptions,
        update: Callable[[dict[str, Any]], None],
    ) -> dict[str, Any]:
        started_at = time.monotonic()
        items = list(result.get("items") or [])
        records = [dict(record) for record in list(result.get("records") or [])]
        account_name = _optional_text(result.get("accountName")) or _optional_text(
            options.account_name
        ) or ""
        target = _target_from_card(options, records[0] if records else {})
        settings = SingleCaptureSettings.from_app_config(self._config)
        tabs = self._window_factory.create_tab_service()
        clicker = self._window_factory.create_clicker()
        control = self._capture_factory.create_in_process_control()
        attempt: Any | None = None
        article_tab: Any | None = None
        attempt_finished = False
        tab_closed = False
        capture_started_at = self._now()
        capture_started_monotonic = started_at
        foreground_lease = self._foreground_lease()
        foreground_lease.__enter__()
        foreground_lease_held = True

        def publish(message: str) -> None:
            update(
                {
                    **result,
                    "ok": False,
                    "status": "running",
                    "message": message,
                    "tone": "info",
                    "items": list(items),
                    "records": records,
                    "accountName": account_name,
                    "captureType": CaptureType.NONE.value,
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
            )

        def append_step(
            label: str,
            value: str,
            *,
            duration_seconds: float | None = None,
            cells: list[dict[str, Any]] | None = None,
        ) -> None:
            item: dict[str, Any] = {"label": label, "value": value}
            if duration_seconds is not None:
                item["durationSeconds"] = round(max(0.0, duration_seconds), 3)
            if cells:
                item["cells"] = cells
            items.append(item)

        try:
            publish("正在启动 MITM 捕获...")
            step_started = time.monotonic()
            baseline = tabs.capture_baseline()
            append_step(
                "标签基线",
                "已记录点击前浏览器标签状态",
                duration_seconds=time.monotonic() - step_started,
            )

            step_started = time.monotonic()
            attempt_id = uuid4().hex[:12]
            capture_started_at = self._now()
            capture_started_monotonic = step_started
            attempt = control.start_attempt(
                task_id=job_id,
                attempt_id=attempt_id,
                proxy_lease_id=f"{job_id}-{attempt_id}",
                proxy_address=settings.proxy_address,
                capture_config=settings.capture_config,
            )
            self._register_active_attempt(job_id, attempt)
            if self._is_cancelled(job_id):
                try:
                    attempt.cancel()
                except Exception:
                    pass
                return {
                    **result,
                    "ok": False,
                    "status": "cancelled",
                    "phase": "finished",
                    "foregroundDone": False,
                    "progressCounted": False,
                    "tabClosed": False,
                    "detailStatus": "cancelled",
                    "message": "单篇任务已取消。",
                }
            attempt.wait_ready(timeout_seconds=settings.ready_timeout_seconds)
            append_step(
                "MITM 捕获",
                "已启动 MITM 并开启系统代理",
                duration_seconds=time.monotonic() - step_started,
            )

            publish("正在点击文章卡片...")
            step_started = time.monotonic()
            click_result = clicker.click(target)
            append_step(
                "点击文章",
                "已发送文章卡片点击",
                duration_seconds=time.monotonic() - step_started,
                cells=[{"label": "点击方式", "value": str(getattr(click_result, "method", ""))}],
            )

            publish("正在确认文章标签打开...")
            step_started = time.monotonic()
            article_tab = tabs.wait_for_opened_article_tab(
                baseline=baseline,
                timeout_seconds=settings.title_timeout_seconds,
                stable_delay_seconds=settings.title_stable_delay_seconds,
                poll_initial_interval_seconds=(
                    settings.title_poll_initial_interval_seconds
                ),
                poll_max_interval_seconds=settings.title_poll_max_interval_seconds,
            )
            append_step(
                "文章标签",
                "已检测到文章标签打开",
                duration_seconds=time.monotonic() - step_started,
                cells=[{"label": "标签名", "value": str(getattr(article_tab, "title", ""))}],
            )

            publish("正在关闭文章标签...")
            step_started = time.monotonic()
            tabs.close_article_tab(article_tab, home_window_handle=target.home_window_handle)
            tab_closed = True
            append_step(
                "关闭标签",
                "已发送 Ctrl+W 关闭文章标签",
                duration_seconds=time.monotonic() - step_started,
            )

            publish("正在停止 MITM 捕获...")
            step_started = time.monotonic()
            capture_result = attempt.stop_capture(timeout_seconds=settings.result_timeout_seconds)
            attempt_finished = True
            append_step(
                "关闭代理",
                "已恢复系统代理并停止 MITM",
                duration_seconds=time.monotonic() - step_started,
            )
            _append_capture_event_items(items, capture_result)
            capture_type = _capture_type_value(capture_result)
            capture_ok = (
                getattr(capture_result, "status", None) is TaskStatus.SUCCESS
                and capture_type != CaptureType.NONE.value
            )
            if not capture_ok:
                error_message = str(getattr(capture_result, "error_message", "") or "未捕获到 HTML 或 reference")
                append_step("捕获结果", error_message)
                return {
                    **result,
                    "ok": False,
                    "status": "failed",
                    "message": error_message,
                    "tone": "error",
                    "items": items,
                    "records": records,
                    "accountName": account_name,
                    "captureType": capture_type,
                    "phase": "foreground",
                    "foregroundDone": bool(tab_closed),
                    "progressCounted": True,
                    "tabClosed": bool(tab_closed),
                    "detailStatus": "failed",
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }

            append_step("捕获结果", f"已得到 {capture_type}")
            # 代理已经恢复、详情标签已经关闭；主流程此时可以安全读取并分发下一篇。
            update(
                {
                    **result,
                    "ok": False,
                    "status": "running",
                    "phase": "foreground",
                    "foregroundDone": True,
                    "progressCounted": True,
                    "tabClosed": bool(tab_closed),
                    "accountName": account_name,
                    "captureType": capture_type,
                    "items": list(items),
                    "records": records,
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                    "message": "前台采集完成，主页可以继续下一篇。",
                }
            )
            # 详情解析和后置任务不得继续占用主页/代理租约。
            foreground_lease.__exit__(None, None, None)
            foreground_lease_held = False
            publish("正在解析 HTML 并存储初始内容...")
            step_started = time.monotonic()
            context = TaskContext(
                task_id=job_id,
                proxy_lease_id=f"{job_id}-storage",
                db_path=self._database_path,
                storage_root=Path(self._config.storage.article_storage_root),
                temp_dir=Path(self._config.storage.temp_dir) / job_id,
                started_at=capture_started_at,
            )
            save = self._html_save.save(
                context=context,
                target=target,
                capture_result=capture_result,
                attempt_started_at=capture_started_at,
                duration_seconds=max(0.0, time.monotonic() - capture_started_monotonic),
                request_timeout_seconds=self._config.request.request_timeout_seconds,
            )
            save_elapsed = time.monotonic() - step_started
            if save.status is TaskStatus.SUCCESS and save.data is not None:
                saved = save.data
                append_step(
                    "解析 HTML 并存储初始内容",
                    f"{save_elapsed:.3f} 秒",
                    cells=[
                        {"label": "结果", "value": "已解析并保存初始内容"},
                        {"label": "HTML 来源", "value": "无" if not saved.html_source else str(saved.html_source)},
                        {"label": "文章ID", "value": str(saved.article_id)},
                        {"label": "公众号ID", "value": str(saved.account_id)},
                        {"label": "历史ID", "value": str(saved.history_id)},
                        {"label": "归档目录", "value": str(saved.archive_dir)},
                        {"label": "详情文件", "value": str(saved.detail_path)},
                        {
                            "label": "资源清单",
                            "value": "，".join(str(value) for value in saved.resource_manifest.to_json_values()) or "无",
                        },
                    ],
                )
                return self._build_save_success_result(
                    job_id=job_id,
                    base_result=result,
                    options=options,
                    update=update,
                    items=items,
                    records=records,
                    account_name=account_name,
                    capture_type=capture_type,
                    started_at=started_at,
                    context=context,
                    saved=saved,
                )

            error_message = str(getattr(save, "message", "") or "初始内容解析保存失败。")
            append_step(
                "解析 HTML 并存储初始内容",
                f"{save_elapsed:.3f} 秒",
                cells=[
                    {"label": "结果", "value": "解析或保存失败"},
                    {"label": "失败原因", "value": error_message},
                ],
            )
            return {
                **result,
                "ok": False,
                "status": "save-failed",
                "message": error_message,
                "tone": "error",
                "items": items,
                "records": records,
                "accountName": account_name,
                "captureType": capture_type,
                "phase": "detail",
                "foregroundDone": True,
                "progressCounted": True,
                "tabClosed": bool(tab_closed),
                "detailStatus": "failed",
                "articleSaved": False,
                "totalSeconds": round(time.monotonic() - started_at, 3),
            }
        except Exception as exc:
            append_step("失败原因", str(exc))
            return {
                **result,
                "ok": False,
                "status": "failed",
                "message": f"初始内容存储失败：{exc}",
                "tone": "error",
                "items": items,
                "records": records,
                "accountName": account_name,
                "captureType": CaptureType.NONE.value,
                "phase": "foreground",
                "foregroundDone": bool(tab_closed),
                "progressCounted": True,
                "tabClosed": bool(tab_closed),
                "detailStatus": "failed",
                "totalSeconds": round(time.monotonic() - started_at, 3),
            }
        finally:
            self._unregister_active_attempt(job_id, attempt)
            if foreground_lease_held:
                foreground_lease.__exit__(None, None, None)
            if article_tab is not None and not tab_closed:
                try:
                    tabs.close_article_tab(article_tab, home_window_handle=target.home_window_handle)
                except Exception:
                    pass
            if attempt is not None and not attempt_finished:
                try:
                    attempt.cancel()
                except Exception:
                    pass

    def _is_cancelled(self, job_id: str) -> bool:
        with self._lock:
            return str(job_id or "") in self._cancelled_jobs

    def _apply_collected_lookup(
        self,
        result: dict[str, Any],
        options: InitialContentStorageTaskOptions,
    ) -> dict[str, Any]:
        if not result.get("ok"):
            return result
        records = list(result.get("records") or [])
        account_name = _optional_text(options.account_name) or _optional_text(
            result.get("accountName")
        )
        if not records:
            return {
                **result,
                "status": result.get("status") or "no-visible-card",
                "accountName": account_name or "",
            }

        first_record = dict(records[0])
        published_date = _optional_text(first_record.get("publishedDate")) or _optional_text(
            first_record.get("published_date")
        )
        title_fragment = (
            _optional_text(first_record.get("rawTitle"))
            or _optional_text(first_record.get("raw_title"))
            or _optional_text(first_record.get("title"))
        )
        items = list(result.get("items") or [])
        lookup = {
            "enabled": bool(options.skip_collected_records),
            "matched": False,
            "accountName": account_name or "",
            "publishedDate": published_date or "",
            "titleFragment": title_fragment or "",
        }
        if not options.skip_collected_records:
            items.append({"label": "已采集记录校验", "value": "未启用，当前文章可以继续"})
            return {
                **result,
                "ok": True,
                "status": "ready-to-continue",
                "message": "未启用跳过已采集记录，可以继续。",
                "tone": "success",
                "items": items,
                "records": records,
                "accountName": account_name or "",
                "collectedLookup": lookup,
            }
        if not account_name or not title_fragment:
            items.append({"label": "已采集记录校验", "value": "公众号或标题为空，默认可以继续"})
            return {
                **result,
                "ok": True,
                "status": "ready-to-continue",
                "message": "已采集记录校验信息不足，可以继续。",
                "tone": "warning",
                "items": items,
                "records": records,
                "accountName": account_name or "",
                "collectedLookup": lookup,
            }
        matched = self._lookup_service.find_by_account_date_and_title_fragment(
            database_path=self._database_path,
            account_name=account_name,
            published_date=published_date or "",
            title_fragment=title_fragment,
        )
        if matched is not None:
            lookup.update(
                {
                    "matched": True,
                    "matchedArticleId": matched.id,
                    "matchedTitle": matched.article_title,
                    "matchedPublishedTime": matched.published_article_time,
                }
            )
            items.append(
                {
                    "label": "已采集记录校验",
                    "value": "已采集，停止任务",
                    "cells": [
                        {"label": "匹配标题", "value": matched.article_title},
                        {"label": "发布时间", "value": matched.published_article_time},
                    ],
                }
            )
            return {
                **result,
                "ok": True,
                "status": "skipped-collected",
                "message": "检测到本地已采集记录，初始内容存储任务停止。",
                "tone": "warning",
                "items": items,
                "records": records,
                "accountName": account_name,
                "collectedLookup": lookup,
            }
        items.append({"label": "已采集记录校验", "value": "未发现已采集记录，可以继续"})
        return {
            **result,
            "ok": True,
            "status": "ready-to-continue",
            "message": "未发现已采集记录，可以继续。",
            "tone": "success",
            "items": items,
            "records": records,
            "accountName": account_name,
            "collectedLookup": lookup,
        }

    def _public_options_payload(
        self,
        options: InitialContentStorageTaskOptions,
    ) -> dict[str, Any]:
        return {
            **_public_options_payload(options),
            **self._extra_public_options,
        }

    def _build_save_success_result(
        self,
        *,
        job_id: str,
        base_result: dict[str, Any],
        options: InitialContentStorageTaskOptions,
        update: Callable[[dict[str, Any]], None],
        items: list[dict[str, Any]],
        records: list[dict[str, Any]],
        account_name: str,
        capture_type: str,
        started_at: float,
        context: TaskContext,
        saved: Any,
    ) -> dict[str, Any]:
        total_seconds = time.monotonic() - started_at
        items.append({"label": "总耗时", "value": f"{total_seconds:.3f} 秒"})
        result = {
            **base_result,
            "ok": True,
            "status": "completed",
            "phase": "detail",
            "foregroundDone": True,
            "progressCounted": True,
            "tabClosed": True,
            "articleSaved": True,
            "detailStatus": "success",
            "message": f"初始内容存储完成，HTML 来源：{saved.html_source}。",
            "tone": "success",
            "items": items,
            "records": records,
            "accountName": account_name,
            "captureType": capture_type,
            "totalSeconds": round(total_seconds, 3),
            "htmlSource": saved.html_source,
            "archiveDir": saved.archive_dir,
            "articleId": saved.article_id,
            "accountId": saved.account_id,
            "historyId": saved.history_id,
            "attemptId": saved.attempt_id,
            "articleDirectory": str(saved.article_directory),
            "articleTitle": _record_title(records),
            "resourceManifest": saved.resource_manifest.to_json_values(),
        }
        return result

    def _trim_jobs(self) -> None:
        if len(self._jobs) <= 20:
            return
        removable = [
            job_id
            for job_id, job in self._jobs.items()
            if str(job.get("status") or "") not in _ACTIVE_STATUSES
        ]
        for job_id in removable[: max(0, len(self._jobs) - 20)]:
            self._jobs.pop(job_id, None)


def _manual_card_result(options: InitialContentStorageTaskOptions) -> dict[str, Any]:
    card = _safe_card_payload(options.card) or {}
    card.setdefault("index", options.card_index)
    raw_title = _optional_text(card.get("rawTitle")) or _optional_text(
        card.get("title")
    ) or "未识别标题"
    return {
        "ok": True,
        "status": "completed",
        "message": "已接收单篇文章卡片信息。",
        "tone": "success",
        "items": [{"kind": "article", "label": f"第{options.card_index}条文章", "value": raw_title}],
        "records": [card],
        "accountName": options.account_name or "",
        "captureType": CaptureType.NONE.value,
    }


def _missing_task_input_result(options: InitialContentStorageTaskOptions) -> dict[str, Any]:
    return {
        "ok": False,
        "status": "failed",
        "message": "缺少文章卡片输入，初始内容存储 Huey 任务未执行窗口读取。",
        "tone": "error",
        "items": [
            {
                "label": "任务输入",
                "value": "缺少文章卡片信息",
                "cells": [
                    {"label": "处理方式", "value": "请先由后端同步读取主页首篇卡片，再启动 Huey 任务"},
                ],
            }
        ],
        "records": [],
        "accountName": options.account_name or "",
        "captureType": CaptureType.NONE.value,
    }


def _initial_card_items(options: InitialContentStorageTaskOptions) -> list[dict[str, Any]]:
    if options.card is None:
        return []
    card = _safe_card_payload(options.card) or {}
    raw_title = _optional_text(card.get("rawTitle")) or _optional_text(
        card.get("title")
    ) or "未识别标题"
    return [
        {
            "kind": "article",
            "label": f"第{options.card_index}条文章",
            "value": raw_title,
        }
    ]


def _record_title(records: list[dict[str, Any]]) -> str:
    if not records:
        return ""
    record = records[0]
    return str(
        record.get("rawTitle")
        or record.get("raw_title")
        or record.get("title")
        or ""
    ).strip()


def _saved_article_from_payload(article: Mapping[str, Any]) -> Any:
    """把详情阶段的可序列化结果恢复为后置服务所需的最小对象。"""

    raw_manifest = article.get("resourceManifest") or article.get("resource_manifest") or []
    resource_types = []
    for value in raw_manifest if isinstance(raw_manifest, (list, tuple)) else []:
        try:
            resource_types.append(ResourceType(str(value)))
        except ValueError:
            continue
    return SimpleNamespace(
        article_id=int(article.get("articleId") or article.get("article_id") or 0),
        account_id=int(article.get("accountId") or article.get("account_id") or 0),
        history_id=int(article.get("historyId") or article.get("history_id") or 0),
        article_directory=Path(
            str(article.get("articleDirectory") or article.get("article_directory") or "")
        ).resolve(),
        archive_dir=str(article.get("archiveDir") or article.get("archive_dir") or ""),
        html_source=str(article.get("htmlSource") or article.get("html_source") or ""),
        resource_manifest=ResourceManifest.from_types(resource_types),
        attempt_id=str(article.get("attemptId") or article.get("attempt_id") or ""),
    )


def _detail_card_payload(card: Mapping[str, Any]) -> dict[str, Any]:
    """详情阶段不再点击窗口，缺失坐标时只补保存所需的最小 ArticleTarget 字段。"""

    payload = {str(key): value for key, value in dict(card).items()}
    payload.setdefault("homeWindowHandle", 1)
    payload.setdefault("clickPoint", [0, 0])
    return payload


def _config_path(config: Any, section: str, field: str, fallback: str | Path) -> Path:
    section_value = getattr(config, section, None)
    value = getattr(section_value, field, None)
    if value is None:
        value = fallback
    return Path(value)


def _target_from_card(
    options: InitialContentStorageTaskOptions,
    card: dict[str, Any],
) -> ArticleTarget:
    raw_title = (
        _optional_text(card.get("rawTitle"))
        or _optional_text(card.get("raw_title"))
        or _optional_text(card.get("title"))
        or "未识别标题"
    )
    click_x, click_y = _click_point_from_card(card)
    home_window_handle = int(card.get("homeWindowHandle") or card.get("home_window_handle") or 0)
    if home_window_handle <= 0:
        raise RuntimeError("文章卡片缺少主页窗口句柄，不能执行后台点击")
    account_name = (
        _optional_text(options.account_name)
        or _optional_text(card.get("accountName"))
        or _optional_text(card.get("account_name"))
        or ""
    )
    published_date = _optional_text(card.get("publishedDate")) or _optional_text(
        card.get("published_date")
    ) or ""
    return ArticleTarget(
        account_name=account_name,
        title=_optional_text(card.get("title")) or raw_title,
        raw_title=raw_title,
        click_x=click_x,
        click_y=click_y,
        home_window_handle=home_window_handle,
        fingerprint=f"{published_date}|{raw_title}",
        date_text=_optional_text(card.get("dateText")) or _optional_text(card.get("date_text")) or "",
        published_date=published_date,
        date_rect=_rect_from_card(card, "dateRect", "date_rect"),
        title_rect=_rect_from_card(card, "titleRect", "title_rect"),
        metric_text=_optional_text(card.get("metricText")) or _optional_text(card.get("metric_text")) or "",
        metric_rect=_rect_from_card(card, "metricRect", "metric_rect"),
    )


def _click_point_from_card(card: dict[str, Any]) -> tuple[int, int]:
    point = card.get("clickPoint") or card.get("click_point")
    if isinstance(point, (list, tuple)) and len(point) >= 2:
        return int(point[0]), int(point[1])
    visible_rect = _rect_from_value(card.get("visibleRect") or card.get("visible_rect"))
    if visible_rect is not None:
        left, top, right, bottom = visible_rect
        return int((left + right) / 2), int((top + bottom) / 2)
    raise RuntimeError("文章卡片缺少点击坐标")


def _rect_from_card(
    card: dict[str, Any],
    camel_key: str,
    snake_key: str,
) -> tuple[int, int, int, int] | None:
    return _rect_from_value(card.get(camel_key) or card.get(snake_key))


def _rect_from_value(value: Any) -> tuple[int, int, int, int] | None:
    if not isinstance(value, (list, tuple)) or len(value) < 4:
        return None
    try:
        left, top, right, bottom = (int(value[0]), int(value[1]), int(value[2]), int(value[3]))
    except Exception:
        return None
    if right <= left or bottom <= top:
        return None
    return left, top, right, bottom


def _capture_type_value(capture_result: Any) -> str:
    capture_type = getattr(capture_result, "capture_type", CaptureType.NONE)
    return getattr(capture_type, "value", str(capture_type or CaptureType.NONE.value))


def _append_capture_event_items(
    items: list[dict[str, Any]],
    capture_result: Any,
) -> None:
    for event in getattr(capture_result, "capture_events", ()) or ():
        if not isinstance(event, dict):
            continue
        name = _optional_text(event.get("name")) or _optional_text(event.get("stage")) or "MITM 子步骤"
        cells: list[dict[str, Any]] = []
        elapsed = event.get("elapsed_seconds")
        if elapsed is not None:
            cells.append({"label": "耗时", "value": f"{float(elapsed):.3f} 秒"})
        capture_type = event.get("capture_type_after_event") or event.get("capture_type")
        if capture_type:
            cells.append({"label": "捕获类型", "value": str(capture_type)})
        items.append({"label": "MITM 子步骤", "value": name, "cells": cells})


def _task_payload(options: InitialContentStorageTaskOptions) -> dict[str, Any]:
    return {
        "cardIndex": options.card_index,
        "accountName": options.account_name,
        "card": options.card,
        "skipCollectedRecords": options.skip_collected_records,
        "storeArticleDetail": True,
        "statefulOfflineCache": options.stateful_offline_cache,
    }


def _public_options_payload(options: InitialContentStorageTaskOptions) -> dict[str, Any]:
    return {
        "skipCollectedRecords": options.skip_collected_records,
        "storeArticleDetail": True,
    }


def _options_from_payload(payload: dict[str, Any]) -> InitialContentStorageTaskOptions:
    return InitialContentStorageTaskOptions(
        card_index=max(1, int(payload.get("cardIndex") or 1)),
        account_name=_optional_text(payload.get("accountName")),
        card=_safe_card_payload(payload.get("card")),
        skip_collected_records=bool(payload.get("skipCollectedRecords")),
        store_article_detail=True,
        stateful_offline_cache=bool(payload.get("statefulOfflineCache")),
    )


def _safe_card_payload(value: Any) -> dict[str, Any] | None:
    if not isinstance(value, dict):
        return None
    return {str(key): item for key, item in value.items()}


def _safe_identifier(value: Any, *, fallback: str = "job") -> str:
    normalized = _SAFE_ID_PATTERN.sub("-", str(value or "").strip()).strip("-._")
    return normalized[:64] or fallback


def _optional_text(value: Any) -> str | None:
    normalized = str(value or "").strip()
    return normalized or None


__all__ = [
    "InitialContentStorageHueyService",
    "InitialContentStorageTaskOptions",
]
