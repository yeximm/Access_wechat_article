from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from threading import RLock
import time
from typing import Any, Callable
from uuid import uuid4

from huey import SqliteHuey
from huey.consumer import Consumer

from src.domain.enums import CaptureType, TaskStatus
from src.domain.models import ArticleTarget
from src.services.capture.collected_article_lookup_service import (
    CollectedArticleLookupService,
)
from src.services.capture.single_article_capture_service import SingleCaptureSettings
from src.services.task.huey_runtime_queue import (
    ensure_huey_storage_ready,
    runtime_huey_queue_dir,
)


_SAFE_ID_PATTERN = re.compile(r"[^A-Za-z0-9._-]+")
_ACTIVE_STATUSES = {"running"}


@dataclass(frozen=True, slots=True)
class SingleArticleDetailTaskOptions:
    """单篇详情诊断任务输入，后续点击/MITM流程也从这里扩展。"""

    card_index: int = 1
    account_name: str | None = None
    card: dict[str, Any] | None = None
    skip_collected_records: bool = False


class _EmbeddedThreadConsumer(Consumer):
    """桌面应用内嵌 Huey worker，避免接管主进程信号处理。"""

    def _set_signal_handlers(self) -> None:
        return None


class SingleArticleDetailHueyService:
    """使用会话级 SqliteHuey 执行单篇文章详情诊断任务。"""

    def __init__(
        self,
        *,
        temp_root: str | Path,
        config: Any,
        window_factory: Any,
        capture_factory: Any,
        database_path: str | Path,
        runner: Callable[..., dict[str, Any]] | None = None,
        lookup_service: CollectedArticleLookupService | None = None,
        session_id: str | None = None,
        job_id_factory: Callable[[], str] | None = None,
        now: Callable[[], datetime] = datetime.now,
    ) -> None:
        self._temp_root = Path(temp_root).resolve()
        self._config = config
        self._window_factory = window_factory
        self._capture_factory = capture_factory
        self._database_path = Path(database_path)
        self._runner = runner
        self._lookup_service = lookup_service or CollectedArticleLookupService()
        self._job_id_factory = job_id_factory or (lambda: uuid4().hex[:12])
        self._now = now
        self._jobs: dict[str, dict[str, Any]] = {}
        self._lock = RLock()
        self._consumer_started = False
        self._closed = False

        normalized_session_id = _safe_identifier(
            session_id or uuid4().hex[:12],
            fallback="session",
        )
        queue_dir = runtime_huey_queue_dir(self._temp_root)
        queue_dir.mkdir(parents=True, exist_ok=True)
        self._queue_database_path = (
            queue_dir / f"single-article-detail-{normalized_session_id}.sqlite3"
        )
        self._huey = SqliteHuey(
            "single-article-detail",
            filename=str(self._queue_database_path),
            results=True,
            store_none=False,
        )
        self._task_wrapper = self._huey.task(
            retries=0,
            name="SingleArticleDetailTask",
        )(self._execute_task)
        self._consumer = _EmbeddedThreadConsumer(
            self._huey,
            workers=1,
            worker_type="thread",
            periodic=False,
            check_worker_health=False,
        )

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
    ) -> dict[str, Any]:
        options = SingleArticleDetailTaskOptions(
            card_index=max(1, int(card_index)),
            account_name=_optional_text(account_name),
            card=_safe_card_payload(card),
            skip_collected_records=bool(skip_collected_records),
        )
        with self._lock:
            if self._closed:
                raise RuntimeError("Huey单篇详情服务已经关闭")
            job_id = f"article-detail-{_safe_identifier(self._job_id_factory())}"
            task = self._task_wrapper.s(job_id, _options_payload(options))
            initial = {
                "ok": False,
                "status": "running",
                "jobId": job_id,
                "hueyTaskId": task.id,
                "action": "single-article-detail",
                "title": "详情获取结果",
                "message": (
                    "已读取首篇文章卡片，正在等待Huey执行单篇文章详情任务..."
                    if options.card is not None
                    else "正在等待Huey执行单篇文章详情任务..."
                ),
                "tone": "info",
                "items": [
                    {"label": "流程", "value": "单篇文章详情流程"},
                    {
                        "label": "跳过已采集记录",
                        "value": "开启" if options.skip_collected_records else "关闭",
                    },
                    {"label": "状态", "value": "等待执行"},
                ]
                + _account_name_items(options.account_name)
                + _initial_card_items(options),
                "records": [options.card] if options.card is not None else [],
                "accountName": options.account_name or "",
                "captureType": CaptureType.NONE.value,
                "startedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": _options_payload(options),
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

    def is_active(self) -> bool:
        with self._lock:
            return any(
                str(job.get("status") or "") in _ACTIVE_STATUSES
                for job in self._jobs.values()
            )

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
                "action": "single-article-detail",
                "title": "详情获取结果",
                **payload,
            }
            with self._lock:
                current = self._jobs.get(job_id, {})
                self._jobs[job_id] = {**current, **merged}

        try:
            should_run_capture = False
            if self._runner is not None:
                result = self._runner(options=options, on_update=update)
            elif options.card is not None:
                result = _manual_card_result(options)
                should_run_capture = True
            else:
                result = _missing_task_input_result(options)
            if not isinstance(result, dict):
                raise RuntimeError("单篇详情任务返回了无法识别的结果")
            final = self._apply_collected_lookup(result, options)
            if should_run_capture and final.get("status") == "ready-to-continue":
                final = self._run_capture_flow(
                    job_id=job_id,
                    result=final,
                    options=options,
                    update=update,
                )
            final = {
                "jobId": job_id,
                "action": "single-article-detail",
                "title": "详情获取结果",
                **final,
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            final.setdefault("captureType", CaptureType.NONE.value)
            update(final)
        except Exception as exc:
            update(
                {
                    "ok": False,
                    "status": "failed",
                    "message": f"单篇文章详情任务失败：{exc}",
                    "tone": "error",
                    "items": [{"label": "失败原因", "value": str(exc)}],
                    "records": [],
                    "captureType": CaptureType.NONE.value,
                    "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    def _run_capture_flow(
        self,
        *,
        job_id: str,
        result: dict[str, Any],
        options: SingleArticleDetailTaskOptions,
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
            attempt = control.start_attempt(
                task_id=job_id,
                attempt_id=attempt_id,
                proxy_lease_id=f"{job_id}-{attempt_id}",
                proxy_address=settings.proxy_address,
                capture_config=settings.capture_config,
            )
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
            ok = (
                getattr(capture_result, "status", None) is TaskStatus.SUCCESS
                and capture_type != CaptureType.NONE.value
            )
            if ok:
                append_step("捕获结果", f"已得到 {capture_type}")
                return {
                    **result,
                    "ok": True,
                    "status": "completed",
                    "message": f"单篇详情获取完成，捕获类型：{capture_type}。",
                    "tone": "success",
                    "items": items,
                    "records": records,
                    "accountName": account_name,
                    "captureType": capture_type,
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
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
                "totalSeconds": round(time.monotonic() - started_at, 3),
            }
        except Exception as exc:
            append_step("失败原因", str(exc))
            return {
                **result,
                "ok": False,
                "status": "failed",
                "message": f"单篇详情获取失败：{exc}",
                "tone": "error",
                "items": items,
                "records": records,
                "accountName": account_name,
                "captureType": CaptureType.NONE.value,
                "totalSeconds": round(time.monotonic() - started_at, 3),
            }
        finally:
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

    def _apply_collected_lookup(
        self,
        result: dict[str, Any],
        options: SingleArticleDetailTaskOptions,
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
        items = _with_account_name_item(
            list(result.get("items") or []),
            account_name,
        )
        lookup = {
            "enabled": bool(options.skip_collected_records),
            "matched": False,
            "accountName": account_name or "",
            "publishedDate": published_date or "",
            "titleFragment": title_fragment or "",
            "reason": "not-enabled" if not options.skip_collected_records else "",
        }
        if not options.skip_collected_records:
            items.append(
                {
                    "label": "已采集记录校验",
                    "value": "未启用，当前文章可以继续",
                    "cells": _lookup_cells(account_name, published_date, title_fragment),
                }
            )
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
            lookup["reason"] = "missing-input"
            items.append(
                {
                    "label": "已采集记录校验",
                    "value": "公众号或标题为空，默认可以继续",
                    "cells": _lookup_cells(
                        account_name,
                        published_date,
                        title_fragment,
                        reason="缺少公众号名称或标题",
                    ),
                }
            )
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
        lookup_result = self._lookup_service.lookup_by_account_date_and_title_fragment(
            database_path=self._database_path,
            account_name=account_name,
            published_date=published_date or "",
            title_fragment=title_fragment,
        )
        matched = lookup_result.record
        lookup.update(
            {
                "reason": lookup_result.reason,
                "accountId": lookup_result.account_id,
            }
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
                    "cells": _lookup_cells(
                        account_name,
                        published_date,
                        title_fragment,
                        reason="已命中本地文章索引",
                    )
                    + [
                        {"label": "匹配标题", "value": matched.article_title},
                        {"label": "发布时间", "value": matched.published_article_time},
                    ],
                }
            )
            return {
                **result,
                "ok": True,
                "status": "skipped-collected",
                "message": "检测到本地已采集记录，单篇任务停止。",
                "tone": "warning",
                "items": items,
                "records": records,
                "accountName": account_name,
                "collectedLookup": lookup,
            }
        items.append(
            {
                "label": "已采集记录校验",
                "value": _lookup_miss_message(lookup_result.reason),
                "cells": _lookup_cells(
                    account_name,
                    published_date,
                    title_fragment,
                    reason=_lookup_miss_reason_label(lookup_result.reason),
                ),
            }
        )
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


def _manual_card_result(options: SingleArticleDetailTaskOptions) -> dict[str, Any]:
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


def _missing_task_input_result(options: SingleArticleDetailTaskOptions) -> dict[str, Any]:
    return {
        "ok": False,
        "status": "failed",
        "message": "缺少文章卡片输入，单篇 Huey 任务未执行窗口读取。",
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


def _account_name_items(account_name: str | None) -> list[dict[str, Any]]:
    """把公众号名称作为诊断弹窗的固定展示项，避免只存在于顶层字段。"""
    return [{"label": "公众号名称", "value": _optional_text(account_name) or "未识别"}]


def _with_account_name_item(
    items: list[dict[str, Any]],
    account_name: str | None,
) -> list[dict[str, Any]]:
    if any(str(item.get("label") or "") == "公众号名称" for item in items):
        return items
    return _account_name_items(account_name) + items


def _lookup_cells(
    account_name: str | None,
    published_date: str | None,
    title_fragment: str | None,
    *,
    reason: str | None = None,
) -> list[dict[str, Any]]:
    cells = [
        {"label": "校验公众号", "value": _optional_text(account_name) or "未识别"},
        {"label": "校验日期", "value": _optional_text(published_date) or "未读取"},
        {"label": "校验标题片段", "value": _optional_text(title_fragment) or "未识别"},
    ]
    if reason:
        cells.append({"label": "未命中原因", "value": reason})
    return cells


def _lookup_miss_message(reason: str) -> str:
    if reason == "account-not-found":
        return "公众号索引不存在，未发现已采集记录，可以继续"
    if reason == "article-not-found":
        return "公众号存在，但未匹配到相同日期和标题，可以继续"
    if reason == "missing-input":
        return "公众号或标题为空，默认可以继续"
    return "未发现已采集记录，可以继续"


def _lookup_miss_reason_label(reason: str) -> str:
    if reason == "account-not-found":
        return "公众号索引不存在"
    if reason == "article-not-found":
        return "同公众号下未匹配到相同日期和标题"
    if reason == "missing-input":
        return "缺少公众号名称或标题"
    return "未命中本地文章索引"


def _initial_card_items(options: SingleArticleDetailTaskOptions) -> list[dict[str, Any]]:
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


def _target_from_card(
    options: SingleArticleDetailTaskOptions,
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


def _options_payload(options: SingleArticleDetailTaskOptions) -> dict[str, Any]:
    return {
        "cardIndex": options.card_index,
        "accountName": options.account_name,
        "card": options.card,
        "skipCollectedRecords": options.skip_collected_records,
    }


def _options_from_payload(payload: dict[str, Any]) -> SingleArticleDetailTaskOptions:
    return SingleArticleDetailTaskOptions(
        card_index=max(1, int(payload.get("cardIndex") or 1)),
        account_name=_optional_text(payload.get("accountName")),
        card=_safe_card_payload(payload.get("card")),
        skip_collected_records=bool(payload.get("skipCollectedRecords")),
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
    "SingleArticleDetailHueyService",
    "SingleArticleDetailTaskOptions",
]
