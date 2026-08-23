from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from threading import Event, RLock
from typing import Any, Callable
from uuid import uuid4

from huey import SqliteHuey
from huey.consumer import Consumer

from src.modules.system.window_diagnostic_trace_store import (
    WindowDiagnosticTraceStore,
)
from src.services.runtime.window_click_flow_diagnostic_service import (
    WindowClickFlowDiagnosticService,
)
from src.services.task.huey_runtime_queue import (
    ensure_huey_storage_ready,
    runtime_huey_queue_dir,
)


_SAFE_ID_PATTERN = re.compile(r"[^A-Za-z0-9._-]+")
_ACTIVE_STATUSES = {"running", "stop-requested"}


class WindowClickFlowConflictError(RuntimeError):
    """同一时间已有窗口读取诊断正在操作微信主页。"""


@dataclass(frozen=True, slots=True)
class WindowClickFlowTaskOptions:
    max_records: int = 20
    date_filter_mode: str = "all"
    start_date: str | None = None
    end_date: str | None = None


class _EmbeddedThreadConsumer(Consumer):
    """应用自行管理生命周期，避免Huey覆盖桌面程序的系统信号处理器。"""

    def _set_signal_handlers(self) -> None:
        return None


class WindowClickFlowHueyService:
    """使用会话级SqliteHuey执行独立的窗口读取诊断。"""

    def __init__(
        self,
        *,
        temp_root: str | Path,
        config: Any,
        window_factory: Any,
        runner: Callable[..., dict[str, Any]] | None = None,
        session_id: str | None = None,
        job_id_factory: Callable[[], str] | None = None,
        now: Callable[[], datetime] = datetime.now,
    ) -> None:
        self._temp_root = Path(temp_root).resolve()
        self._config = config
        self._window_factory = window_factory
        self._runner = runner
        self._job_id_factory = job_id_factory or (lambda: uuid4().hex[:12])
        self._now = now
        self._jobs: dict[str, dict[str, Any]] = {}
        self._stop_flags: dict[str, Event] = {}
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
            queue_dir / f"window-click-flow-{normalized_session_id}.sqlite3"
        )
        self._huey = SqliteHuey(
            "window-click-flow",
            filename=str(self._queue_database_path),
            results=True,
            store_none=False,
        )
        self._task_wrapper = self._huey.task(
            retries=0,
            name="WindowClickFlowDiagnosticTask",
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
        max_records: int = 20,
        date_filter_mode: str = "all",
        start_date: str | None = None,
        end_date: str | None = None,
    ) -> dict[str, Any]:
        options = WindowClickFlowTaskOptions(
            max_records=max(0, int(max_records)),
            date_filter_mode=str(date_filter_mode or "all"),
            start_date=_optional_text(start_date),
            end_date=_optional_text(end_date),
        )
        with self._lock:
            if self._closed:
                raise RuntimeError("Huey窗口诊断服务已经关闭")
            active_job = next(
                (
                    job
                    for job in self._jobs.values()
                    if str(job.get("status") or "") in _ACTIVE_STATUSES
                ),
                None,
            )
            if active_job is not None:
                raise WindowClickFlowConflictError(
                    f"窗口读取诊断正在运行：{active_job.get('jobId', '')}"
                )

            job_id = f"window-click-flow-{_safe_identifier(self._job_id_factory())}"
            stop_event = Event()
            trace_store = WindowDiagnosticTraceStore(
                temp_root=self._temp_root,
                job_id=job_id,
            )
            trace_store.append_event(
                {
                    "event": "diagnostic-start",
                    "status": "running",
                    "message": "窗口读取诊断已提交到Huey队列",
                    "details": {"options": _options_payload(options)},
                }
            )
            task = self._task_wrapper.s(
                job_id,
                options.max_records,
                options.date_filter_mode,
                options.start_date,
                options.end_date,
            )
            initial = {
                "ok": False,
                "status": "running",
                "jobId": job_id,
                "hueyTaskId": task.id,
                "action": "window-click-flow",
                "title": "主页内容读取结果",
                "message": "正在等待Huey执行主页内容读取测试...",
                "tone": "info",
                "items": [],
                "records": [],
                "events": [],
                "recognizedCount": 0,
                "skippedCount": 0,
                "stoppedByUser": False,
                "startedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                "options": _options_payload(options),
                **trace_store.path_fields(),
            }
            self._jobs[job_id] = initial
            self._stop_flags[job_id] = stop_event
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
        """供缓存清理等入口判断窗口诊断是否仍在占用运行资源。"""
        with self._lock:
            return any(
                str(job.get("status") or "") in _ACTIVE_STATUSES
                for job in self._jobs.values()
            )

    def stop(self, job_id: str) -> dict[str, Any]:
        with self._lock:
            job = self._jobs.get(str(job_id))
            if job is None:
                raise KeyError(job_id)
            if str(job.get("status") or "") not in _ACTIVE_STATUSES:
                return dict(job)
            stop_event = self._stop_flags.get(str(job_id))
            if stop_event is not None:
                stop_event.set()
            stopping = {
                **job,
                "ok": False,
                "status": "stop-requested",
                "message": "已请求停止主页内容读取，等待当前读取收尾...",
                "tone": "warning",
                "stoppedByUser": True,
            }
            self._jobs[str(job_id)] = stopping
            return dict(stopping)

    def shutdown(self) -> None:
        with self._lock:
            if self._closed:
                return
            self._closed = True
            for stop_event in self._stop_flags.values():
                stop_event.set()
            consumer_started = self._consumer_started
        if consumer_started:
            self._consumer.stop(graceful=True)
        # Consumer只负责停止线程，Windows下还需显式关闭SQLite连接，
        # 否则应用退出后会继续占用本次会话的队列文件。
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

    def _execute_task(
        self,
        job_id: str,
        max_records: int,
        date_filter_mode: str,
        start_date: str | None,
        end_date: str | None,
    ) -> None:
        options = WindowClickFlowTaskOptions(
            max_records=int(max_records),
            date_filter_mode=str(date_filter_mode),
            start_date=_optional_text(start_date),
            end_date=_optional_text(end_date),
        )
        trace_store = WindowDiagnosticTraceStore(
            temp_root=self._temp_root,
            job_id=job_id,
        )
        trace_paths = trace_store.path_fields()

        def update(payload: dict[str, Any]) -> None:
            merged = {
                "jobId": job_id,
                "action": "window-click-flow",
                "title": "主页内容读取结果",
                **trace_paths,
                **payload,
            }
            with self._lock:
                current = self._jobs.get(job_id, {})
                self._jobs[job_id] = {**current, **merged}

        def stop_requested() -> bool:
            with self._lock:
                stop_event = self._stop_flags.get(job_id)
            return bool(stop_event is not None and stop_event.is_set())

        try:
            if stop_requested():
                result = _stopped_before_start_payload()
            elif self._runner is not None:
                result = self._runner(
                    options=options,
                    on_update=update,
                    stop_requested=stop_requested,
                    trace_store=trace_store,
                )
            else:
                result = WindowClickFlowDiagnosticService(
                    config=self._config,
                    window_factory=self._window_factory,
                ).run(
                    max_records=options.max_records,
                    date_filter_mode=options.date_filter_mode,
                    start_date=options.start_date,
                    end_date=options.end_date,
                    stop_requested=stop_requested,
                    on_update=update,
                    trace_store=trace_store,
                )
            if not isinstance(result, dict):
                raise RuntimeError("主页内容读取诊断返回了无法识别的结果")
            final = {
                "jobId": job_id,
                "action": "window-click-flow",
                "title": "主页内容读取结果",
                **result,
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                **trace_paths,
            }
            trace_store.write_result(final)
            update(final)
        except Exception as exc:
            failed = {
                "jobId": job_id,
                "action": "window-click-flow",
                "title": "主页内容读取结果",
                "ok": False,
                "status": "failed",
                "message": f"主页内容读取测试失败：{exc}",
                "tone": "error",
                "items": [],
                "records": [],
                "events": [],
                "recognizedCount": 0,
                "skippedCount": 0,
                "stoppedByUser": stop_requested(),
                "finishedAt": self._now().strftime("%Y-%m-%d %H:%M:%S"),
                **trace_paths,
            }
            try:
                trace_store.append_event(
                    {
                        "event": "diagnostic-failed",
                        "status": "failed",
                        "message": str(exc),
                        "details": {"exceptionType": type(exc).__name__},
                    }
                )
                trace_store.write_result(failed)
            except OSError as trace_exc:
                failed["traceError"] = str(trace_exc)
            update(failed)
        finally:
            with self._lock:
                self._stop_flags.pop(job_id, None)

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
            self._stop_flags.pop(job_id, None)


def _safe_identifier(value: Any, *, fallback: str = "job") -> str:
    normalized = _SAFE_ID_PATTERN.sub("-", str(value or "").strip()).strip("-._")
    return normalized[:64] or fallback


def _optional_text(value: Any) -> str | None:
    normalized = str(value or "").strip()
    return normalized or None


def _options_payload(options: WindowClickFlowTaskOptions) -> dict[str, Any]:
    return {
        "maxRecords": options.max_records,
        "dateFilterMode": options.date_filter_mode,
        "startDate": options.start_date,
        "endDate": options.end_date,
    }


def _stopped_before_start_payload() -> dict[str, Any]:
    return {
        "ok": True,
        "status": "stopped",
        "message": "已停止主页内容读取",
        "tone": "warning",
        "items": [],
        "records": [],
        "events": [],
        "recognizedCount": 0,
        "skippedCount": 0,
        "stoppedByUser": True,
    }


__all__ = [
    "WindowClickFlowConflictError",
    "WindowClickFlowHueyService",
    "WindowClickFlowTaskOptions",
]
