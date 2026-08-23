from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Event, RLock, Thread
from typing import Any, Callable
from uuid import uuid4

from .main_flow_models import (
    MainFlowCommand,
    MainFlowContext,
    MainFlowSnapshot,
)
from .main_flow_state import MainFlowState


class MainFlowConflictError(RuntimeError):
    """同一时间已有一个主流程接管任务。"""

    def __init__(self, owner_task_id: str) -> None:
        super().__init__(f"主流程任务正在运行：{owner_task_id}")
        self.owner_task_id = owner_task_id


@dataclass(slots=True)
class _ManagedMainFlow:
    context: MainFlowContext
    command: MainFlowCommand
    state: MainFlowState
    thread: Thread | None = None
    finished_at: datetime | None = None


class MainFlowService:
    """主服务生命周期入口；具体主页扫描和单篇分发通过 runner 注入。"""

    def __init__(
        self,
        *,
        project_root: str | Path,
        config: Any,
        db_path: str | Path,
        storage_root: str | Path,
        temp_root: str | Path,
        runtime_logger: Any | None = None,
        runner: Callable[[MainFlowContext, MainFlowCommand], None] | None = None,
        now: Callable[[], datetime] = datetime.now,
        id_factory: Callable[[], str] | None = None,
        shutdown_callback: Callable[[], None] | None = None,
    ) -> None:
        self._project_root = Path(project_root)
        self._config = config
        self._db_path = Path(db_path)
        self._storage_root = Path(storage_root)
        self._temp_root = Path(temp_root)
        self._runtime_logger = runtime_logger
        self._runner = runner
        self._now = now
        self._id_factory = id_factory or (lambda: uuid4().hex)
        self._shutdown_callback = shutdown_callback
        self._lock = RLock()
        self._tasks: dict[str, _ManagedMainFlow] = {}
        self._active_task_id: str | None = None
        self._last_task_id: str | None = None
        self._summary_event_keys: set[str] = set()

    @property
    def active_task_id(self) -> str | None:
        with self._lock:
            return self._active_task_id

    def start(self, command: MainFlowCommand) -> MainFlowSnapshot:
        with self._lock:
            if self._active_task_id is not None:
                active = self._tasks.get(self._active_task_id)
                if active is not None and active.state.snapshot()["status"] in {
                    "starting",
                    "running",
                    "stopping",
                }:
                    raise MainFlowConflictError(self._active_task_id)
            task_id = f"main-{self._id_factory()}"
            state = MainFlowState(
                task_id=task_id,
                target_count=command.target_count,
                now=self._now,
            )
            cancel_token = Event()
            context = MainFlowContext(
                task_id=task_id,
                db_path=self._db_path,
                storage_root=self._storage_root,
                temp_dir=self._temp_root / task_id,
                config_snapshot=self._config,
                cancel_token=cancel_token,
                started_at=self._now(),
                state=state,
                # 事件入口同时负责刷新运行状态和写前台摘要日志；
                # 具体采集能力仍由下游服务完成，主流程只做编排和消息转发。
                event_sink=lambda event: self._handle_context_event(state, event),
            )
            managed = _ManagedMainFlow(
                context=context,
                command=command,
                state=state,
            )
            self._tasks[task_id] = managed
            self._active_task_id = task_id
            self._last_task_id = task_id
            thread = Thread(
                target=self._run,
                args=(managed,),
                name=f"awa-main-flow-{task_id}",
                daemon=True,
            )
            managed.thread = thread
            state.set_starting()
            self._write_summary(
                "INFO",
                self._start_message(command),
                source="main-flow",
                channel="main_flow",
                phase="startup",
            )
            initial_snapshot = self._snapshot(managed)
            thread.start()
            return initial_snapshot

    def stop(self, task_id: str | None = None) -> bool:
        with self._lock:
            selected_id = task_id or self._active_task_id
            managed = self._tasks.get(selected_id or "")
            if managed is None:
                return False
            status = managed.state.snapshot()["status"]
            if status in {"completed", "failed", "cancelled"}:
                return False
            managed.context.cancel_token.set()
            managed.state.request_stop()
            return True

    def get(self, task_id: str) -> MainFlowSnapshot:
        with self._lock:
            managed = self._tasks.get(task_id)
            if managed is None:
                raise KeyError(f"主流程任务不存在：{task_id}")
            return self._snapshot(managed)

    def current(self) -> MainFlowSnapshot | None:
        with self._lock:
            selected_id = self._active_task_id or self._last_task_id
            if selected_id is None:
                return None
            managed = self._tasks.get(selected_id)
            return None if managed is None else self._snapshot(managed)

    def wait(self, task_id: str, timeout_seconds: float | None = None) -> MainFlowSnapshot:
        with self._lock:
            managed = self._tasks.get(task_id)
            if managed is None:
                raise KeyError(f"主流程任务不存在：{task_id}")
            thread = managed.thread
        if thread is not None:
            thread.join(timeout_seconds)
        return self.get(task_id)

    def shutdown(self, timeout_seconds: float = 5.0) -> None:
        """停止当前编排线程后关闭其持有的单篇分发资源。"""

        task_id = self.active_task_id
        if task_id:
            self.stop(task_id)
            try:
                self.wait(task_id, timeout_seconds=max(0.0, float(timeout_seconds)))
            except Exception:
                pass
        callback = self._shutdown_callback
        if callable(callback):
            try:
                callback()
            finally:
                self._shutdown_callback = None

    def _run(self, managed: _ManagedMainFlow) -> None:
        managed.state.start()
        try:
            if self._runner is None:
                # 阶段一只验证生命周期，真实主页扫描器在后续阶段注入。
                managed.state.set_action("等待主页扫描模块接入")
                if managed.context.cancel_token.is_set():
                    managed.state.cancel()
                else:
                    managed.state.complete("主流程骨架已运行；主页扫描模块待接入")
            else:
                self._runner(managed.context, managed.command)
                if managed.context.cancel_token.is_set():
                    managed.state.cancel()
                elif managed.state.snapshot()["status"] not in {"failed", "cancelled"}:
                    managed.state.complete()
        except Exception as exc:
            message = f"主流程异常：{type(exc).__name__}: {exc}"
            managed.state.fail(message)
            self._write_error(message, exc)
        finally:
            managed.finished_at = self._now()
            with self._lock:
                if self._active_task_id == managed.context.task_id:
                    self._active_task_id = None

    def _snapshot(self, managed: _ManagedMainFlow) -> MainFlowSnapshot:
        state = managed.state.snapshot()
        return MainFlowSnapshot(
            task_id=managed.context.task_id,
            status=str(state["status"]),
            message=str(state.get("message") or ""),
            runtime_state=state,
            started_at=managed.context.started_at.isoformat(timespec="seconds"),
            finished_at=(
                managed.finished_at.isoformat(timespec="seconds")
                if managed.finished_at is not None
                else None
            ),
        )

    def _write_error(self, message: str, exception: BaseException) -> None:
        method = getattr(self._runtime_logger, "write_error", None)
        if not callable(method):
            return
        try:
            method(
                message,
                source="main-flow",
                channel="main_flow",
                phase="error",
                exception=exception,
                summary=True,
            )
        except Exception:
            return

    def _handle_context_event(self, state: MainFlowState, event: Any) -> bool:
        changed = state.handle_event(event)
        if isinstance(event, dict):
            self._write_context_event_summary(event)
        return changed

    def _write_context_event_summary(self, event: dict[str, Any]) -> None:
        event_type = str(event.get("type") or event.get("eventType") or "").strip().lower()
        if event_type not in {"article", "article_log"}:
            return
        article_task_id = str(event.get("articleTaskId") or event.get("article_task_id") or "")
        phase = str(event.get("phase") or "finished").strip() or "finished"
        status = str(event.get("status") or "running").strip().lower()
        message = str(event.get("message") or "").strip()
        article_title = str(event.get("articleTitle") or event.get("article_title") or "").strip()
        task_index = _event_task_index(event)
        dedupe_key = "|".join(
            (
                article_task_id,
                str(task_index),
                phase,
                status,
                message,
                article_title,
            )
        )
        with self._lock:
            if dedupe_key in self._summary_event_keys:
                return
            self._summary_event_keys.add(dedupe_key)

        self._write_summary(
            _article_log_level(status),
            _article_log_message(
                task_index=task_index,
                phase=phase,
                status=status,
                message=message,
                article_title=article_title,
            ),
            source="article-task",
            channel="article_task",
            phase=phase,
            task_index=task_index,
            article_task_id=article_task_id,
            article_title=article_title,
        )

    def _write_summary(
        self,
        level: str,
        message: str,
        *,
        source: str,
        channel: str,
        phase: str = "",
        task_index: int | str | None = None,
        article_task_id: str | None = None,
        article_title: str | None = None,
    ) -> None:
        method = getattr(self._runtime_logger, "write_summary", None)
        if not callable(method):
            return
        try:
            method(
                level,
                message,
                source=source,
                channel=channel,
                phase=phase,
                task_index=task_index,
                article_task_id=article_task_id,
                article_title=article_title,
            )
        except Exception:
            return

    @staticmethod
    def _start_message(command: MainFlowCommand) -> str:
        content = ["文章详情"]
        if command.collect_comments:
            content.append("评论采集")
        if command.archive_offline:
            offline_label = "离线归档 beta" if command.offline_archive_mode == "beta" else "离线归档"
            content.append(offline_label)
        count_label = "全部符合条件文章" if command.target_count == 0 else f"{command.target_count} 篇"
        return f"主流程启动：目标 {count_label}，开启{'、'.join(content)}"


def _event_task_index(event: dict[str, Any]) -> int | str:
    value = event.get("articleIndex", event.get("article_index", ""))
    if value in {None, ""}:
        value = event.get("cardIndex", event.get("card_index", ""))
    if value in {None, ""}:
        return ""
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return str(value)


def _article_log_level(status: str) -> str:
    if status in {"failed", "error"}:
        return "ERROR"
    if status in {"success", "completed"}:
        return "SUCCESS"
    if status in {"skipped", "skipped_collected", "skipped-collected", "cancelled", "canceled"}:
        return "WARN"
    return "INFO"


def _article_log_message(
    *,
    task_index: int | str,
    phase: str,
    status: str,
    message: str,
    article_title: str,
) -> str:
    prefix = f"任务 {task_index}" if task_index not in {"", 0} else "单篇任务"
    title = article_title or "未识别标题"
    if phase == "precheck":
        return f"{prefix} 接收任务：{title}"
    if phase == "foreground":
        if status in {"success", "completed"}:
            return f"{prefix} 返回主流程：{message or '文章标签已关闭'}"
        return f"{prefix} 前台处理：{message or _status_label(status)}"
    if phase == "detail":
        return f"{prefix} 文章详情：{message or _status_label(status)}"
    if phase == "comments":
        return f"{prefix} 评论信息：{message or _status_label(status)}"
    if phase == "offline":
        return f"{prefix} 离线归档：{message or _status_label(status)}"
    if phase == "finished":
        return f"{prefix} 最终结果：{message or _status_label(status)}"
    return f"{prefix} {phase}：{message or _status_label(status)}"


def _status_label(status: str) -> str:
    return {
        "running": "正在处理",
        "pending": "等待处理",
        "success": "成功",
        "completed": "成功",
        "failed": "失败",
        "error": "失败",
        "skipped_collected": "已采集，跳过",
        "skipped-collected": "已采集，跳过",
        "cancelled": "已取消",
        "canceled": "已取消",
    }.get(status, status or "状态未知")

