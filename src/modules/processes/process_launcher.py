from __future__ import annotations

from dataclasses import dataclass
import multiprocessing
from typing import Any

from src.modules.processes.process_channel import ProcessChannel


_process_observer: Any | None = None


@dataclass(slots=True)
class LaunchedProcess:
    process: Any
    channel: Any
    pid: int | None = None


def set_process_observer(observer: Any | None) -> None:
    """设置进程启动观察器，用于把新建子进程 PID 交给运行状态监控。"""
    global _process_observer
    _process_observer = observer


def notify_process_started(process: Any, *, label: str) -> int | None:
    """在 process.start() 后登记 PID；登记失败不能影响业务子进程。"""
    pid = getattr(process, "pid", None)
    observer = _process_observer
    register = getattr(observer, "register_process", None)
    if callable(register):
        try:
            register(pid, label=label)
        except Exception:
            pass
    try:
        return int(pid) if pid is not None else None
    except (TypeError, ValueError):
        return None


class MultiprocessingProcessLauncher:
    """为每次采集尝试创建全新的 spawn 子进程和双向 Pipe。"""

    def __init__(self, context: Any | None = None) -> None:
        self._context = context or multiprocessing.get_context("spawn")

    def launch(
        self,
        *,
        task_id: str,
        attempt_id: str,
        proxy_lease_id: str,
    ) -> LaunchedProcess:
        parent_connection, child_connection = self._context.Pipe(duplex=True)
        process = self._context.Process(
            target=_run_mitm_child,
            args=(child_connection, task_id, attempt_id, proxy_lease_id),
            name=f"awa-mitm-{attempt_id}",
            daemon=False,
        )
        try:
            process.start()
        except Exception:
            parent_connection.close()
            child_connection.close()
            raise
        pid = notify_process_started(process, label="mitm")
        child_connection.close()
        return LaunchedProcess(
            process=process,
            channel=ProcessChannel(
                parent_connection,
                task_id=task_id,
                attempt_id=attempt_id,
            ),
            pid=pid,
        )


def _run_mitm_child(
    connection: Any,
    task_id: str,
    attempt_id: str,
    proxy_lease_id: str,
) -> None:
    """保持为顶层函数，确保 Windows spawn 可以导入。"""
    try:
        from src.modules.processes.mitm_capture_process import run_mitm_capture_process

        run_mitm_capture_process(
            connection=connection,
            task_id=task_id,
            attempt_id=attempt_id,
            expected_proxy_lease_id=proxy_lease_id,
        )
    finally:
        connection.close()
