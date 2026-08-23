from __future__ import annotations

import multiprocessing
from typing import Any, Mapping

from src.modules.processes.process_channel import ProcessChannel
from src.modules.processes.process_launcher import LaunchedProcess, notify_process_started


class MultiprocessingOfflineCacheProcessLauncher:
    """为每篇文章创建一个独立的 Playwright spawn 子进程。"""

    def __init__(self, context: Any | None = None) -> None:
        self._context = context or multiprocessing.get_context("spawn")

    def launch(
        self,
        *,
        task_id: str,
        attempt_id: str,
        payload: Mapping[str, Any],
    ) -> LaunchedProcess:
        parent_connection, child_connection = self._context.Pipe(duplex=True)
        process = self._context.Process(
            target=_run_offline_cache_child,
            args=(child_connection, task_id, attempt_id, dict(payload)),
            name=f"awa-offline-cache-{attempt_id}",
            daemon=False,
        )
        try:
            process.start()
        except Exception:
            parent_connection.close()
            child_connection.close()
            raise
        pid = notify_process_started(process, label="offline-cache")
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


def _run_offline_cache_child(
    connection: Any,
    task_id: str,
    attempt_id: str,
    payload: dict[str, Any],
) -> None:
    try:
        from src.modules.processes.offline_cache_process import run_offline_cache_process

        run_offline_cache_process(
            connection=connection,
            task_id=task_id,
            attempt_id=attempt_id,
            payload=payload,
        )
    finally:
        connection.close()


__all__ = ["MultiprocessingOfflineCacheProcessLauncher"]
