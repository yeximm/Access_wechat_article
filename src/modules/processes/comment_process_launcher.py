from __future__ import annotations

import multiprocessing
from typing import Any, Mapping

from src.modules.processes.process_channel import ProcessChannel
from src.modules.processes.process_launcher import LaunchedProcess, notify_process_started


class MultiprocessingCommentProcessLauncher:
    """为评论采集创建独立 spawn 子进程。"""

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
            target=_run_comment_child,
            args=(child_connection, task_id, attempt_id, dict(payload)),
            name=f"awa-comment-{attempt_id}",
            daemon=False,
        )
        try:
            process.start()
        except Exception:
            parent_connection.close()
            child_connection.close()
            raise
        pid = notify_process_started(process, label="comment")
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


def _run_comment_child(
    connection: Any,
    task_id: str,
    attempt_id: str,
    payload: dict[str, Any],
) -> None:
    """保持顶层函数，确保 Windows spawn 可以导入。"""
    try:
        from src.modules.processes.comment_collect_process import run_comment_collect_process

        run_comment_collect_process(
            connection=connection,
            task_id=task_id,
            attempt_id=attempt_id,
            payload=payload,
        )
    finally:
        connection.close()


__all__ = ["MultiprocessingCommentProcessLauncher"]
