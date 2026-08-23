from __future__ import annotations

from pathlib import Path
import time
from typing import Any, Callable, Mapping

from src.domain.enums import ProcessMessageType
from src.modules.archive.offline_navigation_state import load_offline_navigation_state
from src.modules.archive.offline_archiver import (
    OfflineArchiveRequest,
    OfflineArchiveResult,
    archive_offline_article,
)
from src.modules.processes.process_channel import ProcessChannel


ArchiveFunc = Callable[..., OfflineArchiveResult]


def run_offline_cache_process(
    *,
    connection: Any,
    task_id: str,
    attempt_id: str,
    payload: Mapping[str, Any],
    archive_func: ArchiveFunc = archive_offline_article,
) -> None:
    """单篇离线缓存子进程入口；只生成暂存文件并返回结果。"""
    channel = ProcessChannel(connection, task_id=task_id, attempt_id=attempt_id)
    started_at = time.monotonic()
    article_id = int(payload["article_id"])
    try:
        navigation_state = load_offline_navigation_state(
            enabled=bool(payload.get("stateful_offline_cache")),
            article_directory=payload.get("article_directory"),
            request_json_path=payload.get("request_json_path"),
        )
        request = OfflineArchiveRequest(
            article_id=article_id,
            article_title=str(payload.get("article_title") or ""),
            article_link=str(payload["article_link"]),
            stage_dir=Path(str(payload["stage_dir"])),
            browser_cache_dir=Path(str(payload["browser_cache_dir"])),
            max_scroll_seconds=float(payload.get("max_scroll_seconds", 30.0)),
            resource_timeout_seconds=float(payload.get("resource_timeout_seconds", 10.0)),
            navigation_url=navigation_state.url,
            navigation_mode=navigation_state.mode,
            navigation_user_agent=navigation_state.user_agent,
            navigation_headers=navigation_state.headers,
            navigation_cookies=navigation_state.cookies,
        )
        channel.send(
            ProcessMessageType.READY,
            {
                "article_id": article_id,
                "stage_dir": str(request.stage_dir),
                "navigation_mode": request.navigation_mode,
            },
        )
        result = archive_func(
            request,
            on_event=lambda event: channel.send(
                ProcessMessageType.PROGRESS,
                {"event": dict(event)},
            ),
        )
        result_payload = _result_payload(
            article_id=article_id,
            result=result,
            elapsed_seconds=time.monotonic() - started_at,
            navigation_mode=request.navigation_mode,
        )
        message_type = ProcessMessageType.RESULT if result.ok else ProcessMessageType.FAILED
        channel.send(message_type, {"offline_cache_result": result_payload})
    except Exception as exc:
        channel.send(
            ProcessMessageType.FAILED,
            {
                "offline_cache_result": {
                    "ok": False,
                    "article_id": article_id,
                    "message": f"离线缓存子进程失败：{type(exc).__name__}: {exc}",
                    "elapsed_seconds": round(time.monotonic() - started_at, 3),
                }
            },
        )


def _result_payload(
    *,
    article_id: int,
    result: OfflineArchiveResult,
    elapsed_seconds: float,
    navigation_mode: str,
) -> dict[str, Any]:
    return {
        "ok": bool(result.ok),
        "article_id": article_id,
        "stage_dir": str(result.stage_dir),
        "index_html_path": "" if result.index_html_path is None else str(result.index_html_path),
        "assets_dir": "" if result.assets_dir is None else str(result.assets_dir),
        "resource_count": int(result.resource_count),
        "download_bytes": int(result.download_bytes),
        "message": result.message,
        "warning": result.warning,
        "elapsed_seconds": round(max(0.0, elapsed_seconds), 3),
        "navigation_mode": navigation_mode,
    }


__all__ = ["run_offline_cache_process"]
