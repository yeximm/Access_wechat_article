from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from src.domain.enums import ProcessMessageType, ResourceType, TaskStatus
from src.domain.models import ResourceManifest, TaskContext
from src.modules.processes.process_channel import ProcessChannel
from src.services.capture.comment_collect_service import CommentCollectData, CommentCollectService
from src.services.capture.html_parse_save_service import ArticleSaveData


def run_comment_collect_process(
    *,
    connection: Any,
    task_id: str,
    attempt_id: str,
    payload: Mapping[str, Any],
) -> None:
    """评论采集子进程入口；只处理一篇文章的评论后退出。"""
    channel = ProcessChannel(connection, task_id=task_id, attempt_id=attempt_id)
    try:
        context = _context_from_payload(payload, task_id=task_id)
        article = _article_from_payload(payload, attempt_id=attempt_id)
        timeout_seconds = float(payload.get("timeout_seconds", 10.0))
        page_interval_seconds = float(payload.get("page_interval_seconds", 0.5))
        max_pages = int(payload.get("max_pages", 50))

        channel.send(
            ProcessMessageType.READY,
            {
                "article_id": article.article_id,
                "archive_dir": article.archive_dir,
                "comment_path": "comments/final.json",
            },
        )

        result = CommentCollectService(
            write_coordinator=payload.get("database_write_coordinator"),
        ).collect(
            context=context,
            article=article,
            timeout_seconds=timeout_seconds,
            page_interval_seconds=page_interval_seconds,
            max_pages=max_pages,
            on_event=lambda event: channel.send(ProcessMessageType.PROGRESS, {"event": event}),
        )
        payload = _result_payload(result_status=result.status, message=result.message, data=result.data)
        payload["duration_seconds"] = round(max(0.0, float(result.duration_seconds)), 3)
        payload["error_code"] = "" if result.error_code is None else result.error_code.value
        channel.send(ProcessMessageType.RESULT, {"comment_result": payload})
    except Exception as exc:
        channel.send(
            ProcessMessageType.FAILED,
            {
                "comment_result": {
                    "status": TaskStatus.FAILED.value,
                    "message": f"评论采集子进程失败：{type(exc).__name__}: {exc}",
                    "duration_seconds": 0,
                    "error_code": "comment_fetch_failed",
                }
            },
        )


def _context_from_payload(payload: Mapping[str, Any], *, task_id: str) -> TaskContext:
    return TaskContext(
        task_id=task_id,
        proxy_lease_id=str(payload.get("proxy_lease_id") or f"{task_id}-comment"),
        db_path=Path(str(payload["db_path"])),
        storage_root=Path(str(payload["storage_root"])),
        temp_dir=Path(str(payload["temp_dir"])),
        started_at=_parse_datetime(payload.get("started_at")),
    )


def _article_from_payload(payload: Mapping[str, Any], *, attempt_id: str) -> ArticleSaveData:
    article_directory = Path(str(payload["article_directory"]))
    resource_values = payload.get("resource_manifest") or []
    resource_types = []
    if isinstance(resource_values, list):
        for value in resource_values:
            try:
                resource_types.append(ResourceType(str(value)))
            except ValueError:
                continue
    return ArticleSaveData(
        article_id=int(payload["article_id"]),
        account_id=int(payload["account_id"]),
        history_id=int(payload.get("history_id") or 0),
        article_directory=article_directory,
        archive_dir=str(payload["archive_dir"]),
        detail_path=article_directory / "article_detail.json",
        resource_manifest=ResourceManifest.from_types(resource_types),
        html_source=str(payload.get("html_source") or ""),
        attempt_id=attempt_id,
    )


def _result_payload(
    *,
    result_status: TaskStatus,
    message: str,
    data: CommentCollectData | None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "status": result_status.value,
        "message": message,
    }
    if data is not None:
        payload.update(
            {
                "article_id": data.article_id,
                "history_id": data.history_id,
                "comment_path": "" if data.comment_path is None else str(data.comment_path),
                "comment_count": data.comment_count,
                "reply_count": data.reply_count,
                "page_count": data.page_count,
                "pagination_complete": data.pagination_complete,
                "stop_reason": data.stop_reason,
                "html_comment_count": data.html_comment_count,
                "resource_manifest": data.resource_manifest.to_json_values(),
                "asset_count": data.asset_count,
                "asset_dir": "" if data.asset_dir is None else str(data.asset_dir),
                "download_bytes": int(data.download_bytes),
            }
        )
    return payload


def _parse_datetime(value: Any):
    from datetime import datetime

    if isinstance(value, datetime):
        return value
    text = str(value or "").strip()
    if not text:
        return datetime.now()
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return datetime.now()


__all__ = ["run_comment_collect_process"]
