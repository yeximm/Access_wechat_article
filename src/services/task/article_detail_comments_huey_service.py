from __future__ import annotations

from datetime import datetime
from pathlib import Path
import time
from typing import Any, Callable, Mapping
from uuid import uuid4

from src.services.capture.comment_process_control_service import (
    CommentProcessControlService,
    CommentProcessError,
)
from src.services.capture.collected_article_lookup_service import (
    CollectedArticleLookupService,
)
from src.services.runtime.database_write_coordinator import DatabaseWriteCoordinator
from src.services.task.initial_content_storage_huey_service import (
    InitialContentStorageHueyService,
    InitialContentStorageTaskOptions,
    _saved_article_from_payload,
)


class ArticleDetailCommentsHueyService(InitialContentStorageHueyService):
    """单篇评论存储任务：先完成文章详情存储，再启动评论采集子进程。"""

    def __init__(
        self,
        *,
        temp_root: str | Path,
        config: Any,
        window_factory: Any,
        capture_factory: Any,
        database_path: str | Path,
        html_save: Any | None = None,
        runner: Callable[..., dict[str, Any]] | None = None,
        lookup_service: CollectedArticleLookupService | None = None,
        comment_process_control: Any | None = None,
        write_coordinator: DatabaseWriteCoordinator | None = None,
        session_id: str | None = None,
        job_id_factory: Callable[[], str] | None = None,
        now: Callable[[], datetime] = datetime.now,
        worker_count: int = 1,
    ) -> None:
        super().__init__(
            temp_root=temp_root,
            config=config,
            window_factory=window_factory,
            capture_factory=capture_factory,
            database_path=database_path,
            html_save=html_save,
            runner=runner,
            lookup_service=lookup_service,
            session_id=session_id,
            job_id_factory=job_id_factory,
            now=now,
            action="article-detail-comments",
            title="详情评论结果",
            flow_label="单篇评论存储测试",
            job_prefix="detail-comments",
            queue_name="article-detail-comments",
            task_name="ArticleDetailCommentsTask",
            wait_message_with_card="已读取首篇文章卡片，正在等待Huey执行单篇评论存储任务...",
            wait_message_without_card="正在等待Huey执行单篇评论存储任务...",
            extra_public_options={"storeCommentInfo": True},
            worker_count=worker_count,
            post_phase="comments",
        )
        self._comment_process_control = comment_process_control or CommentProcessControlService()
        self._write_coordinator = write_coordinator or DatabaseWriteCoordinator()

    def start(
        self,
        *,
        card_index: int = 1,
        account_name: str | None = None,
        card: dict[str, Any] | None = None,
        skip_collected_records: bool = False,
        store_article_detail: bool = True,
        store_comment_info: bool = True,
    ) -> dict[str, Any]:
        # 文章详情和评论信息是本测试的固定步骤；前端锁定只是 UI 保护，后端仍强制开启。
        initial = super().start(
            card_index=card_index,
            account_name=account_name,
            card=card,
            skip_collected_records=skip_collected_records,
            store_article_detail=True,
        )
        comment_item = {"label": "存储评论信息", "value": "开启（锁定）"}
        items = list(initial.get("items") or [])
        if not any(item.get("label") == "存储评论信息" for item in items if isinstance(item, Mapping)):
            insert_at = max(0, len(items) - 1)
            items.insert(insert_at, comment_item)
            initial["items"] = items
            with self._lock:
                current = self._jobs.get(str(initial.get("jobId")), {})
                self._jobs[str(initial.get("jobId"))] = {**current, "items": items}
        initial["options"] = {
            **dict(initial.get("options") or {}),
            "storeArticleDetail": True,
            "storeCommentInfo": True,
        }
        return initial

    def _run_post_task(
        self,
        *,
        job_id: str,
        article: Mapping[str, Any],
        update: Callable[[dict[str, Any]], None],
    ) -> dict[str, Any]:
        started_at = time.monotonic()
        saved = _saved_article_from_payload(article)
        records = [dict(item) for item in article.get("records", []) if isinstance(item, Mapping)]
        account_name = str(article.get("accountName") or "").strip()
        capture_type = str(article.get("captureType") or "none")
        items: list[dict[str, Any]] = []
        result = self._run_comment_process(
            job_id=job_id,
            saved=saved,
            context=None,
            items=items,
            update=update,
            base_result={
                "phase": "comments",
                "records": records,
                "accountName": account_name,
                "captureType": capture_type,
            },
            records=records,
            account_name=account_name,
            capture_type=capture_type,
            started_at=started_at,
        )
        status = str(result.get("status") or "failed").strip().lower()
        skipped = status == "skipped"
        success = status in {"success", "completed"}
        return {
            **dict(result),
            "ok": success or skipped,
            "status": "success" if success else ("skipped" if skipped else "failed"),
            "phase": "comments",
            "commentsStatus": "success" if success else ("skipped" if skipped else "failed"),
            "message": _final_message(result, ok=success, skipped=skipped),
            "totalSeconds": round(time.monotonic() - started_at, 3),
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
        context: Any,
        saved: Any,
    ) -> dict[str, Any]:
        comment_result = self._run_comment_process(
            job_id=job_id,
            saved=saved,
            context=context,
            items=items,
            update=update,
            base_result=base_result,
            records=records,
            account_name=account_name,
            capture_type=capture_type,
            started_at=started_at,
        )
        status = str(comment_result.get("status") or "")
        ok = status == "success"
        skipped = status == "skipped"
        total_seconds = time.monotonic() - started_at
        items.append(
            {
                "label": "总耗时",
                "value": _seconds(total_seconds),
                "cells": [
                    _cell(
                        "结果",
                        "文章详情与评论信息已保存"
                        if ok
                        else ("文章详情已保存，评论信息跳过" if skipped else "文章详情已保存，评论信息失败"),
                    )
                ],
            }
        )
        return {
            **base_result,
            "ok": ok,
            "status": "completed" if ok else ("skipped" if skipped else "failed"),
            "message": _final_message(comment_result, ok=ok, skipped=skipped),
            "tone": "success" if ok else ("warning" if skipped else "error"),
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
            "resourceManifest": list(comment_result.get("resource_manifest") or saved.resource_manifest.to_json_values()),
            "htmlCommentCount": _safe_int(comment_result.get("html_comment_count")),
            "commentCount": _safe_int(comment_result.get("comment_count")),
            "replyCount": _safe_int(comment_result.get("reply_count")),
            "commentPageCount": _safe_int(comment_result.get("page_count")),
            "commentPath": str(comment_result.get("comment_path") or ""),
            "commentAssetCount": _safe_int(comment_result.get("asset_count")),
            "commentAssetDir": str(comment_result.get("asset_dir") or ""),
        }

    def _run_comment_process(
        self,
        *,
        job_id: str,
        saved: Any,
        context: Any,
        items: list[dict[str, Any]],
        update: Callable[[dict[str, Any]], None],
        base_result: dict[str, Any],
        records: list[dict[str, Any]],
        account_name: str,
        capture_type: str,
        started_at: float,
    ) -> dict[str, Any]:
        task_id = f"{job_id}-comments"
        attempt_id = f"{task_id}-attempt-001"
        payload = {
            "proxy_lease_id": f"{task_id}-comment",
            "db_path": str(self._database_path),
            "storage_root": str(self._config.storage.article_storage_root),
            "temp_dir": str(Path(self._config.storage.temp_dir) / task_id),
            "started_at": self._now().isoformat(),
            "article_id": int(saved.article_id),
            "account_id": int(saved.account_id),
            "history_id": int(getattr(saved, "history_id", 0) or 0),
            "archive_dir": str(saved.archive_dir),
            "article_directory": str(saved.article_directory),
            "html_source": str(saved.html_source or ""),
            "resource_manifest": list(saved.resource_manifest.to_json_values()),
            "database_write_coordinator": self._write_coordinator,
            "timeout_seconds": float(self._config.comment.request_timeout_seconds),
            "page_interval_seconds": float(self._config.comment.page_interval_seconds),
            "max_pages": int(self._config.comment.max_pages),
        }

        attempt = None
        try:
            step_started = time.monotonic()
            attempt = self._comment_process_control.start(
                task_id=task_id,
                attempt_id=attempt_id,
                payload=payload,
            )
            self._register_active_attempt(job_id, attempt)
            items.append(
                {
                    "label": "启动评论采集子进程",
                    "value": _seconds(time.monotonic() - step_started),
                    "cells": [
                        _cell("结果", "已创建评论采集子进程"),
                        _cell("评论文件", "comments/final.json"),
                    ],
                }
            )
            update(
                {
                    **base_result,
                    "ok": False,
                    "status": "running",
                    "message": "评论采集子进程已启动，正在等待 READY...",
                    "tone": "info",
                    "items": list(items),
                    "records": records,
                    "accountName": account_name,
                    "captureType": capture_type,
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
            )

            step_started = time.monotonic()
            ready_payload = attempt.wait_ready(timeout_seconds=5.0)
            items.append(
                {
                    "label": "评论子进程 READY",
                    "value": _seconds(time.monotonic() - step_started),
                    "cells": [
                        _cell("结果", "评论子进程已就绪"),
                        _cell("文章ID", ready_payload.get("article_id")),
                        _cell("归档目录", ready_payload.get("archive_dir")),
                    ],
                }
            )

            def on_progress(event: dict[str, Any]) -> None:
                if event.get("internal"):
                    return
                items.append(_event_item(event))
                update(
                    {
                        **base_result,
                        "ok": False,
                        "status": "running",
                        "message": str(event.get("result") or "评论子进程正在执行..."),
                        "tone": "info",
                        "items": list(items),
                        "records": records,
                        "accountName": account_name,
                        "captureType": capture_type,
                        "totalSeconds": round(time.monotonic() - started_at, 3),
                    }
                )

            step_started = time.monotonic()
            comment_result = attempt.wait_result(
                timeout_seconds=_comment_result_timeout_seconds(self._config),
                on_progress=on_progress,
            )
            items.append(
                {
                    "label": "评论子进程返回结果",
                    "value": _seconds(time.monotonic() - step_started),
                    "cells": [
                        _cell("结果", comment_result.get("status")),
                        _cell("HTML 评论数", comment_result.get("html_comment_count")),
                        _cell("评论数", comment_result.get("comment_count")),
                        _cell("回复数", comment_result.get("reply_count")),
                        _cell("页数", comment_result.get("page_count")),
                        _cell("停止原因", _format_stop_reason(comment_result.get("stop_reason"))),
                    ],
                }
            )
            return dict(comment_result)
        except Exception as exc:
            if attempt is not None:
                try:
                    attempt.cancel()
                except Exception:
                    pass
            comment_result = exc.result if isinstance(exc, CommentProcessError) else None
            items.append(
                {
                    "label": "评论采集失败",
                    "value": str(exc),
                    "cells": [
                        _cell("失败原因", str(exc)),
                        _cell("子进程状态", _safe_mapping_value(comment_result, "status")),
                    ],
                }
            )
            if isinstance(comment_result, dict):
                return dict(comment_result)
            return {
                "status": "failed",
                "message": f"评论采集失败：{exc}",
                "comment_count": 0,
                "reply_count": 0,
                "page_count": 0,
                "html_comment_count": 0,
            }
        finally:
            self._unregister_active_attempt(job_id, attempt)


def _comment_result_timeout_seconds(config: Any) -> float:
    timeout = max(1.0, float(config.comment.request_timeout_seconds))
    pages = max(1, int(config.comment.max_pages))
    interval = max(0.0, float(config.comment.page_interval_seconds))
    return max(30.0, timeout * (pages + 2) + interval * pages + 15.0)


def _event_item(event: Mapping[str, Any]) -> dict[str, Any]:
    details = event.get("details")
    cells = [_cell("结果", event.get("result"))]
    if isinstance(details, Mapping):
        for key, value in details.items():
            if str(key) == "stop_reason":
                value = _format_stop_reason(value)
            cells.append(_cell(_detail_label(str(key)), value))
    return {
        "label": str(event.get("name") or "评论子进程步骤"),
        "value": _seconds(float(event.get("elapsed_seconds") or 0)),
        "cells": cells,
    }


def _final_message(comment_result: Mapping[str, Any], *, ok: bool, skipped: bool) -> str:
    if ok:
        if str(comment_result.get("stop_reason") or "").startswith("html_comment_count_"):
            return "评论信息存储完成：HTML 未发现可采集评论，未发起评论接口请求。"
        return (
            "评论信息存储完成："
            f"评论 {comment_result.get('comment_count', 0)} 条，"
            f"回复 {comment_result.get('reply_count', 0)} 条。"
        )
    if skipped:
        return str(comment_result.get("message") or "评论参数不足，已跳过评论采集。")
    return str(comment_result.get("message") or "评论采集失败。")


def _detail_label(key: str) -> str:
    return {
        "required_parameter_count": "必要参数数",
        "html_comment_count": "HTML 评论数",
        "comment_count": "评论数",
        "reply_count": "回复数",
        "page_count": "页数",
        "stop_reason": "停止原因",
        "comment_path": "评论文件",
        "history_id": "历史ID",
    }.get(key, key)


def _format_stop_reason(value: Any) -> str:
    reason = str(value or "").strip()
    if not reason:
        return "无"
    mapping = {
        "continue_flag_false": "接口返回无下一页，分页正常结束",
        "no_new_comments": "本页没有新增评论，停止分页",
        "max_pages_reached": "达到最大页数限制，可能未完全采完",
        "html_comment_count_zero": "HTML 评论数为 0，未发起评论请求",
        "html_comment_count_missing": "未解析到 HTML 评论数，未发起评论请求",
        "reply_total_reached": "回复数已达到接口声明总数",
        "reply_parameters_missing": "回复分页参数不足，停止拉取该楼层回复",
        "no_new_replies": "本页没有新增回复，停止拉取该楼层回复",
        "completed": "已完成",
    }
    return mapping.get(reason, f"未识别停止原因：{reason}")


def _safe_mapping_value(value: Any, key: str) -> str:
    return str(value.get(key) or "") if isinstance(value, Mapping) else ""


def _safe_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _cell(label: str, value: Any) -> dict[str, str]:
    return {"label": label, "value": _format_value(value)}


def _format_value(value: Any) -> str:
    if value is None or value == "":
        return "无"
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def _seconds(value: float) -> str:
    return f"{max(0.0, float(value)):.3f} 秒"


__all__ = ["ArticleDetailCommentsHueyService"]
