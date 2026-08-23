from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
import time
from typing import Any, Callable, Mapping

from src.domain.enums import TaskStatus
from src.modules.archive.resource_manifest_builder import ResourceManifestBuilder
from src.services.archive.offline_cache_job_service import WECHAT_SHORT_LINK_PATTERN
from src.services.archive.offline_cache_process_control_service import (
    OfflineCacheProcessControlService,
    OfflineCacheProcessError,
)
from src.services.archive.resource_commit_service import ResourceCommitService
from src.services.capture.collected_article_lookup_service import (
    CollectedArticleLookupService,
)
from src.services.runtime.database_write_coordinator import DatabaseWriteCoordinator
from src.services.task.initial_content_storage_huey_service import (
    InitialContentStorageHueyService,
    InitialContentStorageTaskOptions,
    _saved_article_from_payload,
)
from src.storage.repositories.article_repository import ArticleRepository
from src.storage.repositories.fetch_history_repository import (
    FetchHistoryRepository,
    FetchHistoryWrite,
)
from src.storage.sqlite.connection import sqlite_connection


class ArticleDetailOfflineCacheHueyService(InitialContentStorageHueyService):
    """单篇离线缓存任务：先保存文章详情，再启动 Playwright 离线缓存子进程。"""

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
        offline_cache_process_control: Any | None = None,
        resource_commit: ResourceCommitService | None = None,
        write_coordinator: DatabaseWriteCoordinator | None = None,
        browser_cache_dir: str | Path | None = None,
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
            action="article-detail-offline-cache",
            title="单篇离线缓存结果",
            flow_label="单篇离线缓存测试",
            job_prefix="detail-offline-cache",
            queue_name="article-detail-offline-cache",
            task_name="ArticleDetailOfflineCacheTask",
            wait_message_with_card="已读取首篇文章卡片，正在等待Huey执行单篇离线缓存任务...",
            wait_message_without_card="正在等待Huey执行单篇离线缓存任务...",
            extra_public_options={"archiveOfflineContent": True},
            worker_count=worker_count,
            post_phase="offline",
        )
        self._offline_cache_process_control = (
            offline_cache_process_control or OfflineCacheProcessControlService()
        )
        self._write_coordinator = write_coordinator or DatabaseWriteCoordinator()
        self._resource_commit = resource_commit or ResourceCommitService(
            write_coordinator=self._write_coordinator,
        )
        self._browser_cache_dir = Path(
            browser_cache_dir
            or Path(database_path).resolve().parents[2] / ".playwright-browsers"
            if len(Path(database_path).resolve().parents) >= 3
            else ".playwright-browsers"
        ).resolve()

    def start(
        self,
        *,
        card_index: int = 1,
        account_name: str | None = None,
        card: dict[str, Any] | None = None,
        skip_collected_records: bool = False,
        store_article_detail: bool = True,
        archive_offline_content: bool = True,
        stateful_offline_cache: bool = False,
    ) -> dict[str, Any]:
        # 离线缓存测试固定包含“保存文章详情”和“离线归档内容”两个步骤。
        initial = super().start(
            card_index=card_index,
            account_name=account_name,
            card=card,
            skip_collected_records=skip_collected_records,
            store_article_detail=True,
            stateful_offline_cache=stateful_offline_cache,
        )
        archive_item = {"label": "离线归档内容", "value": "开启（锁定）"}
        stateful_item = {
            "label": "带状态（bate）",
            "value": "开启" if stateful_offline_cache else "关闭",
        }
        items = list(initial.get("items") or [])
        if not any(item.get("label") == "带状态（bate）" for item in items if isinstance(item, Mapping)):
            insert_at = 2
            for index, item in enumerate(items):
                if isinstance(item, Mapping) and item.get("label") == "跳过已采集记录":
                    insert_at = index + 1
                    break
            items.insert(insert_at, stateful_item)
        if not any(item.get("label") == "离线归档内容" for item in items if isinstance(item, Mapping)):
            insert_at = max(0, len(items) - 1)
            items.insert(insert_at, archive_item)
        initial["items"] = items
        with self._lock:
            current = self._jobs.get(str(initial.get("jobId")), {})
            self._jobs[str(initial.get("jobId"))] = {**current, "items": items}
        initial["options"] = {
            **dict(initial.get("options") or {}),
            "storeArticleDetail": True,
            "archiveOfflineContent": True,
            "statefulOfflineCache": bool(stateful_offline_cache),
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
        result = self._run_offline_cache_process(
            job_id=job_id,
            saved=saved,
            items=[],
            update=update,
            base_result={
                "phase": "offline",
                "records": records,
                "accountName": account_name,
                "captureType": capture_type,
            },
            records=records,
            account_name=account_name,
            capture_type=capture_type,
            started_at=started_at,
            stateful_offline_cache=bool(article.get("statefulOfflineCache") or article.get("stateful_offline_cache")),
        )
        success = bool(result.get("ok")) and str(result.get("status") or "").lower() in {"success", "completed"}
        return {
            **dict(result),
            "ok": success,
            "status": "success" if success else "failed",
            "phase": "offline",
            "offlineStatus": "success" if success else "failed",
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
        offline_result = self._run_offline_cache_process(
            job_id=job_id,
            saved=saved,
            items=items,
            update=update,
            base_result=base_result,
            records=records,
            account_name=account_name,
            capture_type=capture_type,
            started_at=started_at,
            stateful_offline_cache=bool(options.stateful_offline_cache),
        )
        ok = bool(offline_result.get("ok"))
        total_seconds = time.monotonic() - started_at
        items.append(
            {
                "label": "总耗时",
                "value": _seconds(total_seconds),
                "cells": [
                    _cell("结果", "文章详情与离线归档内容已保存" if ok else "文章详情已保存，离线归档失败")
                ],
            }
        )
        return {
            **base_result,
            "ok": ok,
            "status": "completed" if ok else "failed",
            "message": str(
                offline_result.get("message")
                or ("单篇离线缓存完成。" if ok else "单篇离线缓存失败。")
            ),
            "tone": "success" if ok else "error",
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
            "resourceManifest": list(
                offline_result.get("resource_manifest")
                or saved.resource_manifest.to_json_values()
            ),
            "offlineIndexPath": str(offline_result.get("index_html_path") or ""),
            "offlineResourceCount": _safe_int(offline_result.get("resource_count")),
            "offlineAssetsDir": str(offline_result.get("assets_dir") or ""),
            "offlineWarning": str(offline_result.get("warning") or ""),
        }

    def _run_offline_cache_process(
        self,
        *,
        job_id: str,
        saved: Any,
        items: list[dict[str, Any]],
        update: Callable[[dict[str, Any]], None],
        base_result: dict[str, Any],
        records: list[dict[str, Any]],
        account_name: str,
        capture_type: str,
        started_at: float,
        stateful_offline_cache: bool,
    ) -> dict[str, Any]:
        article_directory = Path(saved.article_directory).resolve()
        article_short_link = self._read_article_short_link(int(saved.article_id))
        if not article_short_link:
            message = "数据库中缺少文章短链，已终止离线缓存。"
            items.append({"label": "离线缓存准备", "value": message})
            return {"ok": False, "status": "failed", "message": message}

        task_id = f"{job_id}-offline-cache"
        attempt_id = f"{task_id}-attempt-001"
        attempt_root = self._attempt_root(job_id, int(saved.article_id), article_directory)
        stage_dir = attempt_root / "stage"
        backup_dir = attempt_root / "backup"
        payload = {
            # 这两个是离线缓存子进程的业务输入；下面字段只服务于现有子进程运行和暂存目录。
            "article_directory": str(article_directory),
            "article_short_link": article_short_link,
            "article_id": int(saved.article_id),
            "article_title": _record_title(records),
            "article_link": article_short_link,
            "stage_dir": str(stage_dir),
            "browser_cache_dir": str(self._browser_cache_dir),
            "max_scroll_seconds": float(self._config.offline_cache.max_scroll_seconds),
            "resource_timeout_seconds": float(
                self._config.offline_cache.resource_timeout_seconds
            ),
            "stateful_offline_cache": bool(stateful_offline_cache),
            "request_json_path": str(article_directory / "origin" / "request.json"),
        }
        attempt = None
        started_time = self._now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            step_started = time.monotonic()
            attempt = self._offline_cache_process_control.start(
                task_id=task_id,
                attempt_id=attempt_id,
                payload=payload,
            )
            self._register_active_attempt(job_id, attempt)
            items.append(
                {
                    "label": "启动离线缓存子进程",
                    "value": _seconds(time.monotonic() - step_started),
                    "cells": [
                        _cell("结果", "已创建 Playwright 离线缓存子进程"),
                        _cell("文章短链", article_short_link),
                        _cell("带状态", "开启（bate）" if stateful_offline_cache else "关闭"),
                    ],
                }
            )
            update(
                {
                    **base_result,
                    "ok": False,
                    "status": "running",
                    "message": "离线缓存子进程已启动，正在等待 READY...",
                    "tone": "info",
                    "items": list(items),
                    "records": records,
                    "accountName": account_name,
                    "captureType": capture_type,
                    "totalSeconds": round(time.monotonic() - started_at, 3),
                }
            )

            step_started = time.monotonic()
            attempt.wait_ready(timeout_seconds=10.0)
            items.append(
                {
                    "label": "离线缓存子进程 READY",
                    "value": _seconds(time.monotonic() - step_started),
                    "cells": [_cell("结果", "Playwright 子进程已就绪")],
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
                        "message": str(event.get("name") or "离线缓存子进程正在执行..."),
                        "tone": "info",
                        "items": list(items),
                        "records": records,
                        "accountName": account_name,
                        "captureType": capture_type,
                        "totalSeconds": round(time.monotonic() - started_at, 3),
                    }
                )

            step_started = time.monotonic()
            raw_result = attempt.wait_result(
                timeout_seconds=max(
                    30.0,
                    float(self._config.offline_cache.max_scroll_seconds) + 60.0,
                ),
                on_progress=on_progress,
            )
            items.append(
                {
                    "label": "离线缓存子进程返回结果",
                    "value": _seconds(time.monotonic() - step_started),
                    "cells": [
                        _cell("结果", raw_result.get("message")),
                        _cell("访问模式", _navigation_mode_label(raw_result.get("navigation_mode"))),
                        _cell("资源数量", raw_result.get("resource_count")),
                        _cell("警告", raw_result.get("warning")),
                    ],
                }
            )
            if not bool(raw_result.get("ok")):
                raise OfflineCacheProcessError(
                    str(raw_result.get("message") or "离线缓存失败"),
                    result=raw_result,
                )
            if not (stage_dir / "index.html").is_file() or not (stage_dir / "assets").is_dir():
                raise RuntimeError("子进程未生成完整的 index.html 和 assets")

            finished_time = self._now().strftime("%Y-%m-%d %H:%M:%S")
            elapsed = time.monotonic() - started_at
            resource_manifest = ResourceManifestBuilder().build(
                article_directory,
                planned_paths=("index.html", "assets"),
            )

            def database_operation(connection):
                ArticleRepository(connection).update_resource_manifest(
                    int(saved.article_id),
                    resource_manifest,
                    collected_time=finished_time,
                )
                return FetchHistoryRepository(connection).append(
                    FetchHistoryWrite(
                        article_id=int(saved.article_id),
                        account_id=int(saved.account_id),
                        target_account_name=account_name,
                        target_title=_record_title(records),
                        target_link=article_short_link,
                        task_type="offline_cache",
                        resource_manifest=resource_manifest,
                        status=TaskStatus.SUCCESS,
                        started_time=started_time,
                        finished_time=finished_time,
                        duration_seconds=elapsed,
                        output_dir=str(saved.archive_dir),
                    )
                )

            history_id = self._resource_commit.commit(
                database_path=self._database_path,
                stage_root=stage_dir,
                target_root=article_directory,
                backup_root=backup_dir,
                resource_paths=("index.html", "assets"),
                database_operation=database_operation,
            )
            return {
                **dict(raw_result),
                "ok": True,
                "status": "success",
                "message": str(raw_result.get("message") or "离线缓存完成"),
                "index_html_path": str(article_directory / "index.html"),
                "assets_dir": str(article_directory / "assets"),
                "history_id": history_id,
                "resource_manifest": resource_manifest.to_json_values(),
            }
        except Exception as exc:
            if attempt is not None:
                try:
                    attempt.cancel()
                except Exception:
                    pass
            raw_result = exc.result if isinstance(exc, OfflineCacheProcessError) else None
            message = str(raw_result.get("message") or exc) if isinstance(raw_result, dict) else str(exc)
            items.append(
                {
                    "label": "离线缓存失败",
                    "value": message,
                    "cells": [_cell("失败原因", message)],
                }
            )
            self._append_failure_history(
                article_id=int(saved.article_id),
                account_id=int(saved.account_id),
                account_name=account_name,
                title=_record_title(records),
                article_short_link=article_short_link,
                archive_dir=str(saved.archive_dir),
                started_time=started_time,
                elapsed_seconds=time.monotonic() - started_at,
                message=message,
            )
            if isinstance(raw_result, dict):
                return {**raw_result, "ok": False, "status": "failed", "message": message}
            return {"ok": False, "status": "failed", "message": f"离线缓存失败：{message}"}
        finally:
            self._unregister_active_attempt(job_id, attempt)
            if attempt_root.exists():
                shutil.rmtree(attempt_root, ignore_errors=True)

    def _read_article_short_link(self, article_id: int) -> str:
        with sqlite_connection(self._database_path, write=False) as connection:
            record = ArticleRepository(connection).get_by_id(article_id)
        if record is None:
            raise LookupError(f"文章索引不存在：{article_id}")
        value = str(record.article_link or "").strip()
        if WECHAT_SHORT_LINK_PATTERN.fullmatch(value) is None:
            return ""
        return value

    def _public_options_payload(
        self,
        options: InitialContentStorageTaskOptions,
    ) -> dict[str, Any]:
        return {
            **super()._public_options_payload(options),
            "archiveOfflineContent": True,
            "statefulOfflineCache": bool(options.stateful_offline_cache),
        }

    def _append_failure_history(
        self,
        *,
        article_id: int,
        account_id: int,
        account_name: str,
        title: str,
        article_short_link: str,
        archive_dir: str,
        started_time: str,
        elapsed_seconds: float,
        message: str,
    ) -> None:
        try:
            finished_time = self._now().strftime("%Y-%m-%d %H:%M:%S")
            with self._write_coordinator.hold():
                with sqlite_connection(self._database_path) as connection:
                    FetchHistoryRepository(connection).append(
                        FetchHistoryWrite.failed(
                            article_id=article_id,
                            account_id=account_id,
                            target_account_name=account_name,
                            target_title=title,
                            target_link=article_short_link,
                            task_type="offline_cache",
                            started_time=started_time,
                            finished_time=finished_time,
                            duration_seconds=max(0.0, elapsed_seconds),
                            error_stage="offline_cache",
                            error_message=message,
                        )
                    )
        except Exception:
            # 诊断弹窗已展示失败原因；历史写入失败不能掩盖原始离线缓存错误。
            pass

    def _attempt_root(self, job_id: str, article_id: int, article_directory: Path) -> Path:
        configured = self._temp_root / "offline-cache" / job_id / str(article_id)
        try:
            if configured.anchor.casefold() == article_directory.anchor.casefold():
                return configured
        except OSError:
            pass
        return article_directory.parent / ".awa-offline-cache" / job_id / str(article_id)


def _record_title(records: list[dict[str, Any]]) -> str:
    if records:
        record = records[0]
        return str(record.get("rawTitle") or record.get("title") or "").strip()
    return ""


def _event_item(event: Mapping[str, Any]) -> dict[str, Any]:
    cells = [_cell("状态", event.get("status"))]
    return {
        "label": str(event.get("name") or "离线缓存子进程步骤"),
        "value": _seconds(float(event.get("elapsed_seconds") or 0)),
        "cells": cells,
    }


def _safe_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _cell(label: str, value: Any) -> dict[str, str]:
    return {"label": label, "value": _format_value(value)}


def _navigation_mode_label(value: Any) -> str:
    return "有状态" if str(value or "").strip().lower() == "stateful" else "无状态"


def _format_value(value: Any) -> str:
    if value is None or value == "":
        return "无"
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def _seconds(value: float) -> str:
    return f"{max(0.0, float(value)):.3f} 秒"


__all__ = ["ArticleDetailOfflineCacheHueyService"]
