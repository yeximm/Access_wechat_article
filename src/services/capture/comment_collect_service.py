from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import shutil
import time
from typing import Any, Callable, Mapping, Protocol

from src.domain.enums import ErrorCode, ResourceType, TaskStatus
from src.domain.models import ResourceManifest, TaskContext
from src.domain.results import ServiceResult
from src.modules.archive.resource_manifest_builder import ResourceManifestBuilder
from src.modules.request.article_parser import extract_html_comment_count
from src.modules.request.comment_requester import (
    CommentFetchData,
    CommentFetchError,
    CommentParametersMissing,
    REQUIRED_IDENTITY_KEYS,
    WechatCommentRequester,
    extract_comment_identity,
    normalize_article_url,
)
from src.services.archive.resource_commit_service import ResourceCommitService
from src.services.capture.html_parse_save_service import ArticleSaveData
from src.services.runtime.database_write_coordinator import DatabaseWriteCoordinator
from src.storage.repositories.article_repository import ArticleRecord, ArticleRepository
from src.storage.repositories.fetch_history_repository import (
    FetchHistoryRepository,
    FetchHistoryWrite,
)
from src.storage.sqlite.connection import sqlite_connection


COMMENT_RESOURCE_PATH = Path("comments/final.json")
COMMENT_ASSETS_PATH = Path("comments/assets")


class CommentRequester(Protocol):
    def fetch(self, **kwargs: Any) -> CommentFetchData: ...


@dataclass(frozen=True, slots=True)
class CommentCollectData:
    article_id: int
    history_id: int
    comment_path: Path | None
    comment_count: int
    reply_count: int
    page_count: int
    pagination_complete: bool
    stop_reason: str
    resource_manifest: ResourceManifest
    asset_count: int = 0
    asset_dir: Path | None = None
    html_comment_count: int | None = None
    download_bytes: int = 0


class CommentCollectService:
    """读取单篇正文证据，抓取评论，并只原子替换评论资源。"""

    def __init__(
        self,
        *,
        requester: CommentRequester | None = None,
        manifest_builder: ResourceManifestBuilder | None = None,
        commit_service: ResourceCommitService | None = None,
        write_coordinator: DatabaseWriteCoordinator | None = None,
        now: Callable[[], datetime] = datetime.now,
        monotonic: Callable[[], float] = time.monotonic,
    ) -> None:
        self._requester = requester or WechatCommentRequester()
        self._manifest_builder = manifest_builder or ResourceManifestBuilder()
        self._write_coordinator = write_coordinator or DatabaseWriteCoordinator()
        self._commit_service = commit_service or ResourceCommitService(
            write_coordinator=self._write_coordinator,
        )
        self._now = now
        self._monotonic = monotonic

    def collect(
        self,
        *,
        context: TaskContext,
        article: ArticleSaveData,
        timeout_seconds: float,
        page_interval_seconds: float,
        max_pages: int,
        on_event: Callable[[dict[str, Any]], None] | None = None,
    ) -> ServiceResult[CommentCollectData]:
        started = self._now()
        started_monotonic = self._monotonic()
        record: ArticleRecord | None = None
        try:
            step_started = self._monotonic()
            record = self._load_article(context.db_path, article.article_id)
            html_text, article_url, headers = _load_request_inputs(article.article_directory)
            _emit_event(
                on_event,
                name="读取评论输入文件",
                started_at=step_started,
                result="已读取 origin/main.html 和 origin/request.json",
            )

            step_started = self._monotonic()
            html_comment_count = extract_html_comment_count(html_text)
            if html_comment_count is None or html_comment_count <= 0:
                stop_reason = (
                    "html_comment_count_missing"
                    if html_comment_count is None
                    else "html_comment_count_zero"
                )
                _emit_event(
                    on_event,
                    name="检测 HTML 评论数",
                    started_at=step_started,
                    result=(
                        "未解析到 HTML 评论数，判定无评论，未发起评论接口请求"
                        if html_comment_count is None
                        else "HTML 评论数为 0，未发起评论接口请求"
                    ),
                    details={
                        "html_comment_count": html_comment_count,
                        "stop_reason": stop_reason,
                    },
                )
                return self._finish_without_comment_resources(
                    context=context,
                    article=article,
                    record=record,
                    started=started,
                    started_monotonic=started_monotonic,
                    html_comment_count=html_comment_count,
                    stop_reason=stop_reason,
                    on_event=on_event,
                )
            else:
                _emit_event(
                    on_event,
                    name="检测 HTML 评论数",
                    started_at=step_started,
                    result="已解析到 HTML 评论数，继续请求评论接口",
                    details={"html_comment_count": html_comment_count},
                )

                step_started = self._monotonic()
                identity = extract_comment_identity(article_url, html_text, headers)
                missing = [key for key in REQUIRED_IDENTITY_KEYS if not identity.get(key)]
                if missing:
                    _emit_event(
                        on_event,
                        name="提取评论请求参数",
                        started_at=step_started,
                        result=f"参数不足：{', '.join(missing)}",
                        status="skipped",
                    )
                    raise CommentParametersMissing(f"评论参数不足：{', '.join(missing)}")
                _emit_event(
                    on_event,
                    name="提取评论请求参数",
                    started_at=step_started,
                    result="已从 origin/main.html 提取评论参数",
                    details={
                        "html_comment_count": html_comment_count,
                        "required_parameter_count": len(REQUIRED_IDENTITY_KEYS),
                    },
                )

                step_started = self._monotonic()
                fetched = self._requester.fetch(
                    article_url=article_url,
                    html=html_text,
                    request_headers=headers,
                    timeout_seconds=timeout_seconds,
                    page_interval_seconds=page_interval_seconds,
                    max_pages=max_pages,
                )
                _emit_event(
                    on_event,
                    name="请求评论分页与回复",
                    started_at=step_started,
                    result="已完成评论接口请求",
                    details={
                        "html_comment_count": html_comment_count,
                        "comment_count": fetched.comment_count,
                        "reply_count": fetched.reply_count,
                        "page_count": fetched.page_count,
                        "stop_reason": fetched.stop_reason,
                    },
                )
        except CommentParametersMissing as exc:
            return ServiceResult(
                status=TaskStatus.SKIPPED,
                error_code=None,
                message=str(exc),
                duration_seconds=max(0.0, self._monotonic() - started_monotonic),
            )
        except Exception as exc:
            duration = max(0.0, self._monotonic() - started_monotonic)
            message = _safe_error_message(exc)
            if record is not None:
                self._append_failure_safely(
                    context=context,
                    record=record,
                    started=started,
                    duration_seconds=duration,
                    message=message,
                )
            return ServiceResult.failure(
                ErrorCode.COMMENT_FETCH_FAILED,
                message,
                duration_seconds=duration,
            )

        return self._store_fetched_comments(
            context=context,
            article=article,
            record=record,
            started=started,
            started_monotonic=started_monotonic,
            fetched=fetched,
            html_comment_count=html_comment_count,
            article_url=article_url,
            headers=headers,
            timeout_seconds=timeout_seconds,
            on_event=on_event,
        )

    def _finish_without_comment_resources(
        self,
        *,
        context: TaskContext,
        article: ArticleSaveData,
        record: ArticleRecord,
        started: datetime,
        started_monotonic: float,
        html_comment_count: int | None,
        stop_reason: str,
        on_event: Callable[[dict[str, Any]], None] | None,
    ) -> ServiceResult[CommentCollectData]:
        backup_root = context.temp_dir / "comment_cleanup_backup" / str(article.article_id) / article.attempt_id
        try:
            with self._write_coordinator.hold():
                step_started = self._monotonic()
                removed = _remove_comment_resources(
                    article.article_directory,
                    backup_root=backup_root,
                )
                _emit_event(
                    on_event,
                    name="跳过评论资源处理",
                    started_at=step_started,
                    result=(
                        "未解析到 HTML 评论数，未写入 comments/final.json"
                        if html_comment_count is None
                        else "HTML 评论数为 0，未写入 comments/final.json"
                    ),
                    details={
                        "comment_path": "",
                        "asset_count": 0,
                        "removed_resource_count": len(removed),
                    },
                )

                finished = self._now()
                finished_text = _format_time(finished)
                duration = max(0.0, self._monotonic() - started_monotonic)
                manifest_holder: dict[str, ResourceManifest] = {}

                def database_operation(connection):
                    # 0 评论不产生评论文件，只刷新文章资源清单并记录本次评论检查结果。
                    manifest = self._manifest_builder.build(article.article_directory)
                    manifest_holder["value"] = manifest
                    ArticleRepository(connection).update_resource_manifest(
                        article.article_id,
                        manifest,
                        collected_time=finished_text,
                    )
                    return FetchHistoryRepository(connection).append(
                        FetchHistoryWrite(
                            article_id=article.article_id,
                            account_id=article.account_id,
                            target_account_name=record_account_name(record),
                            target_title=record.article_title,
                            target_link=record.article_link,
                            task_type="comment_fetch",
                            resource_manifest=ResourceManifest.from_types([]),
                            status=TaskStatus.SUCCESS,
                            started_time=_format_time(started),
                            finished_time=finished_text,
                            duration_seconds=duration,
                            output_dir=article.archive_dir,
                        )
                    )

                step_started = self._monotonic()
                try:
                    with sqlite_connection(context.db_path) as connection:
                        history_id = database_operation(connection)
                except Exception:
                    _restore_comment_resources(
                        article.article_directory,
                        backup_root=backup_root,
                    )
                    raise
                _discard_backup(backup_root)
            _emit_event(
                on_event,
                name="提交无评论状态",
                started_at=step_started,
                result="已刷新资源清单并记录无评论结果",
                details={"history_id": history_id},
            )
            return ServiceResult.success(
                CommentCollectData(
                    article_id=article.article_id,
                    history_id=history_id,
                    comment_path=None,
                    comment_count=0,
                    reply_count=0,
                    page_count=0,
                    pagination_complete=True,
                    stop_reason=stop_reason,
                    resource_manifest=manifest_holder.get("value")
                    or self._manifest_builder.build(article.article_directory),
                    asset_count=0,
                    asset_dir=None,
                    html_comment_count=html_comment_count,
                    download_bytes=0,
                ),
                duration_seconds=duration,
            )
        except Exception as exc:
            duration = max(0.0, self._monotonic() - started_monotonic)
            message = _safe_error_message(exc)
            self._append_failure_safely(
                context=context,
                record=record,
                started=started,
                duration_seconds=duration,
                message=message,
            )
            return ServiceResult.failure(
                ErrorCode.COMMENT_FETCH_FAILED,
                message,
                duration_seconds=duration,
            )

    def _store_fetched_comments(
        self,
        *,
        context: TaskContext,
        article: ArticleSaveData,
        record: ArticleRecord,
        started: datetime,
        started_monotonic: float,
        fetched: CommentFetchData,
        html_comment_count: int | None,
        article_url: str,
        headers: Mapping[str, Any],
        timeout_seconds: float,
        on_event: Callable[[dict[str, Any]], None] | None,
    ) -> ServiceResult[CommentCollectData]:
        try:
            stage_root = context.temp_dir / "comment_stage" / str(article.article_id) / article.attempt_id
            backup_root = context.temp_dir / "comment_backup" / str(article.article_id) / article.attempt_id
            _prepare_stage(stage_root)
            asset_count = 0

            step_started = self._monotonic()
            comment_file = stage_root / COMMENT_RESOURCE_PATH
            comment_file.parent.mkdir(parents=True, exist_ok=True)
            (comment_file).write_text(
                json.dumps(fetched.package, ensure_ascii=False, indent=2, default=str),
                encoding="utf-8",
            )
            _emit_event(
                on_event,
                name="准备评论存储文件",
                started_at=step_started,
                result="已写入 comments/final.json 暂存文件",
                details={
                    "comment_path": COMMENT_RESOURCE_PATH.as_posix(),
                },
            )

            resource_paths = [COMMENT_RESOURCE_PATH]
            finished = self._now()
            finished_text = _format_time(finished)
            duration = max(0.0, self._monotonic() - started_monotonic)
            manifest_holder: dict[str, ResourceManifest] = {}

            def database_operation(connection):
                # 资源已先替换到文章目录，此时再按真实文件状态刷新资源清单。
                manifest = self._manifest_builder.build(article.article_directory)
                manifest_holder["value"] = manifest
                ArticleRepository(connection).update_resource_manifest(
                    article.article_id,
                    manifest,
                    collected_time=finished_text,
                )
                return FetchHistoryRepository(connection).append(
                    FetchHistoryWrite(
                        article_id=article.article_id,
                        account_id=article.account_id,
                        target_account_name=record_account_name(record),
                        target_title=record.article_title,
                        target_link=record.article_link,
                        task_type="comment_fetch",
                        resource_manifest=ResourceManifest.from_types([ResourceType.COMMENT_DETAIL]),
                        status=TaskStatus.SUCCESS,
                        started_time=_format_time(started),
                        finished_time=finished_text,
                        duration_seconds=duration,
                        output_dir=article.archive_dir,
                    )
                )

            step_started = self._monotonic()
            with self._write_coordinator.hold():
                # 评论资源下载已停用，提交新 final.json 前清理旧 assets。
                _remove_comment_asset_resources(article.article_directory)
                history_id = self._commit_service.commit(
                    database_path=context.db_path,
                    stage_root=stage_root,
                    target_root=article.article_directory,
                    backup_root=backup_root,
                    resource_paths=tuple(resource_paths),
                    database_operation=database_operation,
                )
            _emit_event(
                on_event,
                name="提交评论文件和数据库记录",
                started_at=step_started,
                result="已保存 comments/final.json 并更新数据库",
                details={"history_id": history_id},
            )
            return ServiceResult.success(
                CommentCollectData(
                    article_id=article.article_id,
                    history_id=history_id,
                    comment_path=article.article_directory / COMMENT_RESOURCE_PATH,
                    comment_count=fetched.comment_count,
                    reply_count=fetched.reply_count,
                    page_count=fetched.page_count,
                    pagination_complete=fetched.pagination_complete,
                    stop_reason=fetched.stop_reason,
                    resource_manifest=manifest_holder.get("value")
                    or self._manifest_builder.build(article.article_directory),
                    asset_count=asset_count,
                    asset_dir=None,
                    html_comment_count=html_comment_count,
                    download_bytes=fetched.download_bytes,
                ),
                duration_seconds=duration,
            )
        except Exception as exc:
            duration = max(0.0, self._monotonic() - started_monotonic)
            message = _safe_error_message(exc)
            self._append_failure_safely(
                context=context,
                record=record,
                started=started,
                duration_seconds=duration,
                message=message,
            )
            return ServiceResult.failure(
                ErrorCode.COMMENT_FETCH_FAILED,
                message,
                duration_seconds=duration,
            )

    @staticmethod
    def _load_article(database_path: Path, article_id: int) -> ArticleRecord:
        with sqlite_connection(database_path, write=False) as connection:
            record = ArticleRepository(connection).get_by_id(article_id)
        if record is None:
            raise LookupError(f"文章索引不存在：{article_id}")
        return record

    def _append_failure_safely(
        self,
        *,
        context: TaskContext,
        record: ArticleRecord,
        started: datetime,
        duration_seconds: float,
        message: str,
    ) -> None:
        try:
            finished = datetime.now()
            with self._write_coordinator.hold():
                with sqlite_connection(context.db_path) as connection:
                    FetchHistoryRepository(connection).append(
                        FetchHistoryWrite.failed(
                            article_id=record.id,
                            account_id=record.account_id,
                            target_account_name=record_account_name(record),
                            target_title=record.article_title,
                            target_link=record.article_link,
                            task_type="comment_fetch",
                            started_time=_format_time(started),
                            finished_time=_format_time(finished),
                            duration_seconds=duration_seconds,
                            error_stage="comment_fetch",
                            error_message=message,
                        )
                    )
        except Exception:
            # 评论失败历史写入失败不能掩盖最初的评论异常。
            pass


def _load_request_inputs(article_directory: Path) -> tuple[str, str, dict[str, Any]]:
    html_path = article_directory / "origin" / "main.html"
    request_path = article_directory / "origin" / "request.json"
    html_text = html_path.read_text(encoding="utf-8")
    evidence = json.loads(request_path.read_text(encoding="utf-8"))
    reference = evidence.get("reference") if isinstance(evidence, Mapping) else None
    if not isinstance(reference, Mapping):
        raise CommentParametersMissing("缺少文章 reference，无法获取评论参数")
    article_url = normalize_article_url(str(reference.get("url") or ""))
    headers = reference.get("request_headers")
    return html_text, article_url, dict(headers) if isinstance(headers, Mapping) else {}


def _prepare_stage(stage_root: Path) -> None:
    if stage_root.exists():
        shutil.rmtree(stage_root)
    stage_root.mkdir(parents=True, exist_ok=False)


def _remove_comment_resources(article_directory: Path, *, backup_root: Path) -> list[Path]:
    if backup_root.exists():
        shutil.rmtree(backup_root)
    removed: list[Path] = []
    for relative in (
        COMMENT_RESOURCE_PATH,
        COMMENT_ASSETS_PATH,
        Path("comments_final.json"),
        Path("comment_assets"),
    ):
        target = article_directory / relative
        if not target.exists():
            continue
        backup = backup_root / relative
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(target), str(backup))
        removed.append(relative)
    _remove_empty_comment_dirs(article_directory)
    return removed


def _restore_comment_resources(article_directory: Path, *, backup_root: Path) -> None:
    if not backup_root.exists():
        return
    for backup in sorted(backup_root.rglob("*"), key=lambda path: len(path.parts)):
        if not backup.exists() or backup.is_dir():
            continue
        relative = backup.relative_to(backup_root)
        target = article_directory / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            if target.is_dir():
                shutil.rmtree(target)
            else:
                target.unlink()
        shutil.move(str(backup), str(target))
    for backup in sorted(
        (path for path in backup_root.rglob("*") if path.is_dir()),
        key=lambda path: len(path.parts),
        reverse=True,
    ):
        relative = backup.relative_to(backup_root)
        target = article_directory / relative
        if target.exists():
            continue
        target.mkdir(parents=True, exist_ok=True)
    _discard_backup(backup_root)


def _discard_backup(backup_root: Path) -> None:
    if backup_root.exists():
        shutil.rmtree(backup_root)


def _remove_empty_comment_dirs(article_directory: Path) -> None:
    for directory in (
        article_directory / "comments" / "assets",
        article_directory / "comments",
        article_directory / "comment_assets",
    ):
        try:
            directory.rmdir()
        except OSError:
            pass


def _remove_comment_asset_resources(article_directory: Path) -> list[Path]:
    removed: list[Path] = []
    for relative in (
        COMMENT_ASSETS_PATH,
        Path("comment_assets"),
    ):
        target = article_directory / relative
        if not target.exists():
            continue
        if target.is_dir() and not target.is_symlink():
            shutil.rmtree(target)
        else:
            target.unlink(missing_ok=True)
        removed.append(relative)
    _remove_empty_comment_dirs(article_directory)
    return removed


def _emit_event(
    callback: Callable[[dict[str, Any]], None] | None,
    *,
    name: str,
    started_at: float,
    result: str,
    status: str = "success",
    details: Mapping[str, Any] | None = None,
) -> None:
    if callback is None:
        return
    callback(
        {
            "name": name,
            "elapsed_seconds": round(max(0.0, time.monotonic() - started_at), 3),
            "status": status,
            "result": result,
            "details": dict(details or {}),
        }
    )


def _safe_error_message(exc: Exception) -> str:
    if isinstance(exc, CommentFetchError):
        return str(exc)
    if isinstance(exc, (OSError, ValueError, LookupError, json.JSONDecodeError)):
        return f"评论采集失败：{type(exc).__name__}: {exc}"
    return f"评论采集失败：{type(exc).__name__}"


def record_account_name(record: ArticleRecord) -> str:
    # article 表只保留 account_id；评论历史的账号快照由归档相对路径首段提供。
    path = Path(record.archive_dir)
    return path.parts[0] if path.parts else ""


def _format_time(value: datetime) -> str:
    return value.strftime("%Y-%m-%d %H:%M:%S")
