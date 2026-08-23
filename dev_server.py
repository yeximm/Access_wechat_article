from __future__ import annotations

from contextlib import asynccontextmanager
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime
import importlib.metadata
import json
from pathlib import Path
import platform
import re
import socket
import subprocess
import sys
import tempfile
from threading import Lock
from typing import Any

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from pydantic import model_validator
from ruamel.yaml import YAML
import uvicorn

from src.app.main_orchestrator import (
    ApplicationRuntime,
    build_application_runtime,
    create_capture_task_manager,
    load_application_runtime,
)
from src.config.config_loader import build_app_config, load_config_mapping
from src.domain.enums import TaskStatus
from src.domain.models import TaskCommand
from src.modules.window.article_date_filter import ArticleDateFilter
from src.modules.processes.process_launcher import set_process_observer
from src.modules.proxy.capture_buffer import CaptureBuffer
from src.modules.proxy.mitmproxy_listener import MitmproxyListener, MitmproxyListenerError
from src.modules.proxy.proxy_state import ProxySnapshot, proxy_points_to
from src.modules.request.chrome_headers import build_chrome_document_headers
from src.modules.system.windows_system_proxy import WindowsSystemProxy
from src.services.archive.archive_delete_service import ArchiveDeleteService
from src.services.archive.archive_excel_export_service import ArchiveExcelExportService
from src.services.archive.offline_cache_job_service import (
    OfflineCacheConflictError,
    OfflineCacheJobService,
)
from src.services.capture.capture_runtime_factory import CaptureRuntimeFactory
from src.services.capture.window_runtime_factory import WindowRuntimeFactory
from src.services.history.history_query_service import HistoryQueryService
from src.services.history.history_clear_service import HistoryClearService
from src.services.main_flow.main_flow_models import MainFlowCommand
from src.services.main_flow.main_flow_factory import build_main_flow_service
from src.services.main_flow.main_flow_service import (
    MainFlowConflictError,
    MainFlowService,
)
from src.services.runtime.database_init_service import DatabaseInitService
from src.services.runtime.article_card_probe_service import ArticleCardProbeService
from src.services.runtime.runtime_cache_clear_service import RuntimeCacheClearService
from src.services.runtime.runtime_directory_service import RuntimeDirectoryService
from src.services.runtime.runtime_log_service import RuntimeLogService
from src.services.runtime.startup_self_check_service import StartupSelfCheckService
from src.services.runtime.process_tree_resource_monitor import ProcessTreeResourceMonitor
from src.services.runtime.task_runtime_state import TaskRuntimeTracker
from src.services.runtime.window_diagnostic_service import (
    WINDOW_DIAGNOSTIC_ACTIONS,
    WindowDiagnosticService,
)
from src.services.task.window_click_flow_huey_service import (
    WindowClickFlowConflictError,
    WindowClickFlowHueyService,
)
from src.services.task.single_article_detail_huey_service import (
    SingleArticleDetailHueyService,
)
from src.services.task.initial_content_storage_huey_service import (
    InitialContentStorageHueyService,
)
from src.services.task.article_detail_comments_huey_service import (
    ArticleDetailCommentsHueyService,
)
from src.services.task.article_detail_offline_cache_huey_service import (
    ArticleDetailOfflineCacheHueyService,
)
from src.services.task.huey_runtime_queue import reset_runtime_huey_queue_dir
from src.services.task.task_manager import TaskConflictError
from src.storage.sqlite.connection import sqlite_connection


DEFAULT_API_HOST = "127.0.0.1"
DEFAULT_API_PORT = 8766
DEFAULT_TRAVERSE_ALL_MAX_ATTEMPTS = 1000
WEBVIEW_DIR = Path(__file__).resolve().parent / "src" / "webview"


@dataclass(slots=True)
class DevBackendContext:
    """开发期后端共享上下文，只保存入口层需要的运行时对象。"""

    project_root: Path
    runtime: ApplicationRuntime | Any
    db_path: Path
    task_manager: Any
    main_flow_service: MainFlowService | Any | None = None
    runtime_logger: RuntimeLogService | Any | None = None
    system_resource_monitor: Any | None = None
    started_at: datetime = field(default_factory=datetime.now)
    active_task_id: str | None = None
    logs: list[dict[str, Any]] = field(default_factory=list)
    config_mapping: dict[str, Any] | None = None
    directory_selector: Any | None = None
    command_runner: Any | None = None
    proxy_tester: Any | None = None
    system_proxy: Any | None = None
    port_checker: Any | None = None
    mitm_listener_factory: Any | None = None
    window_diagnostic_runner: Any | None = None
    window_click_flow_huey_service: Any | None = None
    article_detail_card_probe_service: Any | None = None
    article_detail_huey_service: Any | None = None
    main_flow_foreground_huey_service: Any | None = None
    main_flow_detail_huey_service: Any | None = None
    diagnostic_mitm_listener: Any | None = None
    diagnostic_mitm_started_at: float = 0.0
    diagnostic_system_proxy_snapshot: ProxySnapshot | None = None
    initial_content_storage_huey_service: Any | None = None
    article_detail_comments_huey_service: Any | None = None
    article_detail_offline_cache_huey_service: Any | None = None
    health_check_results: dict[str, dict[str, Any]] = field(default_factory=dict)
    health_startup_completed: bool = False
    health_check_lock: Any = field(default_factory=Lock, repr=False)
    startup_self_check_lock: Any = field(default_factory=Lock, repr=False)
    offline_cache_service: Any | None = None

    def append_log(
        self,
        level: str,
        message: str,
        source: str = "dev_server",
        *,
        summary: bool = False,
        channel: str | None = None,
        phase: str | None = None,
        task_index: int | str | None = None,
        article_task_id: str | None = None,
        article_title: str | None = None,
        context: dict[str, Any] | None = None,
        exception: BaseException | None = None,
    ) -> None:
        if self.runtime_logger is not None:
            if str(level).upper() == "ERROR":
                self.runtime_logger.write_error(
                    message,
                    source=source,
                    channel=channel,
                    phase=phase,
                    task_index=task_index,
                    article_task_id=article_task_id,
                    article_title=article_title,
                    context=context,
                    exception=exception,
                    summary=summary,
                )
            elif summary:
                self.runtime_logger.write_summary(
                    level,
                    message,
                    source=source,
                    channel=channel,
                    phase=phase,
                    task_index=task_index,
                    article_task_id=article_task_id,
                    article_title=article_title,
                    context=context,
                )
            else:
                self.runtime_logger.write_detail(
                    level,
                    message,
                    source=source,
                    channel=channel,
                    phase=phase,
                    task_index=task_index,
                    article_task_id=article_task_id,
                    article_title=article_title,
                    context=context,
                    exception=exception,
                )
            return
        if summary:
            self.logs.append(
                {
                    "level": level.upper(),
                    "message": message,
                    "source": source,
                    "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "channel": channel or "system",
                    "phase": phase or "",
                    "taskIndex": task_index or "",
                    "articleTaskId": article_task_id or "",
                    "articleTitle": article_title or "",
                }
            )
            del self.logs[:-100]

    def recent_summary_logs(self, limit: int = 100) -> list[dict[str, Any]]:
        if self.runtime_logger is not None:
            return self.runtime_logger.recent_summary(limit)
        safe_limit = max(1, min(int(limit), 100))
        return [dict(item) for item in self.logs[-safe_limit:]]


class RuntimePathOpenPayload(BaseModel):
    key: str


class RuntimeDirectorySelectPayload(BaseModel):
    configKey: str
    currentPath: str | None = None


class ArchiveArticleOpenDirectoryPayload(BaseModel):
    articleId: int


class ArchiveDeleteArticlesPayload(BaseModel):
    articleIds: list[int]


class ArchiveExcelExportPayload(BaseModel):
    accountIds: list[int]


class ArchiveCacheArticlesPayload(BaseModel):
    articleIds: list[int]


class UnsupportedPayload(BaseModel):
    """占位请求体，避免前端 POST 空对象时触发 422。"""


class HealthCheckPayload(BaseModel):
    target: str


class WindowDiagnosticPayload(BaseModel):
    action: str
    # 仅覆盖本次“滚动页面”诊断，不写回运行配置或 YAML。
    scrollSteps: int | None = Field(default=None, ge=1, le=200)

    @model_validator(mode="after")
    def validate_scroll_steps(self) -> WindowDiagnosticPayload:
        if self.action.strip() == "scroll-page" and self.scrollSteps is None:
            raise ValueError("滚动页面诊断必须提供滚动步长")
        return self


class WindowClickFlowDiagnosticPayload(BaseModel):
    """窗口点击流程诊断条件；数量为零表示由日期边界结束。"""

    maxRecords: int = 20
    dateFilterMode: str = "all"
    startDate: str | None = None
    endDate: str | None = None

    @model_validator(mode="after")
    def validate_options(self) -> WindowClickFlowDiagnosticPayload:
        date_filter = ArticleDateFilter.create(
            mode=self.dateFilterMode,
            start_date=self.startDate,
            end_date=self.endDate,
        )
        self.dateFilterMode = date_filter.mode
        self.startDate = date_filter.start_date.isoformat() if date_filter.start_date else None
        self.endDate = date_filter.end_date.isoformat() if date_filter.end_date else None
        # 范围和截止模式由日期边界停止；其他模式最多读取 20 条。
        self.maxRecords = (
            0
            if date_filter.mode in {"range", "before"}
            else max(1, min(int(self.maxRecords), 20))
        )
        return self


class ArticleDetailDiagnosticPayload(BaseModel):
    """单篇详情诊断输入；当前前端只传跳过已采集记录开关。"""

    cardIndex: int = Field(default=1, ge=1)
    accountName: str | None = None
    card: dict[str, Any] | None = None
    skipCollectedRecords: bool = False


class InitialContentStorageDiagnosticPayload(BaseModel):
    """初始内容存储诊断输入；文章详情存储是该流程的锁定动作。"""

    skipCollectedRecords: bool = False
    storeArticleDetail: bool = True

    @model_validator(mode="after")
    def lock_store_article_detail(self) -> InitialContentStorageDiagnosticPayload:
        # 前端显示为锁定勾选；后端也强制为 True，避免被手动请求绕过。
        self.storeArticleDetail = True
        return self


class ArticleDetailCommentsDiagnosticPayload(BaseModel):
    """评论信息存储诊断输入；文章详情和评论信息都是锁定动作。"""

    skipCollectedRecords: bool = False
    storeArticleDetail: bool = True
    storeCommentInfo: bool = True

    @model_validator(mode="after")
    def lock_store_options(self) -> ArticleDetailCommentsDiagnosticPayload:
        # 前端两个存储项显示为锁定勾选；后端也强制开启，避免手动请求绕过流程。
        self.storeArticleDetail = True
        self.storeCommentInfo = True
        return self


class ArticleDetailOfflineCacheDiagnosticPayload(BaseModel):
    """离线缓存诊断输入；文章详情和离线归档内容都是锁定动作。"""

    skipCollectedRecords: bool = False
    storeArticleDetail: bool = True
    archiveOfflineContent: bool = True
    statefulOfflineCache: bool = False

    @model_validator(mode="after")
    def lock_store_options(self) -> ArticleDetailOfflineCacheDiagnosticPayload:
        # 前端两个存储项显示为锁定勾选；带状态是可选实验项，不在这里锁定。
        self.storeArticleDetail = True
        self.archiveOfflineContent = True
        return self


def create_dev_backend(
    *,
    project_root: str | Path | None = None,
    config_path: str | Path | None = None,
) -> DevBackendContext:
    """按新结构装配开发期 FastAPI 后端，并在启动阶段初始化数据库。"""
    root = Path(project_root or Path(__file__).resolve().parent).resolve()
    runtime = load_application_runtime(project_root=root, config_path=config_path)
    RuntimeDirectoryService().prepare(runtime.config)
    db_path = DatabaseInitService(project_root=root).initialize(runtime.config)
    runtime_logger = RuntimeLogService(
        log_dir=runtime.config.storage.log_dir,
        level=runtime.config.runtime.log_level,
        redactions={str(db_path): db_path.name},
    )
    # Huey 的 SQLite 队列只保存当前进程运行态任务，启动时清理旧队列，
    # 并放在 data/runtime/huey，避免被“清理缓存”误删 data/tmp 时破坏。
    reset_runtime_huey_queue_dir(runtime.config.storage.temp_dir)
    # 诊断工具和主服务使用独立的详情队列；两者仍通过服务内部的进程级
    # foreground lease 互斥接管主页和全局代理。
    initial_content_service = InitialContentStorageHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        worker_count=1,
    )
    main_flow_foreground_service = InitialContentStorageHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        write_coordinator=runtime.database_write_coordinator,
        worker_count=1,
        job_prefix="main-flow-foreground",
        queue_name="main-flow-foreground",
        task_name="MainFlowForegroundTask",
        action="main-flow-foreground",
        title="主服务单篇前台结果",
        flow_label="主服务单篇前台捕获",
    )
    main_flow_detail_service = InitialContentStorageHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        write_coordinator=runtime.database_write_coordinator,
        worker_count=2,
        job_prefix="main-flow-detail",
        queue_name="main-flow-detail",
        task_name="MainFlowDetailTask",
        action="main-flow-detail",
        title="主服务单篇详情结果",
        flow_label="主服务单篇详情",
    )
    # 主服务和诊断工具使用不同的 Huey 队列；两套 service 可以共享运行时
    # DatabaseWriteCoordinator，但不能共享队列、任务记录或关闭生命周期。
    main_flow_comments_service = ArticleDetailCommentsHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        write_coordinator=runtime.database_write_coordinator,
        worker_count=runtime.config.comment.max_concurrent_processes,
    )
    main_flow_offline_service = ArticleDetailOfflineCacheHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        browser_cache_dir=root / ".playwright-browsers",
        write_coordinator=runtime.database_write_coordinator,
        worker_count=runtime.config.offline_cache.max_concurrent_processes,
    )
    diagnostic_comments_service = ArticleDetailCommentsHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        write_coordinator=runtime.database_write_coordinator,
        worker_count=1,
    )
    diagnostic_offline_service = ArticleDetailOfflineCacheHueyService(
        temp_root=runtime.config.storage.temp_dir,
        config=runtime.config,
        window_factory=runtime.window_factory,
        capture_factory=runtime.capture_factory,
        database_path=db_path,
        browser_cache_dir=root / ".playwright-browsers",
        write_coordinator=runtime.database_write_coordinator,
        worker_count=1,
    )
    main_flow_service, _main_flow_dispatcher = build_main_flow_service(
        project_root=root,
        config=runtime.config,
        db_path=db_path,
        storage_root=runtime.config.storage.article_storage_root,
        temp_root=runtime.config.storage.temp_dir,
        window_factory=runtime.window_factory,
        foreground_service=main_flow_foreground_service,
        detail_service=main_flow_detail_service,
        comments_service=main_flow_comments_service,
        offline_service=main_flow_offline_service,
        runtime_logger=runtime_logger,
    )
    # 运行状态区只展示当前软件进程树的资源占用；后台独立采样，前端读取时不再临时扫整机。
    resource_monitor = ProcessTreeResourceMonitor(
        history_sample_count=120,
        sample_interval_seconds=0.5,
    )
    set_process_observer(resource_monitor)
    resource_monitor.start()
    backend = DevBackendContext(
        project_root=root,
        runtime=runtime,
        db_path=db_path,
        task_manager=create_capture_task_manager(
            runtime=runtime,
            db_path=db_path,
            runtime_logger=runtime_logger,
        ),
        main_flow_service=main_flow_service,
        runtime_logger=runtime_logger,
        system_resource_monitor=resource_monitor,
        window_click_flow_huey_service=WindowClickFlowHueyService(
            temp_root=runtime.config.storage.temp_dir,
            config=runtime.config,
            window_factory=runtime.window_factory,
        ),
        article_detail_card_probe_service=ArticleCardProbeService(
            config=runtime.config,
            window_factory=runtime.window_factory,
        ),
        article_detail_huey_service=SingleArticleDetailHueyService(
            temp_root=runtime.config.storage.temp_dir,
            config=runtime.config,
            window_factory=runtime.window_factory,
            capture_factory=runtime.capture_factory,
            database_path=db_path,
        ),
        main_flow_foreground_huey_service=main_flow_foreground_service,
        main_flow_detail_huey_service=main_flow_detail_service,
        initial_content_storage_huey_service=initial_content_service,
        article_detail_comments_huey_service=diagnostic_comments_service,
        article_detail_offline_cache_huey_service=diagnostic_offline_service,
        offline_cache_service=OfflineCacheJobService(
            database_path=db_path,
            storage_root=runtime.config.storage.article_storage_root,
            temp_root=runtime.config.storage.temp_dir,
            browser_cache_dir=root / ".playwright-browsers",
            max_concurrent_processes=runtime.config.offline_cache.max_concurrent_processes,
            max_scroll_seconds=runtime.config.offline_cache.max_scroll_seconds,
            resource_timeout_seconds=runtime.config.offline_cache.resource_timeout_seconds,
            write_coordinator=runtime.database_write_coordinator,
        ),
    )
    backend.append_log(
        "INFO",
        f"程序已就绪，数据库：{db_path.name}",
        source="startup",
        summary=True,
        channel="system",
        phase="startup",
    )
    return backend


def create_backend_app(backend: DevBackendContext) -> FastAPI:
    """创建开发期 FastAPI 应用；业务路由只调用 Service / TaskManager。"""

    @asynccontextmanager
    async def lifespan(_app: FastAPI):
        try:
            yield
        finally:
            shutdown_backend(backend)

    app = FastAPI(
        title="Access WeChat Article Dev API",
        version=str(backend.runtime.config.software.version),
        lifespan=lifespan,
    )
    app.state.backend = backend
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.exception_handler(Exception)
    async def handle_unexpected_error(_request: Request, exc: Exception):
        backend.append_log(
            "ERROR",
            f"程序运行异常：{type(exc).__name__}: {exc}",
            source="api",
            summary=True,
            exception=exc,
        )
        return JSONResponse(
            {
                "ok": False,
                "status": "error",
                "message": f"{type(exc).__name__}: {exc}",
            },
            status_code=500,
        )

    @app.get("/api/status")
    def get_status() -> dict[str, Any]:
        return {
            "ok": True,
            "status": "ready",
            "webviewExists": False,
            "serverTime": datetime.now().isoformat(timespec="seconds"),
            "environment": _environment_payload(backend.runtime.config),
            **_runtime_status_fields(backend),
        }

    @app.post("/api/health/startup")
    def run_startup_health_checks(_payload: UnsupportedPayload | None = None) -> dict[str, Any]:
        return _startup_health_payload(backend)

    @app.get("/api/startup-self-check/status")
    def get_startup_self_check_status() -> dict[str, Any]:
        return _startup_self_check_status_payload(backend)

    @app.post("/api/startup-self-check/run")
    def run_startup_self_check(_payload: UnsupportedPayload | None = None) -> dict[str, Any]:
        return _startup_self_check_run_payload(backend)

    @app.post("/api/health/check")
    def run_health_check(payload: HealthCheckPayload):
        target = payload.target.strip().lower()
        if target not in HEALTH_CHECK_TARGETS:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "invalid-target",
                    "message": f"不支持的健康检测项：{payload.target}",
                },
                status_code=400,
            )
        result = _run_health_check(backend, target)
        backend.health_check_results[target] = result
        return result

    @app.post("/api/task/start")
    def start_task(payload: dict[str, Any] | None = None):
        if backend.main_flow_service is not None:
            try:
                command = _main_flow_command_from_payload(
                    payload or {},
                    backend.runtime.config,
                )
                snapshot = backend.main_flow_service.start(command)
            except MainFlowConflictError as exc:
                return JSONResponse(
                    {
                        "ok": False,
                        "status": "conflict",
                        "message": str(exc),
                        "ownerTaskId": exc.owner_task_id,
                    },
                    status_code=409,
                )
            except Exception as exc:
                backend.append_log(
                    "ERROR",
                    f"启动主流程失败：{exc}",
                    source="main-flow-api",
                    summary=True,
                    exception=exc,
                )
                return JSONResponse(
                    {"ok": False, "status": "failed", "message": str(exc)},
                    status_code=400,
                )
            backend.active_task_id = snapshot.task_id
            return _main_flow_snapshot_payload(snapshot, backend)

        try:
            command = _task_command_from_payload(payload or {}, backend.runtime.config)
            snapshot = backend.task_manager.start_capture(command)
        except TaskConflictError as exc:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "conflict",
                    "message": str(exc),
                    "ownerTaskId": exc.owner_task_id,
                },
                status_code=409,
            )
        except Exception as exc:
            backend.append_log(
                "ERROR",
                f"启动采集任务失败：{exc}",
                source="task-api",
                summary=True,
                exception=exc,
            )
            return JSONResponse(
                {"ok": False, "status": "failed", "message": str(exc)},
                status_code=400,
            )

        backend.active_task_id = snapshot.task_id
        return _task_snapshot_payload(snapshot, backend)

    @app.post("/api/task/stop")
    def stop_task() -> dict[str, Any]:
        if backend.main_flow_service is not None:
            task_id = backend.active_task_id or backend.main_flow_service.active_task_id
            if not task_id:
                return _idle_task_payload(backend)
            stopped = backend.main_flow_service.stop(task_id)
            if not stopped:
                backend.append_log(
                    "WARN",
                    "当前主流程已经结束，无法再次停止",
                    source="main-flow-api",
                    summary=True,
                )
            return _current_task_payload(backend)

        task_id = backend.active_task_id
        if not task_id:
            return _idle_task_payload(backend)
        cancelled = bool(backend.task_manager.cancel(task_id))
        if not cancelled:
            backend.append_log(
                "WARN",
                "当前任务已经结束，无法再次停止",
                source="task-api",
                summary=True,
            )
        return _current_task_payload(backend)

    @app.get("/api/task/status")
    def get_task_status() -> dict[str, Any]:
        return _current_task_payload(backend)

    @app.get("/api/runtime/hardware")
    def get_runtime_hardware() -> dict[str, Any]:
        return _hardware_status_payload(backend)

    @app.get("/api/task/logs")
    def get_task_logs(limit: int = 100) -> dict[str, Any]:
        safe_limit = max(1, min(int(limit), 100))
        return {"ok": True, "items": backend.recent_summary_logs(safe_limit)}

    @app.get("/api/runtime/paths")
    def get_runtime_paths() -> dict[str, Any]:
        paths = _runtime_paths(backend)
        return {
            "ok": True,
            "status": "ok",
            "paths": {key: str(path) for key, path in paths.items()},
        }

    @app.post("/api/runtime/paths/open")
    def open_runtime_path(payload: RuntimePathOpenPayload):
        paths = _runtime_paths(backend)
        path = paths.get(payload.key)
        if path is None:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "invalid-key",
                    "key": payload.key,
                    "message": "未知运行目录 key。",
                },
                status_code=400,
            )
        opened = _open_directory(path)
        return {
            "ok": opened,
            "status": "opened" if opened else "open-failed",
            "key": payload.key,
            "path": str(path),
            "message": "已打开目录。" if opened else "打开目录失败。",
        }

    @app.post("/api/runtime/paths/select-directory")
    def select_runtime_directory(payload: RuntimeDirectorySelectPayload):
        return _select_runtime_directory(
            backend,
            config_key=payload.configKey,
            current_path=payload.currentPath,
        )

    @app.get("/api/archive/summary")
    def get_archive_summary() -> dict[str, Any]:
        return _archive_summary_payload(backend)

    @app.get("/api/archive/accounts")
    def list_archive_accounts() -> dict[str, Any]:
        items = _archive_account_items(backend)
        return {
            "ok": True,
            "status": "ok",
            "items": items,
            "total": len(items),
            "dbPath": str(backend.db_path),
        }

    @app.get("/api/archive/accounts/{account_id}/articles")
    def list_archive_account_articles(
        account_id: int,
        page: int = 1,
        pageSize: int = 10,
    ) -> dict[str, Any]:
        return _archive_account_articles_payload(backend, account_id, page, pageSize)

    @app.post("/api/archive/articles/open-directory")
    def open_archive_article_directory(payload: ArchiveArticleOpenDirectoryPayload):
        return _open_archive_article_directory(backend, payload.articleId)

    @app.delete("/api/archive/articles")
    def delete_archive_articles(payload: ArchiveDeleteArticlesPayload):
        return ArchiveDeleteService().delete_articles(
            database_path=backend.db_path,
            storage_root=backend.runtime.config.storage.article_storage_root,
            article_ids=payload.articleIds,
        )

    @app.delete("/api/archive/accounts/{account_id}")
    def delete_archive_account(account_id: int):
        return ArchiveDeleteService().delete_account(
            database_path=backend.db_path,
            storage_root=backend.runtime.config.storage.article_storage_root,
            account_id=account_id,
        )

    @app.delete("/api/archive")
    def delete_archive_all():
        return ArchiveDeleteService().delete_all(
            database_path=backend.db_path,
            storage_root=backend.runtime.config.storage.article_storage_root,
        )

    @app.post("/api/archive/export/accounts")
    def export_archive_accounts(payload: ArchiveExcelExportPayload):
        storage_root = backend.runtime.config.storage.article_storage_root
        return ArchiveExcelExportService().export_accounts(
            database_path=backend.db_path,
            storage_root=storage_root,
            account_ids=payload.accountIds,
            target_dir=storage_root,
        )

    @app.get("/api/history/records")
    def list_history_records(
        page: int = 1,
        pageSize: int = 15,
        keyword: str = "",
        collectType: str = "",
        status: str = "",
        collectDate: str = "",
        collectStartDate: str = "",
        collectEndDate: str = "",
    ) -> dict[str, Any]:
        return HistoryQueryService(backend.db_path).list_records(
            page=page,
            page_size=pageSize,
            keyword=keyword,
            collect_type=collectType,
            status=status,
            collect_date=collectDate,
            collect_start_date=collectStartDate,
            collect_end_date=collectEndDate,
        )

    @app.delete("/api/history/records")
    def clear_history_records():
        if getattr(backend, "active_task_id", None):
            return JSONResponse(
                {
                    "ok": False,
                    "status": "busy",
                    "message": "当前采集任务仍在运行，不能清空采集历史。",
                },
                status_code=409,
            )
        result = HistoryClearService(backend.db_path).clear_all()
        backend.append_log(
            "INFO",
            result["message"],
            source="history-clear",
            summary=True,
        )
        return result

    @app.get("/api/history/summary")
    def get_history_summary() -> dict[str, Any]:
        return HistoryQueryService(backend.db_path).get_summary()

    @app.get("/api/history/suggestions")
    def list_history_suggestions(keyword: str = "", limit: int = 20) -> dict[str, Any]:
        return HistoryQueryService(backend.db_path).list_suggestions(keyword=keyword, limit=limit)

    @app.get("/api/ca/status")
    def check_ca_certificate() -> dict[str, Any]:
        return _ca_certificate_status_payload(backend)

    @app.post("/api/proxy/test")
    def test_proxy_connection():
        return _proxy_connection_test_payload(backend)

    @app.post("/api/config/save")
    def save_runtime_config(payload: dict[str, Any] | None = None):
        try:
            _update_runtime_config_memory(backend, payload or {})
            config_path = _write_runtime_config_yaml(backend)
        except Exception as exc:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "save-failed",
                    "message": f"保存配置失败：{exc}",
                },
                status_code=400,
            )
        backend.append_log("INFO", f"配置已保存到 custom.yaml：{config_path}")
        return {
            "ok": True,
            "status": "saved",
            "message": "配置已保存到 custom.yaml。",
            "configPath": str(config_path),
            "taskStatus": _current_task_payload(backend),
        }

    @app.post("/api/config/update")
    def update_runtime_config(payload: dict[str, Any] | None = None):
        try:
            _update_runtime_config_memory(backend, payload or {})
        except Exception as exc:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "update-failed",
                    "message": f"更新运行配置失败：{exc}",
                },
                status_code=400,
            )
        backend.append_log("INFO", "配置已更新到当前进程内存，尚未写入 custom.yaml。")
        return {
            "ok": True,
            "status": "updated",
            "message": "配置已同步到当前进程内存。",
            "taskStatus": _current_task_payload(backend),
        }

    @app.post("/api/config/reset")
    def reset_runtime_config(_payload: UnsupportedPayload | None = None):
        busy_reason = _runtime_cache_busy_reason(backend)
        if busy_reason:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "busy",
                    "message": f"{busy_reason}，暂不能恢复系统默认配置。",
                },
                status_code=409,
            )

        config_service = getattr(backend.runtime, "config_service", None)
        if config_service is None or not hasattr(config_service, "restore_system_defaults"):
            return JSONResponse(
                {
                    "ok": False,
                    "status": "reset-unavailable",
                    "message": "当前运行时未提供系统默认配置恢复能力。",
                },
                status_code=400,
            )

        try:
            result = config_service.restore_system_defaults()
            backend.config_mapping = load_config_mapping(result.config_path)
            _replace_runtime_config(backend, result.config)
        except Exception as exc:
            backend.append_log(
                "ERROR",
                f"恢复系统默认配置失败：{exc}",
                source="config-api",
                summary=True,
                exception=exc,
            )
            return JSONResponse(
                {
                    "ok": False,
                    "status": "reset-failed",
                    "message": f"恢复系统默认配置失败：{exc}",
                },
                status_code=400,
            )

        backend.append_log(
            "INFO",
            f"已使用 system.yaml 恢复 custom.yaml：{result.config_path}",
            source="config-api",
            summary=True,
        )
        return {
            "ok": True,
            "status": "restored",
            "message": "已恢复系统默认配置。",
            "configPath": str(result.config_path),
            "backupPath": str(result.backup_path) if result.backup_path else None,
            "taskStatus": _current_task_payload(backend),
        }

    @app.post("/api/proxy/mitm/start")
    def start_mitm_proxy(_payload: UnsupportedPayload | None = None):
        return _set_diagnostic_mitm_payload(backend, enabled=True)

    @app.post("/api/proxy/mitm/stop")
    def stop_mitm_proxy(_payload: UnsupportedPayload | None = None):
        return _set_diagnostic_mitm_payload(backend, enabled=False)

    @app.post("/api/proxy/system/enable")
    def enable_system_proxy(_payload: UnsupportedPayload | None = None):
        return _set_system_proxy_payload(backend, enabled=True)

    @app.post("/api/proxy/system/disable")
    def disable_system_proxy(_payload: UnsupportedPayload | None = None):
        return _set_system_proxy_payload(backend, enabled=False)

    @app.post("/api/diagnostics/window")
    def run_window_diagnostic(payload: WindowDiagnosticPayload):
        action = payload.action.strip()
        if action not in WINDOW_DIAGNOSTIC_ACTIONS:
            return JSONResponse(
                {
                    "ok": False,
                    "status": "invalid-action",
                    "action": action,
                    "message": f"不支持的窗口诊断动作：{payload.action}",
                },
                status_code=400,
            )
        try:
            return _window_diagnostic_payload(
                backend,
                action,
                scroll_steps=payload.scrollSteps,
            )
        except Exception as exc:
            backend.append_log("ERROR", f"窗口诊断失败：{exc}")
            return JSONResponse(
                {
                    "ok": False,
                    "status": "failed",
                    "action": action,
                    "title": "窗口诊断失败",
                    "message": str(exc),
                    "tone": "error",
                    "items": [{"label": "失败原因", "value": str(exc)}],
                },
                status_code=400,
            )

    @app.post("/api/diagnostics/window-click-flow")
    def start_window_click_flow_diagnostic(payload: WindowClickFlowDiagnosticPayload):
        return _start_window_click_flow_diagnostic_job(backend, payload)

    @app.get("/api/diagnostics/window-click-flow/{job_id}")
    def get_window_click_flow_diagnostic(job_id: str):
        return _window_click_flow_diagnostic_job_payload(backend, job_id)

    @app.post("/api/diagnostics/window-click-flow/{job_id}/stop")
    def stop_window_click_flow_diagnostic(job_id: str, _payload: UnsupportedPayload | None = None):
        return _stop_window_click_flow_diagnostic_job(backend, job_id)

    @app.post("/api/diagnostics/article-detail")
    def start_article_detail_diagnostic(payload: ArticleDetailDiagnosticPayload):
        return _start_article_detail_diagnostic_job(backend, payload)

    @app.get("/api/diagnostics/article-detail/{job_id}")
    def get_article_detail_diagnostic(job_id: str):
        return _article_detail_diagnostic_job_payload(backend, job_id)

    @app.post("/api/diagnostics/initial-content-storage")
    def start_initial_content_storage_diagnostic(
        payload: InitialContentStorageDiagnosticPayload | None = None,
    ):
        return _start_initial_content_storage_diagnostic_job(backend, payload)

    @app.get("/api/diagnostics/initial-content-storage/{job_id}")
    def get_initial_content_storage_diagnostic(job_id: str):
        return _initial_content_storage_diagnostic_job_payload(backend, job_id)

    @app.post("/api/diagnostics/article-detail-comments")
    def start_article_detail_comments_diagnostic(
        payload: ArticleDetailCommentsDiagnosticPayload | None = None,
    ):
        return _start_article_detail_comments_diagnostic_job(backend, payload)

    @app.get("/api/diagnostics/article-detail-comments/{job_id}")
    def get_article_detail_comments_diagnostic(job_id: str):
        return _article_detail_comments_diagnostic_job_payload(backend, job_id)

    @app.post("/api/diagnostics/article-detail-offline-cache")
    def start_article_detail_offline_cache_diagnostic(
        payload: ArticleDetailOfflineCacheDiagnosticPayload | None = None,
    ):
        return _start_article_detail_offline_cache_diagnostic_job(backend, payload)

    @app.get("/api/diagnostics/article-detail-offline-cache/{job_id}")
    def get_article_detail_offline_cache_diagnostic(job_id: str):
        return _article_detail_offline_cache_diagnostic_job_payload(backend, job_id)

    @app.get("/api/ca/mitm/list")
    def list_mitm_ca_certificates():
        return _list_mitm_ca_certificates_payload(backend)

    @app.post("/api/ca/install")
    def install_ca_certificate(_payload: UnsupportedPayload | None = None):
        return _install_ca_certificate_payload(backend)

    @app.post("/api/ca/mitm/delete")
    def delete_mitm_ca_certificates(_payload: dict[str, Any] | None = None):
        return _delete_mitm_ca_certificates_payload(backend, _payload or {})

    @app.post("/api/cache/runtime/clear")
    def clear_runtime_cache(_payload: UnsupportedPayload | None = None):
        result = RuntimeCacheClearService(
            project_root=backend.project_root,
            config=backend.runtime.config,
            busy_check=lambda: _runtime_cache_busy_reason(backend),
        ).clear()
        backend.append_log(
            "INFO" if result.ok else "WARN",
            result.message,
            source="runtime-cache-clear",
        )
        return JSONResponse(result.to_payload(), status_code=result.http_status)

    _register_archive_cache_routes(app, backend)
    mount_webview_static_files(app, WEBVIEW_DIR)
    return app


def shutdown_backend(backend: DevBackendContext) -> None:
    """后端退出时只取消当前采集任务；8766 端口由 uvicorn 生命周期关闭。"""
    monitor = backend.system_resource_monitor
    stop_monitor = getattr(monitor, "stop", None)
    if callable(stop_monitor):
        try:
            stop_monitor()
        except Exception:
            pass
    if monitor is not None:
        set_process_observer(None)
    if backend.main_flow_service is not None:
        try:
            backend.main_flow_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止主流程编排器。")
        except Exception as exc:
            backend.append_log("ERROR", f"主流程编排器停止失败：{exc}")
    if backend.main_flow_foreground_huey_service is not None:
        try:
            backend.main_flow_foreground_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止主服务前台Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"主服务前台Huey队列停止失败：{exc}")
    if backend.main_flow_detail_huey_service is not None:
        try:
            backend.main_flow_detail_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止主服务详情Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"主服务详情Huey队列停止失败：{exc}")
    if backend.window_click_flow_huey_service is not None:
        try:
            backend.window_click_flow_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止窗口诊断Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"窗口诊断Huey队列停止失败：{exc}")
    if backend.article_detail_huey_service is not None:
        try:
            backend.article_detail_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止单篇详情Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"单篇详情Huey队列停止失败：{exc}")
    if backend.initial_content_storage_huey_service is not None:
        try:
            backend.initial_content_storage_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止初始内容存储Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"初始内容存储Huey队列停止失败：{exc}")
    if backend.article_detail_comments_huey_service is not None:
        try:
            backend.article_detail_comments_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止单篇评论存储Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"单篇评论存储Huey队列停止失败：{exc}")
    if backend.article_detail_offline_cache_huey_service is not None:
        try:
            backend.article_detail_offline_cache_huey_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止单篇离线缓存Huey队列。")
        except Exception as exc:
            backend.append_log("ERROR", f"单篇离线缓存Huey队列停止失败：{exc}")
    if backend.offline_cache_service is not None:
        try:
            backend.offline_cache_service.shutdown()
            backend.append_log("INFO", "后端退出，已停止离线缓存子进程。")
        except Exception as exc:
            backend.append_log("ERROR", f"离线缓存子进程停止失败：{exc}")
    if backend.diagnostic_mitm_listener is not None:
        try:
            backend.diagnostic_mitm_listener.stop()
            backend.append_log("INFO", "后端退出，已停止诊断 MITM 代理。")
        except Exception as exc:
            backend.append_log("ERROR", f"诊断 MITM 代理停止失败：{exc}")
        finally:
            backend.diagnostic_mitm_listener = None
    if backend.active_task_id:
        try:
            if backend.main_flow_service is not None:
                backend.main_flow_service.stop(backend.active_task_id)
            else:
                backend.task_manager.cancel(backend.active_task_id)
            backend.append_log("INFO", f"后端退出，已请求停止任务：{backend.active_task_id}")
        except Exception as exc:
            backend.append_log("ERROR", f"后端退出清理失败：{exc}")


def ensure_dev_server_port_available(
    host: str = DEFAULT_API_HOST,
    port: int = DEFAULT_API_PORT,
) -> None:
    """启动前检查开发 API 端口，避免多个后端进程同时接管前端请求。"""
    address = (host, int(port))
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.settimeout(0.2)
        if probe.connect_ex(address) == 0:
            raise RuntimeError(
                f"FastAPI 开发后端端口已被占用：{host}:{port}。请先停止旧的 dev_server.py。"
            )


def mount_webview_static_files(app: FastAPI, webview_dir: Path = WEBVIEW_DIR) -> None:
    """让开发后端也能直接返回已经构建好的 Vue 静态页面。"""
    index_html = webview_dir / "index.html"
    if not index_html.exists():
        return
    assets_dir = webview_dir / "assets"
    if assets_dir.exists():
        app.mount("/assets", StaticFiles(directory=str(assets_dir)), name="assets")

    @app.get("/")
    def serve_root():
        return FileResponse(index_html)

    @app.get("/index.html")
    def serve_index():
        return FileResponse(index_html)

    app.mount("/", StaticFiles(directory=str(webview_dir)), name="webview-static")


def main() -> None:
    ensure_dev_server_port_available()
    backend = create_dev_backend()
    app = create_backend_app(backend)
    print(f"AWA 开发后端已启动：http://{DEFAULT_API_HOST}:{DEFAULT_API_PORT}")
    print(f"API 文档：http://{DEFAULT_API_HOST}:{DEFAULT_API_PORT}/docs")
    print("按 Ctrl+C 停止开发后端。")
    uvicorn.run(app, host=DEFAULT_API_HOST, port=DEFAULT_API_PORT)


def _task_command_from_payload(payload: dict[str, Any], config: Any) -> TaskCommand:
    record_limit = _safe_int(payload.get("recordLimit"), default=1, minimum=0)
    selections = payload.get("selections")
    if not isinstance(selections, dict):
        selections = {}
    max_attempts = payload.get("maxAttempts")
    if max_attempts is None:
        max_attempts = (
            DEFAULT_TRAVERSE_ALL_MAX_ATTEMPTS
            if record_limit == 0
            else max(1, record_limit)
        )
    command = TaskCommand(
        target_success_count=record_limit,
        max_attempts=_safe_int(max_attempts, default=1, minimum=1),
        collect_comments=bool(
            selections.get("commentInfo", config.comment.enabled_by_default)
        ),
        build_offline_cache=bool(
            selections.get("offlineArchive", config.offline_cache.enabled_by_default)
        ),
        skip_collected_records=bool(
            selections.get("skipCollectedRecords", False)
        ),
        request_interval_seconds=float(config.request.request_interval_seconds),
    )
    return command


def _main_flow_command_from_payload(payload: dict[str, Any], config: Any) -> MainFlowCommand:
    """解析主服务页面命令；默认值来自已加载的内存配置。"""
    raw = dict(payload)
    selections = raw.get("selections")
    if not isinstance(selections, dict):
        selections = {}
    if "singleTaskIntervalSeconds" not in raw:
        raw["singleTaskIntervalSeconds"] = config.request.request_interval_seconds
    if "commentInfo" not in selections:
        selections["commentInfo"] = config.comment.enabled_by_default
    if "offlineArchive" not in selections:
        selections["offlineArchive"] = config.offline_cache.enabled_by_default
    raw["selections"] = selections
    return MainFlowCommand.from_payload(raw)


def _current_task_payload(backend: DevBackendContext) -> dict[str, Any]:
    if backend.main_flow_service is not None and backend.active_task_id:
        try:
            return _main_flow_snapshot_payload(
                backend.main_flow_service.get(backend.active_task_id),
                backend,
            )
        except KeyError:
            backend.active_task_id = None
            return _idle_task_payload(backend)
    if backend.main_flow_service is not None:
        current = backend.main_flow_service.current()
        if current is not None:
            backend.active_task_id = current.task_id
            return _main_flow_snapshot_payload(current, backend)
        return _idle_task_payload(backend)

    task_id = backend.active_task_id
    if not task_id:
        return _idle_task_payload(backend)
    try:
        snapshot = backend.task_manager.get(task_id)
    except KeyError:
        backend.active_task_id = None
        return _idle_task_payload(backend)
    return _task_snapshot_payload(snapshot, backend)


def _idle_task_payload(backend: DevBackendContext) -> dict[str, Any]:
    return {
        "ok": True,
        "status": "idle",
        "message": "当前没有运行中的采集任务。",
        **_runtime_status_fields(backend),
    }


def _main_flow_snapshot_payload(snapshot: Any, backend: DevBackendContext) -> dict[str, Any]:
    """把主流程领域快照转换成现有前端状态协议。"""
    payload = snapshot.to_payload()
    runtime_fields = _runtime_status_fields(backend)
    payload.update(runtime_fields)
    payload["runtimeState"] = dict(snapshot.runtime_state)
    state = snapshot.runtime_state
    payload.update(
        {
            "completedCount": int(state.get("progressDone") or 0),
            "failedCount": int(state.get("errorCount") or 0),
            "skippedCount": int(state.get("skippedCount") or 0),
            "totalAttempts": int(state.get("totalWorkerCount") or 0),
        }
    )
    return payload


def _task_snapshot_payload(snapshot: Any, backend: DevBackendContext) -> dict[str, Any]:
    status = _frontend_task_status(snapshot.status)
    result = getattr(snapshot, "result", None)
    data = getattr(result, "data", None)
    runtime_fields = _runtime_status_fields(backend)
    payload = {
        "ok": status not in {"failed"},
        "status": status,
        "taskId": str(getattr(snapshot, "task_id", "")),
        "proxyLeaseId": str(getattr(snapshot, "proxy_lease_id", "")),
        "message": str(getattr(snapshot, "message", "") or ""),
        **runtime_fields,
    }
    payload["runtimeState"] = _snapshot_runtime_state(
        snapshot,
        fallback=runtime_fields.get("runtimeState"),
        proxy_status_label=_proxy_status_label(payload.get("proxy", {})),
    )
    if data is not None:
        payload.update(
            {
                "completedCount": int(getattr(data, "success_count", 0)),
                "failedCount": int(getattr(data, "failed_attempt_count", 0)),
                "skippedCount": int(getattr(data, "skipped_count", 0)),
                "totalAttempts": int(getattr(data, "total_attempts", 0)),
            }
        )
    return payload


def _runtime_status_fields(backend: DevBackendContext) -> dict[str, Any]:
    config = backend.runtime.config
    system_proxy = _system_proxy_status_payload(backend)
    mitm_proxy = _diagnostic_mitm_status_payload(backend)
    proxy_payload = {
        "host": config.proxy.host,
        "port": config.proxy.port,
        "enabled": bool(config.proxy.enable_system_proxy),
        "systemProxyEnabled": bool(config.proxy.enable_system_proxy),
        **system_proxy,
        **mitm_proxy,
    }
    return {
        "proxy": proxy_payload,
        "config": {
            "autoSaveContent": True,
            "autoCleanTempFiles": bool(config.runtime.auto_clean_temp_files),
            "autoStartProxy": False,
            "enableSystemProxy": bool(config.proxy.enable_system_proxy),
            "logLevel": config.runtime.log_level,
            "proxyHost": config.proxy.host,
            "proxyPort": config.proxy.port,
            "requestIntervalSeconds": config.request.request_interval_seconds,
            "startupDelaySeconds": config.proxy.startup_delay_seconds,
            "verificationUrl": config.proxy.verification_url,
            "values": _custom_yaml_config_values(backend),
        },
        "workers": [],
        "dbPath": str(backend.db_path),
        "appStartedAt": backend.started_at.isoformat(timespec="seconds"),
        "uptimeSeconds": int((datetime.now() - backend.started_at).total_seconds()),
        "hardware": _hardware_status_payload(backend),
        "runtimeState": TaskRuntimeTracker.default_snapshot(
            proxy_status_label=_proxy_status_label(proxy_payload)
        ),
    }


def _hardware_status_payload(backend: DevBackendContext) -> dict[str, Any]:
    """读取当前 CPU / 物理内存占用，供运行状态区两行折线图使用。"""

    monitor = backend.system_resource_monitor
    if monitor is None:
        monitor = ProcessTreeResourceMonitor(history_sample_count=120, sample_interval_seconds=0.5)
        set_process_observer(monitor)
        start_monitor = getattr(monitor, "start", None)
        if callable(start_monitor):
            start_monitor()
        backend.system_resource_monitor = monitor
    try:
        snapshot = monitor.snapshot()
    except Exception:
        timestamp = datetime.now().isoformat(timespec="milliseconds")
        return {
            "cpuPercent": 0.0,
            "memoryPercent": 0.0,
            "memoryUsedBytes": 0,
            "cpuLabel": "0.0%",
            "memoryLabel": "0.0%",
            "processCount": 0,
            "historySampleCount": 120,
            "sampleIntervalSeconds": 0.5,
            "history": [],
            "updatedAt": timestamp,
        }
    return {
        "cpuPercent": float(snapshot.get("cpuPercent") or 0.0),
        "memoryPercent": float(snapshot.get("memoryPercent") or 0.0),
        "memoryUsedBytes": int(snapshot.get("memoryUsedBytes") or 0),
        "cpuLabel": str(snapshot.get("cpuLabel") or "0.0%"),
        "memoryLabel": str(snapshot.get("memoryLabel") or "0.0%"),
        "processCount": int(snapshot.get("processCount") or 0),
        "historySampleCount": int(snapshot.get("historySampleCount") or 120),
        "sampleIntervalSeconds": float(snapshot.get("sampleIntervalSeconds") or 0.5),
        "history": list(snapshot.get("history") or []),
        "updatedAt": str(snapshot.get("updatedAt") or datetime.now().isoformat(timespec="milliseconds")),
    }


def _snapshot_runtime_state(
    snapshot: Any,
    *,
    fallback: Any,
    proxy_status_label: str,
) -> dict[str, Any]:
    base = dict(fallback) if isinstance(fallback, dict) else TaskRuntimeTracker.default_snapshot(
        proxy_status_label=proxy_status_label
    )
    state = getattr(snapshot, "runtime_state", None)
    if isinstance(state, dict):
        base.update(state)
    elif hasattr(state, "snapshot"):
        try:
            raw_state = state.snapshot()
        except Exception:
            raw_state = None
        if isinstance(raw_state, dict):
            base.update(raw_state)
    base["proxyStatusLabel"] = proxy_status_label
    return base


def _proxy_status_label(proxy_payload: Any) -> str:
    if not isinstance(proxy_payload, dict):
        return "空闲"
    if bool(proxy_payload.get("mitmEnabled")):
        return "正在使用"
    if not bool(proxy_payload.get("mitmPortOccupied")):
        return "空闲"
    if str(proxy_payload.get("mitmPortOwner") or "").lower() == "application":
        return "正在使用"
    return "端口被占用"


HEALTH_CHECK_TARGETS = ("storage", "ca", "proxy-port", "https")


def _startup_self_check_service(backend: DevBackendContext) -> StartupSelfCheckService:
    return StartupSelfCheckService(
        project_root=backend.project_root,
        ca_status_checker=lambda _config: _startup_self_check_ca_status_payload(backend),
    )


def _startup_self_check_ca_status_payload(backend: DevBackendContext) -> dict[str, Any]:
    try:
        return _ca_certificate_status_payload(backend)
    except Exception as exc:
        append_log = getattr(backend, "append_log", None)
        if callable(append_log):
            append_log("WARN", f"启动自检读取 CA 状态失败：{exc}", source="startup_self_check")
        return {
            "ok": False,
            "caFileExists": False,
            "projectCertificateInstalled": False,
            "message": f"CA 状态未确认：{exc}",
        }


def _startup_self_check_status_payload(backend: DevBackendContext) -> dict[str, Any]:
    return _startup_self_check_service(backend).get_status(backend.runtime.config)


def _startup_self_check_run_payload(backend: DevBackendContext) -> dict[str, Any]:
    # 自检会读取数据库、目录和系统证书状态，串行执行避免重复点击导致结果互相覆盖。
    with backend.startup_self_check_lock:
        try:
            return _startup_self_check_service(backend).run(backend.runtime.config)
        except Exception as exc:
            append_log = getattr(backend, "append_log", None)
            if callable(append_log):
                append_log("ERROR", f"启动自检失败：{exc}", source="startup_self_check", exception=exc)
            return {
                "ok": False,
                "status": "failed",
                "fatalCount": 1,
                "warningCount": 0,
                "items": [
                    {
                        "key": "startup_self_check_error",
                        "group": "runtime",
                        "label": "启动自检",
                        "status": "failed",
                        "severity": "fatal",
                        "ok": False,
                        "message": f"{type(exc).__name__}: {exc}",
                        "action": "请查看运行日志中的 ERROR 详情。",
                    }
                ],
                "checkedAt": datetime.now().isoformat(timespec="seconds"),
                "durationSeconds": 0,
                "statePath": "data/runtime/startup_self_check.json",
            }


def _startup_health_payload(backend: DevBackendContext) -> dict[str, Any]:
    """串行执行一次启动健康检测，后续请求只返回进程内缓存。"""
    with backend.health_check_lock:
        if backend.health_startup_completed:
            return _health_summary_payload(backend, cached=True)

        for target in HEALTH_CHECK_TARGETS:
            backend.health_check_results[target] = _run_health_check(backend, target)
        backend.health_startup_completed = True
        return _health_summary_payload(backend, cached=False)


def _health_summary_payload(backend: DevBackendContext, *, cached: bool) -> dict[str, Any]:
    results = {
        target: backend.health_check_results[target]
        for target in HEALTH_CHECK_TARGETS
        if target in backend.health_check_results
    }
    return {
        "ok": len(results) == len(HEALTH_CHECK_TARGETS)
        and all(bool(result.get("ok")) for result in results.values()),
        "status": "completed",
        "cached": cached,
        "results": results,
        "checkedAt": datetime.now().isoformat(timespec="seconds"),
    }


def _run_health_check(backend: DevBackendContext, target: str) -> dict[str, Any]:
    checkers = {
        "storage": _storage_health_payload,
        "ca": _ca_health_payload,
        "proxy-port": _proxy_port_health_payload,
        "https": _https_health_payload,
    }
    try:
        return checkers[target](backend)
    except Exception as exc:
        backend.append_log("ERROR", f"健康检测失败（{target}）：{exc}")
        return _health_result(
            target=target,
            ok=False,
            label="检测失败",
            message=f"{type(exc).__name__}: {exc}",
        )


def _health_result(
    *,
    target: str,
    ok: bool,
    label: str,
    message: str,
    items: list[dict[str, Any]] | None = None,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    return {
        "ok": ok,
        "target": target,
        "tone": "success" if ok else "danger",
        "label": label,
        "message": message,
        "items": items or [],
        "details": details or {},
        "checkedAt": datetime.now().isoformat(timespec="seconds"),
    }


def _storage_health_payload(backend: DevBackendContext) -> dict[str, Any]:
    storage = backend.runtime.config.storage
    configured_paths = (
        ("article_storage_root", "文章数据目录", storage.article_storage_root),
        ("db_dir", "数据库目录", storage.db_dir),
        ("temp_dir", "临时目录", storage.temp_dir),
        ("log_dir", "日志目录", storage.log_dir),
    )
    items = [
        _directory_health_item(backend, key=key, label=label, configured_path=value)
        for key, label, value in configured_paths
    ]
    ok = all(bool(item["ok"]) for item in items)
    failed_labels = [str(item["label"]) for item in items if not item["ok"]]
    message = (
        "配置涉及的 4 个数据目录均可正常读取和写入。"
        if ok
        else f"以下目录读写异常：{'、'.join(failed_labels)}。"
    )
    return _health_result(
        target="storage",
        ok=ok,
        label="读写正常" if ok else "读写异常",
        message=message,
        items=items,
        details={"checkedCount": len(items), "failedCount": len(failed_labels)},
    )


def _directory_health_item(
    backend: DevBackendContext,
    *,
    key: str,
    label: str,
    configured_path: Any,
) -> dict[str, Any]:
    path = _resolve_project_path(backend, configured_path)
    exists = path.exists()
    is_directory = path.is_dir()
    readable = False
    writable = False
    errors: list[str] = []

    if is_directory:
        try:
            next(path.iterdir(), None)
            readable = True
        except OSError as exc:
            errors.append(f"读取失败：{exc}")

        probe_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                prefix=".awa-health-",
                dir=path,
                delete=False,
            ) as probe:
                probe.write(b"health-check")
                probe.flush()
                probe_path = Path(probe.name)
            writable = True
        except OSError as exc:
            errors.append(f"写入失败：{exc}")
        finally:
            if probe_path is not None:
                try:
                    probe_path.unlink(missing_ok=True)
                except OSError as exc:
                    writable = False
                    errors.append(f"清理探测文件失败：{exc}")
    elif not exists:
        errors.append("目录不存在")
    else:
        errors.append("配置路径不是目录")

    ok = exists and is_directory and readable and writable
    return {
        "key": key,
        "label": label,
        "path": str(configured_path),
        "pathAbsolute": str(path),
        "exists": exists,
        "isDirectory": is_directory,
        "readable": readable,
        "writable": writable,
        "ok": ok,
        "message": "读取、写入正常" if ok else "；".join(errors),
    }


def _ca_health_payload(backend: DevBackendContext) -> dict[str, Any]:
    status = _ca_certificate_status_payload(backend)
    project_certificate = status.get("projectCertificate")
    project_certificate = project_certificate if isinstance(project_certificate, dict) else {}
    certificates = status.get("certificates")
    certificates = certificates if isinstance(certificates, list) else []
    matching_certificate = next(
        (
            certificate
            for certificate in certificates
            if isinstance(certificate, dict) and certificate.get("matchesProject")
        ),
        {},
    )
    project_thumbprint = _normalize_certificate_thumbprint(project_certificate.get("thumbprint"))
    system_thumbprint = _normalize_certificate_thumbprint(matching_certificate.get("thumbprint"))
    installed = bool(status.get("ok") and status.get("projectCertificateInstalled"))
    items = [
        {
            "key": "project-certificate",
            "label": "项目证书",
            "value": status.get("currentCaRelativePath") or status.get("currentCaPath") or "未找到",
            "ok": bool(status.get("caFileExists")),
        },
        {
            "key": "project-thumbprint",
            "label": "项目指纹",
            "value": project_thumbprint or "未读取",
            "ok": bool(project_thumbprint),
        },
        {
            "key": "system-certificate",
            "label": "系统证书",
            "value": status.get("storePath") or "Cert:\\CurrentUser\\Root",
            "ok": installed,
        },
        {
            "key": "system-thumbprint",
            "label": "系统指纹",
            "value": system_thumbprint or "未找到相同指纹证书",
            "ok": installed,
        },
        {
            "key": "issuer",
            "label": "颁发者",
            "value": project_certificate.get("issuer") or "未知",
            "ok": bool(project_certificate.get("issuer")),
        },
        {
            "key": "validity",
            "label": "有效期",
            "value": (
                f"{project_certificate.get('notBefore') or '未知'} 至 "
                f"{project_certificate.get('notAfter') or '未知'}"
            ),
            "ok": bool(project_certificate.get("notBefore") and project_certificate.get("notAfter")),
        },
    ]
    return _health_result(
        target="ca",
        ok=installed,
        label="已安装" if installed else "未安装",
        message=str(status.get("message") or "CA 证书检测完成。"),
        items=items,
        details={
            "projectThumbprint": project_thumbprint,
            "systemThumbprint": system_thumbprint,
            "matchesProject": installed,
            "status": status.get("status"),
        },
    )


def _proxy_port_health_payload(backend: DevBackendContext) -> dict[str, Any]:
    host, port = _proxy_host_port(backend)
    application_owned = _application_owns_proxy_port(backend)
    occupied = True if application_owned else _port_is_occupied(backend, host, port)
    owner = "application" if application_owned else ("external" if occupied else "free")
    available = application_owned or not occupied
    owner_labels = {
        "application": "本程序 MITM 正在监听",
        "external": "被外部进程占用",
        "free": "端口当前空闲",
    }
    address = f"{host}:{port}"
    return _health_result(
        target="proxy-port",
        ok=available,
        label="可用" if available else "不可用",
        message=(
            f"代理端口可用：{address}（{owner_labels[owner]}）。"
            if available
            else f"代理端口不可用：{address} 已被外部进程占用。"
        ),
        items=[
            {"key": "address", "label": "监听地址", "value": address, "ok": True},
            {"key": "owner", "label": "端口状态", "value": owner_labels[owner], "ok": available},
        ],
        details={"host": host, "port": port, "occupied": occupied, "owner": owner},
    )


def _application_owns_proxy_port(backend: DevBackendContext) -> bool:
    if backend.diagnostic_mitm_listener is not None:
        return True
    if not backend.active_task_id:
        return False
    try:
        snapshot = backend.task_manager.get(backend.active_task_id)
    except Exception:
        return False
    return _frontend_task_status(getattr(snapshot, "status", "")) in {"starting", "running"}


def _runtime_cache_busy_reason(backend: DevBackendContext) -> str | None:
    """缓存目录可能被采集或诊断流程使用，运行期间不允许整体清理。"""
    if backend.active_task_id:
        try:
            snapshot = backend.task_manager.get(backend.active_task_id)
        except Exception:
            snapshot = None
        if snapshot is not None and _frontend_task_status(
            getattr(snapshot, "status", "")
        ) in {"starting", "running"}:
            return "采集任务正在运行"

    if backend.diagnostic_mitm_listener is not None:
        return "MITM 诊断正在运行"

    if backend.offline_cache_service is not None and backend.offline_cache_service.is_busy():
        return "离线缓存任务正在运行"

    if (
        backend.main_flow_foreground_huey_service is not None
        and backend.main_flow_foreground_huey_service.is_active()
    ):
        return "主服务前台任务正在运行"
    if (
        backend.main_flow_detail_huey_service is not None
        and backend.main_flow_detail_huey_service.is_active()
    ):
        return "主服务详情任务正在运行"

    if (
        backend.window_click_flow_huey_service is not None
        and backend.window_click_flow_huey_service.is_active()
    ):
        return "诊断任务正在运行"
    if (
        backend.article_detail_huey_service is not None
        and backend.article_detail_huey_service.is_active()
    ):
        return "诊断任务正在运行"
    if (
        backend.initial_content_storage_huey_service is not None
        and backend.initial_content_storage_huey_service.is_active()
    ):
        return "诊断任务正在运行"
    if (
        backend.article_detail_comments_huey_service is not None
        and backend.article_detail_comments_huey_service.is_active()
    ):
        return "诊断任务正在运行"
    if (
        backend.article_detail_offline_cache_huey_service is not None
        and backend.article_detail_offline_cache_huey_service.is_active()
    ):
        return "诊断任务正在运行"

    diagnostic_groups = (
    )
    for jobs, lock in diagnostic_groups:
        with lock:
            if any(
                str(job.get("status") or "").lower() in {"pending", "starting", "running"}
                for job in jobs.values()
            ):
                return "诊断任务正在运行"
    return None


def _https_health_payload(backend: DevBackendContext) -> dict[str, Any]:
    """按 MITM、系统代理、HTTPS、系统代理恢复、MITM 停止的顺序执行检测。"""
    steps: list[dict[str, Any]] = []
    failure_messages: list[str] = []
    probe_result: dict[str, Any] = {}
    had_mitm_listener = backend.diagnostic_mitm_listener is not None
    prior_diagnostic_snapshot = backend.diagnostic_system_proxy_snapshot
    prior_system_snapshot = _system_proxy_controller(backend).current()
    started_mitm = False
    system_proxy_enabled = False

    try:
        mitm_result = _set_diagnostic_mitm_payload(backend, enabled=True)
        mitm_ok = bool(mitm_result.get("ok"))
        steps.append({"key": "mitm-start", "label": "MITM 代理", "value": mitm_result.get("message", ""), "ok": mitm_ok})
        if not mitm_ok:
            failure_messages.append(str(mitm_result.get("message") or "MITM 代理启动失败"))
        else:
            started_mitm = not had_mitm_listener

            system_result = _set_system_proxy_payload(backend, enabled=True)
            system_ok = bool(system_result.get("ok"))
            steps.append({"key": "system-enable", "label": "系统代理", "value": system_result.get("message", ""), "ok": system_ok})
            if not system_ok:
                failure_messages.append(str(system_result.get("message") or "系统代理开启失败"))
            else:
                system_proxy_enabled = True
                probe_result = _proxy_connection_test_payload(backend)
                probe_ok = bool(probe_result.get("ok"))
                status_code = probe_result.get("statusCode")
                probe_value = f"HTTP {status_code}" if status_code is not None else str(probe_result.get("message") or "请求失败")
                steps.append({"key": "https-request", "label": "HTTPS 校验", "value": probe_value, "ok": probe_ok})
                if not probe_ok:
                    failure_messages.append(str(probe_result.get("message") or "HTTPS 校验失败"))
    finally:
        if system_proxy_enabled or (
            prior_diagnostic_snapshot is None
            and backend.diagnostic_system_proxy_snapshot is not None
        ):
            try:
                if prior_diagnostic_snapshot is None:
                    restore_result = _set_system_proxy_payload(backend, enabled=False)
                    restore_ok = bool(restore_result.get("ok"))
                    restore_message = str(restore_result.get("message") or "系统代理恢复完成")
                else:
                    _system_proxy_controller(backend).restore(prior_system_snapshot)
                    restore_ok = True
                    restore_message = (
                        "系统代理已恢复到检测前状态："
                        f"{_format_system_proxy_snapshot_for_message(prior_system_snapshot)}"
                    )
                steps.append({"key": "system-restore", "label": "系统代理恢复", "value": restore_message, "ok": restore_ok})
                if not restore_ok:
                    failure_messages.append(restore_message)
            except Exception as exc:
                message = f"系统代理恢复失败：{exc}"
                failure_messages.append(message)
                steps.append({"key": "system-restore", "label": "系统代理恢复", "value": message, "ok": False})

        if started_mitm:
            stop_result = _set_diagnostic_mitm_payload(backend, enabled=False)
            stop_ok = bool(stop_result.get("ok"))
            stop_message = str(stop_result.get("message") or "MITM 代理关闭完成")
            steps.append({"key": "mitm-stop", "label": "MITM 代理清理", "value": stop_message, "ok": stop_ok})
            if not stop_ok:
                failure_messages.append(stop_message)

    ok = bool(probe_result.get("ok")) and not failure_messages
    proxy = str(probe_result.get("proxy") or _configured_proxy_server(backend))
    url = str(probe_result.get("url") or backend.runtime.config.proxy.verification_url)
    status_code = probe_result.get("statusCode")
    items = [
        {"key": "proxy", "label": "代理地址", "value": proxy, "ok": ok},
        {"key": "url", "label": "验证地址", "value": url, "ok": bool(probe_result)},
        {
            "key": "http-status",
            "label": "HTTP 状态",
            "value": str(status_code) if status_code is not None else "未返回",
            "ok": bool(probe_result.get("ok")),
        },
        *steps,
    ]
    return _health_result(
        target="https",
        ok=ok,
        label="HTTPS 正常" if ok else "HTTPS 异常",
        message="HTTPS 代理校验通过，代理状态已恢复。" if ok else "；".join(failure_messages),
        items=items,
        details={"proxy": proxy, "url": url, "statusCode": status_code, "steps": steps},
    )


def _proxy_host_port(backend: DevBackendContext) -> tuple[str, int]:
    config = backend.runtime.config
    return str(config.proxy.host), int(config.proxy.port)


def _configured_proxy_server(backend: DevBackendContext) -> str:
    host, port = _proxy_host_port(backend)
    return f"{host}:{port}"


def _port_is_occupied(backend: DevBackendContext, host: str, port: int) -> bool:
    if callable(backend.port_checker):
        return bool(backend.port_checker(host, port))
    try:
        with socket.create_connection((host, int(port)), timeout=0.25):
            return True
    except OSError:
        return False


def _diagnostic_mitm_status_payload(backend: DevBackendContext) -> dict[str, Any]:
    host, port = _proxy_host_port(backend)
    enabled = backend.diagnostic_mitm_listener is not None
    occupied = True if enabled else _port_is_occupied(backend, host, port)
    owner = "diagnostic" if enabled else ("external" if occupied else "free")
    return {
        "mitmEnabled": enabled,
        "mitmPort": port,
        "mitmListenHost": host,
        "mitmListenAddress": f"{host}:{port}",
        "mitmPortOccupied": occupied,
        "mitmPortAvailable": enabled or not occupied,
        "mitmPortOwner": owner,
        "mitmStartedAt": backend.diagnostic_mitm_started_at if enabled else None,
    }


def _create_diagnostic_mitm_listener(backend: DevBackendContext) -> Any:
    config = backend.runtime.config
    host, port = _proxy_host_port(backend)
    listener_factory = backend.mitm_listener_factory or MitmproxyListener
    return listener_factory(
        host=host,
        port=port,
        confdir=_resolve_project_path(backend, config.proxy.confdir),
        ssl_insecure=bool(config.proxy.ssl_insecure),
        buffer=CaptureBuffer(
            task_id="diagnostic-mitm",
            attempt_id="diagnostic-mitm",
        ),
        ready_timeout_seconds=max(0.1, float(config.mitm_capture.ready_timeout_seconds)),
    )


def _set_diagnostic_mitm_payload(backend: DevBackendContext, *, enabled: bool) -> dict[str, Any]:
    host, port = _proxy_host_port(backend)
    action = "start-mitm" if enabled else "stop-mitm"
    if enabled:
        if backend.diagnostic_mitm_listener is not None:
            return {
                "ok": True,
                "status": "mitm-proxy-running",
                "action": action,
                "message": f"MITM 代理已在监听端口 {port}。",
                **_runtime_status_fields(backend),
            }
        if _port_is_occupied(backend, host, port):
            return {
                "ok": False,
                "status": "mitm-port-unavailable",
                "action": action,
                "message": f"代理端口不可用：{host}:{port} 已被占用。",
                **_runtime_status_fields(backend),
            }
        try:
            listener = _create_diagnostic_mitm_listener(backend)
            backend.diagnostic_mitm_started_at = float(listener.start())
            backend.diagnostic_mitm_listener = listener
        except (MitmproxyListenerError, OSError, RuntimeError, ValueError) as exc:
            backend.diagnostic_mitm_listener = None
            backend.diagnostic_mitm_started_at = 0.0
            backend.append_log("ERROR", f"MITM 代理启动失败：{exc}")
            return {
                "ok": False,
                "status": "mitm-proxy-start-failed",
                "action": action,
                "message": f"MITM 代理启动失败：{exc}",
                **_runtime_status_fields(backend),
            }
        message = f"MITM 代理已开启，监听端口：{port}"
        backend.append_log("INFO", message)
        return {
            "ok": True,
            "status": "mitm-proxy-started",
            "action": action,
            "message": message,
            **_runtime_status_fields(backend),
        }

    listener = backend.diagnostic_mitm_listener
    if listener is None:
        return {
            "ok": True,
            "status": "mitm-proxy-stopped",
            "action": action,
            "message": "MITM 代理当前未开启。",
            **_runtime_status_fields(backend),
        }
    try:
        listener.stop()
    except Exception as exc:
        backend.append_log("ERROR", f"MITM 代理关闭失败：{exc}")
        return {
            "ok": False,
            "status": "mitm-proxy-stop-failed",
            "action": action,
            "message": f"MITM 代理关闭失败：{exc}",
            **_runtime_status_fields(backend),
        }
    finally:
        backend.diagnostic_mitm_listener = None
        backend.diagnostic_mitm_started_at = 0.0

    message = f"MITM 代理已关闭，释放端口：{port}"
    backend.append_log("INFO", message)
    return {
        "ok": True,
        "status": "mitm-proxy-stopped",
        "action": action,
        "message": message,
        **_runtime_status_fields(backend),
    }


def _system_proxy_controller(backend: DevBackendContext) -> Any:
    if backend.system_proxy is None:
        backend.system_proxy = WindowsSystemProxy()
    return backend.system_proxy


def _system_proxy_status_payload(backend: DevBackendContext) -> dict[str, Any]:
    configured_server = _configured_proxy_server(backend)
    try:
        snapshot = _system_proxy_controller(backend).current()
    except Exception as exc:
        return {
            "systemProxyActive": False,
            "systemProxyReadable": False,
            "systemProxyServer": "",
            "systemProxyReadError": f"{type(exc).__name__}: {exc}",
            "configuredProxyServer": configured_server,
        }

    return {
        "systemProxyActive": bool(snapshot.enabled),
        "systemProxyReadable": True,
        "systemProxyServer": snapshot.server,
        "systemProxyBypass": snapshot.bypass,
        "systemProxyAutoConfigUrl": snapshot.auto_config_url,
        "configuredProxyServer": configured_server,
    }


def _format_system_proxy_snapshot_for_message(snapshot: ProxySnapshot) -> str:
    if not snapshot.enabled:
        return "已关闭"

    if snapshot.server.strip():
        return snapshot.server.strip()

    if snapshot.auto_config_url.strip():
        return f"PAC：{snapshot.auto_config_url.strip()}"

    return "已开启（未配置代理地址）"


def _set_system_proxy_payload(backend: DevBackendContext, *, enabled: bool) -> dict[str, Any]:
    controller = _system_proxy_controller(backend)
    try:
        previous = controller.current()
        target_server = _configured_proxy_server(backend)
        if enabled:
            # 诊断页手动接管系统代理前保存原状态，后续关闭时用于恢复用户原代理。
            if backend.diagnostic_system_proxy_snapshot is None:
                backend.diagnostic_system_proxy_snapshot = previous
            controller.enable(target_server)
            status = "system-proxy-enabled"
            message = f"系统代理已开启，当前代理：{target_server}"
            action = "enable-system-proxy"
        else:
            snapshot = backend.diagnostic_system_proxy_snapshot
            if snapshot is not None:
                if proxy_points_to(previous, target_server):
                    controller.restore(snapshot)
                    status = "system-proxy-restored"
                    message = f"系统代理已恢复到诊断开启前状态：{_format_system_proxy_snapshot_for_message(snapshot)}"
                else:
                    # 用户或其他程序已经改走系统代理时，不覆盖外部新状态。
                    status = "system-proxy-external-changed"
                    message = "系统代理已被外部修改，未覆盖当前系统代理。"
                backend.diagnostic_system_proxy_snapshot = None
            else:
                # 没有诊断快照时保留旧行为：只关闭开关，保留 ProxyServer / PAC 文本。
                controller.restore(
                    ProxySnapshot(
                        enabled=False,
                        server=previous.server,
                        bypass=previous.bypass,
                        auto_config_url=previous.auto_config_url,
                    )
                )
                status = "system-proxy-disabled"
                message = "系统代理已关闭。"
            action = "disable-system-proxy"
    except Exception as exc:
        backend.append_log("ERROR", f"系统代理修改失败：{exc}")
        return {
            "ok": False,
            "status": "system-proxy-failed",
            "action": "enable-system-proxy" if enabled else "disable-system-proxy",
            "message": f"系统代理修改失败：{exc}",
            **_runtime_status_fields(backend),
        }

    backend.append_log("INFO", message)
    return {
        "ok": True,
        "status": status,
        "action": action,
        "message": message,
        "previousSystemProxyEnabled": previous.enabled,
        "previousSystemProxyServer": previous.server,
        **_runtime_status_fields(backend),
    }


def _frontend_task_status(status: Any) -> str:
    value = str(getattr(status, "value", status)).lower()
    if value == TaskStatus.PENDING.value:
        return "starting"
    if value == TaskStatus.RUNNING.value:
        return "running"
    return value


def _environment_payload(config: Any) -> dict[str, str]:
    return {
        "appName": "Access WeChat Article",
        "appVersion": str(config.software.version),
        "systemLabel": f"{platform.system()} {platform.release()}",
        "pythonVersion": platform.python_version(),
        "mitmproxyVersion": _package_version("mitmproxy"),
        "playwrightVersion": _package_version("playwright"),
        "pywebviewVersion": _package_version("pywebview"),
    }


def _package_version(name: str) -> str:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return "未安装"


def _runtime_paths(backend: DevBackendContext) -> dict[str, Path]:
    config = backend.runtime.config
    return {
        "projectDir": backend.project_root,
        "outputDir": config.storage.article_storage_root,
        "storageDir": config.storage.article_storage_root,
        "logDir": config.storage.log_dir,
    }


def _select_runtime_directory(
    backend: DevBackendContext,
    *,
    config_key: str,
    current_path: str | None,
):
    if "." not in config_key:
        return JSONResponse(
            {
                "ok": False,
                "status": "invalid-key",
                "configKey": config_key,
                "message": "配置 key 必须使用 section.name 格式。",
            },
            status_code=400,
        )

    initial_dir = _existing_directory_for_config_path(backend.project_root, current_path)
    try:
        selected = _pick_directory(backend, initial_dir)
    except Exception as exc:
        return {
            "ok": False,
            "status": "unavailable",
            "configKey": config_key,
            "initialDir": str(initial_dir),
            "message": f"系统目录选择器不可用：{exc}",
        }

    if selected is None:
        return {
            "ok": False,
            "status": "cancelled",
            "configKey": config_key,
            "initialDir": str(initial_dir),
            "message": "已取消目录选择。",
        }

    selected_path = Path(selected).resolve()
    # 目录选择结果需要保留绝对路径，便于页面直接显示并让当前进程立即使用。
    selected_path_value = str(selected_path)
    _update_runtime_config_memory(backend, {"values": {config_key: selected_path_value}})
    return {
        "ok": True,
        "status": "selected",
        "configKey": config_key,
        "initialDir": str(initial_dir),
        "path": str(selected_path),
        "selectedPath": selected_path_value,
        "message": "目录已同步到当前进程内存，点击保存配置后写入 custom.yaml。",
        "taskStatus": _current_task_payload(backend),
    }


def _pick_directory(backend: DevBackendContext, initial_dir: Path) -> Path | None:
    if backend.directory_selector is not None:
        selected = backend.directory_selector(initial_dir)
        return Path(selected).resolve() if selected else None

    # 本地开发后端直接调用系统目录选择器；失败时向前端返回 unavailable，避免误写配置。
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        selected_dir = filedialog.askdirectory(
            initialdir=str(initial_dir),
            title="选择配置目录",
            mustexist=False,
        )
    finally:
        root.destroy()

    return Path(selected_dir).resolve() if selected_dir else None


def _existing_directory_for_config_path(project_root: Path, current_path: str | None) -> Path:
    raw_value = (current_path or "").strip()
    target = Path(raw_value) if raw_value else project_root
    if not target.is_absolute():
        target = project_root / target
    target = target.resolve()

    if target.is_file():
        return target.parent
    if target.is_dir():
        return target

    # 目录尚未创建时，从当前相对路径逐级回退到最近存在的父目录，保证选择器能打开。
    parent = target
    while parent != parent.parent:
        parent = parent.parent
        if parent.is_dir():
            return parent
    return project_root


def _display_config_path(project_root: Path, selected_path: Path) -> str:
    try:
        return selected_path.relative_to(project_root).as_posix()
    except ValueError:
        return str(selected_path)


def _update_runtime_config_memory(
    backend: DevBackendContext,
    payload: dict[str, Any],
) -> None:
    mapping = _current_runtime_config_mapping(backend)
    _merge_runtime_config_payload(mapping, payload)
    config = build_app_config(_mapping_for_app_config(mapping), project_root=backend.project_root)
    backend.config_mapping = mapping
    _replace_runtime_config(backend, config)


def _write_runtime_config_yaml(backend: DevBackendContext) -> Path:
    config_path = _runtime_config_path(backend)
    if config_path is None:
        raise RuntimeError("当前运行时没有可写入的 custom.yaml 路径")

    config_path.parent.mkdir(parents=True, exist_ok=True)
    yaml = YAML()
    yaml.default_flow_style = False
    yaml.indent(mapping=2, sequence=4, offset=2)
    with config_path.open("w", encoding="utf-8") as file:
        yaml.dump(backend.config_mapping or _current_runtime_config_mapping(backend), file)
    return config_path


def _runtime_config_path(backend: DevBackendContext) -> Path | None:
    config_path = getattr(getattr(backend.runtime, "config_service", None), "config_path", None)
    return Path(config_path) if config_path is not None else None


def _current_runtime_config_mapping(backend: DevBackendContext) -> dict[str, Any]:
    if backend.config_mapping is not None:
        return deepcopy(backend.config_mapping)

    config_path = _runtime_config_path(backend)
    if config_path is not None and config_path.is_file():
        loaded = YAML(typ="safe").load(config_path.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            return _plain_config_mapping(loaded)

    return _runtime_config_mapping(backend)


def _plain_config_mapping(mapping: dict[str, Any]) -> dict[str, Any]:
    plain: dict[str, Any] = {}
    for key, value in mapping.items():
        plain[str(key)] = _plain_config_mapping(value) if isinstance(value, dict) else value
    return plain


def _runtime_config_mapping(backend: DevBackendContext) -> dict[str, Any]:
    mapping: dict[str, Any] = {}
    for key, value in _flatten_runtime_config(backend.runtime.config).items():
        _set_nested_config_value(
            mapping,
            key,
            _config_value_for_yaml(backend.project_root, value),
        )
    return mapping


def _config_value_for_yaml(project_root: Path, value: Any) -> Any:
    if isinstance(value, Path):
        return _display_config_path(project_root, value.resolve())
    return value


def _merge_runtime_config_payload(mapping: dict[str, Any], payload: dict[str, Any]) -> None:
    values = payload.get("values")
    if isinstance(values, dict):
        for key, value in values.items():
            _set_config_value_from_payload(mapping, str(key), value)

    legacy_values = {
        "runtime.auto_clean_temp_files": payload.get("autoCleanTempFiles"),
        "runtime.log_level": payload.get("logLevel"),
        "proxy.enable_system_proxy": payload.get("enableSystemProxy"),
        "request.request_interval_seconds": payload.get("requestIntervalSeconds"),
    }
    proxy = payload.get("proxy")
    if isinstance(proxy, dict):
        legacy_values.update(
            {
                "proxy.host": proxy.get("host"),
                "proxy.port": proxy.get("port"),
                "proxy.startup_delay_seconds": proxy.get("startupDelaySeconds"),
                "proxy.verification_url": proxy.get("verificationUrl"),
            }
        )

    for key, value in legacy_values.items():
        if value is not None:
            _set_config_value_from_payload(mapping, key, value)


def _set_config_value_from_payload(mapping: dict[str, Any], flat_key: str, value: Any) -> None:
    current_value = _get_nested_config_value(mapping, flat_key)
    _set_nested_config_value(mapping, flat_key, _coerce_config_value(flat_key, value, current_value))


def _get_nested_config_value(mapping: dict[str, Any], flat_key: str) -> Any:
    current: Any = mapping
    for part in flat_key.split("."):
        if not isinstance(current, dict) or part not in current:
            return None
        current = current[part]
    return current


def _set_nested_config_value(mapping: dict[str, Any], flat_key: str, value: Any) -> None:
    parts = flat_key.split(".")
    current = mapping
    for part in parts[:-1]:
        nested = current.setdefault(part, {})
        if not isinstance(nested, dict):
            nested = {}
            current[part] = nested
        current = nested
    current[parts[-1]] = value


def _coerce_config_value(flat_key: str, value: Any, current_value: Any) -> Any:
    if isinstance(value, str):
        normalized = value.strip()
        # 数值配置优先按字段类型转换，避免字符串 "0"/"1" 被误判为布尔值。
        if flat_key in _integer_config_keys():
            return int(float(normalized))
        if flat_key in _float_config_keys():
            return float(normalized)
        if isinstance(current_value, int) and not isinstance(current_value, bool):
            return int(float(normalized))
        if isinstance(current_value, float):
            return float(normalized)
        if normalized in {"开启", "开", "true", "True", "TRUE", "1", "yes", "on"}:
            return True
        if normalized in {"关闭", "关", "false", "False", "FALSE", "0", "no", "off"}:
            return False
        return normalized

    if isinstance(current_value, bool):
        return bool(value)
    if isinstance(current_value, int) and not isinstance(current_value, bool):
        return int(value)
    if isinstance(current_value, float):
        return float(value)
    return value


def _integer_config_keys() -> set[str]:
    return {
        "proxy.port",
        "basic_settings.proxy_settings.port",
        "comment.max_pages",
        "comment.max_concurrent_processes",
        "offline_cache.max_concurrent_processes",
        "runtime.temp_retention_days",
        "runtime.log_retention_days",
        "window.scroll_wheel_steps",
        "window.date_seek_max_steps",
        "window.bounce_attempts",
        "window.bounce_up_steps",
        "window.bounce_down_steps",
        "single_article_task.comment_collection.top_level_max_pages",
        "single_article_task.comment_collection.max_concurrent_processes",
        "single_article_task.offline_cache.max_concurrent_processes",
        "main_flow.home_scroll.bounce_attempts",
        "main_flow.home_scroll.bounce_up_steps",
        "main_flow.home_scroll.bounce_down_steps",
    }


def _float_config_keys() -> set[str]:
    return {
        "proxy.startup_delay_seconds",
        "basic_settings.proxy_settings.startup_delay_seconds",
        "request.request_interval_seconds",
        "request.request_timeout_seconds",
        "reference.request_timeout_seconds",
        "mitm_capture.ready_timeout_seconds",
        "mitm_capture.capture_timeout_seconds",
        "mitm_capture.result_timeout_seconds",
        "mitm_capture.listener_shutdown_timeout_seconds",
        "comment.request_timeout_seconds",
        "comment.page_interval_seconds",
        "offline_cache.max_scroll_seconds",
        "offline_cache.resource_timeout_seconds",
        "window.activation_wait_seconds",
        "window.home_find_timeout_seconds",
        "window.article_open_timeout_seconds",
        "window.article_title_poll_interval_seconds",
        "window.article_title_stable_delay_seconds",
        "window.article_close_confirm_timeout_seconds",
        "window.scroll_initial_delay_seconds",
        "window.scroll_probe_interval_seconds",
        "window.scroll_probe_max_interval_seconds",
        "window.lazy_load_timeout_seconds",
        "window.unchanged_before_bounce_seconds",
        "window.bounce_pause_seconds",
        "main_flow.home_window.activation_wait_seconds",
        "main_flow.home_window.home_find_timeout_seconds",
        "main_flow.home_scroll.scroll_initial_delay_seconds",
        "main_flow.home_scroll.scroll_probe_interval_seconds",
        "main_flow.home_scroll.scroll_probe_max_interval_seconds",
        "main_flow.home_scroll.lazy_load_timeout_seconds",
        "main_flow.home_scroll.unchanged_before_bounce_seconds",
        "main_flow.home_scroll.bounce_pause_seconds",
        "main_flow.dispatch_control.single_task_interval_seconds",
        "single_article_task.article_tab.article_open_timeout_seconds",
        "single_article_task.article_tab.article_title_poll_interval_seconds",
        "single_article_task.article_tab.article_title_stable_delay_seconds",
        "single_article_task.article_tab.article_close_confirm_timeout_seconds",
        "single_article_task.mitm_capture.ready_timeout_seconds",
        "single_article_task.mitm_capture.capture_timeout_seconds",
        "single_article_task.mitm_capture.result_timeout_seconds",
        "single_article_task.mitm_capture.listener_shutdown_timeout_seconds",
        "single_article_task.html_storage.request_timeout_seconds",
        "single_article_task.comment_collection.request_timeout_seconds",
        "single_article_task.comment_collection.page_interval_seconds",
        "single_article_task.offline_cache.max_scroll_seconds",
        "single_article_task.offline_cache.resource_timeout_seconds",
    }


def _mapping_for_app_config(mapping: dict[str, Any]) -> dict[str, Any]:
    app_mapping = deepcopy(mapping)
    reference = app_mapping.get("reference")
    if isinstance(reference, dict) and "request_timeout_seconds" in reference:
        request = app_mapping.setdefault("request", {})
        if isinstance(request, dict):
            request["request_timeout_seconds"] = reference["request_timeout_seconds"]
    return app_mapping


def _replace_runtime_config(backend: DevBackendContext, config: Any) -> None:
    config_service = getattr(backend.runtime, "config_service", None)
    if config_service is not None:
        lock = getattr(config_service, "_lock", None)
        if lock is not None:
            with lock:
                setattr(config_service, "_current", config)
        elif hasattr(config_service, "_current"):
            setattr(config_service, "_current", config)

    if _can_rebuild_application_runtime(backend.runtime):
        backend.runtime = build_application_runtime(
            config_service=config_service,
            config=config,
        )
        if backend.active_task_id is None:
            backend.task_manager = create_capture_task_manager(
                runtime=backend.runtime,
                db_path=backend.db_path,
                runtime_logger=backend.runtime_logger,
            )
        return

    # 单元测试会传入轻量 runtime；这里仍更新 config，保持接口行为可验证。
    setattr(backend.runtime, "config", config)


def _can_rebuild_application_runtime(runtime: Any) -> bool:
    return all(
        hasattr(runtime, name)
        for name in (
            "config_service",
            "window_factory",
            "capture_factory",
            "single_capture_settings",
        )
    )


def _custom_yaml_config_values(backend: DevBackendContext) -> dict[str, str]:
    """读取启动时 YAML 配置值，供系统配置页按 configKey 动态展示。"""
    if backend.config_mapping is not None:
        return {
            key: _format_config_display_value(value)
            for key, value in _flatten_config_mapping(backend.config_mapping).items()
        }

    config_path = getattr(getattr(backend.runtime, "config_service", None), "config_path", None)
    if config_path is not None and Path(config_path).is_file():
        yaml = YAML(typ="safe")
        raw_config = yaml.load(Path(config_path).read_text(encoding="utf-8"))
        if isinstance(raw_config, dict):
            return {
                key: _format_config_display_value(value)
                for key, value in _flatten_config_mapping(raw_config).items()
            }

    return {
        key: _format_config_display_value(value)
        for key, value in _flatten_runtime_config(backend.runtime.config).items()
    }


def _flatten_config_mapping(
    mapping: dict[str, Any],
    *,
    prefix: str = "",
) -> dict[str, Any]:
    flattened: dict[str, Any] = {}
    for key, value in mapping.items():
        path = f"{prefix}.{key}" if prefix else str(key)
        if isinstance(value, dict):
            flattened.update(_flatten_config_mapping(value, prefix=path))
            continue
        flattened[path] = value
    return flattened


def _flatten_runtime_config(config: Any) -> dict[str, Any]:
    sections = (
        "software",
        "storage",
        "proxy",
        "mitm_capture",
        "window",
        "request",
        "comment",
        "offline_cache",
        "runtime",
    )
    flattened: dict[str, Any] = {}
    for section in sections:
        section_value = getattr(config, section, None)
        if section_value is None:
            continue
        for name in dir(section_value):
            if name.startswith("_"):
                continue
            value = getattr(section_value, name)
            if callable(value):
                continue
            flattened[f"{section}.{name}"] = value
    return flattened


def _format_config_display_value(value: Any) -> str:
    if isinstance(value, bool):
        return "开启" if value else "关闭"
    if isinstance(value, Path):
        return value.as_posix()
    if value is None:
        return ""
    return str(value)


def _archive_summary_payload(backend: DevBackendContext) -> dict[str, Any]:
    with sqlite_connection(backend.db_path, write=False) as connection:
        account_count = connection.execute(
            "SELECT COUNT(*) FROM awa_public_accounts"
        ).fetchone()[0]
        article_count = connection.execute(
            "SELECT COUNT(*) FROM awa_public_articles"
        ).fetchone()[0]
    storage_root = Path(backend.runtime.config.storage.article_storage_root)
    size_bytes = _directory_size_bytes(storage_root)
    return {
        "ok": True,
        "status": "ok",
        "accountCount": int(account_count),
        "articleCount": int(article_count),
        "dataType": "SQLite + JSON",
        "storageSizeBytes": size_bytes,
        "storageSizeLabel": _format_size(size_bytes),
        "storageRoot": str(storage_root),
        "dbPath": str(backend.db_path),
    }


def _archive_account_items(backend: DevBackendContext) -> list[dict[str, Any]]:
    with sqlite_connection(backend.db_path, write=False) as connection:
        rows = connection.execute(
            """
            SELECT
                a.id,
                a.account_name,
                a.created_time,
                a.updated_time,
                COUNT(DISTINCT ar.id) AS article_count,
                COALESCE(MAX(ar.last_collected_time), '') AS latest_collect_time
            FROM awa_public_accounts a
            LEFT JOIN awa_public_articles ar ON ar.account_id = a.id
            GROUP BY a.id
            ORDER BY latest_collect_time DESC, a.updated_time DESC, a.id DESC
            """
        ).fetchall()
    return [
        {
            "id": int(row["id"]),
            "accountName": str(row["account_name"]),
            "createdTime": str(row["created_time"]),
            "updatedTime": str(row["updated_time"]),
            "latestCollectTime": str(row["latest_collect_time"] or ""),
            "articleCount": int(row["article_count"] or 0),
            "savedCount": int(row["article_count"] or 0),
            "failedCount": 0,
            "sizeLabel": f"{int(row['article_count'] or 0)} 条",
        }
        for row in rows
    ]


def _archive_account_articles_payload(
    backend: DevBackendContext,
    account_id: int,
    page: int,
    page_size: int,
) -> dict[str, Any]:
    safe_page_size = max(1, min(int(page_size), 100))
    with sqlite_connection(backend.db_path, write=False) as connection:
        total = int(
            connection.execute(
                "SELECT COUNT(*) FROM awa_public_articles WHERE account_id = ?",
                (int(account_id),),
            ).fetchone()[0]
        )
        total_pages = max(1, (total + safe_page_size - 1) // safe_page_size)
        safe_page = max(1, min(int(page), total_pages))
        rows = connection.execute(
            """
            SELECT *
            FROM awa_public_articles
            WHERE account_id = ?
            ORDER BY published_article_time DESC, last_collected_time DESC, id DESC
            LIMIT ? OFFSET ?
            """,
            (int(account_id), safe_page_size, (safe_page - 1) * safe_page_size),
        ).fetchall()
    storage_root = Path(backend.runtime.config.storage.article_storage_root)
    return {
        "ok": True,
        "status": "ok",
        "accountId": int(account_id),
        "page": safe_page,
        "pageSize": safe_page_size,
        "items": [_archive_article_item(row, storage_root) for row in rows],
        "total": total,
        "dbPath": str(backend.db_path),
    }


def _archive_article_item(row: Any, storage_root: Path) -> dict[str, Any]:
    archive_dir = str(row["archive_dir"] or "")
    resolved = _resolve_archive_dir(storage_root, archive_dir)
    size_bytes = _directory_size_bytes(resolved) if resolved else 0
    return {
        "id": int(row["id"]),
        "accountId": int(row["account_id"]),
        "title": str(row["article_title"] or "未命名文章"),
        "publishedArticleTime": str(row["published_article_time"] or ""),
        "articleLink": str(row["article_link"] or ""),
        "recordType": "文章详情",
        "collectTime": str(row["last_collected_time"] or ""),
        "durationSeconds": 0,
        "collectStatus": "saved",
        "statusLabel": "已保存",
        "archiveDir": str(resolved) if resolved else "",
        "archiveDirs": [str(resolved)] if resolved and resolved.exists() else [],
        "sizeBytes": size_bytes,
        "sizeLabel": _format_size(size_bytes),
    }


def _open_archive_article_directory(backend: DevBackendContext, article_id: int):
    storage_root = Path(backend.runtime.config.storage.article_storage_root).resolve()
    with sqlite_connection(backend.db_path, write=False) as connection:
        row = connection.execute(
            "SELECT archive_dir FROM awa_public_articles WHERE id = ?",
            (int(article_id),),
        ).fetchone()
    if row is None:
        return JSONResponse(
            {
                "ok": False,
                "status": "not-found",
                "articleId": int(article_id),
                "message": "未找到对应文章记录。",
            },
            status_code=404,
        )
    archive_dir = _resolve_archive_dir(storage_root, str(row["archive_dir"] or ""))
    if archive_dir is None or not archive_dir.exists():
        return JSONResponse(
            {
                "ok": False,
                "status": "missing",
                "articleId": int(article_id),
                "message": "该文章没有可打开的本地归档目录。",
            },
            status_code=404,
        )
    opened = _open_directory(archive_dir)
    return {
        "ok": opened,
        "status": "opened" if opened else "open-failed",
        "articleId": int(article_id),
        "path": str(archive_dir),
        "message": "已打开文章归档目录。" if opened else "打开文章归档目录失败。",
    }


def _register_archive_cache_routes(app: FastAPI, backend: Any) -> None:
    @app.post("/api/archive/cache/articles")
    def cache_archive_articles(payload: ArchiveCacheArticlesPayload):
        try:
            job = backend.offline_cache_service.create_articles_job(payload.articleIds)
        except OfflineCacheConflictError as exc:
            return JSONResponse(
                {"ok": False, "status": "conflict", "message": str(exc)},
                status_code=409,
            )
        return job.to_payload()

    @app.post("/api/archive/accounts/{account_id}/cache")
    def cache_archive_account(account_id: int):
        try:
            job = backend.offline_cache_service.create_account_job(account_id)
        except OfflineCacheConflictError as exc:
            return JSONResponse(
                {"ok": False, "status": "conflict", "message": str(exc)},
                status_code=409,
            )
        return job.to_payload()

    @app.get("/api/archive/cache/jobs/{job_id}")
    def get_archive_cache_job(job_id: str):
        try:
            return backend.offline_cache_service.get_job(job_id).to_payload()
        except KeyError:
            return JSONResponse(
                {
                    "ok": False,
                    "jobId": job_id,
                    "status": "missing",
                    "total": 0,
                    "finished": 0,
                    "running": 0,
                    "skipped": 0,
                    "requestedTotal": 0,
                    "processed": 0,
                    "queued": 0,
                    "failed": 0,
                    "activeProcesses": [],
                    "concurrency": 0,
                    "results": [],
                    "message": "离线缓存任务不存在。",
                },
                status_code=404,
            )


def _ca_certificate_status_payload(backend: DevBackendContext) -> dict[str, Any]:
    config = backend.runtime.config
    ca_path = _resolve_project_path(backend, config.proxy.ca_cert_path)
    list_payload = _list_mitm_ca_certificates_payload(backend)
    certificates = list_payload.get("certificates", [])
    project_certificate = list_payload.get("projectCertificate")
    store_count = len(certificates) if isinstance(certificates, list) else 0
    ca_file_exists = ca_path.is_file()
    project_thumbprint = _normalize_certificate_thumbprint(
        project_certificate.get("thumbprint") if isinstance(project_certificate, dict) else ""
    )
    normalized_certificates = [
        {
            **certificate,
            "matchesProject": bool(
                project_thumbprint
                and _normalize_certificate_thumbprint(certificate.get("thumbprint")) == project_thumbprint
            ),
        }
        for certificate in certificates
        if isinstance(certificate, dict)
    ]
    installed = any(certificate["matchesProject"] for certificate in normalized_certificates)

    if installed:
        status = "installed"
        label = "CA 证书已安装"
        message = "项目 CA 证书已按指纹安装到当前用户根证书库。"
    elif ca_file_exists:
        status = "file-present"
        label = "CA 文件存在，证书未安装"
        message = "已找到项目 CA 证书文件，但系统证书库中没有相同指纹的证书。"
    else:
        status = "missing"
        label = "CA 证书文件不存在"
        message = f"未找到 CA 证书文件：{ca_path}"

    return {
        "ok": bool(list_payload.get("ok", True)),
        "status": status,
        "installed": installed,
        "label": label,
        "message": message if list_payload.get("ok", True) else list_payload.get("message", message),
        "currentCaPath": str(ca_path),
        "currentCaRelativePath": _project_relative_path(backend, ca_path),
        "caFileExists": ca_file_exists,
        "storePath": "Cert:\\CurrentUser\\Root",
        "storeCertificateCount": store_count,
        "projectCertificate": project_certificate,
        "projectCertificateInstalled": installed,
        "certificates": normalized_certificates,
    }


def _list_mitm_ca_certificates_payload(backend: DevBackendContext) -> dict[str, Any]:
    ca_path = _resolve_project_path(backend, backend.runtime.config.proxy.ca_cert_path)
    relative_ca_path = _project_relative_path(backend, ca_path)
    if backend.command_runner is None and platform.system() != "Windows":
        return {
            "ok": False,
            "status": "unsupported",
            "count": 0,
            "certificates": [],
            "projectCertificate": None,
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "message": "MITM CA 证书库检测仅支持 Windows 当前用户证书库。",
        }

    project_path_literal = _powershell_single_quote(str(ca_path))
    relative_path_literal = _powershell_single_quote(relative_ca_path)
    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        (
            f"$projectPath = '{project_path_literal}'; $projectCertificate = $null; "
            "if (Test-Path -LiteralPath $projectPath -PathType Leaf) { "
            "try { $project = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($projectPath); "
            "$projectCertificate = [pscustomobject]@{ path='"
            f"{relative_path_literal}"
            "'; thumbprint=$project.Thumbprint; subject=$project.Subject; issuer=$project.Issuer; "
            "friendlyName=$project.FriendlyName; notBefore=$project.NotBefore.ToString('yyyy-MM-dd HH:mm:ss'); "
            "notAfter=$project.NotAfter.ToString('yyyy-MM-dd HH:mm:ss') }; } catch {} }; "
            "$projectThumbprint = if ($projectCertificate) { (($projectCertificate.thumbprint -replace '\\s','').ToUpperInvariant()) } else { '' }; "
            "$certs = @(Get-ChildItem Cert:\\CurrentUser\\Root | Where-Object { "
            "($_.Subject -like '*mitmproxy*' -or $_.FriendlyName -like '*mitmproxy*' -or $_.Issuer -like '*mitmproxy*') "
            "-or ($projectThumbprint -and ((($_.Thumbprint -replace '\\s','').ToUpperInvariant()) -eq $projectThumbprint)) } | "
            "Select-Object "
            "@{Name='storePath';Expression={'Cert:\\CurrentUser\\Root'}},"
            "@{Name='thumbprint';Expression={(($_.Thumbprint -replace '\\s','').ToUpperInvariant())}},"
            "@{Name='subject';Expression={$_.Subject}},"
            "@{Name='issuer';Expression={$_.Issuer}},"
            "@{Name='friendlyName';Expression={$_.FriendlyName}},"
            "@{Name='notBefore';Expression={$_.NotBefore.ToString('yyyy-MM-dd HH:mm:ss')}},"
            "@{Name='notAfter';Expression={$_.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')}}); "
            "[pscustomobject]@{ projectCertificate=$projectCertificate; certificates=$certs } | ConvertTo-Json -Depth 5"
        ),
    ]

    try:
        completed = _run_diagnostic_command(backend, command, timeout=12)
    except Exception as exc:
        return {
            "ok": False,
            "status": "list-failed",
            "count": 0,
            "certificates": [],
            "projectCertificate": None,
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "message": f"检索当前用户根证书库失败：{exc}",
        }

    returncode = int(getattr(completed, "returncode", 0) or 0)
    stdout = str(getattr(completed, "stdout", "") or "")
    stderr = str(getattr(completed, "stderr", "") or "")
    if returncode != 0:
        return {
            "ok": False,
            "status": "list-failed",
            "count": 0,
            "certificates": [],
            "projectCertificate": None,
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "message": stderr.strip() or "检索当前用户根证书库失败。",
        }

    parsed = _parse_mitm_certificate_payload(stdout)
    certificates = parsed["certificates"]
    return {
        "ok": True,
        "status": "ok",
        "count": len(certificates),
        "certificates": certificates,
        "projectCertificate": parsed["projectCertificate"],
        "currentCaPath": str(ca_path),
        "currentCaRelativePath": relative_ca_path,
        "message": f"检索到 {len(certificates)} 张 mitmproxy 相关证书。" if certificates else "未检索到 mitmproxy 相关证书。",
    }


def _install_ca_certificate_payload(backend: DevBackendContext) -> dict[str, Any]:
    ca_path = _resolve_project_path(backend, backend.runtime.config.proxy.ca_cert_path)
    relative_ca_path = _project_relative_path(backend, ca_path)
    if not ca_path.is_file():
        return {
            "ok": False,
            "status": "missing",
            "installed": False,
            "label": "CA 证书文件不存在",
            "message": f"无法安装，未找到 CA 证书文件：{ca_path}",
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "storePath": "Cert:\\CurrentUser\\Root",
        }

    current_status = _ca_certificate_status_payload(backend)
    if current_status.get("projectCertificateInstalled"):
        current_status.update(
            {
                "ok": True,
                "status": "already-installed",
                "installed": True,
                "label": "CA 证书已安装",
                "message": "项目 CA 证书已存在于当前用户根证书库，无需重复安装。",
                "currentCaPath": str(ca_path),
                "currentCaRelativePath": relative_ca_path,
                "storePath": "Cert:\\CurrentUser\\Root",
            }
        )
        return current_status

    command = ["certutil", "-user", "-addstore", "Root", str(ca_path)]
    try:
        completed = _run_diagnostic_command(backend, command, timeout=20)
    except Exception as exc:
        return {
            "ok": False,
            "status": "install-failed",
            "installed": False,
            "label": "CA 证书安装失败",
            "message": f"安装 CA 证书失败：{exc}",
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "storePath": "Cert:\\CurrentUser\\Root",
        }

    returncode = int(getattr(completed, "returncode", 0) or 0)
    stderr = str(getattr(completed, "stderr", "") or "").strip()
    if returncode != 0:
        return {
            "ok": False,
            "status": "install-failed",
            "installed": False,
            "label": "CA 证书安装失败",
            "message": stderr or "certutil 安装 CA 证书失败。",
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "storePath": "Cert:\\CurrentUser\\Root",
        }

    status_payload = _ca_certificate_status_payload(backend)
    status_payload.update(
        {
            "ok": True,
            "status": "installed",
            "installed": True,
            "label": "CA 证书已安装",
            "message": "CA 证书已安装到当前用户根证书库。",
            "currentCaPath": str(ca_path),
            "currentCaRelativePath": relative_ca_path,
            "storePath": "Cert:\\CurrentUser\\Root",
        }
    )
    return status_payload


def _delete_mitm_ca_certificates_payload(backend: DevBackendContext, payload: dict[str, Any]) -> dict[str, Any]:
    raw_thumbprints = payload.get("thumbprints") if isinstance(payload, dict) else None
    thumbprints = _normalize_thumbprints(raw_thumbprints)
    if not thumbprints:
        return {
            "ok": False,
            "status": "missing-thumbprints",
            "deletedCount": 0,
            "skippedCount": 0,
            "deleted": [],
            "skipped": [],
            "message": "未收到需要删除的 MITM 证书指纹。",
        }

    deleted: list[dict[str, str]] = []
    skipped: list[dict[str, str]] = []
    for thumbprint in thumbprints:
        command = ["certutil", "-user", "-delstore", "Root", thumbprint]
        try:
            completed = _run_diagnostic_command(backend, command, timeout=20)
        except Exception as exc:
            skipped.append({"thumbprint": thumbprint, "reason": str(exc)})
            continue

        returncode = int(getattr(completed, "returncode", 0) or 0)
        stderr = str(getattr(completed, "stderr", "") or "").strip()
        if returncode == 0:
            deleted.append({"thumbprint": thumbprint, "storePath": "Cert:\\CurrentUser\\Root"})
        else:
            skipped.append({"thumbprint": thumbprint, "reason": stderr or "certutil 删除失败"})

    ok = not skipped
    result = {
        "ok": ok,
        "status": "deleted" if ok else "partial-failed",
        "deletedCount": len(deleted),
        "skippedCount": len(skipped),
        "deleted": deleted,
        "skipped": skipped,
        "message": f"已删除 {len(deleted)} 张 MITM 证书。" if ok else f"已删除 {len(deleted)} 张，{len(skipped)} 张删除失败。",
    }
    refreshed = _list_mitm_ca_certificates_payload(backend)
    result["remainingCertificates"] = refreshed.get("certificates", [])
    result["remainingCertificateCount"] = len(result["remainingCertificates"])
    return result


def _proxy_connection_test_payload(backend: DevBackendContext) -> dict[str, Any]:
    config = backend.runtime.config
    if callable(backend.proxy_tester):
        result = backend.proxy_tester(config)
        if isinstance(result, dict):
            return result
        return {
            "ok": False,
            "status": "invalid-result",
            "message": "代理检测器返回了无法识别的结果。",
        }

    host = str(config.proxy.host)
    port = int(config.proxy.port)
    url = str(config.proxy.verification_url)
    proxy = f"{host}:{port}"
    proxy_url = f"http://{proxy}"

    try:
        import requests

        response = requests.get(
            url,
            headers=build_chrome_document_headers(),
            timeout=8,
            proxies={"http": proxy_url, "https": proxy_url},
            verify=not bool(getattr(config.proxy, "ssl_insecure", False)),
        )
        data = bytes(getattr(response, "content", b"") or b"")[:4096]
        status_code = int(getattr(response, "status_code", 0) or 0)
        return {
            "ok": 200 <= status_code < 400,
            "status": "ok" if 200 <= status_code < 400 else "http-error",
            "message": "代理连接测试通过。" if 200 <= status_code < 400 else f"代理请求返回 HTTP {status_code}。",
            "url": url,
            "proxy": proxy,
            "statusCode": status_code,
            "bytesRead": len(data),
        }
    except Exception as exc:
        return {
            "ok": False,
            "status": "failed",
            "message": f"代理连接测试失败：{exc}",
            "url": url,
            "proxy": proxy,
            "statusCode": None,
            "bytesRead": 0,
        }


def _diagnostic_managed_result(backend: DevBackendContext, *, action: str, message: str) -> dict[str, Any]:
    return {
        "ok": False,
        "status": "managed-by-capture",
        "action": action,
        "message": message,
        **_runtime_status_fields(backend),
    }


def _window_diagnostic_payload(
    backend: DevBackendContext,
    action: str,
    *,
    scroll_steps: int | None = None,
) -> dict[str, Any]:
    if callable(backend.window_diagnostic_runner):
        result = backend.window_diagnostic_runner(action, backend)
    else:
        window_factory = getattr(backend.runtime, "window_factory", None)
        if window_factory is None:
            window_factory = WindowRuntimeFactory(backend.runtime.config)
        result = WindowDiagnosticService(
            config=backend.runtime.config,
            window_factory=window_factory,
        ).run(action, scroll_steps=scroll_steps)
    if not isinstance(result, dict):
        raise RuntimeError("窗口诊断返回了无法识别的结果")
    result.setdefault("action", action)
    result.setdefault("title", "窗口诊断结果")
    result.setdefault("message", "")
    result.setdefault("items", [])
    result.setdefault("tone", "success" if result.get("ok") else "error")
    return result


def _start_window_click_flow_diagnostic_job(
    backend: DevBackendContext,
    options: WindowClickFlowDiagnosticPayload | None = None,
) -> dict[str, Any] | JSONResponse:
    options = options or WindowClickFlowDiagnosticPayload()
    service = _window_click_flow_huey_service(backend)
    try:
        return service.start(
            max_records=options.maxRecords,
            date_filter_mode=options.dateFilterMode,
            start_date=options.startDate,
            end_date=options.endDate,
        )
    except WindowClickFlowConflictError as exc:
        return JSONResponse(
            {
                "ok": False,
                "status": "conflict",
                "action": "window-click-flow",
                "title": "主页内容读取结果",
                "message": str(exc),
                "tone": "warning",
                "items": [],
            },
            status_code=409,
        )


def _window_click_flow_diagnostic_job_payload(
    backend: DevBackendContext,
    job_id: str,
) -> dict[str, Any] | JSONResponse:
    try:
        return _window_click_flow_huey_service(backend).get(job_id)
    except KeyError:
        return _missing_window_click_flow_job_response(
            job_id,
            message="未找到该主页内容读取诊断记录。",
        )


def _stop_window_click_flow_diagnostic_job(
    backend: DevBackendContext,
    job_id: str,
) -> dict[str, Any] | JSONResponse:
    try:
        return _window_click_flow_huey_service(backend).stop(job_id)
    except KeyError:
        return _missing_window_click_flow_job_response(
            job_id,
            message="未找到该主页内容读取诊断记录，无法停止。",
        )


def _window_click_flow_huey_service(backend: DevBackendContext) -> Any:
    service = backend.window_click_flow_huey_service
    if service is None:
        raise RuntimeError("窗口读取诊断Huey服务尚未初始化")
    return service


def _missing_window_click_flow_job_response(
    job_id: str,
    *,
    message: str,
) -> JSONResponse:
    return JSONResponse(
        {
            "ok": False,
            "status": "missing",
            "jobId": job_id,
            "action": "window-click-flow",
            "title": "主页内容读取结果",
            "message": message,
            "tone": "warning",
            "items": [{"label": "任务ID", "value": job_id}],
        },
        status_code=404,
    )


def _start_article_detail_diagnostic_job(
    backend: DevBackendContext,
    payload: ArticleDetailDiagnosticPayload | None = None,
) -> dict[str, Any] | JSONResponse:
    payload = payload or ArticleDetailDiagnosticPayload()
    try:
        card_index = payload.cardIndex
        account_name = payload.accountName
        card = payload.card
        if card is None or not account_name:
            probe_result = _article_detail_card_probe_service(
                backend
            ).read_first_visible_card()
            records = list(probe_result.get("records") or [])
            if not probe_result.get("ok") or not records:
                return _article_detail_probe_not_started_payload(probe_result)
            if card is None:
                card = dict(records[0])
                card_index = _coerce_positive_int(card.get("index"), card_index)
            if not account_name:
                account_name = str(probe_result.get("accountName") or "")
        return _article_detail_huey_service(backend).start(
            card_index=card_index,
            account_name=account_name,
            card=card,
            skip_collected_records=payload.skipCollectedRecords,
        )
    except Exception as exc:
        backend.append_log("ERROR", f"详情获取Huey任务启动失败：{exc}")
        return JSONResponse(
            {
                "ok": False,
                "status": "failed",
                "action": "single-article-detail",
                "title": "详情获取结果",
                "message": f"详情获取Huey任务启动失败：{exc}",
                "tone": "error",
                "items": [{"label": "失败原因", "value": str(exc)}],
            },
            status_code=400,
        )


def _article_detail_probe_not_started_payload(probe_result: dict[str, Any]) -> dict[str, Any]:
    result = dict(probe_result)
    result.setdefault("ok", False)
    result.setdefault("status", "failed")
    result.setdefault("action", "single-article-detail")
    result.setdefault("title", "详情获取结果")
    result.setdefault("message", "读取首篇文章卡片失败，未启动单篇 Huey 任务。")
    result.setdefault("tone", "warning")
    result.setdefault("items", [])
    result.setdefault("records", [])
    result.setdefault("accountName", "")
    result.setdefault("captureType", "none")
    result["jobId"] = ""
    result["hueyStarted"] = False
    return result


def _article_detail_diagnostic_job_payload(
    backend: DevBackendContext,
    job_id: str,
) -> dict[str, Any] | JSONResponse:
    try:
        return _article_detail_huey_service(backend).get(job_id)
    except KeyError:
        return JSONResponse(
            {
                "ok": False,
                "status": "missing",
                "jobId": job_id,
                "action": "single-article-detail",
                "title": "详情获取结果",
                "message": "未找到该详情获取诊断记录。",
                "tone": "warning",
                "items": [{"label": "任务ID", "value": job_id}],
            },
            status_code=404,
        )


def _article_detail_huey_service(backend: DevBackendContext) -> Any:
    service = backend.article_detail_huey_service
    if service is None:
        raise RuntimeError("单篇详情Huey服务尚未初始化")
    return service


def _article_detail_card_probe_service(backend: DevBackendContext) -> Any:
    service = backend.article_detail_card_probe_service
    if service is not None:
        return service
    return ArticleCardProbeService(
        config=backend.runtime.config,
        window_factory=backend.runtime.window_factory,
    )


def _coerce_positive_int(value: Any, fallback: int) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError):
        return max(1, int(fallback))
    return max(1, result)


def _start_initial_content_storage_diagnostic_job(
    backend: DevBackendContext,
    payload: InitialContentStorageDiagnosticPayload | None = None,
) -> dict[str, Any] | JSONResponse:
    payload = payload or InitialContentStorageDiagnosticPayload()
    try:
        probe_result = _article_detail_card_probe_service(backend).read_first_visible_card()
        records = list(probe_result.get("records") or [])
        if not probe_result.get("ok") or not records:
            return _initial_content_storage_probe_not_started_payload(probe_result)
        card = dict(records[0])
        card_index = _coerce_positive_int(card.get("index"), 1)
        account_name = str(probe_result.get("accountName") or "")
        return _initial_content_storage_huey_service(backend).start(
            card_index=card_index,
            account_name=account_name,
            card=card,
            skip_collected_records=payload.skipCollectedRecords,
            store_article_detail=True,
        )
    except Exception as exc:
        backend.append_log("ERROR", f"初始内容存储Huey任务启动失败：{exc}")
        return JSONResponse(
            {
                "ok": False,
                "status": "failed",
                "action": "initial-content-storage",
                "title": "初始内容存储结果",
                "message": f"初始内容存储Huey任务启动失败：{exc}",
                "tone": "error",
                "items": [{"label": "失败原因", "value": str(exc)}],
                "options": {
                    "skipCollectedRecords": payload.skipCollectedRecords,
                    "storeArticleDetail": True,
                },
            },
            status_code=400,
        )


def _initial_content_storage_diagnostic_job_payload(
    backend: DevBackendContext,
    job_id: str,
) -> dict[str, Any] | JSONResponse:
    try:
        return _initial_content_storage_huey_service(backend).get(job_id)
    except KeyError:
        return JSONResponse(
            {
                "ok": False,
                "status": "missing",
                "jobId": job_id,
                "action": "initial-content-storage",
                "title": "初始内容存储结果",
                "message": "未找到该初始内容存储诊断记录。",
                "tone": "warning",
                "items": [{"label": "任务ID", "value": job_id}],
            },
            status_code=404,
        )


def _initial_content_storage_huey_service(backend: DevBackendContext) -> Any:
    service = backend.initial_content_storage_huey_service
    if service is None:
        raise RuntimeError("初始内容存储Huey服务尚未初始化")
    return service


def _initial_content_storage_probe_not_started_payload(
    probe_result: dict[str, Any],
) -> dict[str, Any]:
    result = dict(probe_result)
    result.setdefault("ok", False)
    result.setdefault("status", "failed")
    result["action"] = "initial-content-storage"
    result["title"] = "初始内容存储结果"
    result.setdefault("message", "读取首篇文章卡片失败，未启动初始内容存储 Huey 任务。")
    result.setdefault("tone", "warning")
    result.setdefault("items", [])
    result.setdefault("records", [])
    result.setdefault("accountName", "")
    result.setdefault("captureType", "none")
    result["jobId"] = ""
    result["hueyStarted"] = False
    return result


def _start_article_detail_comments_diagnostic_job(
    backend: DevBackendContext,
    payload: ArticleDetailCommentsDiagnosticPayload | None = None,
) -> dict[str, Any] | JSONResponse:
    payload = payload or ArticleDetailCommentsDiagnosticPayload()
    try:
        probe_result = _article_detail_card_probe_service(backend).read_first_visible_card()
        records = list(probe_result.get("records") or [])
        if not probe_result.get("ok") or not records:
            return _article_detail_comments_probe_not_started_payload(probe_result)

        card = dict(records[0])
        card_index = int(card.get("index") or 1)
        if card_index <= 0:
            card_index = 1
        account_name = str(probe_result.get("accountName") or "")
        return _article_detail_comments_huey_service(backend).start(
            card_index=card_index,
            account_name=account_name,
            card=card,
            skip_collected_records=payload.skipCollectedRecords,
            store_article_detail=True,
            store_comment_info=True,
        )
    except Exception as exc:
        backend.append_log("ERROR", f"单篇评论存储Huey任务启动失败：{exc}")
        return JSONResponse(
            {
                "ok": False,
                "status": "failed",
                "jobId": "",
                "action": "article-detail-comments",
                "title": "详情评论结果",
                "message": f"单篇评论存储Huey任务启动失败：{exc}",
                "tone": "error",
                "items": [{"label": "失败原因", "value": str(exc)}],
                "hueyStarted": False,
            },
            status_code=400,
        )


def _article_detail_comments_diagnostic_job_payload(
    backend: DevBackendContext,
    job_id: str,
) -> dict[str, Any] | JSONResponse:
    try:
        return _article_detail_comments_huey_service(backend).get(job_id)
    except KeyError:
        return JSONResponse(
            {
                "ok": False,
                "status": "missing",
                "jobId": job_id,
                "action": "article-detail-comments",
                "title": "详情评论结果",
                "message": "未找到该详情评论诊断记录。",
                "tone": "warning",
                "items": [{"label": "任务ID", "value": job_id}],
            },
            status_code=404,
        )


def _article_detail_comments_huey_service(backend: DevBackendContext) -> Any:
    service = backend.article_detail_comments_huey_service
    if service is None:
        raise RuntimeError("单篇评论存储Huey服务尚未初始化")
    return service


def _article_detail_comments_probe_not_started_payload(
    probe_result: dict[str, Any],
) -> dict[str, Any]:
    result = dict(probe_result)
    result["action"] = "article-detail-comments"
    result["title"] = "详情评论结果"
    result["jobId"] = ""
    result["hueyStarted"] = False
    result.setdefault("ok", False)
    result.setdefault("status", "no-visible-card")
    result.setdefault("message", "读取首篇文章卡片失败，未启动单篇评论存储 Huey 任务。")
    result.setdefault("tone", "warning")
    result.setdefault("items", [])
    return result


def _start_article_detail_offline_cache_diagnostic_job(
    backend: DevBackendContext,
    payload: ArticleDetailOfflineCacheDiagnosticPayload | None = None,
) -> dict[str, Any] | JSONResponse:
    payload = payload or ArticleDetailOfflineCacheDiagnosticPayload()
    try:
        probe_result = _article_detail_card_probe_service(backend).read_first_visible_card()
        records = list(probe_result.get("records") or [])
        if not probe_result.get("ok") or not records:
            return _article_detail_offline_cache_probe_not_started_payload(probe_result)

        card = dict(records[0])
        card_index = _coerce_positive_int(card.get("index"), 1)
        account_name = str(probe_result.get("accountName") or "")
        return _article_detail_offline_cache_huey_service(backend).start(
            card_index=card_index,
            account_name=account_name,
            card=card,
            skip_collected_records=payload.skipCollectedRecords,
            store_article_detail=True,
            archive_offline_content=True,
            stateful_offline_cache=payload.statefulOfflineCache,
        )
    except Exception as exc:
        backend.append_log("ERROR", f"单篇离线缓存Huey任务启动失败：{exc}")
        return JSONResponse(
            {
                "ok": False,
                "status": "failed",
                "jobId": "",
                "action": "article-detail-offline-cache",
                "title": "单篇离线缓存结果",
                "message": f"单篇离线缓存Huey任务启动失败：{exc}",
                "tone": "error",
                "items": [{"label": "失败原因", "value": str(exc)}],
                "hueyStarted": False,
                "options": {
                    "skipCollectedRecords": payload.skipCollectedRecords,
                    "storeArticleDetail": True,
                    "archiveOfflineContent": True,
                    "statefulOfflineCache": payload.statefulOfflineCache,
                },
            },
            status_code=400,
        )


def _article_detail_offline_cache_diagnostic_job_payload(
    backend: DevBackendContext,
    job_id: str,
) -> dict[str, Any] | JSONResponse:
    try:
        return _article_detail_offline_cache_huey_service(backend).get(job_id)
    except KeyError:
        return JSONResponse(
            {
                "ok": False,
                "status": "missing",
                "jobId": job_id,
                "action": "article-detail-offline-cache",
                "title": "单篇离线缓存结果",
                "message": "未找到该单篇离线缓存诊断记录。",
                "tone": "warning",
                "items": [{"label": "任务ID", "value": job_id}],
            },
            status_code=404,
        )


def _article_detail_offline_cache_huey_service(backend: DevBackendContext) -> Any:
    service = backend.article_detail_offline_cache_huey_service
    if service is None:
        raise RuntimeError("单篇离线缓存Huey服务尚未初始化")
    return service


def _article_detail_offline_cache_probe_not_started_payload(
    probe_result: dict[str, Any],
) -> dict[str, Any]:
    result = dict(probe_result)
    result["action"] = "article-detail-offline-cache"
    result["title"] = "单篇离线缓存结果"
    result["jobId"] = ""
    result["hueyStarted"] = False
    result.setdefault("ok", False)
    result.setdefault("status", "no-visible-card")
    result.setdefault("message", "读取首篇文章卡片失败，未启动单篇离线缓存 Huey 任务。")
    result.setdefault("tone", "warning")
    result.setdefault("items", [])
    result.setdefault("options", {
        "skipCollectedRecords": False,
        "storeArticleDetail": True,
        "archiveOfflineContent": True,
        "statefulOfflineCache": False,
    })
    return result


def _run_diagnostic_command(backend: DevBackendContext, args: list[str], *, timeout: int) -> Any:
    # 诊断命令集中在这里，测试可注入 fake runner，真实环境只执行明确的系统命令。
    if callable(backend.command_runner):
        return backend.command_runner(
            args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    return subprocess.run(
        args,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )


def _parse_mitm_certificate_payload(stdout: str) -> dict[str, Any]:
    text = stdout.strip()
    if not text:
        return {"projectCertificate": None, "certificates": []}
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return {"projectCertificate": None, "certificates": []}
    if isinstance(parsed, dict) and "certificates" in parsed:
        project_certificate = _normalize_project_certificate(parsed.get("projectCertificate"))
        raw_items = parsed.get("certificates") or []
    else:
        project_certificate = None
        raw_items = parsed if isinstance(parsed, list) else [parsed]
    certificates: list[dict[str, str]] = []
    for item in raw_items:
        if not isinstance(item, dict):
            continue
        certificate = _normalize_mitm_certificate(item)
        if certificate["thumbprint"]:
            certificates.append(certificate)
    return {"projectCertificate": project_certificate, "certificates": certificates}


def _parse_mitm_certificate_json(stdout: str) -> list[dict[str, str]]:
    """兼容已有调用方，仅返回系统证书列表。"""
    return _parse_mitm_certificate_payload(stdout)["certificates"]


def _normalize_project_certificate(value: Any) -> dict[str, str] | None:
    if not isinstance(value, dict):
        return None
    return {
        "path": _certificate_value(value, "path", "Path"),
        "thumbprint": _normalize_certificate_thumbprint(_certificate_value(value, "thumbprint", "Thumbprint")),
        "subject": _certificate_value(value, "subject", "Subject"),
        "issuer": _certificate_value(value, "issuer", "Issuer"),
        "friendlyName": _certificate_value(value, "friendlyName", "FriendlyName"),
        "notBefore": _certificate_value(value, "notBefore", "NotBefore"),
        "notAfter": _certificate_value(value, "notAfter", "NotAfter"),
    }


def _normalize_mitm_certificate(item: dict[str, Any]) -> dict[str, str]:
    return {
        "storePath": _certificate_value(item, "storePath", "StorePath") or "Cert:\\CurrentUser\\Root",
        "thumbprint": _normalize_certificate_thumbprint(_certificate_value(item, "thumbprint", "Thumbprint")),
        "subject": _certificate_value(item, "subject", "Subject"),
        "issuer": _certificate_value(item, "issuer", "Issuer"),
        "friendlyName": _certificate_value(item, "friendlyName", "FriendlyName"),
        "notBefore": _certificate_value(item, "notBefore", "NotBefore"),
        "notAfter": _certificate_value(item, "notAfter", "NotAfter"),
    }


def _normalize_certificate_thumbprint(value: Any) -> str:
    return re.sub(r"[\s-]+", "", str(value or "")).upper()


def _project_relative_path(backend: DevBackendContext, path: Path) -> str:
    try:
        return path.resolve().relative_to(backend.project_root.resolve()).as_posix()
    except ValueError:
        return str(path)


def _powershell_single_quote(value: str) -> str:
    return value.replace("'", "''")


def _certificate_value(item: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = item.get(key)
        if value is not None:
            return str(value).strip()
    return ""


def _normalize_thumbprints(value: Any) -> list[str]:
    if not isinstance(value, list):
        return []
    seen: set[str] = set()
    thumbprints: list[str] = []
    for item in value:
        thumbprint = str(item or "").strip()
        if not thumbprint or thumbprint in seen:
            continue
        seen.add(thumbprint)
        thumbprints.append(thumbprint)
    return thumbprints


def _resolve_project_path(backend: DevBackendContext, value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return (backend.project_root / path).resolve()


def _unsupported(message: str) -> JSONResponse:
    return JSONResponse(
        {"ok": False, "status": "unsupported", "message": message},
        status_code=501,
    )


def _safe_int(value: Any, *, default: int, minimum: int) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        parsed = default
    return max(minimum, parsed)


def _directory_size_bytes(path: Path) -> int:
    if not path.exists():
        return 0
    if path.is_file():
        return path.stat().st_size
    total = 0
    for child in path.rglob("*"):
        if child.is_file():
            try:
                total += child.stat().st_size
            except OSError:
                continue
    return total


def _format_size(size_bytes: int) -> str:
    value = float(max(0, int(size_bytes)))
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{value:.1f} GB"


def _format_duration(seconds: Any) -> str:
    value = max(0, int(float(seconds or 0)))
    minutes, sec = divmod(value, 60)
    hours, minute = divmod(minutes, 60)
    return f"{hours:02d}:{minute:02d}:{sec:02d}"


def _resolve_archive_dir(storage_root: Path, archive_dir: str) -> Path | None:
    if not archive_dir.strip():
        return None
    raw_path = Path(archive_dir)
    resolved = raw_path if raw_path.is_absolute() else storage_root / raw_path
    try:
        resolved = resolved.resolve()
        resolved.relative_to(storage_root.resolve())
    except ValueError:
        return None
    return resolved


def _open_directory(path: Path) -> bool:
    target = Path(path).resolve()
    if not target.exists():
        return False
    if sys.platform.startswith("win"):
        subprocess.Popen(["explorer", str(target)])
        return True
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(target)])
        return True
    subprocess.Popen(["xdg-open", str(target)])
    return True


if __name__ == "__main__":
    main()
