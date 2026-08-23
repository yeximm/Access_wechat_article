from __future__ import annotations

from pathlib import Path
from typing import Any

from .article_dispatcher import ArticleDispatcher
from .home_scan_service import HomeScanService
from .main_flow_coordinator import MainFlowCoordinator
from .main_flow_service import MainFlowService


def build_main_flow_service(
    *,
    project_root: str | Path,
    config: Any,
    db_path: str | Path,
    storage_root: str | Path,
    temp_root: str | Path,
    window_factory: Any,
    detail_service: Any,
    foreground_service: Any | None = None,
    comments_service: Any | None = None,
    offline_service: Any | None = None,
    runtime_logger: Any | None = None,
) -> tuple[MainFlowService, ArticleDispatcher]:
    """装配主服务编排链；入口层不复制窗口或单篇业务逻辑。"""

    scanner = HomeScanService(
        window_factory=window_factory,
        config=config,
        trace=_build_trace(runtime_logger),
    )
    dispatcher = ArticleDispatcher(
        service=detail_service,
        foreground_service=foreground_service,
        post_services={
            "comments": comments_service,
            "offline": offline_service,
        },
    )
    coordinator = MainFlowCoordinator(
        scanner=scanner,
        dispatcher=dispatcher,
        foreground_timeout_seconds=_foreground_timeout(config),
        post_task_timeout_seconds=_post_timeout(config),
        trace=_build_trace(runtime_logger),
    )
    service = MainFlowService(
        project_root=project_root,
        config=config,
        db_path=db_path,
        storage_root=storage_root,
        temp_root=temp_root,
        runtime_logger=runtime_logger,
        runner=coordinator.run,
        shutdown_callback=dispatcher.shutdown,
    )
    return service, dispatcher


def _foreground_timeout(config: Any) -> float:
    window = getattr(config, "window", None)
    mitm = getattr(config, "mitm_capture", None)
    return max(
        30.0,
        float(getattr(window, "article_open_timeout_seconds", 30.0))
        + float(getattr(mitm, "result_timeout_seconds", 30.0))
        + 15.0,
    )


def _post_timeout(config: Any) -> float:
    comment = getattr(config, "comment", None)
    offline = getattr(config, "offline_cache", None)
    comment_timeout = float(getattr(comment, "request_timeout_seconds", 30.0)) * max(
        1, int(getattr(comment, "max_pages", 1))
    )
    offline_timeout = float(getattr(offline, "max_scroll_seconds", 30.0)) + 60.0
    return max(60.0, comment_timeout, offline_timeout)


def _build_trace(runtime_logger: Any | None):
    if runtime_logger is None:
        return None
    seen_summary_keys: set[str] = set()

    def trace(event: dict[str, Any]) -> None:
        message = str(event.get("message") or event.get("event") or "主流程事件")
        event_name = str(event.get("event") or "").strip()
        details = event.get("details")
        try:
            runtime_logger.write_detail(
                "INFO",
                message,
                source="main-flow",
                channel="main_flow",
                phase=event_name,
                context=details,
            )
            summary_message = _summary_trace_message(event_name, message, details)
            if not summary_message:
                return
            dedupe_key = f"{event_name}|{summary_message}"
            if dedupe_key in seen_summary_keys:
                return
            seen_summary_keys.add(dedupe_key)
            runtime_logger.write_summary(
                "WARN" if "stalled" in event_name else "INFO",
                summary_message,
                source="main-flow",
                channel="main_flow",
                phase=event_name,
            )
        except Exception:
            return

    return trace


def _summary_trace_message(
    event_name: str,
    message: str,
    details: Any,
) -> str:
    """把主流程内部 trace 压缩成前台可读的一句话。"""

    if event_name == "action":
        action = ""
        if isinstance(details, dict):
            action = str(details.get("action") or "").strip()
        return f"主流程：{action or message}"
    if event_name == "home-found":
        account_name = ""
        if isinstance(details, dict):
            account_name = str(details.get("accountName") or "").strip()
        return f"主页读取：公众号名称 {account_name}" if account_name else message
    if event_name == "home-scan":
        if isinstance(details, dict):
            return (
                "主页读取："
                f"候选 {details.get('rawCount', 0)} 条，"
                f"可处理 {details.get('targetCount', 0)} 条，"
                f"跳过 {details.get('skippedCount', 0)} 条"
            )
        return message
    if event_name in {"more-trigger-found", "more-trigger-clicked"}:
        remaining = 0
        if isinstance(details, dict):
            remaining = int(details.get("remainingCount") or 0)
        if event_name == "more-trigger-found":
            return f"主页读取：发现余下 {remaining} 篇入口"
        return f"主页读取：已展开余下 {remaining} 篇，重新读取当前页面"
    if event_name == "more-trigger-skipped":
        return message
    if event_name == "more-trigger-not-visible":
        return message
    if event_name == "more-trigger-click-failed":
        return message
    if event_name == "main-flow-failed":
        error = ""
        if isinstance(details, dict):
            error = str(details.get("error") or "").strip()
        return f"主流程异常：{error}" if error else message
    if event_name in {"date-scroll-stalled", "scroll-stalled"}:
        return message
    return ""


__all__ = ["build_main_flow_service"]
