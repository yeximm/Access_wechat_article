from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Any, Literal, Mapping


DateFilterMode = Literal["all", "range", "before", "after"]
OfflineArchiveMode = Literal["standard", "beta"]
ArticleTaskStatus = Literal["running", "success", "skipped_collected", "failed", "cancelled"]
ArticleTaskPhase = Literal["precheck", "foreground", "detail", "comments", "offline", "finished"]


def _text(value: Any) -> str:
    return " ".join(str(value or "").split())


@dataclass(frozen=True, slots=True)
class MainFlowCommand:
    """主服务启动参数；只表达业务选择，不携带底层模块对象。"""

    target_count: int = 0
    date_filter_mode: DateFilterMode = "all"
    start_date: str | None = None
    end_date: str | None = None
    collect_comments: bool = False
    archive_offline: bool = False
    offline_archive_mode: OfflineArchiveMode = "standard"
    skip_collected_records: bool = False
    single_task_interval_seconds: float = 0.0

    def __post_init__(self) -> None:
        if int(self.target_count) < 0:
            raise ValueError("任务数量不能小于 0")
        if self.date_filter_mode not in {"all", "range", "before", "after"}:
            raise ValueError(f"不支持的日期筛选模式：{self.date_filter_mode}")
        if self.offline_archive_mode not in {"standard", "beta"}:
            raise ValueError(f"不支持的离线归档模式：{self.offline_archive_mode}")
        if float(self.single_task_interval_seconds) < 0:
            raise ValueError("单篇任务间隔不能小于 0")
        if self.date_filter_mode == "range" and bool(self.start_date) != bool(self.end_date):
            raise ValueError("日期范围必须同时提供起始日期和截止日期")

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> "MainFlowCommand":
        """把主服务页面当前使用的字段转换为稳定的领域命令。"""
        selections = payload.get("selections")
        if not isinstance(selections, Mapping):
            selections = {}
        mode = str(payload.get("dateFilterMode") or payload.get("date_filter_mode") or "all")
        offline_mode = str(
            selections.get("offlineArchiveMode")
            or payload.get("offlineArchiveMode")
            or "standard"
        )
        return cls(
            target_count=max(0, int(payload.get("recordLimit", payload.get("targetCount", 0)) or 0)),
            date_filter_mode=mode,  # type: ignore[arg-type]
            start_date=payload.get("startDate") or payload.get("start_date"),
            end_date=payload.get("endDate") or payload.get("end_date"),
            collect_comments=bool(selections.get("commentInfo", payload.get("collectComments", False))),
            archive_offline=bool(selections.get("offlineArchive", payload.get("archiveOffline", False))),
            offline_archive_mode=offline_mode,  # type: ignore[arg-type]
            skip_collected_records=bool(
                selections.get("skipCollectedRecords", payload.get("skipCollectedRecords", False))
            ),
            single_task_interval_seconds=float(
                payload.get("singleTaskIntervalSeconds", payload.get("requestIntervalSeconds", 0)) or 0
            ),
        )

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True, slots=True)
class MainFlowContext:
    task_id: str
    db_path: Path
    storage_root: Path
    temp_dir: Path
    config_snapshot: Any
    cancel_token: Event = field(repr=False, compare=False)
    started_at: datetime
    state: Any = field(repr=False, compare=False)
    # 事件入口只负责状态和日志聚合。
    event_sink: Any | None = field(default=None, repr=False, compare=False)


@dataclass(frozen=True, slots=True)
class HomeArticleTarget:
    """主页 UIA 解析产生的单篇文章目标；fingerprint 不包含坐标。"""

    sequence: int
    account_name: str
    article_date: str
    title_raw: str
    title_display: str
    card_rect: tuple[int, int, int, int]
    visible_rect: tuple[int, int, int, int]
    click_point: tuple[int, int]
    source_snapshot_id: str = ""
    home_window_id: str = ""
    fingerprint: str = field(default="", init=False)

    def __post_init__(self) -> None:
        identity = "|".join(
            (
                _text(self.account_name),
                _text(self.article_date),
                _text(self.title_raw) or _text(self.title_display),
            )
        )
        object.__setattr__(self, "fingerprint", identity)


@dataclass(frozen=True, slots=True)
class SingleArticleOptions:
    skip_collected_records: bool = False
    collect_article_detail: bool = True
    collect_comments: bool = False
    archive_offline: bool = False
    offline_archive_mode: OfflineArchiveMode = "standard"


@dataclass(frozen=True, slots=True)
class SingleArticleReceipt:
    task_id: str
    target_fingerprint: str
    status: ArticleTaskStatus
    phase: ArticleTaskPhase = "finished"
    foreground_done: bool = False
    progress_counted: bool = False
    tab_closed: bool = False
    article_saved: bool = False
    detail_status: str = "pending"
    comments_status: str = "not_requested"
    offline_status: str = "not_requested"
    archive_dir: str | None = None
    article_title: str = ""
    message: str = ""
    error_stage: str | None = None
    error_detail: str | None = None
    duration_seconds: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "SingleArticleReceipt":
        return cls(
            task_id=str(data.get("task_id") or ""),
            target_fingerprint=str(data.get("target_fingerprint") or ""),
            status=str(data.get("status") or "failed"),  # type: ignore[arg-type]
            phase=str(data.get("phase") or "finished"),  # type: ignore[arg-type]
            foreground_done=bool(data.get("foreground_done", False)),
            progress_counted=bool(data.get("progress_counted", False)),
            tab_closed=bool(data.get("tab_closed", False)),
            article_saved=bool(data.get("article_saved", False)),
            detail_status=str(data.get("detail_status") or "pending"),
            comments_status=str(data.get("comments_status") or "not_requested"),
            offline_status=str(data.get("offline_status") or "not_requested"),
            archive_dir=data.get("archive_dir"),
            article_title=str(data.get("article_title") or ""),
            message=str(data.get("message") or ""),
            error_stage=data.get("error_stage"),
            error_detail=data.get("error_detail"),
            duration_seconds=float(data.get("duration_seconds") or 0),
        )


@dataclass(frozen=True, slots=True)
class MainFlowSnapshot:
    task_id: str
    status: str
    message: str
    runtime_state: dict[str, Any]
    started_at: str
    finished_at: str | None = None

    def to_payload(self) -> dict[str, Any]:
        payload = {
            "ok": self.status not in {"failed"},
            "status": self.status,
            "taskId": self.task_id,
            "message": self.message,
            "runtimeState": dict(self.runtime_state),
            "startedAt": self.started_at,
        }
        if self.finished_at:
            payload["finishedAt"] = self.finished_at
        return payload
