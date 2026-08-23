from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
import time
from typing import Any, Callable

from src.domain.models import ArticleTarget
from src.modules.window.article_date_filter import (
    ArticleDateFilter,
    DateFilterDecision,
    parse_target_published_date,
)
from src.modules.window.article_card_reader import normalize_window_text
from src.modules.window.window_models import WindowInfo

from .main_flow_models import HomeArticleTarget, MainFlowCommand


_MAX_STALLED_DATE_SEEKS = 3
_MAX_MORE_EXPANSIONS_PER_SCAN = 8


class HomeScanError(RuntimeError):
    """主页无法定位、读取或建立扫描会话。"""


@dataclass(frozen=True, slots=True)
class HomeScanBatch:
    """一次 UIA 读取后交给主流程的候选结果。"""

    targets: tuple[HomeArticleTarget, ...]
    raw_count: int
    skipped_count: int
    boundary_reached: bool = False
    loading: bool = False
    snapshot_id: str = ""


@dataclass(slots=True)
class HomeScanSession:
    window: WindowInfo
    account_name: str
    cursor: Any | None
    scroller: Any
    date_filter: ArticleDateFilter
    command: MainFlowCommand
    date_reader: Any | None = None
    snapshot_reader: Any | None = None
    processed_fingerprints: set[str] = field(default_factory=set)
    scan_number: int = 0
    boundary_reached: bool = False
    last_article_targets: list[ArticleTarget] = field(default_factory=list)
    last_snapshot_id: str = ""
    next_sequence: int = 1
    last_snapshot: Any | None = None
    pending_snapshot: Any | None = None
    last_snapshot_signature: tuple[Any, ...] = ()
    last_date_snapshot: Any | None = None
    pending_date_snapshot: Any | None = None
    last_date_snapshot_signature: tuple[Any, ...] = ()
    expanded_more_keys: set[tuple[str, str]] = field(default_factory=set)


class HomeScanService:
    """封装主页定位、日期边界、UIA 卡片读取和滚动。

    该服务不点击文章、不查询数据库，也不把 UIA 控件对象交给主流程。
    """

    def __init__(
        self,
        *,
        window_factory: Any,
        config: Any,
        trace: Callable[[dict[str, Any]], None] | None = None,
        sleep: Callable[[float], None] = time.sleep,
        monotonic: Callable[[], float] = time.monotonic,
    ) -> None:
        self._window_factory = window_factory
        self._config = config
        self._trace = trace
        self._sleep_fn = sleep
        self._monotonic = monotonic

    def initialize(self, command: MainFlowCommand) -> HomeScanSession:
        reader = self._window_factory.create_reader()
        timeout = float(
            getattr(getattr(self._config, "window", None), "home_find_timeout_seconds", 3.0)
        )
        window = self._window_factory.find_home_window(
            reader=reader,
            timeout_seconds=timeout,
        )
        if window is None:
            raise HomeScanError("未找到公众号主页窗口，请先打开公众号主页")

        try:
            home_info = self._read_home_info_after_activation(window)
            account_name = str(getattr(home_info, "account_name", "") or "").strip()
        except Exception as exc:
            raise HomeScanError(f"读取公众号名称失败：{exc}") from exc
        if not account_name:
            raise HomeScanError("未从公众号主页读取到公众号名称")

        date_filter = ArticleDateFilter.create(
            mode=command.date_filter_mode,
            start_date=command.start_date,
            end_date=command.end_date,
        )
        snapshot_reader = self._create_snapshot_reader()
        cursor = None
        if snapshot_reader is None:
            cursor = self._window_factory.create_cursor(
                reader=reader,
                account_name=account_name,
                trace=self._trace,
            )
        scroller = self._window_factory.create_scroller()
        date_reader = (
            snapshot_reader
            if callable(getattr(snapshot_reader, "read_date_groups", None))
            else self._create_date_reader()
        )
        session = HomeScanSession(
            window=window,
            account_name=account_name,
            cursor=cursor,
            scroller=scroller,
            date_filter=date_filter,
            command=command,
            date_reader=date_reader,
            snapshot_reader=snapshot_reader,
        )
        self._emit(
            "home-found",
            "已定位公众号主页并读取公众号名称",
            accountName=account_name,
            dateFilter=date_filter.label,
        )
        if command.date_filter_mode in {"range", "after"}:
            self.locate_date_boundary(session)
        return session

    def _read_home_info_after_activation(self, window: WindowInfo) -> Any:
        """读取主页身份前先激活窗口，避免 Chromium 未暴露可读 DocumentControl。"""

        last_error: Exception | None = None
        for attempt in range(2):
            self._activate_home_window(window)
            try:
                return self._window_factory.create_home_reader().read(window)
            except Exception as exc:
                last_error = exc
                if attempt == 0:
                    self._emit(
                        "home-read-retry",
                        "主页 UIA 暂不可读，已重新激活主页并准备重试读取公众号名称",
                        error=str(exc),
                    )
        assert last_error is not None
        raise last_error

    def _activate_home_window(self, window: WindowInfo) -> None:
        """复用窗口测试的主页激活动作，使主流程读取 UIA 前处于稳定前台。"""

        guard_factory = getattr(self._window_factory, "create_home_guard", None)
        if not callable(guard_factory):
            return
        guard = guard_factory()
        activate = getattr(guard, "activate", None)
        if not callable(activate):
            return
        activate(window)
        wait_seconds = float(
            getattr(getattr(self._config, "window", None), "activation_wait_seconds", 0.0)
            or 0.0
        )
        if wait_seconds > 0:
            self._sleep_fn(wait_seconds)

    def scan_visible_targets(self, session: HomeScanSession) -> HomeScanBatch:
        if session.boundary_reached:
            return HomeScanBatch((), 0, 0, boundary_reached=True, snapshot_id=session.last_snapshot_id)

        result: list[HomeArticleTarget] = []
        batch_fingerprints: set[str] = set()
        raw_count = 0
        skipped_count = 0
        boundary_reached = False
        loading = False
        expansion_count = 0

        while True:
            snapshot = self._consume_snapshot(session)
            if snapshot is not None:
                raw_cards = list(_snapshot_visible_cards(snapshot))
                raw_targets: list[ArticleTarget] = []
                session.last_article_targets = []
                loading = bool(getattr(snapshot, "loading", False))
            else:
                if session.cursor is None:
                    raise HomeScanError("主页扫描器没有可用的 UIA 快照读取器")
                raw_targets = list(session.cursor.refresh_visible(session.window) or ())
                raw_cards = []
                loading = bool(getattr(session.cursor, "_viewport_loading", False))
            session.scan_number += 1
            session.last_snapshot_id = f"{session.window.handle}:{session.scan_number}"

            candidates = raw_cards if snapshot is not None else raw_targets
            raw_count += len(candidates)
            expanded_more = False
            for target in candidates:
                article_date = parse_target_published_date(target)
                decision = session.date_filter.decide(article_date)
                if decision is DateFilterDecision.STOP:
                    boundary_reached = True
                    break
                if decision is not DateFilterDecision.INCLUDE:
                    skipped_count += 1
                    continue
                if bool(getattr(target, "is_more_trigger", False)):
                    if expansion_count >= _MAX_MORE_EXPANSIONS_PER_SCAN:
                        skipped_count += 1
                        self._emit(
                            "more-trigger-skipped",
                            "余下文章入口展开次数已达到本轮上限，跳过继续展开",
                            remainingCount=int(getattr(target, "remaining_count", 0) or 0),
                        )
                        continue
                    if self._expand_more_articles(session, target):
                        skipped_count += 1
                        expansion_count += 1
                        expanded_more = True
                        break
                    skipped_count += 1
                    continue
                if snapshot is not None:
                    converted = self._to_snapshot_target(
                        target,
                        account_name=session.account_name,
                        home_window_id=str(session.window.handle),
                        snapshot_id=session.last_snapshot_id,
                        sequence=session.next_sequence,
                    )
                else:
                    converted = self._to_home_target(
                        target,
                        account_name=session.account_name,
                        snapshot_id=session.last_snapshot_id,
                        sequence=session.next_sequence,
                    )
                if (
                    converted is None
                    or converted.fingerprint in session.processed_fingerprints
                    or converted.fingerprint in batch_fingerprints
                ):
                    skipped_count += 1
                    continue
                result.append(converted)
                batch_fingerprints.add(converted.fingerprint)
                session.next_sequence += 1
            if boundary_reached or not expanded_more:
                break

        session.boundary_reached = boundary_reached
        self._emit(
            "home-scan",
            f"第 {session.scan_number} 次主页读取：识别 {len(result)} 条可处理文章",
            rawCount=raw_count,
            targetCount=len(result),
            skippedCount=skipped_count,
            boundaryReached=boundary_reached,
            loading=loading,
        )
        return HomeScanBatch(
            tuple(result),
            raw_count=raw_count,
            skipped_count=skipped_count,
            boundary_reached=boundary_reached,
            loading=loading,
            snapshot_id=session.last_snapshot_id,
        )

    def _expand_more_articles(self, session: HomeScanSession, target: Any) -> bool:
        """点击当前日期组的“余下 xx 篇”入口，并让下一轮重新读取 UIA 快照。"""

        trigger_key = _more_trigger_key(target)
        remaining_count = int(getattr(target, "remaining_count", 0) or 0)
        if trigger_key in session.expanded_more_keys:
            self._emit(
                "more-trigger-skipped",
                "当前日期组的剩余文章入口已经点击过",
                dateGroupKey=list(trigger_key),
                remainingCount=remaining_count,
            )
            return False

        click_point = getattr(target, "click_point", None)
        if not (
            isinstance(click_point, (tuple, list))
            and len(click_point) >= 2
        ):
            self._emit(
                "more-trigger-not-visible",
                "剩余文章入口当前没有可点击的可视坐标",
                dateGroupKey=list(trigger_key),
                remainingCount=remaining_count,
            )
            return False

        clicker_factory = getattr(self._window_factory, "create_clicker", None)
        if not callable(clicker_factory):
            self._emit(
                "more-trigger-skipped",
                "当前运行时没有可用的主页点击器，跳过剩余文章入口",
                dateGroupKey=list(trigger_key),
                remainingCount=remaining_count,
            )
            return False

        click_x, click_y = int(click_point[0]), int(click_point[1])
        self._emit(
            "more-trigger-found",
            "识别到当前日期组的剩余文章入口",
            dateGroupKey=list(trigger_key),
            remainingCount=remaining_count,
            clickPoint=[click_x, click_y],
        )
        try:
            clicker = clicker_factory()
            clicker.click_point(int(session.window.handle), click_x, click_y)
        except Exception as exc:
            self._emit(
                "more-trigger-click-failed",
                "点击剩余文章入口失败",
                dateGroupKey=list(trigger_key),
                remainingCount=remaining_count,
                error=str(exc),
            )
            return False

        session.expanded_more_keys.add(trigger_key)
        self._emit(
            "more-trigger-clicked",
            f"已展开余下 {remaining_count} 篇，正在重新读取当前页面",
            dateGroupKey=list(trigger_key),
            remainingCount=remaining_count,
            clickPoint=[click_x, click_y],
        )
        self._sleep(self._scroll_initial_delay_seconds())
        return True

    def mark_processed(self, session: HomeScanSession, target: HomeArticleTarget) -> None:
        session.processed_fingerprints.add(target.fingerprint)
        if session.cursor is None:
            return
        for item in session.last_article_targets:
            if _article_identity(item) == _home_identity(target):
                mark = getattr(session.cursor, "mark_processed", None)
                if callable(mark):
                    mark(item)
                break

    def scroll_next(
        self,
        session: HomeScanSession,
        *,
        date_seek: bool = False,
        wheel_steps: int | None = None,
    ) -> bool:
        if session.boundary_reached:
            return False
        if wheel_steps is None:
            wheel_steps = self._date_seek_steps(session) if date_seek else self._normal_steps()
        scroll = getattr(session.scroller, "scroll", None)
        if not callable(scroll):
            return False
        succeeded = self._scroll_once(
            session,
            direction="down",
            wheel_steps=max(1, int(wheel_steps)),
        )
        if not succeeded:
            return False

        self._sleep(self._scroll_initial_delay_seconds())
        if date_seek and session.date_reader is not None:
            if self._probe_date_after_scroll(session):
                return True
            if self._bounce_enabled():
                for _ in range(self._bounce_attempts()):
                    if not self._scroll_once(
                        session,
                        direction="up",
                        wheel_steps=self._bounce_up_steps(),
                    ):
                        return False
                    self._sleep(self._bounce_pause_seconds())
                    if not self._scroll_once(
                        session,
                        direction="down",
                        wheel_steps=self._bounce_down_steps(),
                    ):
                        return False
                    self._sleep(self._scroll_initial_delay_seconds())
                    if self._probe_date_after_scroll(session):
                        return True
            self._emit(
                "date-scroll-stalled",
                "日期组快照没有变化，回弹后仍未发现新日期",
            )
            return False

        if session.snapshot_reader is not None:
            if self._probe_snapshot_after_scroll(session):
                return True
            if self._bounce_enabled():
                for _ in range(self._bounce_attempts()):
                    if not self._scroll_once(
                        session,
                        direction="up",
                        wheel_steps=self._bounce_up_steps(),
                    ):
                        return False
                    self._sleep(self._bounce_pause_seconds())
                    if not self._scroll_once(
                        session,
                        direction="down",
                        wheel_steps=self._bounce_down_steps(),
                    ):
                        return False
                    self._sleep(self._scroll_initial_delay_seconds())
                    if self._probe_snapshot_after_scroll(session):
                        return True
            # 保留最后一次快照供诊断记录；主流程停止继续滚动，避免无变化死循环。
            self._emit(
                "home-scroll-stalled",
                "滚动后 UIA 快照没有新文章，回弹后仍未变化",
            )
            return False
        invalidate = getattr(session.cursor, "invalidate_snapshot", None)
        if callable(invalidate):
            invalidate()
        self._emit(
            "home-scroll",
            f"向下滚动 {max(1, int(wheel_steps))} 步{'成功' if succeeded else '失败'}",
            wheelSteps=max(1, int(wheel_steps)),
            dateSeek=date_seek,
            succeeded=succeeded,
        )
        return succeeded

    def locate_date_boundary(self, session: HomeScanSession) -> None:
        """只读取轻量日期组，逼近目标日期后再进入文章扫描。"""

        if session.command.date_filter_mode not in {"range", "after"}:
            return
        if session.date_reader is None:
            # 测试替身或旧运行时没有日期组读取器时，让正式扫描器承担兼容定位。
            self._emit("date-locate", "当前运行时没有轻量日期组读取器，使用文章日期兼容定位")
            return

        snapshot = self._consume_date_snapshot(session)
        if snapshot is None:
            snapshot = self._read_date_snapshot(session, stage="date-location-initial")
        session.last_date_snapshot = snapshot
        session.last_date_snapshot_signature = _date_snapshot_signature(snapshot)
        previous_signature = _date_snapshot_signature(snapshot)
        stalled_reads = 0
        while True:
            if self._date_snapshot_reaches_target(session, snapshot):
                self._emit("date-locate", "已定位到目标日期附近")
                return
            wheel_steps = self._date_seek_steps(session, snapshot)
            if not self.scroll_next(session, date_seek=True, wheel_steps=wheel_steps):
                raise HomeScanError("日期定位时主页无法继续滚动")
            next_snapshot = self._consume_date_snapshot(session)
            if next_snapshot is None:
                next_snapshot = self._read_date_snapshot(
                    session,
                    stage="date-location-after-scroll",
                )
            next_signature = _date_snapshot_signature(next_snapshot)
            if next_signature == previous_signature:
                stalled_reads += 1
            else:
                stalled_reads = 0
            if stalled_reads >= _MAX_STALLED_DATE_SEEKS:
                raise HomeScanError("日期定位连续未发现新的日期组，已停止继续滚动")
            snapshot = next_snapshot
            session.last_date_snapshot = snapshot
            session.last_date_snapshot_signature = next_signature
            previous_signature = next_signature

    def _read_date_snapshot(self, session: HomeScanSession, *, stage: str) -> Any | None:
        reader = session.date_reader
        if reader is None:
            return None
        started_at = self._monotonic()
        snapshot = reader.read_date_groups(session.window)
        self._emit(
            "uia-date-snapshot",
            "已读取主页 UIA 日期组",
            stage=stage,
            groupCount=len(tuple(getattr(snapshot, "groups", ()) or ())),
            visibleGroupCount=len(tuple(getattr(snapshot, "visible_groups", ()) or ())),
            contentViewport=tuple(getattr(snapshot, "content_viewport", ()) or (0, 0, 0, 0)),
            nodeCount=int(getattr(snapshot, "node_count", 0) or 0),
            loading=bool(getattr(snapshot, "loading", False)),
            durationSeconds=round(self._monotonic() - started_at, 3),
            visibleDates=[
                {
                    "dateText": str(getattr(group, "date_text", "") or ""),
                    "publishedDate": str(getattr(group, "published_date", "") or ""),
                }
                for group in tuple(getattr(snapshot, "visible_groups", ()) or ())
            ],
        )
        return snapshot

    def _date_snapshot_reaches_target(self, session: HomeScanSession, snapshot: Any) -> bool:
        groups = list(getattr(snapshot, "groups", ()) or ())
        dates: list[date] = []
        for group in groups:
            value = str(getattr(group, "published_date", "") or "")
            try:
                if value:
                    dates.append(date.fromisoformat(value))
            except ValueError:
                continue
        if not dates:
            return bool(getattr(snapshot, "loading", False)) is False and False
        target = (
            session.date_filter.start_date
            if session.command.date_filter_mode == "after"
            else session.date_filter.end_date
        )
        if target is None:
            return False
        return min(dates) <= target

    def _create_date_reader(self) -> Any | None:
        factory = getattr(self._window_factory, "create_window_test_reader", None)
        if not callable(factory):
            return None
        try:
            return factory()
        except Exception:
            return None

    def _create_snapshot_reader(self) -> Any | None:
        factory = getattr(self._window_factory, "create_window_test_reader", None)
        if not callable(factory):
            return None
        try:
            reader = factory()
        except Exception:
            return None
        return reader if callable(getattr(reader, "read", None)) else None

    def _consume_snapshot(self, session: HomeScanSession) -> Any | None:
        if session.snapshot_reader is None:
            return None
        snapshot = session.pending_snapshot
        session.pending_snapshot = None
        if snapshot is None:
            snapshot = session.snapshot_reader.read(session.window)
        session.last_snapshot = snapshot
        session.last_snapshot_signature = _snapshot_signature(snapshot)
        return snapshot

    def _consume_date_snapshot(self, session: HomeScanSession) -> Any | None:
        snapshot = session.pending_date_snapshot
        session.pending_date_snapshot = None
        return snapshot

    def _probe_snapshot_after_scroll(self, session: HomeScanSession) -> bool:
        if session.snapshot_reader is None:
            return False
        deadline = self._monotonic() + self._lazy_load_timeout_seconds()
        interval = self._scroll_probe_interval_seconds()
        while True:
            snapshot = session.snapshot_reader.read(session.window)
            signature = _snapshot_signature(snapshot)
            session.pending_snapshot = snapshot
            if signature != session.last_snapshot_signature:
                return True
            if not bool(getattr(snapshot, "loading", False)):
                return False
            if self._monotonic() >= deadline:
                return False
            self._sleep(interval)
            interval = min(
                self._scroll_probe_max_interval_seconds(),
                max(interval, interval * 1.5),
            )

    def _probe_date_after_scroll(self, session: HomeScanSession) -> bool:
        reader = session.date_reader
        if reader is None:
            return False
        started_at = self._monotonic()
        unchanged_deadline = started_at + self._unchanged_before_bounce_seconds()
        lazy_deadline = started_at + max(
            self._unchanged_before_bounce_seconds(),
            self._lazy_load_timeout_seconds(),
        )
        interval = self._scroll_probe_interval_seconds()
        max_interval = self._scroll_probe_max_interval_seconds()
        loading_observed = False
        probe_count = 0
        latest = self._read_date_snapshot(session, stage="date-location-after-scroll")
        while True:
            snapshot = latest
            signature = _date_snapshot_signature(snapshot)
            session.pending_date_snapshot = snapshot
            if signature != session.last_date_snapshot_signature:
                self._emit(
                    "uia-date-observation-finished",
                    "滚动后的 UIA 日期组观察阶段已结束",
                    stage="date-location-after-scroll",
                    probeCount=probe_count + 1,
                    pageChanged=True,
                    targetReached=self._date_snapshot_reaches_target(session, snapshot),
                    loadingObserved=loading_observed,
                    loadingFinished=loading_observed and not bool(getattr(snapshot, "loading", False)),
                    durationSeconds=round(self._monotonic() - started_at, 3),
                )
                return True

            loading_observed = loading_observed or bool(getattr(snapshot, "loading", False))
            deadline = lazy_deadline if loading_observed else unchanged_deadline
            now = self._monotonic()
            if now >= deadline:
                self._emit(
                    "uia-date-observation-finished",
                    "滚动后的 UIA 日期组观察阶段已结束",
                    stage="date-location-after-scroll",
                    probeCount=probe_count + 1,
                    pageChanged=False,
                    targetReached=self._date_snapshot_reaches_target(session, snapshot),
                    loadingObserved=loading_observed,
                    loadingFinished=loading_observed and not bool(getattr(snapshot, "loading", False)),
                    durationSeconds=round(self._monotonic() - started_at, 3),
                )
                return False
            self._sleep(min(interval, max(0.0, deadline - now)))
            interval = min(max_interval, max(interval, interval * 1.5))
            probe_count += 1
            latest = self._read_date_snapshot(session, stage="date-location-after-scroll")

    def _scroll_once(
        self,
        session: HomeScanSession,
        *,
        direction: str,
        wheel_steps: int,
    ) -> bool:
        scroll = getattr(session.scroller, "scroll", None)
        if not callable(scroll):
            return False
        succeeded = bool(
            scroll(
                session.window,
                visible_targets=list(session.last_article_targets),
                direction=direction,
                wheel_steps=max(1, int(wheel_steps)),
            )
        )
        if succeeded and session.cursor is not None:
            invalidate = getattr(session.cursor, "invalidate_snapshot", None)
            if callable(invalidate):
                invalidate()
        direction_label = "上" if direction == "up" else "下"
        result_label = "成功" if succeeded else "失败"
        self._emit(
            "home-scroll",
            f"向{direction_label}滚动 {max(1, int(wheel_steps))} 步{result_label}",
            wheelSteps=max(1, int(wheel_steps)),
            direction=direction,
            succeeded=succeeded,
        )
        return succeeded

    def _sleep(self, seconds: float) -> None:
        if seconds > 0:
            self._sleep_fn(seconds)

    def _scroll_initial_delay_seconds(self) -> float:
        return max(
            0.0,
            float(
                getattr(
                    getattr(self._config, "window", None),
                    "scroll_initial_delay_seconds",
                    0.05,
                )
            ),
        )

    def _scroll_probe_interval_seconds(self) -> float:
        return max(
            0.01,
            float(
                getattr(
                    getattr(self._config, "window", None),
                    "scroll_probe_interval_seconds",
                    0.1,
                )
            ),
        )

    def _scroll_probe_max_interval_seconds(self) -> float:
        return max(
            self._scroll_probe_interval_seconds(),
            float(
                getattr(
                    getattr(self._config, "window", None),
                    "scroll_probe_max_interval_seconds",
                    0.4,
                )
            ),
        )

    def _lazy_load_timeout_seconds(self) -> float:
        return max(
            0.0,
            float(
                getattr(
                    getattr(self._config, "window", None),
                    "lazy_load_timeout_seconds",
                    3.0,
                )
            ),
        )

    def _unchanged_before_bounce_seconds(self) -> float:
        return max(
            0.0,
            float(
                getattr(
                    getattr(self._config, "window", None),
                    "unchanged_before_bounce_seconds",
                    0.6,
                )
            ),
        )

    def _bounce_enabled(self) -> bool:
        return bool(
            getattr(
                getattr(self._config, "window", None),
                "bounce_enabled",
                True,
            )
        )

    def _bounce_attempts(self) -> int:
        return max(
            0,
            int(
                getattr(
                    getattr(self._config, "window", None),
                    "bounce_attempts",
                    2,
                )
            ),
        )

    def _bounce_up_steps(self) -> int:
        return max(1, int(getattr(getattr(self._config, "window", None), "bounce_up_steps", 2)))

    def _bounce_down_steps(self) -> int:
        return max(1, int(getattr(getattr(self._config, "window", None), "bounce_down_steps", 6)))

    def _bounce_pause_seconds(self) -> float:
        return max(0.0, float(getattr(getattr(self._config, "window", None), "bounce_pause_seconds", 0.2)))

    def _normal_steps(self) -> int:
        return max(1, int(getattr(getattr(self._config, "window", None), "scroll_wheel_steps", 3)))

    def _date_seek_steps(self, session: HomeScanSession, snapshot: Any | None = None) -> int:
        maximum = max(1, int(getattr(getattr(self._config, "window", None), "date_seek_max_steps", 18)))
        normal = self._normal_steps()
        if snapshot is None:
            return min(maximum, max(normal, 6))
        target = (
            session.date_filter.start_date
            if session.command.date_filter_mode == "after"
            else session.date_filter.end_date
        )
        dates = []
        for group in list(getattr(snapshot, "groups", ()) or ()):
            value = str(getattr(group, "published_date", "") or "")
            try:
                if value:
                    dates.append(date.fromisoformat(value))
            except ValueError:
                continue
        if target is None or not dates:
            return normal
        remaining_days = max(0, (min(dates) - target).days)
        near = max(normal, round(maximum / 3))
        medium = max(near, round(maximum * 2 / 3))
        if remaining_days >= 15:
            return maximum
        if remaining_days >= 8:
            return medium
        if remaining_days >= 4:
            return near
        return normal

    @staticmethod
    def _to_home_target(
        target: ArticleTarget,
        *,
        account_name: str,
        snapshot_id: str,
        sequence: int = 1,
    ) -> HomeArticleTarget | None:
        raw_title = normalize_window_text(target.raw_title or target.title)
        published_date = normalize_window_text(target.published_date or target.date_text)
        if not raw_title or not published_date:
            return None
        visible_rect = _visible_rect(target)
        if visible_rect is None or visible_rect[3] - visible_rect[1] < 10:
            return None
        card_rect = _card_rect(target, visible_rect)
        return HomeArticleTarget(
            sequence=sequence,
            account_name=account_name,
            article_date=published_date,
            title_raw=raw_title,
            title_display=normalize_window_text(target.title or raw_title),
            card_rect=card_rect,
            visible_rect=visible_rect,
            click_point=(int(target.click_x), int(target.click_y)),
            source_snapshot_id=snapshot_id,
            home_window_id=str(target.home_window_handle),
        )

    @staticmethod
    def _to_snapshot_target(
        card: Any,
        *,
        account_name: str,
        home_window_id: str,
        snapshot_id: str,
        sequence: int,
    ) -> HomeArticleTarget | None:
        raw_title = normalize_window_text(
            getattr(card, "raw_title", "") or getattr(card, "title", "")
        )
        title_display = normalize_window_text(
            getattr(card, "title", "") or raw_title
        )
        article_date = normalize_window_text(
            getattr(card, "published_date", "") or getattr(card, "date_text", "")
        )
        visible_rect = getattr(card, "visible_rect", None)
        card_rect = getattr(card, "card_rect", None)
        click_point = getattr(card, "click_point", None)
        if (
            not raw_title
            or not article_date
            or not _valid_rect(visible_rect)
            or int(visible_rect[3]) - int(visible_rect[1]) < 10
            or not _valid_rect(card_rect)
            or not isinstance(click_point, (tuple, list))
            or len(click_point) < 2
        ):
            return None
        expected_center = _rect_center(tuple(int(value) for value in visible_rect))
        actual_click = (int(click_point[0]), int(click_point[1]))
        if actual_click != expected_center:
            actual_click = expected_center
        return HomeArticleTarget(
            sequence=sequence,
            account_name=account_name,
            article_date=article_date,
            title_raw=raw_title,
            title_display=title_display,
            card_rect=tuple(int(value) for value in card_rect),
            visible_rect=tuple(int(value) for value in visible_rect),
            click_point=actual_click,
            source_snapshot_id=snapshot_id,
            home_window_id=str(home_window_id or ""),
        )

    def _emit(self, event: str, message: str, **details: Any) -> None:
        if self._trace is None:
            return
        try:
            self._trace({"event": event, "message": message, "details": details})
        except Exception:
            return


def _article_identity(target: ArticleTarget) -> tuple[str, str]:
    return (
        normalize_window_text(target.published_date or target.date_text).casefold(),
        normalize_window_text(target.raw_title or target.title).casefold(),
    )


def _home_identity(target: HomeArticleTarget) -> tuple[str, str]:
    return (
        normalize_window_text(target.article_date).casefold(),
        normalize_window_text(target.title_raw).casefold(),
    )


def _more_trigger_key(target: Any) -> tuple[str, str]:
    """按日期组标识“余下 xx 篇”入口，避免同组展开按钮反复点击。"""

    return (
        normalize_window_text(str(getattr(target, "date_text", "") or "")),
        normalize_window_text(str(getattr(target, "published_date", "") or "")),
    )


def _visible_rect(target: ArticleTarget) -> tuple[int, int, int, int] | None:
    values = [target.metric_rect, target.title_rect, target.date_rect]
    valid = [value for value in values if _valid_rect(value)]
    if not valid:
        return None
    left = max(rect[0] for rect in valid)
    top = max(rect[1] for rect in valid)
    right = min(rect[2] for rect in valid)
    bottom = max(rect[3] for rect in valid)
    # 指标矩形是最可靠的可点击区域；如果多个矩形不相交，退回指标/标题的并集。
    if right <= left or bottom <= top:
        metric = target.metric_rect if _valid_rect(target.metric_rect) else valid[0]
        return metric
    return left, top, right, bottom


def _card_rect(
    target: ArticleTarget,
    visible_rect: tuple[int, int, int, int],
) -> tuple[int, int, int, int]:
    valid = [value for value in (target.date_rect, target.title_rect, target.metric_rect) if _valid_rect(value)]
    if not valid:
        return visible_rect
    return (
        min(rect[0] for rect in valid),
        min(rect[1] for rect in valid),
        max(rect[2] for rect in valid),
        max(rect[3] for rect in valid),
    )


def _valid_rect(value: Any) -> bool:
    return (
        isinstance(value, (tuple, list))
        and len(value) >= 4
        and int(value[2]) > int(value[0])
        and int(value[3]) > int(value[1])
    )


def _date_snapshot_signature(snapshot: Any) -> tuple[tuple[Any, ...], ...]:
    """仅用于判断日期定位是否有新内容，不把坐标作为文章身份。"""

    visible_groups = getattr(snapshot, "visible_groups", None)
    groups = (
        visible_groups
        if visible_groups is not None
        else (getattr(snapshot, "groups", ()) or ())
    )

    return tuple(
        (
            tuple(getattr(group, "runtime_id", ()) or ()),
            str(getattr(group, "published_date", "") or ""),
            str(getattr(group, "date_text", "") or ""),
            tuple(getattr(group, "visible_rect", None) or ()),
            tuple(getattr(group, "group_rect", None) or ()),
        )
        for group in groups
    )


def _snapshot_visible_cards(snapshot: Any) -> tuple[Any, ...]:
    visible = tuple(getattr(snapshot, "visible_cards", ()) or ())
    if visible:
        return visible
    groups = getattr(snapshot, "groups", ()) or ()
    return tuple(
        card
        for group in groups
        for card in (getattr(group, "cards", ()) or ())
        if getattr(card, "visible_rect", None) is not None
    )


def _snapshot_signature(snapshot: Any) -> tuple[Any, ...]:
    """用于滚动后判断 UIA 是否出现新内容；不用于文章身份去重。"""

    cards = tuple(getattr(snapshot, "all_cards", ()) or ())
    if not cards:
        cards = _snapshot_visible_cards(snapshot)
    return tuple(
        (
            normalize_window_text(
                str(getattr(card, "published_date", "") or getattr(card, "date_text", ""))
            ),
            normalize_window_text(
                str(getattr(card, "raw_title", "") or getattr(card, "title", ""))
            ),
            tuple(getattr(card, "visible_rect", ()) or ()),
        )
        for card in cards
    )


def _rect_center(rect: tuple[int, int, int, int]) -> tuple[int, int]:
    return ((rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)


__all__ = [
    "HomeScanBatch",
    "HomeScanError",
    "HomeScanService",
    "HomeScanSession",
]
