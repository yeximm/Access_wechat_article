from __future__ import annotations

from collections.abc import Callable
import ctypes
import platform
import time
from typing import Any

from src.domain.models import ArticleTarget
from src.modules.window.home_window_activation import (
    HomeWindowNotClickableError,
    WindowsHomeWindowGuard,
)
from src.modules.window.window_models import WindowInfo


WM_MOUSEWHEEL = 0x020A
WHEEL_DELTA = 120
GA_ROOT = 2
CHROMIUM_RENDER_WINDOW_CLASS = "Chrome_RenderWidgetHostHWND"


class WechatHomeScroller:
    """滚动前激活微信主页并发送滚轮消息，不移动用户的系统鼠标。"""

    def __init__(
        self,
        *,
        wheel_steps: int = 5,
        user32: Any | None = None,
        kernel32: Any | None = None,
        platform_name: str | None = None,
        activation_wait_seconds: float = 0.05,
        wheel_message_interval_seconds: float = 0.02,
        wheel_dispatch_settle_seconds: float = 0.05,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self._wheel_steps = max(1, int(wheel_steps))
        self._user32 = user32
        self._kernel32 = kernel32
        self._platform_name = platform_name or platform.system()
        self._wheel_message_interval_seconds = max(
            0.0,
            float(wheel_message_interval_seconds),
        )
        self._wheel_dispatch_settle_seconds = max(
            0.0,
            float(wheel_dispatch_settle_seconds),
        )
        self._sleep = sleep
        self._home_guard = WindowsHomeWindowGuard(
            user32=user32,
            kernel32=kernel32,
            platform_name=self._platform_name,
            activation_wait_seconds=activation_wait_seconds,
            sleep=sleep,
        )

    @property
    def wheel_steps(self) -> int:
        """返回普通向下滚动实际使用的默认步数。"""

        return self._wheel_steps

    def scroll_down(
        self,
        home_window: WindowInfo,
        *,
        visible_targets: list[ArticleTarget],
    ) -> bool:
        return self.scroll(
            home_window,
            visible_targets=visible_targets,
            direction="down",
        )

    def scroll_up(
        self,
        home_window: WindowInfo,
        *,
        visible_targets: list[ArticleTarget],
    ) -> bool:
        return self.scroll(
            home_window,
            visible_targets=visible_targets,
            direction="up",
        )

    def scroll(
        self,
        home_window: WindowInfo,
        *,
        visible_targets: list[ArticleTarget],
        direction: str,
        wheel_steps: int | None = None,
    ) -> bool:
        if self._platform_name != "Windows" or home_window.handle <= 0:
            return False
        normalized_direction = str(direction).strip().lower()
        if normalized_direction not in {"up", "down"}:
            raise ValueError(f"不支持的主页滚动方向：{direction}")

        user32 = self._user32 or ctypes.windll.user32
        _enable_per_monitor_dpi_awareness(user32)
        try:
            self._home_guard.activate(home_window)
        except HomeWindowNotClickableError:
            return False
        if not _foreground_belongs_to_window(
            user32,
            home_window.handle,
        ):
            return False

        target_handle = _resolve_home_scroll_target(
            user32,
            home_handle=home_window.handle,
        )
        x, y = _scroll_point(home_window, visible_targets)
        steps = (
            self._wheel_steps
            if wheel_steps is None
            else max(1, int(wheel_steps))
        )
        wheel_delta = WHEEL_DELTA if normalized_direction == "up" else -WHEEL_DELTA

        # Chromium 对一条大 delta 的后台滚轮消息处理不稳定，按标准滚轮刻度逐条发送。
        for index in range(steps):
            posted = bool(
                user32.PostMessageW(
                    target_handle,
                    WM_MOUSEWHEEL,
                    _make_wparam(wheel_delta),
                    _make_lparam(x, y),
                )
            )
            if not posted:
                return False
            if (
                index + 1 < steps
                and self._wheel_message_interval_seconds > 0
            ):
                self._sleep(self._wheel_message_interval_seconds)

        # 滚轮消息刚入队时给 Chromium 留出一次处理时间；滚动后不再恢复原前台窗口，
        # 避免日期定位连续滚动时在微信和用户窗口之间反复抢焦点。
        if self._wheel_dispatch_settle_seconds > 0:
            self._sleep(self._wheel_dispatch_settle_seconds)
        return True


def _scroll_point(
    home_window: WindowInfo,
    visible_targets: list[ArticleTarget],
) -> tuple[int, int]:
    # 文章坐标来自上一次 UIA 快照，滚动前可能已经过期；固定使用主页内容区内的点。
    del visible_targets
    left, top, right, bottom = home_window.rect
    return (left + right) // 2, top + max(1, (bottom - top) * 2 // 3)


def _resolve_home_scroll_target(user32: Any, *, home_handle: int) -> int:
    """只选择微信主页内部的 Chromium 渲染窗口。"""

    try:
        render_handle = int(
            user32.FindWindowExW(
                int(home_handle),
                0,
                CHROMIUM_RENDER_WINDOW_CLASS,
                None,
            )
            or 0
        )
        if render_handle > 0:
            root_handle = int(user32.GetAncestor(render_handle, GA_ROOT) or 0)
            if root_handle == int(home_handle):
                return render_handle
    except Exception:
        pass
    return int(home_handle)


def _get_foreground_window(user32: Any) -> int:
    try:
        return int(user32.GetForegroundWindow() or 0)
    except Exception:
        return 0


def _foreground_belongs_to_window(user32: Any, handle: int) -> bool:
    return _belongs_to_window(
        user32,
        handle,
        _get_foreground_window(user32),
    )


def _belongs_to_window(user32: Any, handle: int, owner_handle: int) -> bool:
    if owner_handle <= 0:
        return False
    if owner_handle == int(handle):
        return True
    try:
        if bool(user32.IsChild(int(handle), int(owner_handle))):
            return True
        return int(user32.GetAncestor(int(owner_handle), GA_ROOT) or 0) == int(handle)
    except Exception:
        return False


def _enable_per_monitor_dpi_awareness(user32: Any) -> None:
    try:
        setter = getattr(user32, "SetThreadDpiAwarenessContext", None)
        if setter is not None:
            setter(ctypes.c_void_p(-4))
    except Exception:
        pass


def _make_wparam(delta: int) -> int:
    return (int(delta) & 0xFFFF) << 16


def _make_lparam(x: int, y: int) -> int:
    return (int(y) & 0xFFFF) << 16 | (int(x) & 0xFFFF)
