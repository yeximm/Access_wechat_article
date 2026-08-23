from __future__ import annotations

from collections import deque
from collections.abc import Callable
from datetime import datetime
from threading import RLock
from typing import Any

import psutil


def _clamp_percent(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = 0.0
    return round(min(100.0, max(0.0, number)), 1)


def _percent_label(value: float) -> str:
    return f"{_clamp_percent(value):.1f}%"


class SystemResourceMonitor:
    """采样当前电脑硬件占用，用于运行状态区域的两行折线图。

    这里只读取 CPU 总使用率和物理内存占比，不参与主流程调度。
    """

    def __init__(
        self,
        *,
        history_sample_count: int = 60,
        now: Callable[[], datetime] = datetime.now,
        cpu_percent_reader: Callable[[], float] | None = None,
        memory_percent_reader: Callable[[], float] | None = None,
    ) -> None:
        self._history_sample_count = max(1, int(history_sample_count or 60))
        self._now = now
        self._cpu_percent_reader = cpu_percent_reader or (lambda: psutil.cpu_percent(interval=None))
        self._memory_percent_reader = memory_percent_reader or (lambda: psutil.virtual_memory().percent)
        self._lock = RLock()
        self._history: deque[dict[str, Any]] = deque(maxlen=self._history_sample_count)

    def snapshot(self) -> dict[str, Any]:
        timestamp = self._now()
        cpu_percent = _clamp_percent(self._cpu_percent_reader())
        memory_percent = _clamp_percent(self._memory_percent_reader())
        sample = {
            "timestamp": timestamp.isoformat(timespec="milliseconds"),
            "cpuPercent": cpu_percent,
            "memoryPercent": memory_percent,
        }
        with self._lock:
            self._history.append(sample)
            history = list(self._history)
        return {
            "cpuPercent": cpu_percent,
            "memoryPercent": memory_percent,
            "cpuLabel": _percent_label(cpu_percent),
            "memoryLabel": _percent_label(memory_percent),
            "historySampleCount": self._history_sample_count,
            "history": history,
            "updatedAt": sample["timestamp"],
        }
