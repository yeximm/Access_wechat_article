from __future__ import annotations

from collections import deque
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
import os
from threading import Event, RLock, Thread
from typing import Any

import psutil


@dataclass(frozen=True, slots=True)
class ProcessResourceSample:
    """单个进程在某个采样时刻的资源快照。"""

    pid: int
    cpu_time_seconds: float
    rss_bytes: int


@dataclass(frozen=True, slots=True)
class _RegisteredProcess:
    pid: int
    label: str
    create_time: float | None


def _clamp_percent(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = 0.0
    return round(min(100.0, max(0.0, number)), 1)


def _percent_label(value: float) -> str:
    return f"{_clamp_percent(value):.1f}%"


class ProcessTreeResourceMonitor:
    """采样当前软件进程树的 CPU 与物理内存占比。

    口径说明：
    - CPU：主进程 + 已登记子进程 + 递归子进程的 CPU 时间差，占整机 CPU 总能力的比例。
    - 内存：上述进程 RSS 物理内存总和，占整机物理内存的比例。
    """

    def __init__(
        self,
        *,
        root_pid: int | None = None,
        history_sample_count: int = 120,
        sample_interval_seconds: float = 0.5,
        now: Callable[[], datetime] = datetime.now,
        process_tree_reader: Callable[[set[int]], list[ProcessResourceSample]] | None = None,
        logical_cpu_count_reader: Callable[[], int] | None = None,
        total_memory_reader: Callable[[], int] | None = None,
        process_create_time_reader: Callable[[int], float | None] | None = None,
        thread_factory: Callable[..., Thread] = Thread,
    ) -> None:
        self._root_pid = int(root_pid or os.getpid())
        self._history_sample_count = max(1, int(history_sample_count or 120))
        self._sample_interval_seconds = max(0.1, float(sample_interval_seconds or 0.5))
        self._now = now
        self._process_tree_reader = process_tree_reader or _read_psutil_process_tree
        self._logical_cpu_count_reader = logical_cpu_count_reader or (
            lambda: psutil.cpu_count(logical=True) or 1
        )
        self._total_memory_reader = total_memory_reader or (
            lambda: int(psutil.virtual_memory().total)
        )
        self._process_create_time_reader = process_create_time_reader or _safe_process_create_time
        self._thread_factory = thread_factory
        self._lock = RLock()
        self._stop_event = Event()
        self._thread: Thread | None = None
        self._history: deque[dict[str, Any]] = deque(maxlen=self._history_sample_count)
        self._registered: dict[int, _RegisteredProcess] = {}
        self._previous_cpu_times: dict[int, float] = {}
        self._previous_timestamp: datetime | None = None

    @property
    def sample_interval_seconds(self) -> float:
        return self._sample_interval_seconds

    def start(self) -> None:
        """启动后台采样线程；重复调用是安全的。"""
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return
            self._stop_event.clear()
            self._thread = self._thread_factory(
                target=self._run_loop,
                name="awa-process-resource-monitor",
                daemon=True,
            )
            self._thread.start()

    def stop(self, *, timeout_seconds: float = 1.0) -> None:
        """停止后台采样线程，供后端退出清理使用。"""
        thread: Thread | None
        with self._lock:
            thread = self._thread
            self._thread = None
        if thread is None:
            return
        self._stop_event.set()
        thread.join(max(0.0, float(timeout_seconds)))

    def register_process(self, pid: int | None, *, label: str = "") -> None:
        """登记由当前软件启动的子进程 PID，便于进程树采样补充跟踪。"""
        try:
            safe_pid = int(pid or 0)
        except (TypeError, ValueError):
            return
        if safe_pid <= 0:
            return
        create_time = self._process_create_time_reader(safe_pid)
        with self._lock:
            self._registered[safe_pid] = _RegisteredProcess(
                pid=safe_pid,
                label=str(label or ""),
                create_time=create_time,
            )

    def sample_once(self) -> dict[str, Any]:
        """同步采样一次；测试和首次状态读取可直接调用。"""
        timestamp = self._now()
        root_pids = self._live_root_pids()
        samples = self._process_tree_reader(root_pids)
        cpu_percent = self._calculate_cpu_percent(timestamp, samples)
        memory_used_bytes = self._calculate_memory_used_bytes(samples)
        memory_percent = self._calculate_memory_percent(memory_used_bytes)
        sample = {
            "timestamp": timestamp.isoformat(timespec="milliseconds"),
            "cpuPercent": cpu_percent,
            "memoryPercent": memory_percent,
            "memoryUsedBytes": memory_used_bytes,
            "processCount": len({item.pid for item in samples}),
        }
        with self._lock:
            self._history.append(sample)
            history = list(self._history)
        return self._snapshot_from_sample(sample, history)

    def snapshot(self) -> dict[str, Any]:
        """返回最近一次后台采样结果；尚无采样时即时补一帧。"""
        with self._lock:
            history = list(self._history)
        if not history:
            return self.sample_once()
        return self._snapshot_from_sample(history[-1], history)

    def _run_loop(self) -> None:
        while not self._stop_event.is_set():
            try:
                self.sample_once()
            except Exception:
                # 状态图不能反向影响主流程；采样异常时跳过本帧，下一轮继续。
                pass
            self._stop_event.wait(self._sample_interval_seconds)

    def _live_root_pids(self) -> set[int]:
        with self._lock:
            registered = list(self._registered.values())
        live = {self._root_pid}
        stale: list[int] = []
        for item in registered:
            if item.pid == self._root_pid:
                continue
            create_time = self._process_create_time_reader(item.pid)
            if create_time is None:
                stale.append(item.pid)
                continue
            if item.create_time is not None and abs(create_time - item.create_time) > 0.001:
                stale.append(item.pid)
                continue
            live.add(item.pid)
        if stale:
            with self._lock:
                for pid in stale:
                    self._registered.pop(pid, None)
        return live

    def _calculate_cpu_percent(
        self,
        timestamp: datetime,
        samples: list[ProcessResourceSample],
    ) -> float:
        current_cpu_times = {item.pid: max(0.0, float(item.cpu_time_seconds)) for item in samples}
        with self._lock:
            previous_timestamp = self._previous_timestamp
            previous_cpu_times = dict(self._previous_cpu_times)
            self._previous_timestamp = timestamp
            self._previous_cpu_times = current_cpu_times
        if previous_timestamp is None:
            return 0.0
        elapsed_seconds = max(0.001, (timestamp - previous_timestamp).total_seconds())
        cpu_delta = 0.0
        for pid, cpu_time in current_cpu_times.items():
            previous = previous_cpu_times.get(pid, cpu_time)
            cpu_delta += max(0.0, cpu_time - previous)
        logical_cpus = max(1, int(self._logical_cpu_count_reader() or 1))
        return _clamp_percent((cpu_delta / (elapsed_seconds * logical_cpus)) * 100.0)

    def _calculate_memory_used_bytes(self, samples: list[ProcessResourceSample]) -> int:
        """汇总当前软件进程树的 RSS，用于悬浮提示展示真实内存占用。"""

        return sum(max(0, int(item.rss_bytes)) for item in samples)

    def _calculate_memory_percent(self, memory_used_bytes: int) -> float:
        total_memory = max(1, int(self._total_memory_reader() or 1))
        return _clamp_percent((max(0, int(memory_used_bytes)) / total_memory) * 100.0)

    def _snapshot_from_sample(
        self,
        sample: dict[str, Any],
        history: list[dict[str, Any]],
    ) -> dict[str, Any]:
        cpu_percent = _clamp_percent(sample.get("cpuPercent"))
        memory_percent = _clamp_percent(sample.get("memoryPercent"))
        memory_used_bytes = max(0, int(sample.get("memoryUsedBytes") or 0))
        return {
            "cpuPercent": cpu_percent,
            "memoryPercent": memory_percent,
            "memoryUsedBytes": memory_used_bytes,
            "cpuLabel": _percent_label(cpu_percent),
            "memoryLabel": _percent_label(memory_percent),
            "processCount": int(sample.get("processCount") or 0),
            "historySampleCount": self._history_sample_count,
            "sampleIntervalSeconds": self._sample_interval_seconds,
            "history": history,
            "updatedAt": str(sample.get("timestamp") or self._now().isoformat(timespec="milliseconds")),
        }


def _read_psutil_process_tree(root_pids: set[int]) -> list[ProcessResourceSample]:
    result: list[ProcessResourceSample] = []
    visited: set[int] = set()
    for root_pid in root_pids:
        try:
            root = psutil.Process(int(root_pid))
            processes = [root, *root.children(recursive=True)]
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, OSError):
            continue
        for process in processes:
            try:
                pid = int(process.pid)
            except (TypeError, ValueError):
                continue
            if pid in visited:
                continue
            try:
                cpu_times = process.cpu_times()
                memory = process.memory_info()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, OSError):
                continue
            visited.add(pid)
            result.append(
                ProcessResourceSample(
                    pid=pid,
                    cpu_time_seconds=float(cpu_times.user) + float(cpu_times.system),
                    rss_bytes=int(memory.rss),
                )
            )
    return result


def _safe_process_create_time(pid: int) -> float | None:
    try:
        return float(psutil.Process(int(pid)).create_time())
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, OSError, ValueError):
        return None


__all__ = ["ProcessResourceSample", "ProcessTreeResourceMonitor"]
