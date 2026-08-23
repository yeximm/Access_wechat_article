from __future__ import annotations

from collections.abc import Mapping
from datetime import datetime
from threading import RLock
from typing import Any, Callable

from .main_flow_models import SingleArticleReceipt


class MainFlowState:
    """主流程线程安全状态；不负责窗口、代理或数据库操作。"""

    def __init__(
        self,
        *,
        task_id: str,
        target_count: int,
        now: Callable[[], datetime] = datetime.now,
    ) -> None:
        self.task_id = task_id
        self.target_count = max(0, int(target_count))
        self._now = now
        self._lock = RLock()
        self._status = "idle"
        self._message = ""
        self._current_action = "点击开始运行后，将从桌面主页窗口读取"
        self._account_name = "等待识别"
        self._task_info = "待获取"
        self._progress_done = 0
        self._skipped_count = 0
        self._error_count = 0
        self._latest_error = ""
        # 进程生命周期按唯一 ID 维护，避免重复回执导致计数漂移。
        self._active_processes: dict[str, str] = {}
        self._started_process_ids: set[str] = set()
        self._seen_event_ids: set[str] = set()
        self._average_total_seconds = 0.0
        self._average_count = 0
        self._detail_success_count = 0
        self._progress_fingerprints: set[str] = set()
        self._detail_fingerprints: set[str] = set()
        self._error_fingerprints: set[str] = set()
        self._skipped_fingerprints: set[str] = set()
        self._handled_fingerprints: set[str] = set()
        self._proxy_status_label = "空闲"

    def start(self) -> None:
        with self._lock:
            self._status = "running"
            self._current_action = "正在准备主流程"
            self._message = "主流程已启动"
            self._proxy_status_label = "空闲"

    def set_starting(self) -> None:
        with self._lock:
            self._status = "starting"
            self._current_action = "正在准备主流程"
            self._message = "主流程正在启动"

    def request_stop(self) -> None:
        with self._lock:
            if self._status in {"running", "starting"}:
                self._status = "stopping"
                self._current_action = "正在清理运行资源"
                self._message = "已请求停止主流程"

    def complete(self, message: str = "主流程已完成") -> None:
        with self._lock:
            self._status = "completed"
            self._current_action = "主流程已完成"
            self._message = message
            self._proxy_status_label = "已释放"

    def fail(self, message: str, *, count_error: bool = True) -> None:
        with self._lock:
            self._status = "failed"
            self._current_action = "主流程异常"
            self._message = message
            self._latest_error = message
            if count_error:
                self._error_count += 1
            self._proxy_status_label = "已释放"

    def cancel(self, message: str = "主流程已停止") -> None:
        with self._lock:
            self._status = "cancelled"
            self._current_action = "主流程已停止"
            self._message = message
            self._proxy_status_label = "已释放"

    def set_action(self, action: str) -> None:
        value = str(action or "").strip()
        if value and not value.startswith("正在") and not value.startswith("主流程"):
            value = f"正在{value}"
        with self._lock:
            self._current_action = value

    def set_account_name(self, account_name: str) -> None:
        with self._lock:
            self._account_name = str(account_name or "").strip() or "等待识别"

    def set_task_info(self, task_info: str) -> None:
        with self._lock:
            self._task_info = str(task_info or "").strip() or "待获取"

    def set_proxy_status(self, label: str) -> None:
        with self._lock:
            self._proxy_status_label = str(label or "空闲").strip() or "空闲"

    def begin_child_process(
        self,
        process_id: str | None = None,
        *,
        process_type: str = "article",
    ) -> bool:
        process_key = str(process_id or "").strip()
        if not process_key:
            return False
        process_kind = str(process_type or "article").strip() or "article"
        with self._lock:
            if process_key in self._active_processes:
                return False
            self._active_processes[process_key] = process_kind
            self._started_process_ids.add(process_key)
            return True

    def end_child_process(
        self,
        process_id: str | None = None,
        *,
        process_type: str | None = None,
    ) -> bool:
        del process_type
        process_key = str(process_id or "").strip()
        if not process_key:
            return False
        with self._lock:
            if process_key not in self._active_processes:
                return False
            self._active_processes.pop(process_key, None)
            return True

    def handle_event(self, event: Mapping[str, Any] | Any) -> bool:
        """聚合主流程事件；坏事件只被丢弃，不影响采集主线。"""

        if isinstance(event, SingleArticleReceipt):
            return self.handle_receipt(event)
        if not isinstance(event, Mapping):
            return False

        event_type = str(event.get("type") or event.get("eventType") or "").strip().lower()
        event_id = str(event.get("eventId") or event.get("event_id") or "").strip()
        if event_id:
            with self._lock:
                if event_id in self._seen_event_ids:
                    return False
                self._seen_event_ids.add(event_id)

        if event_type == "process":
            action = str(event.get("action") or "").strip().lower()
            process_id = event.get("processId") or event.get("process_id")
            process_type = str(
                event.get("processType") or event.get("process_type") or "article"
            )
            if action in {"started", "start", "begin"}:
                changed = self.begin_child_process(process_id, process_type=process_type)
                if process_type == "foreground" and changed:
                    self.set_proxy_status("正在使用")
                return changed
            if action in {"finished", "finish", "stopped", "ended", "end"}:
                changed = self.end_child_process(process_id, process_type=process_type)
                if process_type == "foreground" and changed:
                    self.set_proxy_status("已释放")
                return changed
            return False

        if event_type == "article":
            receipt = _receipt_from_event(event, default_task_id=self.task_id)
            return self.handle_receipt(receipt)

        if event_type == "flow":
            action = event.get("currentAction") or event.get("current_action")
            if action:
                self.set_action(str(action))
            return True
        return False

    def handle_receipt(self, receipt: SingleArticleReceipt) -> bool:
        """处理单篇的任一阶段回执；同一文章的不同阶段不会互相覆盖。"""
        with self._lock:
            # 主流程进入终态后，异步 Huey 回执可能仍在路上。终态是一次性
            # 结论，迟到的成功、失败或取消回执不能重新打开状态或改写统计。
            if self._status in {"completed", "failed", "cancelled"}:
                return False
            fingerprint = str(receipt.target_fingerprint or "")
            if not fingerprint:
                return False
            self._handled_fingerprints.add(fingerprint)

            progress_counted = bool(receipt.progress_counted)
            # 兼容旧的单次最终回执：没有阶段字段时，详情成功仍代表已做一篇。
            if (
                not progress_counted
                and receipt.phase == "finished"
                and receipt.status == "success"
                and receipt.article_saved
            ):
                progress_counted = True
            if progress_counted and fingerprint not in self._progress_fingerprints:
                self._progress_fingerprints.add(fingerprint)
                self._progress_done += 1

            if (
                receipt.status == "success"
                and receipt.article_saved
                and receipt.detail_status in {"success", "pending"}
                and fingerprint not in self._detail_fingerprints
            ):
                self._detail_fingerprints.add(fingerprint)
                self._detail_success_count += 1
                duration = max(0.0, float(receipt.duration_seconds))
                self._average_total_seconds += duration
                self._average_count += 1
            elif receipt.status == "skipped_collected":
                if fingerprint not in self._skipped_fingerprints:
                    self._skipped_fingerprints.add(fingerprint)
                    self._skipped_count += 1
            if (
                receipt.status == "failed"
                or receipt.comments_status == "failed"
                or receipt.offline_status == "failed"
            ) and fingerprint not in self._error_fingerprints:
                self._error_fingerprints.add(fingerprint)
                self._error_count += 1
                self._latest_error = (
                    receipt.error_detail or receipt.message or "单篇任务失败"
                )
            elif receipt.status == "cancelled":
                self._status = "stopping"
            return True

    def snapshot(self) -> dict[str, Any]:
        with self._lock:
            average = (
                self._average_total_seconds / self._average_count
                if self._average_count
                else None
            )
            total_label = "全部" if self.target_count == 0 else str(self.target_count)
            percent = (
                min(100, round(self._progress_done / self.target_count * 100))
                if self.target_count
                else 0
            )
            return {
                "status": self._status,
                "message": self._message,
                "currentAction": self._current_action,
                "accountName": self._account_name,
                "taskInfo": self._task_info,
                "proxyStatusLabel": self._proxy_status_label,
                "progressDone": self._progress_done,
                "processedCount": self._progress_done,
                "detailSuccessCount": self._detail_success_count,
                "progressTotalLabel": total_label,
                "progressPercent": percent,
                "averageArticleSeconds": average,
                "averageArticleDurationLabel": "待统计" if average is None else f"{average:.1f} 秒",
                "activeWorkerCount": len(self._active_processes),
                "totalWorkerCount": len(self._started_process_ids),
                "activeProcessCount": len(self._active_processes),
                "totalProcessCount": len(self._started_process_ids),
                "activeProcessTypes": _process_type_counts(self._active_processes),
                "errorCount": self._error_count,
                "latestError": self._latest_error,
                "skippedCount": self._skipped_count,
                "handledFingerprints": sorted(self._handled_fingerprints),
                "updatedAt": self._now().isoformat(timespec="milliseconds"),
            }


def _process_type_counts(processes: Mapping[str, str]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for process_type in processes.values():
        counts[process_type] = counts.get(process_type, 0) + 1
    return counts


def _receipt_from_event(
    event: Mapping[str, Any],
    *,
    default_task_id: str,
) -> SingleArticleReceipt:
    def value(snake: str, camel: str, default: Any = None) -> Any:
        return event.get(snake, event.get(camel, default))

    status = str(value("status", "status", "failed"))
    if status in {"completed", "ready-to-continue"}:
        status = "success"
    if status in {"skipped", "skipped-collected"}:
        status = "skipped_collected"
    if status == "canceled":
        status = "cancelled"
    return SingleArticleReceipt(
        task_id=str(value("task_id", "taskId", default_task_id) or default_task_id),
        target_fingerprint=str(
            value("target_fingerprint", "targetFingerprint", "") or ""
        ),
        status=status,  # type: ignore[arg-type]
        phase=str(value("phase", "phase", "finished")),  # type: ignore[arg-type]
        foreground_done=bool(value("foreground_done", "foregroundDone", False)),
        progress_counted=bool(value("progress_counted", "progressCounted", False)),
        tab_closed=bool(value("tab_closed", "tabClosed", False)),
        article_saved=bool(value("article_saved", "articleSaved", False)),
        detail_status=str(value("detail_status", "detailStatus", "pending")),
        comments_status=str(value("comments_status", "commentsStatus", "not_requested")),
        offline_status=str(value("offline_status", "offlineStatus", "not_requested")),
        archive_dir=value("archive_dir", "archiveDir"),
        article_title=str(value("article_title", "articleTitle", "") or ""),
        message=str(value("message", "message", "") or ""),
        error_stage=value("error_stage", "errorStage"),
        error_detail=value("error_detail", "errorDetail"),
        duration_seconds=float(value("duration_seconds", "durationSeconds", 0) or 0),
    )
