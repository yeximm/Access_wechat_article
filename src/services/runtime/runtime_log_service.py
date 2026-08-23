from __future__ import annotations

from collections import deque
from collections.abc import Mapping
from datetime import datetime
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from threading import RLock
from typing import Any, Callable


SUMMARY_LOG_FILE_NAME = "runtime-summary.log"
DETAIL_LOG_FILE_NAME = "runtime-detail.log"
DEFAULT_SUMMARY_LIMIT = 100
DEFAULT_MAX_BYTES = 10 * 1024 * 1024
DEFAULT_BACKUP_COUNT = 5


class RuntimeLogService:
    """把主界面摘要和排错明细写入同一配置目录下的两个日志文件。"""

    def __init__(
        self,
        *,
        log_dir: str | Path,
        level: str = "INFO",
        summary_limit: int = DEFAULT_SUMMARY_LIMIT,
        max_bytes: int = DEFAULT_MAX_BYTES,
        backup_count: int = DEFAULT_BACKUP_COUNT,
        redactions: Mapping[str, str] | None = None,
        now: Callable[[], datetime] = datetime.now,
    ) -> None:
        self.log_dir = Path(log_dir).resolve()
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.summary_log_path = self.log_dir / SUMMARY_LOG_FILE_NAME
        self.detail_log_path = self.log_dir / DETAIL_LOG_FILE_NAME
        self._level = _normalize_level(level)
        self._summary_entries: deque[dict[str, Any]] = deque(
            maxlen=max(1, int(summary_limit))
        )
        self._redactions = tuple(
            sorted(
                (
                    (str(source), str(replacement))
                    for source, replacement in (redactions or {}).items()
                    if str(source)
                ),
                key=lambda item: len(item[0]),
                reverse=True,
            )
        )
        self._now = now
        self._error_second = ""
        self._error_sequence = 0
        self._lock = RLock()
        self._closed = False
        self._summary_logger = self._build_logger(
            name=f"awa.runtime.summary.{id(self)}",
            path=self.summary_log_path,
            formatter=_RedactingFormatter(
                "%(asctime)s | %(display_level)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
                redactions=self._redactions,
            ),
            max_bytes=max_bytes,
            backup_count=backup_count,
        )
        self._detail_logger = self._build_logger(
            name=f"awa.runtime.detail.{id(self)}",
            path=self.detail_log_path,
            formatter=_RedactingFormatter(
                "%(asctime)s | %(display_level)s | %(source)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
                redactions=self._redactions,
            ),
            max_bytes=max_bytes,
            backup_count=backup_count,
        )

    def write_summary(
        self,
        level: str,
        message: str,
        *,
        source: str = "runtime",
        channel: str | None = None,
        phase: str | None = None,
        task_index: int | str | None = None,
        article_task_id: str | None = None,
        article_title: str | None = None,
        context: Mapping[str, Any] | None = None,
        exception: BaseException | None = None,
        error_id: str | None = None,
    ) -> dict[str, Any]:
        return self._write(
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
            error_id=error_id,
            summary=True,
            generate_error_id=True,
        )

    def write_detail(
        self,
        level: str,
        message: str,
        *,
        source: str = "runtime",
        channel: str | None = None,
        phase: str | None = None,
        task_index: int | str | None = None,
        article_task_id: str | None = None,
        article_title: str | None = None,
        context: Mapping[str, Any] | None = None,
        exception: BaseException | None = None,
        error_id: str | None = None,
    ) -> dict[str, Any]:
        return self._write(
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
            error_id=error_id,
            summary=False,
            generate_error_id=False,
        )

    def write_error(
        self,
        message: str,
        *,
        source: str = "runtime",
        channel: str | None = None,
        phase: str | None = None,
        task_index: int | str | None = None,
        article_task_id: str | None = None,
        article_title: str | None = None,
        context: Mapping[str, Any] | None = None,
        exception: BaseException | None = None,
        summary: bool = True,
    ) -> dict[str, Any]:
        return self._write(
            "ERROR",
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
            generate_error_id=True,
        )

    def recent_summary(self, limit: int = DEFAULT_SUMMARY_LIMIT) -> list[dict[str, Any]]:
        safe_limit = max(1, min(int(limit), self._summary_entries.maxlen or 1))
        with self._lock:
            return [dict(item) for item in list(self._summary_entries)[-safe_limit:]]

    def close(self) -> None:
        with self._lock:
            if self._closed:
                return
            self._closed = True
            for logger in (self._summary_logger, self._detail_logger):
                for handler in tuple(logger.handlers):
                    handler.flush()
                    handler.close()
                    logger.removeHandler(handler)

    def _write(
        self,
        level: str,
        message: str,
        *,
        source: str,
        channel: str | None,
        phase: str | None,
        task_index: int | str | None,
        article_task_id: str | None,
        article_title: str | None,
        context: Mapping[str, Any] | None,
        exception: BaseException | None,
        error_id: str | None = None,
        summary: bool,
        generate_error_id: bool,
    ) -> dict[str, Any]:
        normalized_level = _normalize_level(level)
        created_at = self._now()
        with self._lock:
            resolved_error_id = error_id
            if normalized_level == "ERROR" and generate_error_id and not resolved_error_id:
                resolved_error_id = self._next_error_id(created_at)
            clean_message = self._sanitize(str(message or "").strip() or "未提供日志内容")
            display_message = (
                f"[{resolved_error_id}] {clean_message}"
                if resolved_error_id
                else clean_message
            )
            entry = {
                "level": normalized_level,
                "message": display_message,
                "source": str(source or "runtime"),
                "createdAt": created_at.strftime("%Y-%m-%d %H:%M:%S"),
                "errorId": resolved_error_id or "",
                # channel/phase/task 字段只服务前台展示；旧调用不传时默认进左侧软件活动栏。
                "channel": _normalize_channel(channel),
                "phase": str(phase or "").strip(),
                "taskIndex": _normalize_task_index(task_index),
                "articleTaskId": str(article_task_id or "").strip(),
                "articleTitle": self._sanitize(str(article_title or "").strip()),
            }
            if not self._is_enabled(normalized_level):
                return entry

            logging_level = _logging_level(normalized_level)
            extra = {
                "display_level": normalized_level,
                "source": entry["source"],
            }
            detail_message = display_message + self._format_context(context)
            exc_info = (
                (type(exception), exception, exception.__traceback__)
                if exception is not None
                else None
            )
            self._detail_logger.log(
                logging_level,
                detail_message,
                extra=extra,
                exc_info=exc_info,
            )
            if summary:
                self._summary_entries.append(entry)
                self._summary_logger.log(
                    logging_level,
                    display_message,
                    extra=extra,
                )
            return entry

    def _format_context(self, context: Mapping[str, Any] | None) -> str:
        if not context:
            return ""
        items = []
        for key, value in context.items():
            if isinstance(value, (dict, list, tuple)):
                rendered = json.dumps(value, ensure_ascii=False, default=str)
            else:
                rendered = str(value)
            rendered = self._sanitize(rendered.replace("\r", " ").replace("\n", " "))
            items.append(f"{key}={rendered}")
        return " | " + " | ".join(items)

    def _sanitize(self, value: str) -> str:
        result = value
        for source, replacement in self._redactions:
            result = result.replace(source, replacement)
        return result

    def _next_error_id(self, created_at: datetime) -> str:
        second = created_at.strftime("%Y%m%d-%H%M%S")
        if second != self._error_second:
            self._error_second = second
            self._error_sequence = 0
        self._error_sequence += 1
        return f"ERR-{second}-{self._error_sequence:03d}"

    def _is_enabled(self, level: str) -> bool:
        return _logging_level(level) >= _logging_level(self._level)

    @staticmethod
    def _build_logger(
        *,
        name: str,
        path: Path,
        formatter: logging.Formatter,
        max_bytes: int,
        backup_count: int,
    ) -> logging.Logger:
        logger = logging.getLogger(name)
        logger.setLevel(logging.DEBUG)
        logger.propagate = False
        handler = RotatingFileHandler(
            path,
            maxBytes=max(1, int(max_bytes)),
            backupCount=max(1, int(backup_count)),
            encoding="utf-8",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        return logger


def _normalize_level(level: str) -> str:
    normalized = str(level or "INFO").strip().upper()
    if normalized == "WARNING":
        normalized = "WARN"
    return normalized if normalized in {"DEBUG", "INFO", "SUCCESS", "WARN", "ERROR"} else "INFO"


def _logging_level(level: str) -> int:
    return {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "SUCCESS": logging.INFO,
        "WARN": logging.WARNING,
        "ERROR": logging.ERROR,
    }[_normalize_level(level)]


def _normalize_channel(channel: str | None) -> str:
    """前台日志分栏字段：未知值统一归到软件活动栏，避免日志丢失。"""

    normalized = str(channel or "system").strip().lower().replace("-", "_")
    return normalized if normalized in {"system", "main_flow", "article_task"} else "system"


def _normalize_task_index(value: int | str | None) -> int | str:
    if value is None or value == "":
        return ""
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return str(value)


class _RedactingFormatter(logging.Formatter):
    """在最终格式化阶段再次脱敏，覆盖异常堆栈中的路径文本。"""

    def __init__(
        self,
        fmt: str,
        *,
        datefmt: str,
        redactions: tuple[tuple[str, str], ...],
    ) -> None:
        super().__init__(fmt, datefmt=datefmt)
        self._redactions = redactions

    def format(self, record: logging.LogRecord) -> str:
        rendered = super().format(record)
        for source, replacement in self._redactions:
            rendered = rendered.replace(source, replacement)
        return rendered


__all__ = [
    "DETAIL_LOG_FILE_NAME",
    "RuntimeLogService",
    "SUMMARY_LOG_FILE_NAME",
]
