from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
from uuid import uuid4


class RuntimeDirectoryError(RuntimeError):
    """应用运行目录无法创建或不可写时抛出，阻止后续数据库/任务启动。"""


@dataclass(frozen=True, slots=True)
class PreparedRuntimeDirectory:
    key: str
    label: str
    path: Path
    created: bool


@dataclass(frozen=True, slots=True)
class RuntimeDirectoryPrepareResult:
    items: tuple[PreparedRuntimeDirectory, ...]

    @property
    def created_count(self) -> int:
        return sum(1 for item in self.items if item.created)


class RuntimeDirectoryService:
    """启动阶段准备 Python 程序自己的运行目录，不接管 mitmproxy 自有目录。"""

    def prepare(self, config: Any) -> RuntimeDirectoryPrepareResult:
        storage = getattr(config, "storage", None)
        temp_dir = self._path(getattr(storage, "temp_dir", None), "temp_dir")
        runtime_dir = temp_dir.parent / "runtime"

        specs = (
            (
                "article_storage_root",
                "文章归档目录",
                getattr(storage, "article_storage_root", None),
            ),
            ("db_dir", "SQLite 数据库目录", getattr(storage, "db_dir", None)),
            ("temp_dir", "临时文件目录", temp_dir),
            ("log_dir", "日志目录", getattr(storage, "log_dir", None)),
            ("runtime_dir", "运行时状态目录", runtime_dir),
            ("huey_queue_dir", "Huey 临时队列目录", runtime_dir / "huey"),
        )
        return RuntimeDirectoryPrepareResult(
            items=tuple(self._prepare_directory(key, label, path) for key, label, path in specs)
        )

    def _prepare_directory(
        self,
        key: str,
        label: str,
        raw_path: Any,
    ) -> PreparedRuntimeDirectory:
        path = self._path(raw_path, key)
        existed = path.is_dir()
        try:
            path.mkdir(parents=True, exist_ok=True)
            if not path.is_dir():
                raise RuntimeDirectoryError(f"{label}不是目录：{path}")
            self._assert_writable(path)
        except Exception as exc:
            if isinstance(exc, RuntimeDirectoryError):
                raise
            raise RuntimeDirectoryError(f"{label}不可用：{path}，原因：{exc}") from exc
        return PreparedRuntimeDirectory(
            key=key,
            label=label,
            path=path,
            created=not existed,
        )

    def _path(self, raw_path: Any, key: str) -> Path:
        if raw_path is None:
            raise RuntimeDirectoryError(f"{key} 目录配置为空")
        return Path(raw_path).resolve()

    def _assert_writable(self, path: Path) -> None:
        # 用一次临时探针确认目录可写，避免启动后才在采集/日志阶段暴露权限问题。
        probe = path / f".awa-runtime-write-probe-{uuid4().hex}"
        try:
            probe.write_text("ok", encoding="utf-8")
        finally:
            probe.unlink(missing_ok=True)


__all__ = [
    "PreparedRuntimeDirectory",
    "RuntimeDirectoryError",
    "RuntimeDirectoryPrepareResult",
    "RuntimeDirectoryService",
]
