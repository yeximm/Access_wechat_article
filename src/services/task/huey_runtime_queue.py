from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

from src.modules.system.safe_directory_cleaner import SafeDirectoryCleaner


def runtime_huey_queue_dir(temp_root: str | Path) -> Path:
    """根据临时目录推导运行态 Huey 队列目录，避免被普通缓存清理删除。"""

    resolved_temp_root = Path(temp_root).resolve()
    return (resolved_temp_root.parent / "runtime" / "huey").resolve()


def reset_runtime_huey_queue_dir(temp_root: str | Path) -> Path:
    """后端启动时清理旧 Huey 队列文件；队列只保存运行态任务，不长期保留。"""

    queue_dir = runtime_huey_queue_dir(temp_root)
    queue_dir.mkdir(parents=True, exist_ok=True)
    SafeDirectoryCleaner().clear_contents(queue_dir)
    queue_dir.mkdir(parents=True, exist_ok=True)
    return queue_dir


def ensure_huey_storage_ready(huey: Any, queue_database_path: str | Path) -> None:
    """提交任务前确保 Huey 队列目录和 SQLite 表结构存在。"""

    database_path = Path(queue_database_path).resolve()
    database_path.parent.mkdir(parents=True, exist_ok=True)
    if _has_huey_task_table(database_path):
        return

    # 如果运行缓存被手动清理，旧连接可能仍指向已删除文件；先关闭再让
    # Huey 自己重建内部 schema，避免我们手写它的私有表结构。
    close = getattr(getattr(huey, "storage", None), "close", None)
    if callable(close):
        close()

    database_path.parent.mkdir(parents=True, exist_ok=True)
    initialize_schema = getattr(getattr(huey, "storage", None), "initialize_schema", None)
    if not callable(initialize_schema):
        raise RuntimeError("Huey队列存储不支持自动初始化")
    initialize_schema()


def _has_huey_task_table(database_path: Path) -> bool:
    if not database_path.exists():
        return False
    try:
        with sqlite3.connect(f"{database_path.as_uri()}?mode=ro", uri=True) as connection:
            row = connection.execute(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'task'"
            ).fetchone()
            return row is not None
    except sqlite3.DatabaseError:
        return False


__all__ = [
    "ensure_huey_storage_ready",
    "reset_runtime_huey_queue_dir",
    "runtime_huey_queue_dir",
]
