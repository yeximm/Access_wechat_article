from __future__ import annotations

from datetime import datetime
import importlib.util
import json
from pathlib import Path
import sqlite3
import sys
import time
from typing import Any, Callable

from src.storage.sqlite.database_initializer import REQUIRED_TABLES, read_required_tables


SELF_CHECK_STATE_RELATIVE_PATH = Path("data/runtime/startup_self_check.json")
SELF_CHECK_FILE_SCHEMA_VERSION = 1
SELF_CHECK_OK_STATUSES = {"passed", "passed_with_warnings"}
SELF_CHECK_RECHECK_STATUSES = {"failed", "interrupted"}

PLAYWRIGHT_INSTALL_COMMAND = (
    '$env:PLAYWRIGHT_BROWSERS_PATH="$PWD\\.playwright-browsers"; '
    "uv run playwright install chromium"
)
CA_INSTALL_GUIDE = "系统配置 -> 诊断工具 -> MITM 管理 -> CA证书安装"

DependencyChecker = Callable[[str], bool]
CaStatusChecker = Callable[[Any], dict[str, Any]]
PlaywrightChromiumChecker = Callable[[Path], bool]


def _default_dependency_checker(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def _default_playwright_chromium_checker(project_root: Path) -> bool:
    browser_root = project_root / ".playwright-browsers"
    if not browser_root.is_dir():
        return False
    return any(path.is_dir() and path.name.startswith("chromium") for path in browser_root.iterdir())


class StartupSelfCheckService:
    """按程序版本记录一次启动自检结果，避免每次启动都执行耗时检查。"""

    def __init__(
        self,
        *,
        project_root: str | Path,
        state_relative_path: str | Path = SELF_CHECK_STATE_RELATIVE_PATH,
        dependency_checker: DependencyChecker | None = None,
        ca_status_checker: CaStatusChecker | None = None,
        playwright_chromium_checker: PlaywrightChromiumChecker | None = None,
    ) -> None:
        self.project_root = Path(project_root).resolve()
        self.state_relative_path = Path(state_relative_path)
        self.state_path = (self.project_root / self.state_relative_path).resolve()
        self._dependency_checker = dependency_checker or _default_dependency_checker
        self._ca_status_checker = ca_status_checker
        self._playwright_chromium_checker = (
            playwright_chromium_checker or _default_playwright_chromium_checker
        )

    def get_status(self, config: Any) -> dict[str, Any]:
        state = self._read_state()
        return {
            "ok": True,
            "needsSelfCheck": self._needs_self_check(config, state),
            "currentVersion": self._software_version(config),
            "currentDataSchemaVersion": self._data_schema_version(config),
            "state": state,
            "statePath": self._display_path(self.state_path),
        }

    def run(self, config: Any) -> dict[str, Any]:
        started_at = time.perf_counter()
        checked_at = datetime.now().isoformat(timespec="seconds")
        items = self._collect_items(config)
        fatal_count = sum(1 for item in items if item["status"] == "failed")
        warning_count = sum(1 for item in items if item["status"] == "warning")
        if fatal_count:
            status = "failed"
        elif warning_count:
            status = "passed_with_warnings"
        else:
            status = "passed"
        duration_seconds = round(time.perf_counter() - started_at, 3)
        state = {
            "file_schema_version": SELF_CHECK_FILE_SCHEMA_VERSION,
            "checked_version": self._software_version(config),
            "checked_data_schema_version": self._data_schema_version(config),
            "checked_at": checked_at,
            "status": status,
            "fatal_count": fatal_count,
            "warning_count": warning_count,
            "duration_seconds": duration_seconds,
            "items": items,
        }
        self._write_state(state)
        return self._state_to_result(state)

    def _needs_self_check(self, config: Any, state: dict[str, Any] | None) -> bool:
        if not state:
            return True
        if state.get("file_schema_version") != SELF_CHECK_FILE_SCHEMA_VERSION:
            return True
        if state.get("checked_version") != self._software_version(config):
            return True
        if state.get("checked_data_schema_version") != self._data_schema_version(config):
            return True
        if state.get("status") in SELF_CHECK_RECHECK_STATUSES:
            return True
        return state.get("status") not in SELF_CHECK_OK_STATUSES

    def _collect_items(self, config: Any) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        items.append(self._check_python_runtime())
        items.extend(self._check_dependencies())
        items.extend(self._check_config_files())
        items.extend(self._check_storage_paths(config))
        items.append(self._check_database(config))
        items.extend(self._check_mitm(config))
        items.append(self._check_playwright_chromium())
        return items

    def _check_python_runtime(self) -> dict[str, Any]:
        version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        ok = sys.version_info >= (3, 13)
        return self._item(
            key="python_runtime",
            group="runtime",
            label="Python 运行环境",
            ok=ok,
            severity="fatal",
            message=f"当前 Python {version}",
            action="请使用 uv sync 安装并通过 uv run python main.py 启动程序。" if not ok else "",
        )

    def _check_dependencies(self) -> list[dict[str, Any]]:
        dependencies = [
            ("fastapi", "FastAPI", "fatal"),
            ("uvicorn", "Uvicorn", "fatal"),
            ("webview", "pywebview", "fatal"),
            ("uiautomation", "uiautomation", "fatal"),
            ("mitmproxy", "mitmproxy", "warning"),
            ("playwright", "Playwright", "warning"),
        ]
        items = []
        for module_name, label, severity in dependencies:
            ok = bool(self._dependency_checker(module_name))
            items.append(
                self._item(
                    key=f"dependency_{module_name}",
                    group="runtime",
                    label=f"{label} 依赖",
                    ok=ok,
                    severity=severity,
                    message="已安装" if ok else "未检测到依赖",
                    action="请在项目根目录执行 uv sync 后重新启动程序。" if not ok else "",
                )
            )
        return items

    def _check_config_files(self) -> list[dict[str, Any]]:
        files = [
            ("system_config", "系统配置文件", self.project_root / "src/config/system.yaml"),
            ("custom_config", "用户配置文件", self.project_root / "data/custom.yaml"),
        ]
        items = []
        for key, label, path in files:
            ok = path.is_file()
            message = self._display_path(path) if ok else f"未找到：{self._display_path(path)}"
            items.append(
                self._item(
                    key=key,
                    group="config",
                    label=label,
                    ok=ok,
                    severity="fatal",
                    message=message,
                    action="请确认项目文件完整，或重新拉取/解压程序文件。" if not ok else "",
                )
            )
        return items

    def _check_storage_paths(self, config: Any) -> list[dict[str, Any]]:
        storage = getattr(config, "storage", None)
        temp_dir = getattr(storage, "temp_dir", None)
        runtime_dir = Path(temp_dir).parent / "runtime" if temp_dir is not None else None
        paths = [
            ("article_storage_root", "文章存储目录", getattr(storage, "article_storage_root", None)),
            ("db_dir", "数据库目录", getattr(storage, "db_dir", None)),
            ("temp_dir", "临时文件目录", getattr(storage, "temp_dir", None)),
            ("log_dir", "日志目录", getattr(storage, "log_dir", None)),
            ("runtime_dir", "运行时状态目录", runtime_dir),
            (
                "huey_queue_dir",
                "Huey 临时队列目录",
                runtime_dir / "huey" if runtime_dir is not None else None,
            ),
        ]
        return [self._check_writable_directory(key, label, path) for key, label, path in paths]

    def _check_writable_directory(self, key: str, label: str, raw_path: Any) -> dict[str, Any]:
        path = Path(raw_path) if raw_path is not None else None
        try:
            if path is None:
                raise OSError("配置值为空")
            path.mkdir(parents=True, exist_ok=True)
            probe = path / ".startup_self_check_write_probe"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            return self._item(
                key=key,
                group="storage",
                label=label,
                ok=True,
                severity="fatal",
                message=self._display_path(path),
            )
        except OSError as exc:
            return self._item(
                key=key,
                group="storage",
                label=label,
                ok=False,
                severity="fatal",
                message=str(exc),
                action="请检查目录是否存在、是否有写入权限，或在系统配置中调整存储路径。",
            )

    def _check_database(self, config: Any) -> dict[str, Any]:
        database_path = Path(getattr(getattr(config, "storage", None), "database_path", ""))
        try:
            existing_tables = set(read_required_tables(database_path))
            missing = sorted(set(REQUIRED_TABLES) - existing_tables)
            if missing:
                raise sqlite3.DatabaseError(f"缺少必要表：{', '.join(missing)}")
            return self._item(
                key="database_tables",
                group="database",
                label="SQLite 数据表",
                ok=True,
                severity="fatal",
                message=f"{database_path.name} 表结构正常",
            )
        except Exception as exc:
            return self._item(
                key="database_tables",
                group="database",
                label="SQLite 数据表",
                ok=False,
                severity="fatal",
                message=f"{database_path.name or '数据库'} 校验失败：{exc}",
                action="请确认数据库文件可打开；首次安装可重新启动程序让数据库初始化流程创建表结构。",
            )

    def _check_mitm(self, config: Any) -> list[dict[str, Any]]:
        proxy = getattr(config, "proxy", None)
        confdir = Path(getattr(proxy, "confdir", self.project_root / ".mitmproxy"))
        ca_cert_path = Path(getattr(proxy, "ca_cert_path", confdir / "mitmproxy-ca-cert.cer"))
        ca_status = self._ca_status_checker(config) if self._ca_status_checker else {}
        ca_file_exists = bool(ca_status.get("caFileExists", ca_cert_path.is_file()))
        ca_installed = bool(ca_status.get("projectCertificateInstalled", False))
        return [
            self._item(
                key="mitm_confdir",
                group="mitm",
                label="MITM 配置目录",
                ok=confdir.is_dir(),
                severity="warning",
                message=self._display_path(confdir) if confdir.is_dir() else "未检测到 MITM 配置目录",
                action="首次运行 MITM 或执行证书安装后会生成该目录。" if not confdir.is_dir() else "",
            ),
            self._item(
                key="mitm_ca_file",
                group="mitm",
                label="MITM CA 文件",
                ok=ca_file_exists,
                severity="warning",
                message=self._display_path(ca_cert_path) if ca_file_exists else "未检测到项目 CA 证书",
                action=CA_INSTALL_GUIDE if not ca_file_exists else "",
            ),
            self._item(
                key="mitm_ca_installed",
                group="mitm",
                label="系统 CA 证书",
                ok=ca_installed,
                severity="warning",
                message="已安装当前项目 CA 证书" if ca_installed else "系统中未确认安装当前项目 CA 证书",
                action=CA_INSTALL_GUIDE if not ca_installed else "",
            ),
        ]

    def _check_playwright_chromium(self) -> dict[str, Any]:
        browser_root = self.project_root / ".playwright-browsers"
        ok = bool(self._playwright_chromium_checker(self.project_root))
        return self._item(
            key="playwright_chromium",
            group="playwright",
            label="Playwright Chromium",
            ok=ok,
            severity="warning",
            message=self._display_path(browser_root) if ok else "未检测到项目目录下的 Chromium",
            action=PLAYWRIGHT_INSTALL_COMMAND if not ok else "",
        )

    def _item(
        self,
        *,
        key: str,
        group: str,
        label: str,
        ok: bool,
        severity: str,
        message: str,
        action: str = "",
    ) -> dict[str, Any]:
        status = "success" if ok else ("failed" if severity == "fatal" else "warning")
        return {
            "key": key,
            "group": group,
            "label": label,
            "status": status,
            "severity": "info" if ok else severity,
            "ok": ok,
            "message": message,
            "action": action,
        }

    def _state_to_result(self, state: dict[str, Any]) -> dict[str, Any]:
        return {
            "ok": state["fatal_count"] == 0,
            "status": state["status"],
            "fatalCount": state["fatal_count"],
            "warningCount": state["warning_count"],
            "items": state["items"],
            "checkedAt": state["checked_at"],
            "durationSeconds": state["duration_seconds"],
            "statePath": self._display_path(self.state_path),
        }

    def _read_state(self) -> dict[str, Any] | None:
        if not self.state_path.is_file():
            return None
        try:
            raw = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None
        return raw if isinstance(raw, dict) else None

    def _write_state(self, state: dict[str, Any]) -> None:
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.state_path.with_suffix(self.state_path.suffix + ".tmp")
        temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(self.state_path)

    def _software_version(self, config: Any) -> str:
        return str(getattr(getattr(config, "software", None), "version", "") or "")

    def _data_schema_version(self, config: Any) -> str:
        return str(getattr(getattr(config, "software", None), "data_schema_version", "") or "")

    def _display_path(self, path: str | Path) -> str:
        resolved = Path(path).resolve()
        try:
            return str(resolved.relative_to(self.project_root))
        except ValueError:
            return str(resolved)


__all__ = [
    "CA_INSTALL_GUIDE",
    "PLAYWRIGHT_INSTALL_COMMAND",
    "SELF_CHECK_STATE_RELATIVE_PATH",
    "StartupSelfCheckService",
]
