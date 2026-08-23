from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any

from ruamel.yaml import YAML

from src.config.app_config import (
    AppConfig,
    CommentConfig,
    MitmCaptureConfig,
    OfflineCacheConfig,
    ProxyConfig,
    RequestConfig,
    RuntimeConfig,
    SoftwareConfig,
    StorageConfig,
    WindowConfig,
)
from src.config.config_validator import ConfigValidationError, validate_app_config


DEFAULT_SYSTEM_CONFIG_PATH = Path(__file__).with_name("system.yaml")


def load_app_config(
    config_path: str | Path,
    *,
    project_root: str | Path,
    system_config_path: str | Path | None = None,
) -> AppConfig:
    mapping = load_layered_config_mapping(
        config_path,
        system_config_path=system_config_path,
    )
    return build_app_config(mapping, project_root=project_root)


def load_layered_config_mapping(
    config_path: str | Path,
    *,
    system_config_path: str | Path | None = None,
) -> dict[str, Any]:
    """以系统 YAML 为基础，再使用 custom.yaml 中存在的字段覆盖。"""
    system_path = Path(system_config_path or DEFAULT_SYSTEM_CONFIG_PATH)
    system_mapping = load_config_mapping(system_path)
    custom_path = Path(config_path)
    custom_mapping = load_config_mapping(custom_path) if custom_path.is_file() else {}
    return _deep_merge(system_mapping, custom_mapping)


def load_system_app_config(
    system_config_path: str | Path,
    *,
    project_root: str | Path,
) -> AppConfig:
    return build_app_config(
        load_config_mapping(system_config_path),
        project_root=project_root,
    )


def load_config_mapping(config_path: str | Path) -> dict[str, Any]:
    path = Path(config_path)
    if not path.is_file():
        raise ConfigValidationError(f"配置文件不存在：{path}")
    try:
        loaded = YAML(typ="safe").load(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ConfigValidationError(f"配置文件解析失败：{exc}") from exc
    if not isinstance(loaded, Mapping):
        raise ConfigValidationError("YAML 根节点必须是映射")
    return _plain_mapping(loaded)


def build_app_config(mapping: Mapping[str, Any], *, project_root: str | Path) -> AppConfig:
    root = Path(project_root).resolve()
    raw_mapping = _plain_mapping(mapping)
    data = _deep_merge(raw_mapping, _latest_menu_mapping(raw_mapping))

    software = _section(data, "software")
    storage = _section(data, "storage")
    proxy = _section(data, "proxy")
    mitm_capture = _section(data, "mitm_capture")
    window = _section(data, "window")
    request = _section(data, "request")
    comment = _section(data, "comment")
    offline_cache = _section(data, "offline_cache")
    runtime = _section(data, "runtime")
    data_schema_version = _as_string(software, "data_schema_version")
    (
        article_title_poll_initial_interval_seconds,
        article_title_poll_max_interval_seconds,
    ) = _window_article_title_poll_intervals(window)
    config = AppConfig(
        software=SoftwareConfig(
            version=_as_string(software, "version"),
            data_schema_version=data_schema_version,
        ),
        storage=StorageConfig(
            article_storage_root=_resolve_path(storage, "article_storage_root", root),
            db_dir=_resolve_path(storage, "db_dir", root),
            db_file_name=_database_file_name(data_schema_version),
            temp_dir=_resolve_path(storage, "temp_dir", root),
            log_dir=_resolve_path(storage, "log_dir", root),
        ),
        proxy=ProxyConfig(
            host=_as_string(proxy, "host"),
            port=_as_int(proxy, "port"),
            startup_delay_seconds=_as_float(proxy, "startup_delay_seconds"),
            verification_url=_as_string(proxy, "verification_url"),
            confdir=_resolve_path(proxy, "confdir", root),
            ca_cert_path=_resolve_path(proxy, "ca_cert_path", root),
            enable_system_proxy=_as_bool(proxy, "enable_system_proxy"),
            ssl_insecure=_as_bool(proxy, "ssl_insecure"),
        ),
        mitm_capture=MitmCaptureConfig(
            ready_timeout_seconds=_as_float(mitm_capture, "ready_timeout_seconds"),
            capture_timeout_seconds=_as_float(mitm_capture, "capture_timeout_seconds"),
            result_timeout_seconds=_as_float(mitm_capture, "result_timeout_seconds"),
            listener_shutdown_timeout_seconds=_as_float(
                mitm_capture,
                "listener_shutdown_timeout_seconds",
            ),
            close_as_capture_deadline=_as_bool(mitm_capture, "close_as_capture_deadline"),
        ),
        window=WindowConfig(
            activation_wait_seconds=_as_float(window, "activation_wait_seconds"),
            home_find_timeout_seconds=_as_float(window, "home_find_timeout_seconds"),
            restore_focus_after_close=_as_bool(window, "restore_focus_after_close"),
            article_open_timeout_seconds=_as_float(
                window, "article_open_timeout_seconds"
            ),
            article_title_poll_initial_interval_seconds=(
                article_title_poll_initial_interval_seconds
            ),
            article_title_poll_max_interval_seconds=(
                article_title_poll_max_interval_seconds
            ),
            article_title_poll_interval_seconds=_as_float(
                window, "article_title_poll_interval_seconds"
            ),
            article_title_stable_delay_seconds=_as_float(
                window, "article_title_stable_delay_seconds"
            ),
            article_close_confirm_timeout_seconds=_as_float(
                window, "article_close_confirm_timeout_seconds"
            ),
            scroll_wheel_steps=_as_int(window, "scroll_wheel_steps"),
            date_seek_max_steps=_as_int(window, "date_seek_max_steps"),
            scroll_initial_delay_seconds=_as_float(
                window, "scroll_initial_delay_seconds"
            ),
            scroll_probe_interval_seconds=_as_float(
                window, "scroll_probe_interval_seconds"
            ),
            scroll_probe_max_interval_seconds=_as_float(
                window, "scroll_probe_max_interval_seconds"
            ),
            lazy_load_timeout_seconds=_as_float(window, "lazy_load_timeout_seconds"),
            unchanged_before_bounce_seconds=_as_float(
                window, "unchanged_before_bounce_seconds"
            ),
            bounce_enabled=_as_bool(window, "bounce_enabled"),
            bounce_attempts=_as_int(window, "bounce_attempts"),
            bounce_up_steps=_as_int(window, "bounce_up_steps"),
            bounce_down_steps=_as_int(window, "bounce_down_steps"),
            bounce_pause_seconds=_as_float(window, "bounce_pause_seconds"),
        ),
        request=RequestConfig(
            request_interval_seconds=_as_float(request, "request_interval_seconds"),
            request_timeout_seconds=_as_float(request, "request_timeout_seconds"),
        ),
        comment=CommentConfig(
            enabled_by_default=_as_bool(comment, "enabled_by_default"),
            request_timeout_seconds=_as_float(comment, "request_timeout_seconds"),
            page_interval_seconds=_as_float(comment, "page_interval_seconds"),
            max_pages=_as_int(comment, "max_pages"),
            max_concurrent_processes=_as_int(comment, "max_concurrent_processes"),
        ),
        offline_cache=OfflineCacheConfig(
            enabled_by_default=_as_bool(offline_cache, "enabled_by_default"),
            max_scroll_seconds=_as_float(offline_cache, "max_scroll_seconds"),
            resource_timeout_seconds=_as_float(offline_cache, "resource_timeout_seconds"),
            max_concurrent_processes=_as_int(
                offline_cache,
                "max_concurrent_processes",
            ),
        ),
        runtime=RuntimeConfig(
            log_level=_as_string(runtime, "log_level").upper(),
            auto_clean_temp_files=_as_bool(runtime, "auto_clean_temp_files"),
            temp_retention_days=_as_int(runtime, "temp_retention_days"),
            log_retention_days=_as_int(runtime, "log_retention_days"),
        ),
    )
    validate_app_config(config)
    return config


def _deep_merge(base: dict[str, Any], override: Mapping[str, Any]) -> dict[str, Any]:
    for key, value in override.items():
        if isinstance(value, Mapping) and isinstance(base.get(key), dict):
            base[key] = _deep_merge(dict(base[key]), value)
        else:
            base[key] = value
    return base


def _plain_mapping(mapping: Mapping[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in mapping.items():
        result[str(key)] = _plain_mapping(value) if isinstance(value, Mapping) else value
    return result


def _latest_menu_mapping(mapping: Mapping[str, Any]) -> dict[str, Any]:
    """把新版一二级菜单 YAML 转成当前 AppConfig 使用的运行时结构。"""
    result: dict[str, Any] = {}

    basic_settings = _mapping_value(mapping, "basic_settings")
    runtime_maintenance = _mapping_value(basic_settings, "runtime_maintenance")
    project_storage = _mapping_value(basic_settings, "project_storage")
    database_settings = _mapping_value(basic_settings, "database_settings")
    basic_proxy_settings = _mapping_value(basic_settings, "proxy_settings")
    main_flow = _mapping_value(mapping, "main_flow")
    dispatch_control = _mapping_value(main_flow, "dispatch_control")
    single_article_task = _mapping_value(mapping, "single_article_task")
    _copy_keys(
        runtime_maintenance,
        result.setdefault("runtime", {}),
        (
            "log_level",
            "auto_clean_temp_files",
            "temp_retention_days",
            "log_retention_days",
        ),
    )
    _copy_keys(
        runtime_maintenance,
        result.setdefault("request", {}),
        ("request_interval_seconds",),
    )
    if "single_task_interval_seconds" in dispatch_control:
        result.setdefault("request", {})["request_interval_seconds"] = dispatch_control[
            "single_task_interval_seconds"
        ]
    _copy_keys(
        project_storage,
        result.setdefault("storage", {}),
        ("article_storage_root", "temp_dir", "log_dir"),
    )
    _copy_keys(database_settings, result.setdefault("storage", {}), ("db_dir",))
    _copy_keys(
        database_settings,
        result.setdefault("software", {}),
        ("data_schema_version",),
    )

    single_task_mitm_capture = _mapping_value(single_article_task, "mitm_capture")
    _copy_keys(
        basic_proxy_settings,
        result.setdefault("proxy", {}),
        (
            "host",
            "port",
            "startup_delay_seconds",
            "verification_url",
            "confdir",
            "ca_cert_path",
            "enable_system_proxy",
            "ssl_insecure",
        ),
    )
    _copy_keys(
        single_task_mitm_capture,
        result.setdefault("mitm_capture", {}),
        (
            "ready_timeout_seconds",
            "capture_timeout_seconds",
            "result_timeout_seconds",
            "listener_shutdown_timeout_seconds",
            "close_as_capture_deadline",
        ),
    )

    main_home_window = _mapping_value(main_flow, "home_window")
    main_home_scroll = _mapping_value(main_flow, "home_scroll")
    single_task_article_tab = _mapping_value(single_article_task, "article_tab")
    window_runtime = result.setdefault("window", {})
    _copy_keys(
        single_task_article_tab,
        window_runtime,
        (
            "restore_focus_after_close",
            "article_open_timeout_seconds",
            "article_title_poll_interval_seconds",
            "article_title_stable_delay_seconds",
            "article_close_confirm_timeout_seconds",
        ),
    )
    if "article_title_poll_interval_seconds_range" in single_task_article_tab:
        initial_interval, max_interval = _number_pair(
            single_task_article_tab,
            "article_title_poll_interval_seconds_range",
        )
        window_runtime["article_title_poll_initial_interval_seconds"] = initial_interval
        window_runtime["article_title_poll_max_interval_seconds"] = max_interval
        window_runtime["article_title_poll_interval_seconds"] = max_interval

    _copy_keys(
        main_home_window,
        window_runtime,
        (
            "activation_wait_seconds",
            "home_find_timeout_seconds",
        ),
    )
    _copy_keys(
        main_home_scroll,
        window_runtime,
        (
            "scroll_wheel_steps",
            "date_seek_max_steps",
            "scroll_initial_delay_seconds",
            "scroll_probe_interval_seconds",
            "scroll_probe_max_interval_seconds",
            "unchanged_before_bounce_seconds",
            "lazy_load_timeout_seconds",
            "bounce_enabled",
            "bounce_up_steps",
            "bounce_pause_seconds",
            "bounce_down_steps",
            "bounce_attempts",
        ),
    )
    if "date_seek_scroll_steps_range" in main_home_scroll:
        normal_steps, max_steps = _number_pair(
            main_home_scroll,
            "date_seek_scroll_steps_range",
        )
        window_runtime["scroll_wheel_steps"] = int(normal_steps)
        window_runtime["date_seek_max_steps"] = int(max_steps)
    if "scroll_probe_interval_seconds_range" in main_home_scroll:
        initial_interval, max_interval = _number_pair(
            main_home_scroll,
            "scroll_probe_interval_seconds_range",
        )
        window_runtime["scroll_probe_interval_seconds"] = initial_interval
        window_runtime["scroll_probe_max_interval_seconds"] = max_interval

    html_storage = _mapping_value(single_article_task, "html_storage")
    single_task_comment_collection = _mapping_value(
        single_article_task,
        "comment_collection",
    )
    single_task_offline_cache = _mapping_value(single_article_task, "offline_cache")
    _copy_keys(
        html_storage,
        result.setdefault("request", {}),
        ("request_timeout_seconds",),
    )
    _copy_keys(
        single_task_comment_collection,
        result.setdefault("comment", {}),
        (
            "enabled_by_default",
            "request_timeout_seconds",
            "page_interval_seconds",
            "max_concurrent_processes",
        ),
    )
    if "top_level_max_pages" in single_task_comment_collection:
        result.setdefault("comment", {})["max_pages"] = single_task_comment_collection[
            "top_level_max_pages"
        ]
    _copy_keys(
        single_task_offline_cache,
        result.setdefault("offline_cache", {}),
        (
            "enabled_by_default",
            "max_scroll_seconds",
            "resource_timeout_seconds",
            "max_concurrent_processes",
        ),
    )

    return result


def _mapping_value(mapping: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    value = mapping.get(key) if isinstance(mapping, Mapping) else None
    return value if isinstance(value, Mapping) else {}


def _copy_keys(
    source: Mapping[str, Any],
    target: dict[str, Any],
    keys: tuple[str, ...],
) -> None:
    for key in keys:
        if key in source:
            target[key] = source[key]


def _number_pair(source: Mapping[str, Any], key: str) -> tuple[float, float]:
    value = source.get(key)
    if (
        isinstance(value, str)
        or not isinstance(value, (list, tuple))
        or len(value) != 2
    ):
        raise ConfigValidationError(f"{key} 必须是包含两个数值的数组")

    try:
        first = float(value[0])
        second = float(value[1])
    except (TypeError, ValueError) as exc:
        raise ConfigValidationError(f"{key} 必须是包含两个数值的数组") from exc

    if second < first:
        raise ConfigValidationError(f"{key} 的结束值不能小于起始值")

    return first, second


def _window_article_title_poll_intervals(
    window: Mapping[str, Any],
) -> tuple[float, float]:
    """标题检测轮询使用起始值和最大值；旧标量字段兼容为“最大值”。"""
    if (
        "article_title_poll_initial_interval_seconds" in window
        and "article_title_poll_max_interval_seconds" in window
    ):
        return (
            _as_float(window, "article_title_poll_initial_interval_seconds"),
            _as_float(window, "article_title_poll_max_interval_seconds"),
        )
    max_interval = _as_float(window, "article_title_poll_interval_seconds")
    return min(0.05, max_interval), max_interval


def _section(data: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    value = data.get(key)
    if not isinstance(value, Mapping):
        raise ConfigValidationError(f"{key} 必须是映射")
    return value


def _value(section: Mapping[str, Any], key: str) -> Any:
    if key not in section:
        raise ConfigValidationError(f"缺少配置字段：{key}")
    return section[key]


def _as_string(section: Mapping[str, Any], key: str) -> str:
    value = str(_value(section, key)).strip()
    if not value:
        raise ConfigValidationError(f"{key} 不能为空")
    return value


def _as_int(section: Mapping[str, Any], key: str) -> int:
    value = _value(section, key)
    if isinstance(value, bool):
        raise ConfigValidationError(f"{key} 必须是整数")
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ConfigValidationError(f"{key} 必须是整数") from exc


def _as_float(section: Mapping[str, Any], key: str) -> float:
    value = _value(section, key)
    if isinstance(value, bool):
        raise ConfigValidationError(f"{key} 必须是数值")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ConfigValidationError(f"{key} 必须是数值") from exc


def _as_bool(section: Mapping[str, Any], key: str) -> bool:
    value = _value(section, key)
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "yes", "1", "on"}:
            return True
        if normalized in {"false", "no", "0", "off"}:
            return False
    raise ConfigValidationError(f"{key} 必须是布尔值")


def _resolve_path(section: Mapping[str, Any], key: str, project_root: Path) -> Path:
    raw_path = Path(_as_string(section, key))
    return raw_path.resolve() if raw_path.is_absolute() else (project_root / raw_path).resolve()


def _database_file_name(data_schema_version: str) -> str:
    """数据库文件名由数据结构版本确定，不再作为独立配置项维护。"""
    return f"awa-{data_schema_version}.sqlite3"
