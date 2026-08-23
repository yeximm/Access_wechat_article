from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Mapping
from urllib.parse import urlsplit, urlunsplit

from src.domain.enums import CaptureType, TaskStatus
from src.domain.models import MitmCaptureResult
from src.modules.request.chrome_headers import (
    DIRECT_REQUEST_PROXIES,
    build_wechat_document_headers,
    direct_requests_get,
)


REPLAY_HEADER_NAMES = frozenset(
    {
        "accept",
        "accept-language",
        "cache-control",
        "cookie",
        "pragma",
        "referer",
        "user-agent",
    }
)
HTTP_URL_SCHEMES = {"http", "https"}
WECHAT_ARTICLE_HOST = "mp.weixin.qq.com"


class ArticleHtmlRequestError(RuntimeError):
    """捕获结果无法形成可解析 HTML；异常信息不得带临时凭据。"""


@dataclass(frozen=True, slots=True)
class PreparedArticleHtml:
    html: str
    html_source: str
    request_evidence: dict[str, Any]


class ArticleHtmlRequester:
    """把 HTML/reference 捕获结果统一转换为 HTML，不负责解析或落盘。"""

    def __init__(
        self,
        *,
        request_get: Callable[..., Any] = direct_requests_get,
    ) -> None:
        self._request_get = request_get

    def prepare(
        self,
        capture: MitmCaptureResult,
        *,
        timeout_seconds: float,
    ) -> PreparedArticleHtml:
        if capture.status is not TaskStatus.SUCCESS:
            raise ArticleHtmlRequestError("MITM 捕获结果不是成功状态")
        if capture.capture_type is CaptureType.HTML and str(capture.html or "").strip():
            return PreparedArticleHtml(
                html=str(capture.html),
                html_source="mitm_response",
                request_evidence=_build_request_evidence(capture),
            )
        if capture.capture_type is not CaptureType.REFERENCE or not capture.reference:
            raise ArticleHtmlRequestError("MITM 捕获结果没有可用 HTML 或 reference")

        reference = dict(capture.reference)
        url = normalize_reference_url(str(reference.get("url", "")))
        reference["url"] = url
        _validate_reference_url(url)
        timeout = float(timeout_seconds)
        if timeout <= 0:
            raise ValueError("timeout_seconds 必须大于 0")

        headers = _normalize_replay_headers(reference.get("request_headers"))
        try:
            # 禁用环境变量和系统代理；此时本次 MITM 已完成代理恢复并退出。
            response = self._request_get(
                url,
                headers=headers,
                timeout=timeout,
                proxies=DIRECT_REQUEST_PROXIES,
            )
            status = int(getattr(response, "status_code", 0) or 0)
            if not 200 <= status < 300:
                raise ArticleHtmlRequestError(f"reference 请求失败：HTTP {status}")
            raw = bytes(getattr(response, "content", b"") or b"")
            html = _decode_response(raw, getattr(response, "headers", {}))
        except ArticleHtmlRequestError:
            raise
        except Exception as exc:
            # 不拼接底层异常文本，避免 URL 中 key/pass_ticket 泄露到日志或前端。
            raise ArticleHtmlRequestError(
                f"reference 请求失败：{type(exc).__name__}"
            ) from exc
        if not html.strip():
            raise ArticleHtmlRequestError("reference 请求成功但响应 HTML 为空")

        evidence = _build_request_evidence(capture, reference=reference)
        evidence["reference_response"] = {
            "status_code": status,
            "body_bytes": len(raw),
        }
        return PreparedArticleHtml(
            html=html,
            html_source="reference_request",
            request_evidence=evidence,
        )


def normalize_reference_url(url: str) -> str:
    """把微信内置浏览器捕获到的 reference URL 统一成 requests 可用的 HTTPS 文章地址。"""
    raw = str(url or "").strip().strip("'\"")
    if raw.startswith("//"):
        raw = f"https:{raw}"
    elif raw.startswith(WECHAT_ARTICLE_HOST):
        raw = f"https://{raw}"

    parsed = urlsplit(raw)
    scheme = parsed.scheme.lower()
    if scheme in HTTP_URL_SCHEMES and parsed.netloc:
        return urlunsplit(("https", parsed.netloc, parsed.path or "/", parsed.query, ""))

    # 微信内置浏览器可能暴露 weixin://mp.weixin.qq.com/s?... 这类 carrier URL，
    # 真正补取正文 HTML 时仍要转换成 HTTPS。
    netloc = parsed.netloc or WECHAT_ARTICLE_HOST
    path = parsed.path or "/"
    if not path.startswith("/"):
        path = f"/{path}"
    return urlunsplit(("https", netloc, path, parsed.query, ""))


def _validate_reference_url(url: str) -> None:
    parsed = urlsplit(url)
    if parsed.scheme.lower() != "https":
        raise ArticleHtmlRequestError("reference URL 必须使用 HTTPS")
    if (parsed.hostname or "").lower() != "mp.weixin.qq.com":
        raise ArticleHtmlRequestError("reference URL 不是微信文章域名")
    if parsed.path.rstrip("/") != "/s":
        raise ArticleHtmlRequestError("reference URL 不是微信文章主请求")


def _normalize_replay_headers(raw_headers: object) -> dict[str, str]:
    if not isinstance(raw_headers, Mapping):
        raw_headers = {}
    replay_headers = {
        str(key).strip().lower(): str(value).strip()
        for key, value in raw_headers.items()
        if str(key).strip().lower() in REPLAY_HEADER_NAMES
        and str(value).strip()
        and "\r" not in str(value)
        and "\n" not in str(value)
    }
    return build_wechat_document_headers(replay_headers)


def _decode_response(raw: bytes, headers: object) -> str:
    content_type = ""
    try:
        content_type = str(headers.get("Content-Type", ""))
    except Exception:
        pass
    charset = "utf-8"
    marker = "charset="
    if marker in content_type.lower():
        charset = content_type.lower().split(marker, 1)[1].split(";", 1)[0].strip()
    try:
        return raw.decode(charset, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


def _build_request_evidence(
    capture: MitmCaptureResult,
    *,
    reference: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    evidence: dict[str, Any] = {
        "request_summary": dict(capture.request_summary),
    }
    reference_data = reference if reference is not None else capture.reference
    if reference_data is not None:
        # 临时凭据只保存在本地 origin/request.json，不用于日志或公开结果。
        evidence["reference"] = dict(reference_data)
    return evidence

