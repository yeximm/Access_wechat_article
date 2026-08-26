from __future__ import annotations

from dataclasses import dataclass
import hashlib
import re
import time
from typing import Any, Mapping
from urllib.parse import parse_qs, parse_qsl, urlencode, urlparse, urlunparse


WECHAT_ARTICLE_HOST = "mp.weixin.qq.com"
ARTICLE_REQUIRED_KEYS = frozenset({"mid", "idx", "sn"})
SHORT_ARTICLE_PATH_PATTERN = re.compile(r"^/s/[A-Za-z0-9_-]{16,128}$")
SENSITIVE_QUERY_KEYS = frozenset(
    {
        "key",
        "pass_ticket",
        "appmsg_token",
        "uin",
        "wxtoken",
        "poc_token",
        "exportkey",
        "sessionid",
    }
)
ALLOWED_REQUEST_HEADERS = frozenset(
    {
        "accept",
        "accept-language",
        "cache-control",
        "content-type",
        "cookie",
        "pragma",
        "referer",
        "user-agent",
    }
)


@dataclass(frozen=True, slots=True)
class ArticleReferenceMatch:
    url: str
    url_redacted: str
    url_source: str
    carrier_url: str
    method: str
    request_headers: dict[str, str]
    query_keys: tuple[str, ...]
    observed_at: float

    def to_reference(self) -> dict[str, Any]:
        """返回后续直连补取 HTML 所需的本地临时证据。"""
        return {
            "url": self.url,
            "url_redacted": self.url_redacted,
            "url_source": self.url_source,
            "carrier_url": self.carrier_url,
            "method": self.method,
            "request_headers": dict(self.request_headers),
            "query_keys": list(self.query_keys),
            "observed_at": self.observed_at,
        }

    def to_request_summary(self) -> dict[str, Any]:
        """生成可用于日志和结果展示的脱敏摘要。"""
        return {
            "source": self.url_source,
            "url_redacted": self.url_redacted,
            "query_keys": list(self.query_keys),
            "observed_at": self.observed_at,
        }


@dataclass(frozen=True, slots=True)
class ArticleHtmlMatch:
    html: str
    reference: ArticleReferenceMatch
    request_summary: dict[str, Any]


class WechatRequestMatcher:
    """识别本次监听开始后出现的微信文章主请求和主 HTML。"""

    def __init__(self, *, listen_started_at: float) -> None:
        self.listen_started_at = float(listen_started_at)

    def match_reference(
        self,
        request_url: str,
        *,
        method: str = "GET",
        headers: Mapping[str, Any] | None = None,
        observed_at: float | None = None,
    ) -> ArticleReferenceMatch | None:
        captured_at = time.monotonic() if observed_at is None else float(observed_at)
        if captured_at < self.listen_started_at:
            return None

        compacted_headers = compact_request_headers(headers)
        direct = _analyze_article_url(request_url)
        if direct is not None:
            return ArticleReferenceMatch(
                url=request_url,
                url_redacted=redact_sensitive_url(request_url),
                url_source="request",
                carrier_url="",
                method=str(method or "GET").upper(),
                request_headers=compacted_headers,
                query_keys=direct,
                observed_at=captured_at,
            )

        referer = compacted_headers.get("referer", "")
        referred = _analyze_article_url(referer)
        if referred is None:
            return None
        return ArticleReferenceMatch(
            url=referer,
            url_redacted=redact_sensitive_url(referer),
            url_source="referer",
            carrier_url=request_url,
            method=str(method or "GET").upper(),
            request_headers=compacted_headers,
            query_keys=referred,
            observed_at=captured_at,
        )

    def match_html_response(
        self,
        request_url: str,
        *,
        html_text: str,
        status_code: int,
        response_headers: Mapping[str, Any] | None = None,
        request_headers: Mapping[str, Any] | None = None,
        method: str = "GET",
        observed_at: float | None = None,
    ) -> ArticleHtmlMatch | None:
        """只接受带关键参数的文章主请求所返回的有效 HTML。"""
        reference = self.match_reference(
            request_url,
            method=method,
            headers=request_headers,
            observed_at=observed_at,
        )
        if reference is None or reference.url_source != "request":
            return None

        html = str(html_text or "")
        if not _looks_like_html(html):
            return None
        if not 200 <= int(status_code) < 400:
            return None

        normalized_response_headers = _lower_headers(response_headers)
        content_type = normalized_response_headers.get("content-type", "").lower()
        if content_type and not any(token in content_type for token in ("html", "text/plain")):
            return None

        encoded = html.encode("utf-8", errors="ignore")
        summary = reference.to_request_summary()
        summary.update(
            {
                "source": "mitm_response",
                "status_code": int(status_code),
                "content_type": content_type,
                "body_bytes": len(encoded),
                "body_sha256": hashlib.sha256(encoded).hexdigest(),
            }
        )
        return ArticleHtmlMatch(html=html, reference=reference, request_summary=summary)


def compact_request_headers(headers: Mapping[str, Any] | None) -> dict[str, str]:
    normalized = _lower_headers(headers)
    return {
        key: value
        for key, value in normalized.items()
        if key in ALLOWED_REQUEST_HEADERS and value
    }


def redact_sensitive_url(url: str) -> str:
    """移除短时有效的敏感查询参数，供日志和状态消息使用。"""
    parsed = urlparse(str(url or ""))
    safe_pairs = [
        (key, value)
        for key, value in parse_qsl(parsed.query, keep_blank_values=True)
        if key.lower() not in SENSITIVE_QUERY_KEYS
    ]
    return urlunparse(
        (parsed.scheme, parsed.netloc, parsed.path, parsed.params, urlencode(safe_pairs), parsed.fragment)
    )


def _analyze_article_url(url: str) -> tuple[str, ...] | None:
    parsed = urlparse(str(url or ""))
    if parsed.scheme.lower() not in {"http", "https"}:
        return None
    if (parsed.hostname or "").lower() != WECHAT_ARTICLE_HOST:
        return None

    normalized_path = parsed.path.rstrip("/")
    query = parse_qs(parsed.query, keep_blank_values=True)
    keys = set(query)

    # 微信主页中的贴图/图片消息可能直接导航到公开短链，而不是先暴露
    # 带 __biz、mid、idx、sn 和临时 key 的传统文章长链。监听器在单篇
    # 点击前才启动，因此可以安全地把同域的 /s/<token> 主导航作为候选；
    # 后续仍会校验响应正文或使用该 reference 补取并解析页面。
    if SHORT_ARTICLE_PATH_PATTERN.fullmatch(normalized_path):
        return tuple(sorted(keys))

    if normalized_path != "/s":
        return None
    if not ({"__biz", "biz"} & keys):
        return None
    if not ARTICLE_REQUIRED_KEYS.issubset(keys):
        return None
    if not str((query.get("key") or [""])[0]).strip():
        return None
    return tuple(sorted(keys))


def _lower_headers(headers: Mapping[str, Any] | None) -> dict[str, str]:
    if not headers:
        return {}
    return {
        str(key).strip().lower(): str(value).strip()
        for key, value in headers.items()
        if str(key).strip()
    }


def _looks_like_html(html: str) -> bool:
    # 普通错误页、验证页同样有 <html>；只有文章正文容器存在时才允许覆盖 reference。
    return bool(
        re.search(
            r"(?is)<[^>]+\bid\s*=\s*(['\"])js_content\1",
            str(html or ""),
        )
    )
