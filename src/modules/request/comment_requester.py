from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import html as html_module
import json
import re
import time
from typing import Any, Callable, Mapping
from urllib.parse import parse_qs, urlencode, urlsplit, urlunsplit

from src.modules.request.chrome_headers import (
    DIRECT_REQUEST_PROXIES,
    build_wechat_json_headers,
    direct_requests_get,
)


COMMENT_LIST_KEYS = (
    "comment",
    "comment_list",
    "elected_comment",
    "elected_comment_list",
    "friend_comment",
    "my_comment",
)
REQUIRED_IDENTITY_KEYS = ("biz", "mid", "idx", "comment_id")
HTTP_URL_SCHEMES = {"http", "https"}
WECHAT_ARTICLE_HOST = "mp.weixin.qq.com"


class CommentFetchError(RuntimeError):
    """评论接口请求失败；异常消息不得携带临时凭据。"""


class CommentParametersMissing(RuntimeError):
    """文章没有形成评论接口所需参数，调用方应将任务标记为跳过。"""


@dataclass(frozen=True, slots=True)
class CommentFetchData:
    package: dict[str, Any]
    comment_count: int
    reply_count: int
    page_count: int
    pagination_complete: bool
    stop_reason: str
    download_bytes: int = 0


class WechatCommentRequester:
    """在 MITM 退出后直接请求微信评论接口，并在内存中形成评论包。"""

    def __init__(
        self,
        *,
        request_get: Callable[..., Any] = direct_requests_get,
        sleep: Callable[[float], None] = time.sleep,
        now: Callable[[], datetime] = datetime.now,
    ) -> None:
        self._request_get = request_get
        self._sleep = sleep
        self._now = now

    def fetch(
        self,
        *,
        article_url: str,
        html: str,
        request_headers: Mapping[str, Any] | None,
        timeout_seconds: float,
        page_interval_seconds: float,
        max_pages: int,
    ) -> CommentFetchData:
        article_url = normalize_article_url(article_url)
        timeout = float(timeout_seconds)
        if timeout <= 0:
            raise ValueError("timeout_seconds 必须大于 0")
        page_limit = int(max_pages)
        if page_limit <= 0:
            raise ValueError("max_pages 必须大于 0")
        interval = max(0.0, float(page_interval_seconds))

        identity = extract_comment_identity(article_url, html, request_headers)
        missing = [key for key in REQUIRED_IDENTITY_KEYS if not identity.get(key)]
        if missing:
            raise CommentParametersMissing(f"评论参数不足：{', '.join(missing)}")

        headers = _normalize_headers(request_headers, referer=article_url)
        # 显式禁用环境变量和系统代理。此时单篇 MITM 已经恢复代理并退出。
        pages: list[dict[str, Any]] = []
        page_summaries: list[dict[str, Any]] = []
        seen_ids: set[str] = set()
        offset = 0
        buffer = ""
        stop_reason = "max_pages_reached"
        pagination_complete = False
        download_bytes = 0

        for page_index in range(page_limit):
            url = _build_comment_page_url(
                article_url,
                identity,
                offset=offset,
                limit=100,
                buffer=buffer,
            )
            payload, response_bytes = self._request_json(
                self._request_get,
                url,
                headers=headers,
                timeout=timeout,
            )
            download_bytes += response_bytes
            pages.append(payload)
            new_count = _count_new_comments(payload, seen_ids)
            page_summaries.append(
                {
                    "page_index": page_index,
                    "offset": offset,
                    "new_comment_count": new_count,
                    "continue_flag": _safe_bool(payload.get("continue_flag")),
                }
            )
            continue_flag = _safe_bool(payload.get("continue_flag"))
            if payload.get("buffer") not in (None, ""):
                buffer = str(payload.get("buffer"))
            if continue_flag is False:
                stop_reason = "continue_flag_false"
                pagination_complete = True
                break
            if new_count <= 0:
                stop_reason = "no_new_comments"
                pagination_complete = True
                break
            offset += new_count
            if interval and page_index + 1 < page_limit:
                self._sleep(interval)

        merged = _merge_comment_pages(pages)
        reply_summary, reply_download_bytes = self._fetch_missing_replies(
            request_get=self._request_get,
            article_url=article_url,
            identity=identity,
            merged=merged,
            headers=headers,
            timeout=timeout,
            interval=interval,
            max_pages=page_limit,
        )
        download_bytes += reply_download_bytes
        comments = _normalize_comments(merged)
        reply_count = sum(_comment_reply_count(item) for item in comments)
        package = {
            "schema_version": "wechat_comments_raw_v1",
            "collect_time": self._now().strftime("%Y-%m-%d %H:%M:%S"),
            "article": {"url_redacted": _redact_url(article_url)},
            "summary": {
                "top_level_comment_count": len(comments),
                "reply_count": reply_count,
                "total_message_count": len(comments) + reply_count,
                "comment_page_count": len(pages),
                "pagination_complete": pagination_complete,
                "stop_reason": stop_reason,
                "page_summaries": page_summaries,
                "reply_fetch": reply_summary,
            },
            "comments": comments,
        }
        return CommentFetchData(
            package=package,
            comment_count=len(comments),
            reply_count=reply_count,
            page_count=len(pages),
            pagination_complete=pagination_complete,
            stop_reason=stop_reason,
            download_bytes=download_bytes,
        )

    def _fetch_missing_replies(
        self,
        *,
        request_get: Callable[..., Any],
        article_url: str,
        identity: Mapping[str, str],
        merged: dict[str, Any],
        headers: dict[str, str],
        timeout: float,
        interval: float,
        max_pages: int,
    ) -> tuple[dict[str, Any], int]:
        targets = [item for item in _raw_comments(merged) if _reply_missing_count(item) > 0]
        results: list[dict[str, Any]] = []
        download_bytes = 0
        for target_index, comment in enumerate(targets):
            reply_new = _ensure_reply_new(comment)
            expected = _safe_int(
                reply_new.get("reply_total_cnt", comment.get("reply_total_cnt")),
                len(_dict_list(reply_new.get("reply_list"))),
            )
            replies = _unique_replies(_dict_list(reply_new.get("reply_list")))
            offset = _safe_int(reply_new.get("offset"), len(replies))
            buffer = str(reply_new.get("buffer") or "")
            max_reply_id = str(reply_new.get("max_reply_id") or "")
            pages = 0
            stop_reason = "max_pages_reached"
            for page_index in range(max_pages):
                if expected and len(replies) >= expected:
                    stop_reason = "reply_total_reached"
                    break
                filter_reply_ids = _reply_ids(replies)
                request_offset = 0 if filter_reply_ids else offset
                url = _build_reply_url(
                    article_url,
                    identity,
                    comment,
                    offset=request_offset,
                    limit=20,
                    buffer=buffer,
                    max_reply_id=max_reply_id,
                    filter_reply_ids=filter_reply_ids,
                )
                if not url:
                    stop_reason = "reply_parameters_missing"
                    break
                payload, response_bytes = self._request_json(
                    request_get,
                    url,
                    headers=headers,
                    timeout=timeout,
                )
                download_bytes += response_bytes
                pages += 1
                reply_payload = payload.get("reply_list")
                reply_payload = reply_payload if isinstance(reply_payload, Mapping) else {}
                before = len(replies)
                replies = _unique_replies([*replies, *_dict_list(reply_payload.get("reply_list"))])
                added = len(replies) - before
                reply_new["reply_list"] = replies
                # filter_reply_list 会让微信接口排除已保存回复；此时请求 offset 必须从 0
                # 开始，否则会在“过滤后列表”上再跳过一段，导致每个楼层少存回复。
                offset = 0 if filter_reply_ids else offset + max(0, added)
                reply_new["offset"] = offset
                if payload.get("buffer") not in (None, ""):
                    buffer = str(payload.get("buffer"))
                    reply_new["buffer"] = buffer
                if reply_payload.get("max_reply_id") not in (None, ""):
                    max_reply_id = str(reply_payload.get("max_reply_id"))
                    reply_new["max_reply_id"] = max_reply_id
                if expected and len(replies) >= expected:
                    stop_reason = "reply_total_reached"
                    break
                if _safe_bool(payload.get("continue_flag")) is False:
                    stop_reason = "continue_flag_false"
                    break
                if added <= 0:
                    stop_reason = "no_new_replies"
                    break
                if interval and page_index + 1 < max_pages:
                    self._sleep(interval)
            results.append(
                {
                    "comment_id": str(comment.get("id") or comment.get("comment_id") or ""),
                    "expected_reply_count": expected,
                    "saved_reply_count": len(replies),
                    "page_count": pages,
                    "stop_reason": stop_reason,
                }
            )
            if interval and target_index + 1 < len(targets):
                self._sleep(interval)
        return (
            {
                "target_comment_count": len(targets),
                "reply_page_count": sum(item["page_count"] for item in results),
                "reply_added_count": sum(item["saved_reply_count"] for item in results),
                "targets": results,
            },
            download_bytes,
        )

    @staticmethod
    def _request_json(
        request_get: Callable[..., Any],
        url: str,
        *,
        headers: Mapping[str, str],
        timeout: float,
    ) -> tuple[dict[str, Any], int]:
        try:
            response = request_get(
                url,
                headers=dict(headers),
                timeout=timeout,
                proxies=DIRECT_REQUEST_PROXIES,
            )
            status = int(getattr(response, "status_code", 0) or 0)
            raw = bytes(getattr(response, "content", b"") or b"")
        except Exception as exc:
            raise CommentFetchError(f"评论接口请求失败：{type(exc).__name__}") from exc
        if not 200 <= status < 300:
            raise CommentFetchError(f"评论接口请求失败：HTTP {status}")
        try:
            payload = json.loads(raw.decode("utf-8", errors="replace"))
        except (UnicodeError, json.JSONDecodeError) as exc:
            raise CommentFetchError("评论接口没有返回有效 JSON") from exc
        if not isinstance(payload, dict):
            raise CommentFetchError("评论接口 JSON 结构无效")
        base_resp = payload.get("base_resp")
        base_resp = base_resp if isinstance(base_resp, Mapping) else {}
        ret = _safe_int(base_resp.get("ret", payload.get("ret", 0)), 0)
        if ret != 0:
            raise CommentFetchError(f"评论接口返回失败状态：ret={ret}")
        return payload, len(raw)


def normalize_article_url(article_url: str) -> str:
    """把微信内置浏览器/相对地址统一成可作为 Referer 使用的 http(s) 地址。"""
    raw = str(article_url or "").strip()
    if not raw:
        raise CommentParametersMissing("缺少带临时参数的文章 URL，无法获取评论")
    if raw.startswith("//"):
        raw = f"https:{raw}"
    elif raw.startswith(WECHAT_ARTICLE_HOST):
        raw = f"https://{raw}"

    parsed = urlsplit(raw)
    if parsed.scheme in HTTP_URL_SCHEMES and parsed.netloc:
        return urlunsplit((parsed.scheme, parsed.netloc, parsed.path or "/", parsed.query, parsed.fragment))

    # 微信内置浏览器有时会给出 weixin://... 这类地址。评论接口仍然必须走 HTTPS。
    path = parsed.path or "/"
    if not path.startswith("/"):
        path = f"/{path}"
    return urlunsplit(("https", WECHAT_ARTICLE_HOST, path, parsed.query, ""))


def extract_comment_identity(
    article_url: str,
    html: str,
    request_headers: Mapping[str, Any] | None = None,
) -> dict[str, str]:
    try:
        normalized_url = normalize_article_url(article_url)
    except CommentParametersMissing:
        normalized_url = ""
    query = parse_qs(urlsplit(normalized_url).query, keep_blank_values=True)
    cookies = _cookie_map(_header_value(request_headers, "cookie"))
    return {
        "biz": _query_value(query, "__biz") or _extract_js_value(html, "biz"),
        "mid": (
            _query_value(query, "mid")
            or _query_value(query, "appmsgid")
            or _extract_js_value(html, "mid")
            or _extract_js_value(html, "appmsgid")
        ),
        "idx": _query_value(query, "idx") or _extract_js_value(html, "idx"),
        "comment_id": _extract_comment_id(html),
        "uin": _query_value(query, "uin") or cookies.get("wxuin", ""),
        "key": _query_value(query, "key"),
        "pass_ticket": _query_value(query, "pass_ticket") or cookies.get("pass_ticket", ""),
        "wxtoken": _query_value(query, "wxtoken") or "777",
        "devicetype": _query_value(query, "devicetype"),
        "clientversion": _query_value(query, "clientversion") or _query_value(query, "version"),
        "appmsg_token": (
            _query_value(query, "appmsg_token")
            or cookies.get("appmsg_token", "")
            or _extract_js_value(html, "appmsg_token")
        ),
        "comment_scene": _query_value(query, "comment_scene") or _extract_js_value(html, "comment_scene"),
        "scene": _query_value(query, "scene") or _extract_js_value(html, "source"),
        "subscene": _query_value(query, "subscene"),
        "send_time": _query_value(query, "send_time") or _extract_js_value(html, "ct"),
        "sessionid": _query_value(query, "sessionid"),
        "enterid": _query_value(query, "enterid"),
        "ascene": _query_value(query, "ascene"),
        "lang": _query_value(query, "lang") or _extract_js_value(html, "lang"),
        "countrycode": _query_value(query, "countrycode"),
        "x5": _query_value(query, "x5") or "0",
    }


def _build_comment_page_url(
    article_url: str,
    identity: Mapping[str, str],
    *,
    offset: int,
    limit: int,
    buffer: str,
) -> str:
    params = _base_params(identity)
    params.extend(
        [("action", "getcomment"), ("offset", str(offset)), ("limit", str(limit)), ("f", "json")]
    )
    if buffer:
        params.append(("buffer", buffer))
    return _comment_endpoint(article_url, params)


def _build_reply_url(
    article_url: str,
    identity: Mapping[str, str],
    comment: Mapping[str, Any],
    *,
    offset: int,
    limit: int,
    buffer: str,
    max_reply_id: str,
    filter_reply_ids: list[str],
) -> str:
    content_id = str(comment.get("content_id") or "")
    comment_item_id = str(comment.get("id") or comment.get("comment_id") or "")
    if not content_id or not comment_item_id:
        return ""
    params = _base_params(identity)
    params.extend(
        [
            ("action", "getcommentreply"),
            ("content_id", content_id),
            ("id", comment_item_id),
            ("r", str(time.time())),
            ("offset", str(offset)),
            ("limit", str(limit)),
            ("buffer", buffer),
            ("is_first", "0" if offset else "1"),
            ("max_reply_id", max_reply_id),
            ("comment_nickname", str(comment.get("nick_name") or comment.get("nickname") or "")),
            ("comment_headurl", str(comment.get("logo_url") or comment.get("avatar_url") or "")),
            ("f", "json"),
        ]
    )
    params.extend(("filter_reply_list", reply_id) for reply_id in filter_reply_ids)
    return _comment_endpoint(article_url, params)


def _base_params(identity: Mapping[str, str]) -> list[tuple[str, str]]:
    mapping = (
        ("__biz", "biz"),
        ("appmsgid", "mid"),
        ("idx", "idx"),
        ("comment_id", "comment_id"),
        ("uin", "uin"),
        ("key", "key"),
        ("pass_ticket", "pass_ticket"),
        ("wxtoken", "wxtoken"),
        ("devicetype", "devicetype"),
        ("clientversion", "clientversion"),
        ("appmsg_token", "appmsg_token"),
        ("comment_scene", "comment_scene"),
        ("scene", "scene"),
        ("subscene", "subscene"),
        ("send_time", "send_time"),
        ("sessionid", "sessionid"),
        ("enterid", "enterid"),
        ("ascene", "ascene"),
        ("lang", "lang"),
        ("countrycode", "countrycode"),
        ("x5", "x5"),
    )
    return [(output, str(identity.get(source) or "")) for output, source in mapping]


def _comment_endpoint(article_url: str, params: list[tuple[str, str]]) -> str:
    parsed = urlsplit(normalize_article_url(article_url))
    compact = [(key, value) for key, value in params if value != ""]
    return urlunsplit(
        (
            parsed.scheme if parsed.scheme in HTTP_URL_SCHEMES else "https",
            parsed.netloc or WECHAT_ARTICLE_HOST,
            "/mp/appmsg_comment",
            urlencode(compact),
            "",
        )
    )


def _normalize_headers(raw: Mapping[str, Any] | None, *, referer: str) -> dict[str, str]:
    return build_wechat_json_headers(raw, referer=referer)


def _merge_comment_pages(pages: list[dict[str, Any]]) -> dict[str, Any]:
    merged: dict[str, Any] = {}
    for key in COMMENT_LIST_KEYS:
        merged[key] = []
    seen_by_key = {key: set() for key in COMMENT_LIST_KEYS}
    for page in pages:
        for key in COMMENT_LIST_KEYS:
            target = merged[key]
            for item in _dict_list(page.get(key)):
                identity = _raw_identity(item)
                if identity in seen_by_key[key]:
                    continue
                seen_by_key[key].add(identity)
                target.append(item)
    if pages:
        for key in ("continue_flag", "total_count", "elected_comment_total_cnt", "reply_flag", "buffer"):
            if key in pages[-1]:
                merged[key] = pages[-1][key]
    return merged


def _normalize_comments(data: Mapping[str, Any]) -> list[dict[str, Any]]:
    sections = (
        ("normal", "comment"),
        ("normal", "comment_list"),
        ("elected", "elected_comment"),
        ("elected", "elected_comment_list"),
        ("friend", "friend_comment"),
        ("mine", "my_comment"),
    )
    result: list[dict[str, Any]] = []
    seen: dict[str, dict[str, Any]] = {}
    for section, source_key in sections:
        for item in _dict_list(data.get(source_key)):
            identity = _raw_identity(item)
            if identity in seen:
                # 同一条评论可能同时出现在普通/精选列表里；去重时保留一份 raw，
                # 只合并来源标记和缺失字段，不裁剪微信接口原始内容。
                _merge_missing_raw_fields(seen[identity], item)
                source_keys = seen[identity]["_source_keys"]
                if source_key not in source_keys:
                    source_keys.append(source_key)
                if section == "elected":
                    seen[identity]["_section"] = "elected"
                    seen[identity]["_computed"]["is_elected"] = True
                continue
            reply_new = item.get("reply_new")
            reply_new = reply_new if isinstance(reply_new, Mapping) else {}
            replies = [
                _normalize_reply(reply, item)
                for reply in _unique_replies(_dict_list(reply_new.get("reply_list") or item.get("reply_list")))
            ]
            expected = _safe_int(reply_new.get("reply_total_cnt", item.get("reply_total_cnt")), len(replies))
            normalized = dict(item)
            normalized_reply_new = dict(reply_new)
            normalized_reply_new["reply_list"] = replies
            normalized["reply_new"] = normalized_reply_new
            normalized["_level"] = 1
            normalized["_section"] = section
            normalized["_source_keys"] = [source_key]
            normalized["_computed"] = {
                "comment_identity": identity,
                "comment_id": item.get("comment_id") or item.get("id"),
                "content_id": item.get("content_id"),
                "reply_total_count": expected,
                "reply_available_count": len(replies),
                "reply_missing_count": max(expected - len(replies), 0),
                "is_elected": bool(item.get("is_elected") or section == "elected"),
            }
            seen[identity] = normalized
            result.append(normalized)
    return result


def _normalize_reply(reply: Mapping[str, Any], parent: Mapping[str, Any]) -> dict[str, Any]:
    normalized = dict(reply)
    normalized["_level"] = 2
    normalized["_parent_comment_id"] = parent.get("comment_id") or parent.get("id") or ""
    normalized["_parent_content_id"] = parent.get("content_id") or ""
    return normalized


def _comment_reply_count(comment: Mapping[str, Any]) -> int:
    reply_new = comment.get("reply_new")
    if not isinstance(reply_new, Mapping):
        return 0
    return len(_dict_list(reply_new.get("reply_list")))


def _merge_missing_raw_fields(target: dict[str, Any], source: Mapping[str, Any]) -> None:
    for key, value in source.items():
        if key == "reply_new":
            _merge_reply_new(target, value)
            continue
        if key not in target or _is_empty_raw_value(target.get(key)):
            target[key] = value


def _merge_reply_new(target: dict[str, Any], source: Any) -> None:
    if not isinstance(source, Mapping):
        return
    current = target.get("reply_new")
    current_reply_new = current if isinstance(current, dict) else {}
    for key, value in source.items():
        if key == "reply_list":
            replies = [
                _normalize_reply(reply, target)
                for reply in _unique_replies(
                    [
                        *_dict_list(current_reply_new.get("reply_list")),
                        *_dict_list(value),
                    ]
                )
            ]
            current_reply_new["reply_list"] = replies
        elif key not in current_reply_new or _is_empty_raw_value(current_reply_new.get(key)):
            current_reply_new[key] = value
    target["reply_new"] = current_reply_new
    computed = target.get("_computed")
    if isinstance(computed, dict):
        replies = _dict_list(current_reply_new.get("reply_list"))
        expected = _safe_int(current_reply_new.get("reply_total_cnt", target.get("reply_total_cnt")), len(replies))
        computed["reply_total_count"] = expected
        computed["reply_available_count"] = len(replies)
        computed["reply_missing_count"] = max(expected - len(replies), 0)


def _is_empty_raw_value(value: Any) -> bool:
    return value in (None, "", [], {})


def _raw_comments(data: Mapping[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for key in COMMENT_LIST_KEYS:
        for item in _dict_list(data.get(key)):
            identity = _raw_identity(item)
            if identity not in seen:
                seen.add(identity)
                result.append(item)
    return result


def _count_new_comments(payload: Mapping[str, Any], seen: set[str]) -> int:
    count = 0
    for item in _raw_comments(payload):
        identity = _raw_identity(item)
        if identity not in seen:
            seen.add(identity)
            count += 1
    return count


def _reply_missing_count(comment: dict[str, Any]) -> int:
    reply_new = _ensure_reply_new(comment)
    replies = _unique_replies(_dict_list(reply_new.get("reply_list")))
    expected = _safe_int(reply_new.get("reply_total_cnt", comment.get("reply_total_cnt")), len(replies))
    return max(expected - len(replies), 0)


def _ensure_reply_new(comment: dict[str, Any]) -> dict[str, Any]:
    value = comment.get("reply_new")
    if not isinstance(value, dict):
        value = {}
        comment["reply_new"] = value
    value["reply_list"] = _dict_list(value.get("reply_list"))
    return value


def _unique_replies(replies: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for reply in replies:
        identity = str(
            reply.get("reply_id")
            or reply.get("id")
            or reply.get("content_id")
            or json.dumps(reply, ensure_ascii=False, sort_keys=True)
        )
        if identity not in seen:
            seen.add(identity)
            result.append(reply)
    return result


def _reply_ids(replies: list[dict[str, Any]]) -> list[str]:
    return [str(value) for item in replies if (value := item.get("reply_id") or item.get("id")) not in (None, "")]


def _dict_list(value: Any) -> list[dict[str, Any]]:
    if isinstance(value, dict):
        return [value]
    if isinstance(value, list):
        return [item for item in value if isinstance(item, dict)]
    return []


def _raw_identity(item: Mapping[str, Any]) -> str:
    return str(item.get("content_id") or item.get("id") or item.get("comment_id") or json.dumps(item, ensure_ascii=False, sort_keys=True))


def _extract_js_value(source: str, name: str) -> str:
    text = str(source or "")
    string_match = re.search(
        rf"(?<![\w$])(?:var\s+|window\.)?{re.escape(name)}(?![\w$])\s*=\s*(['\"])(?P<value>.*?)(?<!\\)\1",
        text,
    )
    if string_match:
        return html_module.unescape(
            string_match.group("value").replace("\\/", "/").replace("\\x26", "&").replace("\\u0026", "&")
        )
    number_match = re.search(
        rf"(?<![\w$])(?:var\s+|window\.)?{re.escape(name)}(?![\w$])\s*=\s*(?P<value>-?\d+(?:\.\d+)?)(?:\s*[;,])",
        text,
    )
    return number_match.group("value") if number_match else ""


def _extract_comment_id(source: str) -> str:
    """提取普通评论列表使用的 comment_id，避免误取 segment/extra_comment_id。"""
    text = str(source or "")
    string_match = re.search(
        r"(?<![\w$])comment_id(?![\w$])\s*[:=]\s*(['\"])(?P<value>.*?)(?<!\\)\1",
        text,
    )
    if string_match:
        return html_module.unescape(
            string_match.group("value").replace("\\/", "/").replace("\\x26", "&").replace("\\u0026", "&")
        )
    number_match = re.search(
        r"(?<![\w$])comment_id(?![\w$])\s*[:=]\s*(?P<value>\d+)(?:\s*[;,])",
        text,
    )
    return number_match.group("value") if number_match else ""


def _clean_text(value: Any) -> str:
    text = html_module.unescape(str(value or ""))
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())


def _header_value(headers: Mapping[str, Any] | None, name: str) -> str:
    target = name.casefold()
    for key, value in dict(headers or {}).items():
        if str(key).casefold() == target:
            return str(value)
    return ""


def _cookie_map(cookie_header: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for part in str(cookie_header or "").split(";"):
        if "=" in part:
            key, value = part.split("=", 1)
            result[key.strip()] = value.strip()
    return result


def _query_value(query: Mapping[str, list[str]], key: str) -> str:
    values = query.get(key) or []
    return str(values[0]) if values else ""


def _safe_bool(value: Any) -> bool | None:
    if isinstance(value, bool):
        return value
    if value in (None, ""):
        return None
    if str(value).strip().casefold() in {"0", "false", "no"}:
        return False
    if str(value).strip().casefold() in {"1", "true", "yes"}:
        return True
    return None


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _redact_url(url: str) -> str:
    parsed = urlsplit(str(url or ""))
    query = parse_qs(parsed.query, keep_blank_values=True)
    safe_names = ("__biz", "mid", "idx", "sn")
    safe_query = [(name, query[name][0]) for name in safe_names if query.get(name)]
    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urlencode(safe_query), ""))
