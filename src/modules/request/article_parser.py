from __future__ import annotations

from datetime import datetime, timedelta, timezone
import html as html_module
from html.parser import HTMLParser
import re
from urllib.parse import urlsplit

from src.domain.models import ArticleDetail
from src.modules.archive.archive_path_builder import canonicalize_article_link


class ArticleParseError(RuntimeError):
    """微信文章 HTML 无效或缺少必要字段。"""


METRIC_KEYS = {
    "audience_count": ("tts_heard_person_cnt", "audience_count", "ori_read_num"),
    "read_count": ("read_num", "read_num_new", "real_read_num"),
    "like_count": ("old_like_num", "old_like_count"),
    "share_count": ("share_num", "share_count"),
    "recommend_count": ("like_count", "like_num", "recommend_count", "friend_like_num"),
    "comment_count": ("comment_count", "comment_num", "elected_comment_total_cnt"),
}

_TEXT_SHARE_ITEM_SHOW_TYPES = frozenset({"10"})
_IMAGE_SHARE_ITEM_SHOW_TYPES = frozenset({"8"})


def extract_html_comment_count(html_text: str) -> int | None:
    """只从文章 HTML 中读取页面暴露的评论数，不触发任何评论接口请求。"""
    return _first_metric(str(html_text or ""), METRIC_KEYS["comment_count"])


class WechatArticleParser:
    """从微信文章 HTML 提取结构化字段，并在覆盖前完成必要校验。"""

    def parse_and_validate(
        self,
        html_text: str,
        *,
        fallback_link: str,
    ) -> ArticleDetail:
        source = str(html_text or "")
        if not _has_valid_article_body(source):
            raise ArticleParseError("HTML 不包含有效微信文章正文")

        account_name = _extract_account_name(source)
        title = _extract_title(source)
        published_time = _extract_publish_time(source)
        article_link = _extract_article_link(source, fallback_link=fallback_link)
        article_short_link = (
            _extract_article_short_link(source, fallback_link=fallback_link)
            or article_link
        )
        missing = [
            name
            for name, value in (
                ("公众号名称", account_name),
                ("文章标题", title),
                ("发布时间", published_time),
                ("文章链接", article_link),
                ("文章短链", article_short_link),
            )
            if not value
        ]
        if missing:
            raise ArticleParseError(f"HTML 缺少必要字段：{'、'.join(missing)}")

        metrics = {
            output_name: _first_metric(source, keys)
            for output_name, keys in METRIC_KEYS.items()
        }
        return ArticleDetail(
            account_name=account_name,
            article_title=title,
            published_article_time=published_time,
            article_link=article_link,
            article_short_link=article_short_link,
            ip_location=_extract_ip_location(source) or None,
            **metrics,
        )


def _has_valid_article_body(source: str) -> bool:
    parser = _ArticleBodyProbe()
    try:
        parser.feed(source)
        parser.close()
    except Exception:
        return False
    body_text = re.sub(r"\s+", "", html_module.unescape("".join(parser.text_parts)))
    if parser.found and (len(body_text) >= 4 or parser.media_count > 0):
        return True
    return (
        _has_valid_embedded_article_body(source)
        or _has_valid_text_share_article(source)
        or _has_valid_image_share_article(source)
    )


def _has_valid_image_share_article(source: str) -> bool:
    """兼容微信贴图页：正文可能只存在于页面脚本的隐藏字段中。"""
    if not _is_image_share_article(source):
        return False
    if not (_extract_title(source) and _extract_publish_time(source)):
        return False
    return _has_image_share_metadata(source)


def _is_image_share_article(source: str) -> bool:
    return _extract_item_show_type(source) in _IMAGE_SHARE_ITEM_SHOW_TYPES


def _has_valid_embedded_article_body(source: str) -> bool:
    """安全识别微信隐藏字段正文，不执行页面 JavaScript。"""
    content_match = re.search(
        r'''\bcontent_noencode\b\s*[:=]\s*(["'])(?P<value>.*?)\1''',
        source,
        re.S | re.I,
    )
    if content_match:
        content = _decode_embedded_body_for_probe(content_match.group("value"))
        has_media = bool(
            re.search(r"<(?:img|video|audio|iframe)\b", content, re.I)
        )
        visible_text = re.sub(r"(?is)<[^>]+>", "", content)
        visible_text = re.sub(
            r"\s+", "", html_module.unescape(visible_text)
        )
        return has_media or len(visible_text) >= 4
    return False


def _decode_embedded_body_for_probe(value: str) -> str:
    """只解码字符串转义；不求值、不执行微信页面脚本。"""
    text = re.sub(
        r"\\x([0-9a-fA-F]{2})",
        lambda item: chr(int(item.group(1), 16)),
        str(value or ""),
    )
    text = re.sub(
        r"\\u([0-9a-fA-F]{4})",
        lambda item: chr(int(item.group(1), 16)),
        text,
    )
    return html_module.unescape(text)


def _has_image_share_metadata(source: str) -> bool:
    """贴图没有隐藏正文时，只接受微信公开的原图清单作为正文证据。"""
    return bool(
        re.search(
            r"\b(?:share_imageinfo|picture_page_info_list)\b",
            source,
            re.I,
        )
        and re.search(
            r"\bcdn_url\b.{0,1200}?https?(?::|\\u003a|\\x3a)(?:/|\\/|\\u002f|\\x2f){2}mmbiz\.qpic\.cn(?:/|\\/|\\u002f|\\x2f)",
            source,
            re.S | re.I,
        )
    )


def _has_valid_text_share_article(source: str) -> bool:
    """兼容微信文字分享页：这类页面没有传统 js_content 正文容器。"""
    if not _is_text_share_article(source):
        return False
    title = _extract_title(source)
    published_time = _extract_publish_time(source)
    return bool(title and published_time)


def _is_text_share_article(source: str) -> bool:
    item_show_type = _extract_item_show_type(source)
    if item_show_type in _TEXT_SHARE_ITEM_SHOW_TYPES:
        return True
    return bool(
        re.search(
            r"\b(?:window\.)?(?:real_)?item_show_type\s*[:=]\s*['\"]?10\b",
            source,
        )
    )


def _extract_item_show_type(source: str) -> str:
    return _first_text(
        source,
        (
            r"\bwindow\.real_item_show_type\s*=\s*['\"]?(?P<value>\d+)",
            r"\bwindow\.item_show_type\s*=\s*['\"]?(?P<value>\d+)",
            r"\breal_item_show_type\s*[:=]\s*['\"]?(?P<value>\d+)",
            r"['\"]?item_show_type['\"]?\s*[:=]\s*['\"]?(?P<value>\d+)",
        ),
    )


class _ArticleBodyProbe(HTMLParser):
    """只累计 js_content 的完整后代，避免正则在第一个内层闭合标签处截断。"""

    _VOID_TAGS = frozenset({"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"})
    _MEDIA_TAGS = frozenset({"img", "video", "audio", "iframe"})

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.depth = 0
        self.found = False
        self.media_count = 0
        self.text_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        lowered = tag.casefold()
        if self.depth:
            if lowered in self._MEDIA_TAGS:
                self.media_count += 1
            if lowered not in self._VOID_TAGS:
                self.depth += 1
            return
        attributes = {str(key).casefold(): str(value or "") for key, value in attrs}
        if attributes.get("id") == "js_content":
            self.found = True
            self.depth = 1

    def handle_startendtag(self, tag: str, attrs) -> None:
        if self.depth and tag.casefold() in self._MEDIA_TAGS:
            self.media_count += 1

    def handle_endtag(self, tag: str) -> None:
        if self.depth and tag.casefold() not in self._VOID_TAGS:
            self.depth -= 1

    def handle_data(self, data: str) -> None:
        if self.depth:
            self.text_parts.append(data)


def _extract_account_name(source: str) -> str:
    value = _first_text(
        source,
        (
            r"\bwindow\.nickname\s*=\s*['\"](?P<value>.*?)['\"]",
            r"\bvar\s+nickname\s*=\s*['\"](?P<value>.*?)['\"]",
            r"(?is)\bwindow\.cgiDataNew\s*=\s*\{.{0,8000}?\bnick_name\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            r"\bprofile_nickname\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            r"(?is)<span[^>]+class=['\"][^'\"]*\bnickNameSpan\b[^'\"]*['\"][^>]*>(?P<value>.*?)</span>",
            r"(?is)<[^>]+id=['\"]js_wx_follow_nickname[^'\"]*['\"][^>]*>.*?<span[^>]+class=['\"][^'\"]*\bnickNameSpan\b[^'\"]*['\"][^>]*>(?P<value>.*?)</span>",
            r"(?is)<[^>]+id=['\"]js_name['\"][^>]*>(?P<value>.*?)</[^>]+>",
            r"(?i)<meta[^>]+property=['\"]og:article:author['\"][^>]+content=['\"](?P<value>.*?)['\"]",
        ),
    )
    if value.lower() in {"data-miniprogram-nickname", "nickname", "author", "undefined", "null"}:
        return ""
    return value


def _extract_title(source: str) -> str:
    return _first_text(
        source,
        (
            r"\bwindow\.msg_title\s*=\s*(?:window\.title\s*=\s*)?['\"](?P<value>.*?)['\"]",
            r"\bvar\s+msg_title\s*=\s*['\"](?P<value>.*?)['\"]",
            r"\bvar\s+appmsg_title\s*=\s*['\"](?P<value>.*?)['\"]",
            r"\bmsg_title\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            r"(?i)<meta[^>]+property=['\"]og:title['\"][^>]+content=['\"](?P<value>.*?)['\"]",
            r"(?is)<h1[^>]+id=['\"]activity-name['\"][^>]*>(?P<value>.*?)</h1>",
        ),
    )


def _extract_publish_time(source: str) -> str:
    text = _first_text(
        source,
        (
            r"\bvar\s+publish_time\s*=\s*['\"](?P<value>[^'\"]+)['\"]",
            r"\bpublish_time\s*:\s*['\"](?P<value>[^'\"]+)['\"]",
        ),
    )
    if text:
        normalized = text.replace("T", " ")
        match = re.fullmatch(r"(\d{4}-\d{2}-\d{2})[ ](\d{2})[:-](\d{2})(?::\d{2})?.*", normalized)
        if match:
            return f"{match.group(1)} {match.group(2)}:{match.group(3)}"

    epoch_text = _first_text(
        source,
        (
            r"\bwindow\.ct\s*=\s*['\"]?(?P<value>\d{9,13})['\"]?",
            r"\bvar\s+ct\s*=\s*['\"]?(?P<value>\d{9,13})['\"]?",
            r"\boriCreateTime\s*[:=]\s*['\"]?(?P<value>\d{9,13})['\"]?",
            r"\bori_create_time\s*[:=]\s*['\"]?(?P<value>\d{9,13})['\"]?",
        ),
    )
    if not epoch_text:
        return ""
    try:
        epoch = int(epoch_text)
        if epoch > 10_000_000_000:
            epoch //= 1000
        # 微信文章时间按中国标准时间展示；固定 UTC+8 避免 Windows 缺少 tzdata。
        china_timezone = timezone(timedelta(hours=8))
        return datetime.fromtimestamp(epoch, tz=china_timezone).strftime(
            "%Y-%m-%d %H:%M"
        )
    except (OSError, OverflowError, ValueError):
        return ""


def _extract_article_link(source: str, *, fallback_link: str) -> str:
    return _first_valid_article_link(
        source,
        (
            r"\bvar\s+msg_link\s*=\s*['\"](?P<value>.*?)['\"]",
            r"\bmsg_link\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            r"(?i)<meta[^>]+property=['\"]og:url['\"][^>]+content=['\"](?P<value>.*?)['\"]",
            r"(?i)<link[^>]+rel=['\"]canonical['\"][^>]+href=['\"](?P<value>.*?)['\"]",
        ),
        fallback_values=(fallback_link,),
    )


def _extract_article_short_link(source: str, *, fallback_link: str) -> str:
    return _first_valid_article_link(
        source,
        (
            r"\bvar\s+short_link\s*=\s*['\"](?P<value>.*?)['\"]",
            r"\bshort_link\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            r"(?i)<meta[^>]+property=['\"]og:url['\"][^>]+content=['\"](?P<value>.*?)['\"]",
            r"(?i)<link[^>]+rel=['\"]canonical['\"][^>]+href=['\"](?P<value>.*?)['\"]",
        ),
        fallback_values=(fallback_link,),
        require_short_path=True,
    )


def _first_valid_article_link(
    source: str,
    patterns: tuple[str, ...],
    *,
    fallback_values: tuple[str, ...] = (),
    require_short_path: bool = False,
) -> str:
    candidates = [*_iter_text(source, patterns), *fallback_values]
    for value in candidates:
        try:
            link = canonicalize_article_link(value)
        except ValueError:
            continue
        parsed = urlsplit(link)
        if (parsed.hostname or "").lower() != "mp.weixin.qq.com":
            continue
        if not parsed.path.startswith("/s"):
            continue
        if require_short_path and not parsed.path.startswith("/s/"):
            continue
        return link
    return ""


def _extract_ip_location(source: str) -> str:
    object_match = re.search(r"\bip_wording\s*[:=]\s*\{(?P<value>.*?)\}", source, re.S)
    if object_match:
        object_body = object_match.group("value")
        province = _first_text(
            object_body,
            (
                r"\bprovince_name\s*[:=]\s*['\"](?P<value>.*?)['\"]",
                r"\bprovinceName\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            ),
        )
        if province:
            return province
        country = _first_text(
            object_body,
            (
                r"\bcountry_name\s*[:=]\s*['\"](?P<value>.*?)['\"]",
                r"\bcountryName\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            ),
        )
        if country:
            return country

    return _first_text(
        source,
        (
            r"\bip_wording\s*[:=]\s*['\"](?P<value>.*?)['\"]",
            r"\bip_location\s*[:=]\s*['\"](?P<value>.*?)['\"]",
        ),
    )


def _first_metric(source: str, keys: tuple[str, ...]) -> int | None:
    for key in keys:
        escaped = re.escape(key)
        patterns = (
            rf"(?<![\w])['\"]?{escaped}['\"]?\s*[:=]\s*['\"](?P<value>\d+)['\"]",
            rf"(?<![\w])['\"]?{escaped}['\"]?\s*[:=]\s*(?P<value>\d+)\b",
        )
        for pattern in patterns:
            match = re.search(pattern, source)
            if match:
                return int(match.group("value"))
    return None


def _first_text(source: str, patterns: tuple[str, ...]) -> str:
    for cleaned in _iter_text(source, patterns):
        return cleaned
    return ""


def _iter_text(source: str, patterns: tuple[str, ...]):
    for pattern in patterns:
        for match in re.finditer(pattern, source):
            value = match.group("value")
            cleaned = _clean_text(value)
            if cleaned:
                yield cleaned


def _clean_text(value: str) -> str:
    text = re.sub(r"(?is)<[^>]+>", "", str(value or ""))
    text = html_module.unescape(text)
    text = re.sub(r"\\x([0-9a-fA-F]{2})", lambda item: chr(int(item.group(1), 16)), text)
    text = re.sub(r"\\u([0-9a-fA-F]{4})", lambda item: chr(int(item.group(1), 16)), text)
    # 微信特殊页标题常以内嵌 JS 字符串保存，普通 \n 应视为换行而不是字母 n。
    text = (
        text.replace("\\r\\n", " ")
        .replace("\\n", " ")
        .replace("\\r", " ")
        .replace("\\t", " ")
    )
    text = text.replace("\\/", "/").replace("\\'", "'").replace('\\"', '"')
    return re.sub(r"\s+", " ", text).strip()

