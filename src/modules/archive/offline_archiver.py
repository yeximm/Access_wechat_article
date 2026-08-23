from __future__ import annotations

from html import escape as escape_html_attribute
from html import unescape as unescape_html
import os
from pathlib import Path
import re
import time
from typing import Any
from dataclasses import dataclass, field
from urllib.parse import unquote, urlparse

from src.modules.archive.offline_html_rewriter import (
    rewrite_css_resource_links,
    rewrite_html_resource_links,
)
from src.modules.archive.offline_media_downloader import (
    MediaCandidate,
    download_media_candidates,
)
from src.modules.archive.offline_resource_store import save_offline_resource


CAPTURED_RESOURCE_TYPES = {"image", "media", "font", "stylesheet"}
READ_ORIGINAL_URL_PATTERNS = (
    re.compile(
        r"\b(?:window\.)?msg_source_url\s*(?:=|:)\s*"
        r"(?P<quote>['\"])(?P<value>.*?)(?P=quote)",
        re.IGNORECASE | re.DOTALL,
    ),
    re.compile(
        r"\bsource_url\s*:\s*(?P<quote>['\"])(?P<value>.*?)(?P=quote)",
        re.IGNORECASE | re.DOTALL,
    ),
)
ANCHOR_OPEN_TAG_PATTERN = re.compile(r"<a\b(?P<attrs>[^>]*)>", re.IGNORECASE | re.DOTALL)
JS_HEX_ESCAPE_PATTERN = re.compile(r"\\x(?P<value>[0-9a-fA-F]{2})")
JS_UNICODE_ESCAPE_PATTERN = re.compile(r"\\u(?P<value>[0-9a-fA-F]{4})")
EMBEDDED_MEDIA_URL_PATTERN = re.compile(
    r"https?://mpvideo\.qpic\.cn/[^\s'\"<>]+?\."
    r"(?:mp4|m3u8|m4a|mp3|webm|ogg|wav)(?:\?[^\s'\"<>]+)?",
    re.IGNORECASE,
)


@dataclass(frozen=True, slots=True)
class OfflineArchiveRequest:
    article_id: int
    article_title: str
    article_link: str
    stage_dir: Path
    browser_cache_dir: Path
    max_scroll_seconds: float
    resource_timeout_seconds: float
    navigation_url: str = ""
    navigation_mode: str = "stateless"
    navigation_user_agent: str = ""
    navigation_headers: dict[str, str] = field(default_factory=dict)
    navigation_cookies: tuple[dict[str, Any], ...] = ()


@dataclass(frozen=True, slots=True)
class OfflineArchiveResult:
    ok: bool
    stage_dir: Path
    index_html_path: Path | None = None
    assets_dir: Path | None = None
    resource_count: int = 0
    download_bytes: int = 0
    message: str = ""
    warning: str = ""


class CapturedResponseStore:
    """保存 Playwright 页面本次加载产生的响应，不发起补充网络请求。"""

    def __init__(
        self,
        assets_dir: str | Path,
    ) -> None:
        self.assets_dir = Path(assets_dir)
        self.assets_dir.mkdir(parents=True, exist_ok=True)
        self.resource_map: dict[str, str] = {}
        self.warnings: list[str] = []
        self.css_sources: dict[str, str] = {}
        self.media_candidates: list[MediaCandidate] = []
        self._media_urls: set[str] = set()
        self.download_bytes = 0

    def capture(self, response: Any) -> bool:
        url = str(getattr(response, "url", "") or "").strip()
        if not url or url in self.resource_map or url in self._media_urls:
            return False
        request = getattr(response, "request", None)
        resource_type = str(getattr(request, "resource_type", "") or "").lower()
        headers = {
            str(key).lower(): str(value)
            for key, value in dict(getattr(response, "headers", {}) or {}).items()
        }
        content_type = headers.get("content-type", "").split(";", 1)[0].strip().lower()
        if not _should_capture(resource_type, content_type):
            return False
        if _is_media_candidate(url, resource_type, content_type):
            self.register_media_candidate(
                url,
                content_type=content_type,
                request_headers=_read_request_headers(request),
            )
            return False
        try:
            body = response.body()
        except Exception as exc:
            self.warnings.append(f"读取页面响应失败：{url}：{exc}")
            return False
        if not body:
            return False
        try:
            saved = save_offline_resource(
                self.assets_dir,
                url=url,
                body=bytes(body),
                content_type=content_type,
            )
        except Exception as exc:
            if resource_type == "media" or content_type.startswith(("audio/", "video/")):
                self.warnings.append(f"媒体资源暂未离线保存，已继续保存图文：{url}：{exc}")
            else:
                self.warnings.append(f"保存页面资源失败，已继续归档：{url}：{exc}")
            return False
        self.resource_map[url] = saved.relative_path
        self.add_download_bytes(len(body))
        if content_type == "text/css":
            self.css_sources[saved.relative_path] = url
        return True

    def add_download_bytes(self, value: int) -> None:
        """累计本次离线缓存已经确认保存的下载字节，仅作为任务结果统计。"""

        bytes_value = max(0, int(value or 0))
        self.download_bytes += bytes_value

    def register_media_candidate(
        self,
        url: str,
        *,
        content_type: str,
        request_headers: dict[str, str],
    ) -> bool:
        source_url = str(url or "").strip()
        if not source_url or source_url in self.resource_map or source_url in self._media_urls:
            return False
        self.media_candidates.append(
            MediaCandidate(
                url=source_url,
                content_type=str(content_type or ""),
                request_headers=dict(request_headers or {}),
            )
        )
        self._media_urls.add(source_url)
        return True

    def rewrite_saved_css(self) -> None:
        root = self.assets_dir.parent
        for relative_path, source_url in self.css_sources.items():
            css_path = root / relative_path
            try:
                css_text = css_path.read_text(encoding="utf-8", errors="ignore")
                # CSS 内的 URL 以 CSS 文件目录为基准，不能直接复用 HTML 的 assets/... 路径。
                css_resource_map = {
                    url: Path(
                        os.path.relpath(root / local_path, start=css_path.parent)
                    ).as_posix()
                    for url, local_path in self.resource_map.items()
                }
                css_path.write_text(
                    rewrite_css_resource_links(
                        css_text,
                        css_resource_map,
                        base_url=source_url,
                    ),
                    encoding="utf-8",
                )
            except OSError as exc:
                self.warnings.append(f"重写样式资源失败：{source_url}：{exc}")


def archive_offline_article(
    request: OfflineArchiveRequest,
    *,
    on_event=None,
) -> OfflineArchiveResult:
    """使用独立 Playwright 浏览器归档一篇文章，所有资源来自本次页面响应。"""
    stage_dir = Path(request.stage_dir)
    assets_dir = stage_dir / "assets"
    stage_dir.mkdir(parents=True, exist_ok=True)
    assets_dir.mkdir(parents=True, exist_ok=True)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(request.browser_cache_dir)
    _emit(on_event, "启动浏览器", 0.0)

    try:
        from playwright.sync_api import sync_playwright
    except ModuleNotFoundError:
        return OfflineArchiveResult(
            ok=False,
            stage_dir=stage_dir,
            assets_dir=assets_dir,
            message="当前环境未安装 Playwright",
        )

    started_at = time.monotonic()
    resource_store = CapturedResponseStore(assets_dir)
    browser = None
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            context_options = _build_playwright_context_options(request)
            navigation_cookies = list(context_options.pop("cookies", []))
            context = browser.new_context(**context_options)
            if navigation_cookies:
                context.add_cookies(navigation_cookies)
            page = context.new_page()
            page.set_default_timeout(max(500, int(request.resource_timeout_seconds * 1000)))
            page.on("response", resource_store.capture)

            target_url = _navigation_target(request)
            _emit(
                on_event,
                "访问模式",
                time.monotonic() - started_at,
                status="有状态" if request.navigation_mode == "stateful" else "无状态",
            )
            _emit(on_event, "打开文章", time.monotonic() - started_at)
            goto_options: dict[str, Any] = {
                "wait_until": "domcontentloaded",
                "timeout": max(30000, int(request.resource_timeout_seconds * 1000)),
            }
            navigation_referer = _navigation_referer(request)
            if navigation_referer:
                goto_options["referer"] = navigation_referer
            page.goto(target_url, **goto_options)
            page.wait_for_timeout(500)
            _load_lazy_article_images(
                page,
                started_at=started_at,
                on_event=on_event,
            )
            _scroll_page(
                page,
                resource_store=resource_store,
                max_scroll_seconds=request.max_scroll_seconds,
                started_at=started_at,
                on_event=on_event,
            )
            _load_lazy_article_images(
                page,
                started_at=started_at,
                on_event=on_event,
            )
            page.wait_for_timeout(300)

            _pause_page_media(page)
            _download_explicit_video_posters(
                page,
                resource_store=resource_store,
                timeout_seconds=request.resource_timeout_seconds,
                started_at=started_at,
                on_event=on_event,
            )
            _download_wechat_videosnap_assets(
                page,
                resource_store=resource_store,
                timeout_seconds=request.resource_timeout_seconds,
                started_at=started_at,
                on_event=on_event,
            )
            _emit(on_event, "整理离线页面", time.monotonic() - started_at)
            _normalize_lazy_loaded_resource_attributes(page)
            original_html = str(page.content() or "")
            page_url = str(getattr(page, "url", "") or target_url)
            _register_embedded_media_candidates_from_html(
                original_html,
                resource_store,
                request_headers=request.navigation_headers,
                started_at=started_at,
                on_event=on_event,
            )
            _download_registered_media(
                resource_store,
                context=context,
                timeout_seconds=request.resource_timeout_seconds,
                started_at=started_at,
                on_event=on_event,
            )
            rewritten_html = _prepare_offline_document_html(
                original_html,
                resource_map=resource_store.resource_map,
                base_url=page_url,
            )
            resource_store.rewrite_saved_css()
            index_path = stage_dir / "index.html"
            index_path.write_text(rewritten_html, encoding="utf-8")
            browser.close()
            browser = None

        warning = "；".join(resource_store.warnings[:10])
        return OfflineArchiveResult(
            ok=True,
            stage_dir=stage_dir,
            index_html_path=index_path,
            assets_dir=assets_dir,
            resource_count=len(resource_store.resource_map),
            download_bytes=resource_store.download_bytes,
            message="离线缓存完成",
            warning=warning,
        )
    except Exception as exc:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass
        return OfflineArchiveResult(
            ok=False,
            stage_dir=stage_dir,
            assets_dir=assets_dir,
            resource_count=len(resource_store.resource_map),
            download_bytes=resource_store.download_bytes,
            message=f"离线缓存失败：{type(exc).__name__}: {exc}",
            warning="；".join(resource_store.warnings[:10]),
        )


def _navigation_target(request: OfflineArchiveRequest) -> str:
    return str(request.navigation_url or "").strip() or request.article_link


def _navigation_referer(request: OfflineArchiveRequest) -> str:
    for key, value in dict(request.navigation_headers or {}).items():
        if str(key).strip().lower() == "referer":
            return str(value).strip()
    return ""


def _build_playwright_context_options(request: OfflineArchiveRequest) -> dict[str, Any]:
    """把离线缓存导航状态转换为 Playwright context 参数，Cookie 单独 add_cookies。"""
    options: dict[str, Any] = {"viewport": {"width": 1365, "height": 1600}}
    if request.navigation_user_agent:
        options["user_agent"] = request.navigation_user_agent
    if request.navigation_headers:
        options["extra_http_headers"] = dict(request.navigation_headers)
    if request.navigation_cookies:
        options["cookies"] = [dict(cookie) for cookie in request.navigation_cookies]
    return options


def _prepare_offline_document_html(
    html_text: str,
    *,
    resource_map: dict[str, str],
    base_url: str,
) -> str:
    """保留微信原始 DOM 结构，只清理联网脚本并重写已捕获资源路径。"""
    original_page_url = _extract_read_original_url(html_text)
    html_text = _restore_read_original_link(html_text, original_page_url)
    return rewrite_html_resource_links(
        html_text,
        resource_map,
        base_url=base_url,
    )


def _extract_read_original_url(html_text: str) -> str:
    """从微信页面运行时数据中提取“阅读原文”对应的真实外部链接。"""
    for pattern in READ_ORIGINAL_URL_PATTERNS:
        for match in pattern.finditer(html_text):
            candidate = _decode_javascript_url(match.group("value"))
            if _is_http_url(candidate):
                return candidate
    return ""


def _decode_javascript_url(value: str) -> str:
    decoded = JS_HEX_ESCAPE_PATTERN.sub(
        lambda match: chr(int(match.group("value"), 16)),
        value,
    )
    decoded = JS_UNICODE_ESCAPE_PATTERN.sub(
        lambda match: chr(int(match.group("value"), 16)),
        decoded,
    )
    decoded = decoded.replace("\\/", "/").replace("\\'", "'").replace('\\"', '"')
    return unescape_html(decoded).strip()


def _restore_read_original_link(html_text: str, original_page_url: str) -> str:
    def replace_anchor(match: re.Match[str]) -> str:
        opening_tag = match.group(0)
        attrs = match.group("attrs")
        anchor_id = _read_html_attribute(attrs, "id").lower()
        class_names = set(_read_html_attribute(attrs, "class").lower().split())
        is_read_original = anchor_id == "js_view_source" or {
            "meta_primary",
            "js_wx_tap_highlight",
            "wx_tap_link",
        }.issubset(class_names)
        if not is_read_original:
            return opening_tag

        current_href = unescape_html(_read_html_attribute(attrs, "href")).strip()
        target_url = current_href if _is_http_url(current_href) else original_page_url
        if not _is_http_url(target_url):
            return opening_tag
        opening_tag = _upsert_html_attribute(opening_tag, "href", target_url)
        opening_tag = _upsert_html_attribute(opening_tag, "target", "_blank")
        return _upsert_html_attribute(opening_tag, "rel", "noopener noreferrer")

    return ANCHOR_OPEN_TAG_PATTERN.sub(replace_anchor, html_text)


def _read_html_attribute(attrs: str, name: str) -> str:
    match = re.search(
        rf"\b{re.escape(name)}\s*=\s*(?P<quote>['\"])(?P<value>.*?)(?P=quote)",
        attrs,
        re.IGNORECASE | re.DOTALL,
    )
    return "" if match is None else str(match.group("value") or "")


def _upsert_html_attribute(opening_tag: str, name: str, value: str) -> str:
    escaped_value = escape_html_attribute(value, quote=True)
    pattern = re.compile(
        rf"\b{re.escape(name)}\s*=\s*(?P<quote>['\"])(?P<value>.*?)(?P=quote)",
        re.IGNORECASE | re.DOTALL,
    )
    replacement = f'{name}="{escaped_value}"'
    if pattern.search(opening_tag):
        return pattern.sub(replacement, opening_tag, count=1)
    return f"{opening_tag[:-1]} {replacement}>"


def _is_http_url(value: str) -> bool:
    parsed = urlparse(value.strip())
    return parsed.scheme.lower() in {"http", "https"} and bool(parsed.netloc)


def _normalize_lazy_loaded_resource_attributes(page: Any) -> None:
    """把浏览器已解析出的真实资源地址写回 DOM，方便后续按响应缓存重写。"""
    page.evaluate(
        """() => {
            document.querySelectorAll('img').forEach(node => {
                const normalizeUrl = (value) => {
                    const text = String(value || '').trim();
                    if (!text || text.startsWith('data:') || text.startsWith('blob:')) {
                        return '';
                    }
                    return text;
                };
                const current = normalizeUrl(node.currentSrc || node.getAttribute('src') || '');
                const src = String(node.getAttribute('src') || '').trim();
                const isPlaceholderImage = !current || src.startsWith('data:')
                    || node.classList.contains('js_img_placeholder')
                    || node.classList.contains('wx_img_placeholder');
                const lazyCandidate = normalizeUrl(node.getAttribute('data-src'))
                    || normalizeUrl(node.getAttribute('data-original-src'))
                    || normalizeUrl(node.getAttribute('data-original'))
                    || normalizeUrl(node.getAttribute('data-lazy-src'))
                    || normalizeUrl(node.getAttribute('data-actualsrc'))
                    || normalizeUrl(node.getAttribute('data-before-oversubscription-url'));
                const candidate = isPlaceholderImage ? (lazyCandidate || current) : (current || lazyCandidate);
                if (candidate) {
                    node.setAttribute('src', candidate);
                    node.classList.remove('js_img_placeholder');
                    node.classList.remove('wx_img_placeholder');
                }
            });
            document.querySelectorAll('source').forEach(node => {
                const source = node.getAttribute('src') || node.getAttribute('data-src') || '';
                if (source) {
                    node.setAttribute('src', source);
                }
                const candidate = node.getAttribute('srcset') || node.getAttribute('data-srcset') || '';
                if (candidate) {
                    node.setAttribute('srcset', candidate);
                }
            });
            document.querySelectorAll('video, audio').forEach(node => {
                const candidate = node.currentSrc || node.getAttribute('src')
                    || node.getAttribute('data-src') || '';
                if (candidate && !candidate.startsWith('data:') && !candidate.startsWith('blob:')) {
                    node.setAttribute('src', candidate);
                }
                const poster = node.poster || node.getAttribute('poster')
                    || node.getAttribute('data-poster') || '';
                if (poster && !poster.startsWith('data:') && !poster.startsWith('blob:')) {
                    node.setAttribute('poster', poster);
                }
            });
        }"""
    )


def _load_lazy_article_images(page: Any, *, started_at: float, on_event) -> None:
    """在 Playwright 会话仍有效时触发微信正文懒加载图片，避免最终 HTML 只留下占位图。"""
    _emit(on_event, "加载正文图文资源", time.monotonic() - started_at)
    _normalize_lazy_loaded_resource_attributes(page)
    try:
        page.wait_for_load_state("networkidle", timeout=2500)
    except Exception:
        # 微信页面可能存在长连接或后台请求，networkidle 失败不应阻断图文归档。
        try:
            page.wait_for_timeout(900)
        except Exception:
            pass


def _pause_page_media(page: Any) -> None:
    """下载媒体期间保持 Playwright 会话存活，同时停止播放器继续消耗带宽。"""
    page.evaluate(
        """() => {
            document.querySelectorAll('video, audio').forEach(node => {
                try { node.pause(); } catch (_error) {}
            });
        }"""
    )


def _download_explicit_video_posters(
    page: Any,
    *,
    resource_store: CapturedResponseStore,
    timeout_seconds: float,
    started_at: float,
    on_event,
) -> None:
    """用当前 Playwright 会话显式保存视频封面，避免 poster 只留外链。"""
    poster_urls = _read_video_poster_urls(page)
    if not poster_urls:
        return
    request_context = getattr(getattr(page, "context", None), "request", None)
    if request_context is None:
        resource_store.warnings.append("视频封面离线保存失败：缺少 Playwright 请求上下文")
        return

    _emit(on_event, "下载视频封面", time.monotonic() - started_at)
    timeout_ms = max(500, int(max(0.0, timeout_seconds) * 1000))
    for url in poster_urls:
        if url in resource_store.resource_map:
            continue
        try:
            response = request_context.get(url, timeout=timeout_ms)
            if getattr(response, "ok", True) is False:
                resource_store.warnings.append(
                    f"视频封面离线保存失败：HTTP {getattr(response, 'status', '')}"
                )
                continue
            headers = {
                str(key).lower(): str(value)
                for key, value in dict(getattr(response, "headers", {}) or {}).items()
            }
            content_type = headers.get("content-type", "").split(";", 1)[0].strip().lower()
            if not content_type:
                content_type = "image/jpeg"
            if not content_type.startswith("image/"):
                resource_store.warnings.append("视频封面离线保存失败：响应不是图片")
                continue
            body = response.body()
            if not body:
                resource_store.warnings.append("视频封面离线保存失败：响应为空")
                continue
            saved = save_offline_resource(
                resource_store.assets_dir,
                url=url,
                body=bytes(body),
                content_type=content_type,
            )
            resource_store.resource_map[url] = saved.relative_path
            resource_store.add_download_bytes(len(body))
        except Exception as exc:
            resource_store.warnings.append(
                f"视频封面离线保存失败：{type(exc).__name__}"
            )


def _download_wechat_videosnap_assets(
    page: Any,
    *,
    resource_store: CapturedResponseStore,
    timeout_seconds: float,
    started_at: float,
    on_event,
) -> None:
    """识别视频号自定义节点，登记视频地址并在当前 Playwright 会话内保存附属图片。"""
    try:
        entries = page.evaluate(
            r"""() => {
                const readBackgroundUrl = (value) => {
                    const text = String(value || '');
                    const match = text.match(/url\((['"]?)(.*?)\1\)/i);
                    return match ? match[2] : '';
                };
                const items = [];
                document.querySelectorAll('mp-common-videosnap').forEach(node => {
                    items.push({
                        coverUrls: [
                            node.getAttribute('data-coverurl') || '',
                            node.getAttribute('data-cover') || '',
                            node.getAttribute('data-thumburl') || '',
                            node.getAttribute('data-poster') || '',
                            node.getAttribute('data-url') || '',
                        ],
                        imageUrls: [
                            node.getAttribute('data-headimgurl') || '',
                            node.getAttribute('data-authiconurl') || '',
                        ],
                    });
                });
                document.querySelectorAll('.wxw_wechannel_video_context').forEach(node => {
                    items.push({
                        coverUrls: [
                            readBackgroundUrl(node.style.backgroundImage),
                            readBackgroundUrl(window.getComputedStyle(node).backgroundImage),
                        ],
                        imageUrls: [],
                    });
                });
                return items;
            }"""
        )
    except Exception:
        return
    if not entries:
        return

    _emit(on_event, "识别视频号资源", time.monotonic() - started_at)
    image_urls: list[str] = []
    for entry in entries or []:
        if not isinstance(entry, dict):
            continue
        for raw_cover_url in entry.get("coverUrls") or []:
            cover_url = _normalize_external_resource_url(raw_cover_url)
            if cover_url and cover_url not in image_urls:
                image_urls.append(cover_url)
        for raw_image_url in entry.get("imageUrls") or []:
            image_url = _normalize_external_resource_url(raw_image_url)
            if image_url and image_url not in image_urls:
                image_urls.append(image_url)

    _download_context_image_urls(
        page,
        image_urls,
        resource_store=resource_store,
        timeout_seconds=timeout_seconds,
        started_at=started_at,
        on_event=on_event,
    )


def _download_context_image_urls(
    page: Any,
    urls: list[str],
    *,
    resource_store: CapturedResponseStore,
    timeout_seconds: float,
    started_at: float,
    on_event,
) -> None:
    if not urls:
        return
    request_context = getattr(getattr(page, "context", None), "request", None)
    if request_context is None:
        resource_store.warnings.append("视频号附属图片未保存：缺少 Playwright 请求上下文")
        return
    _emit(on_event, "保存视频号附属图片", time.monotonic() - started_at)
    timeout_ms = max(500, int(max(0.0, timeout_seconds) * 1000))
    for url in urls:
        if url in resource_store.resource_map:
            continue
        try:
            response = request_context.get(url, timeout=timeout_ms)
            if getattr(response, "ok", True) is False:
                resource_store.warnings.append(
                    f"视频号附属图片请求失败：HTTP {getattr(response, 'status', '')}"
                )
                continue
            headers = {
                str(key).lower(): str(value)
                for key, value in dict(getattr(response, "headers", {}) or {}).items()
            }
            content_type = headers.get("content-type", "").split(";", 1)[0].strip().lower()
            if not content_type.startswith("image/"):
                resource_store.warnings.append("视频号附属图片响应不是图片")
                continue
            body = response.body()
            if not body:
                resource_store.warnings.append("视频号附属图片响应为空")
                continue
            saved = save_offline_resource(
                resource_store.assets_dir,
                url=url,
                body=bytes(body),
                content_type=content_type,
            )
            resource_store.resource_map[url] = saved.relative_path
            resource_store.add_download_bytes(len(body))
        except Exception as exc:
            resource_store.warnings.append(
                f"视频号附属图片保存失败：{type(exc).__name__}"
            )


def _normalize_external_resource_url(value: Any) -> str:
    url = unescape_html(str(value or "").strip())
    if "%" in url and not urlparse(url).scheme:
        # 微信普通视频 iframe 的 data-cover 会把整个 URL 百分号编码。
        url = unquote(url)
    if url.startswith("//"):
        url = f"https:{url}"
    parsed = urlparse(url)
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
        return ""
    return url


def _read_video_poster_urls(page: Any) -> tuple[str, ...]:
    try:
        raw_urls = page.evaluate(
            """() => {
                const urls = [];
                document.querySelectorAll('video').forEach(node => {
                    urls.push(node.poster || node.getAttribute('poster')
                        || node.getAttribute('data-poster') || '');
                });
                document.querySelectorAll('iframe.video_iframe, iframe[data-mpvid], iframe[data-cover]').forEach(node => {
                    urls.push(node.getAttribute('data-cover') || node.getAttribute('data-coverurl')
                        || node.getAttribute('data-poster') || '');
                });
                return urls.filter(Boolean);
            }"""
        )
    except Exception:
        return ()

    urls: list[str] = []
    seen: set[str] = set()
    for item in raw_urls or []:
        url = _normalize_external_resource_url(item)
        if not url or url.lower().startswith(("data:", "blob:")):
            continue
        if url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return tuple(urls)


def _register_embedded_media_candidates_from_html(
    html_text: str,
    resource_store: CapturedResponseStore,
    *,
    request_headers: dict[str, str],
    started_at: float,
    on_event,
) -> None:
    """从微信原始 HTML 运行数据里提取普通视频真实地址，避免删脚本后丢失 MP4。"""
    urls = _extract_embedded_media_urls(html_text)
    registered = 0
    for url in urls:
        if resource_store.register_media_candidate(
            url,
            content_type=_embedded_media_content_type(url),
            request_headers=dict(request_headers or {}),
        ):
            registered += 1
    if registered:
        _emit(
            on_event,
            f"识别原始媒体资源 {registered} 个",
            time.monotonic() - started_at,
        )


def _extract_embedded_media_urls(html_text: str) -> tuple[str, ...]:
    decoded_text = _decode_embedded_resource_text(html_text)
    urls: list[str] = []
    seen_groups: set[str] = set()
    for match in EMBEDDED_MEDIA_URL_PATTERN.finditer(decoded_text):
        url = str(match.group(0) or "").strip().rstrip("),;")
        if not _is_http_url(url):
            continue
        group_key = _embedded_media_group_key(url)
        if group_key in seen_groups:
            continue
        seen_groups.add(group_key)
        urls.append(url)
    return tuple(urls)


def _decode_embedded_resource_text(html_text: str) -> str:
    decoded = JS_HEX_ESCAPE_PATTERN.sub(
        lambda match: chr(int(match.group("value"), 16)),
        str(html_text or ""),
    )
    decoded = JS_UNICODE_ESCAPE_PATTERN.sub(
        lambda match: chr(int(match.group("value"), 16)),
        decoded,
    )
    decoded = decoded.replace("\\/", "/")
    # 微信视频 URL 常以 \x26amp; 形式嵌入脚本，需要两次反转义才能还原成 &。
    for _ in range(2):
        decoded = unescape_html(decoded)
    return decoded


def _embedded_media_group_key(url: str) -> str:
    parsed = urlparse(url)
    path = re.sub(r"\.f\d+(?=\.)", ".f", parsed.path)
    return f"{parsed.netloc.lower()}{path}"


def _embedded_media_content_type(url: str) -> str:
    suffix = Path(urlparse(url).path).suffix.lower()
    if suffix == ".m3u8":
        return "application/vnd.apple.mpegurl"
    if suffix in {".m4a", ".mp3", ".ogg", ".wav"}:
        return {
            ".m4a": "audio/mp4",
            ".mp3": "audio/mpeg",
            ".ogg": "audio/ogg",
            ".wav": "audio/wav",
        }[suffix]
    if suffix == ".webm":
        return "video/webm"
    return "video/mp4"


def _download_registered_media(
    resource_store: CapturedResponseStore,
    *,
    context: Any,
    timeout_seconds: float,
    started_at: float,
    on_event,
) -> None:
    candidates = tuple(resource_store.media_candidates)
    if not candidates:
        return

    _emit(on_event, "下载音视频资源", time.monotonic() - started_at)
    try:
        cookies = context.cookies([candidate.url for candidate in candidates])
    except Exception as exc:
        cookies = []
        resource_store.warnings.append(
            f"读取浏览器媒体会话失败，尝试无 Cookie 下载：{type(exc).__name__}"
        )

    results = download_media_candidates(
        candidates,
        resource_store.assets_dir,
        cookies=cookies,
        timeout_seconds=timeout_seconds,
    )
    for result in results:
        if result.ok and result.relative_path:
            resource_store.resource_map[result.source_url] = result.relative_path
            resource_store.add_download_bytes(max(0, int(result.bytes_downloaded)))
            continue
        # 不附带候选 URL，避免鉴权参数进入日志或诊断弹窗。
        resource_store.warnings.append(result.message or "媒体资源下载失败")


def _read_request_headers(request: Any) -> dict[str, str]:
    headers = getattr(request, "headers", {})
    if callable(headers):
        headers = headers()
    try:
        return {str(name): str(value) for name, value in dict(headers or {}).items()}
    except (TypeError, ValueError):
        return {}


def _scroll_page(
    page: Any,
    *,
    resource_store: CapturedResponseStore,
    max_scroll_seconds: float,
    started_at: float,
    on_event,
) -> None:
    previous_snapshot: tuple[int, int, int] | None = None
    scroll_count = 0
    while True:
        elapsed = time.monotonic() - started_at
        if elapsed >= max(0.0, float(max_scroll_seconds)):
            resource_store.warnings.append("已达到最长滚动加载时间")
            break
        scroll_count += 1
        snapshot = _scroll_down_once(page)
        page.wait_for_timeout(180)
        top, height, viewport = _normalize_scroll_snapshot(snapshot)
        current = (top, height, len(resource_store.resource_map))
        _emit(
            on_event,
            f"页面滚动 第 {scroll_count} 次",
            time.monotonic() - started_at,
        )
        if top + viewport < height - 10:
            previous_snapshot = current
            continue
        if previous_snapshot != current:
            previous_snapshot = current
            continue

        _emit(on_event, "页面无变化，执行回弹滚动", time.monotonic() - started_at)
        _bounce_scroll_for_lazy_load(page)
        page.wait_for_timeout(500)
        bounced = _read_scroll_snapshot(page)
        bounced_top, bounced_height, _bounced_viewport = _normalize_scroll_snapshot(bounced)
        bounced_current = (
            bounced_top,
            bounced_height,
            len(resource_store.resource_map),
        )
        if bounced_current == current:
            _emit(on_event, "回弹后页面仍无变化", time.monotonic() - started_at)
            break
        _emit(on_event, "回弹后检测到新内容", time.monotonic() - started_at)
        previous_snapshot = bounced_current


def _scroll_down_once(page: Any) -> dict[str, int]:
    return dict(
        page.evaluate(
            """() => {
                const root = document.scrollingElement || document.documentElement;
                const viewport = window.innerHeight || document.documentElement.clientHeight || 1;
                window.scrollBy(0, Math.max(240, Math.floor(viewport * 0.8)));
                const height = Math.max(root.scrollHeight, document.body ? document.body.scrollHeight : 0);
                const top = root.scrollTop || window.scrollY || 0;
                return { top, height, viewport };
            }"""
        )
        or {}
    )


def _bounce_scroll_for_lazy_load(page: Any) -> None:
    """页面到底但无变化时，先上滚再下滚，给微信文章懒加载一次重新触发机会。"""
    page.evaluate(
        """() => {
            const root = document.scrollingElement || document.documentElement;
            const viewport = window.innerHeight || document.documentElement.clientHeight || 1;
            window.scrollBy(0, -Math.max(240, Math.floor(viewport * 0.35)));
            const height = Math.max(root.scrollHeight, document.body ? document.body.scrollHeight : 0);
            const top = root.scrollTop || window.scrollY || 0;
            return { top, height, viewport };
        }"""
    )
    page.wait_for_timeout(220)
    page.evaluate(
        """() => {
            const root = document.scrollingElement || document.documentElement;
            const viewport = window.innerHeight || document.documentElement.clientHeight || 1;
            window.scrollBy(0, Math.max(240, Math.floor(viewport * 0.9)));
            const height = Math.max(root.scrollHeight, document.body ? document.body.scrollHeight : 0);
            const top = root.scrollTop || window.scrollY || 0;
            return { top, height, viewport };
        }"""
    )


def _read_scroll_snapshot(page: Any) -> dict[str, int]:
    return dict(
        page.evaluate(
            """() => {
                const root = document.scrollingElement || document.documentElement;
                const viewport = window.innerHeight || document.documentElement.clientHeight || 1;
                const height = Math.max(root.scrollHeight, document.body ? document.body.scrollHeight : 0);
                const top = root.scrollTop || window.scrollY || 0;
                return { top, height, viewport };
            }"""
        )
        or {}
    )


def _normalize_scroll_snapshot(snapshot: dict[str, Any]) -> tuple[int, int, int]:
    top = int(snapshot.get("top") or 0)
    height = int(snapshot.get("height") or 0)
    viewport = int(snapshot.get("viewport") or 1)
    return top, height, viewport


def _emit(callback, name: str, elapsed_seconds: float, *, status: str = "running") -> None:
    if callback is None:
        return
    callback(
        {
            "name": name,
            "status": status,
            "elapsed_seconds": round(max(0.0, elapsed_seconds), 3),
        }
    )


def _should_capture(resource_type: str, content_type: str) -> bool:
    return resource_type in CAPTURED_RESOURCE_TYPES or content_type.startswith(
        ("image/", "video/", "audio/", "font/")
    ) or content_type == "text/css"


def _is_media_candidate(url: str, resource_type: str, content_type: str) -> bool:
    if resource_type == "media" or content_type.startswith(("video/", "audio/")):
        return True
    return Path(urlparse(url).path).suffix.lower() in {
        ".m3u8",
        ".m4a",
        ".mp3",
        ".mp4",
        ".ogg",
        ".wav",
        ".webm",
    }


__all__ = [
    "CapturedResponseStore",
    "OfflineArchiveRequest",
    "OfflineArchiveResult",
    "archive_offline_article",
]
