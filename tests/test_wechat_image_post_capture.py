from __future__ import annotations

import unittest
import sqlite3
from pathlib import Path

from src.domain.models import ResourceManifest
from src.modules.proxy.wechat_request_matcher import WechatRequestMatcher
from src.modules.request.article_parser import WechatArticleParser
from src.storage.repositories.article_repository import ArticleIndexWrite, ArticleRepository


SAMPLE_SHORT_LINKS = (
    "https://mp.weixin.qq.com/s/o11oBVMCW8UKq14hIvS3_g",
    "https://mp.weixin.qq.com/s/5Xv4RFZac0lXJm5uoN-wNg",
)


def _image_only_html(short_link: str) -> str:
    return f"""
    <html>
      <head>
        <meta property="og:title" content="贴图测试记录">
        <meta property="og:article:author" content="测试公众号">
        <meta property="og:url" content="{short_link}">
      </head>
      <body>
        <script>
          window.nickname = '测试公众号';
          window.msg_title = '贴图测试记录';
          window.ct = '1787587200';
          var msg_link = '{short_link}';
          var short_link = '{short_link}';
        </script>
        <div id="js_content"><img data-src="https://mmbiz.qpic.cn/example.jpg"></div>
      </body>
    </html>
    """


def _script_backed_image_post_html(
    short_link: str,
    *,
    embedded_fields: str,
) -> str:
    return f"""
    <html>
      <head>
        <meta property="og:title" content="脚本型贴图测试">
        <meta property="og:article:author" content="测试公众号">
        <meta property="og:url" content="{short_link}">
      </head>
      <body>
        <div id="js_content"></div>
        <script>
          window.nickname = '测试公众号';
          window.msg_title = '脚本型贴图测试';
          window.ct = '1787587200';
          window.item_show_type = 8;
          var msg_link = '{short_link}';
          var short_link = '{short_link}';
          {embedded_fields}
        </script>
      </body>
    </html>
    """


class WechatImagePostCaptureTests(unittest.TestCase):
    def test_public_short_links_are_accepted_as_references(self) -> None:
        matcher = WechatRequestMatcher(listen_started_at=10.0)

        for link in SAMPLE_SHORT_LINKS:
            with self.subTest(link=link):
                matched = matcher.match_reference(link, observed_at=11.0)
                self.assertIsNotNone(matched)
                assert matched is not None
                self.assertEqual(matched.url, link)
                self.assertEqual(matched.query_keys, ())

    def test_short_link_html_with_image_only_body_is_captured(self) -> None:
        link = SAMPLE_SHORT_LINKS[0]
        matcher = WechatRequestMatcher(listen_started_at=10.0)

        matched = matcher.match_html_response(
            link,
            html_text=_image_only_html(link),
            status_code=200,
            response_headers={"content-type": "text/html; charset=utf-8"},
            observed_at=11.0,
        )

        self.assertIsNotNone(matched)
        assert matched is not None
        self.assertEqual(matched.reference.url, link)
        self.assertEqual(matched.request_summary["source"], "mitm_response")

    def test_image_only_short_link_page_parses_into_article_detail(self) -> None:
        link = SAMPLE_SHORT_LINKS[1]

        detail = WechatArticleParser().parse_and_validate(
            _image_only_html(link),
            fallback_link=link,
        )

        self.assertEqual(detail.account_name, "测试公众号")
        self.assertEqual(detail.article_title, "贴图测试记录")
        self.assertEqual(detail.article_link, link)
        self.assertEqual(detail.article_short_link, link)

    def test_image_post_with_content_noencode_parses_without_js_content(self) -> None:
        link = SAMPLE_SHORT_LINKS[0]
        html = _script_backed_image_post_html(
            link,
            embedded_fields=(
                "window.cgiData = {content_noencode: "
                r"'\x3csection\x3e\x3cp\x3e贴图正文\x3c/p\x3e\x3c/section\x3e'};"
            ),
        )

        detail = WechatArticleParser().parse_and_validate(html, fallback_link=link)

        self.assertEqual(detail.article_title, "脚本型贴图测试")
        self.assertEqual(detail.article_short_link, link)

    def test_content_noencode_remains_valid_when_type_marker_is_missing(self) -> None:
        link = SAMPLE_SHORT_LINKS[0]
        html = _script_backed_image_post_html(
            link,
            embedded_fields=(
                "window.cgiData = {content_noencode: "
                r"'\x3csection\x3e\x3cimg data-src=\x22https://mmbiz.qpic.cn/example.jpg\x22\x3e\x3c/section\x3e'};"
            ),
        ).replace("window.item_show_type = 8;", "")

        detail = WechatArticleParser().parse_and_validate(html, fallback_link=link)

        self.assertEqual(detail.article_title, "脚本型贴图测试")

    def test_plain_text_content_noencode_matches_real_image_post_shape(self) -> None:
        link = SAMPLE_SHORT_LINKS[0]
        html = _script_backed_image_post_html(
            link,
            embedded_fields=(
                "window.cgiData = {content_noencode: "
                r"'贴图正文第一行\x0a第二行\x0a\x3ca href=\x22#\x22\x3e#话题\x3c/a\x3e'};"
            ),
        )

        detail = WechatArticleParser().parse_and_validate(html, fallback_link=link)

        self.assertEqual(detail.article_short_link, link)

    def test_image_post_with_share_imageinfo_parses_as_one_record(self) -> None:
        link = SAMPLE_SHORT_LINKS[1]
        html = _script_backed_image_post_html(
            link,
            embedded_fields=(
                "window.share_imageinfo = [{cdn_url: "
                "'https://mmbiz.qpic.cn/mmbiz_png/example/0'}];"
            ),
        )

        detail = WechatArticleParser().parse_and_validate(html, fallback_link=link)

        self.assertEqual(detail.account_name, "测试公众号")
        self.assertEqual(detail.article_link, link)

    def test_image_post_with_picture_page_info_list_parses(self) -> None:
        link = SAMPLE_SHORT_LINKS[1]
        html = _script_backed_image_post_html(
            link,
            embedded_fields=(
                "window.cgiData = {picture_page_info_list: [{cdn_url: "
                "'https://mmbiz.qpic.cn/mmbiz_jpg/example/0'}]};"
            ),
        )

        detail = WechatArticleParser().parse_and_validate(html, fallback_link=link)

        self.assertEqual(detail.article_title, "脚本型贴图测试")

    def test_empty_item_show_type_eight_shell_is_rejected(self) -> None:
        link = SAMPLE_SHORT_LINKS[0]
        html = _script_backed_image_post_html(link, embedded_fields="")

        with self.assertRaisesRegex(Exception, "不包含有效微信文章正文"):
            WechatArticleParser().parse_and_validate(html, fallback_link=link)

    def test_image_post_short_link_is_upserted_as_one_article_record(self) -> None:
        link = SAMPLE_SHORT_LINKS[0]
        detail = WechatArticleParser().parse_and_validate(
            _image_only_html(link),
            fallback_link=link,
        )
        schema_path = (
            Path(__file__).resolve().parents[1]
            / "data"
            / "sql"
            / "create_script"
            / "create_awa_v2_1.sql"
        )
        connection = sqlite3.connect(":memory:")
        connection.row_factory = sqlite3.Row
        try:
            connection.executescript(schema_path.read_text(encoding="utf-8"))
            account_id = int(
                connection.execute(
                    "INSERT INTO awa_public_accounts(account_name) VALUES (?) RETURNING id",
                    (detail.account_name,),
                ).fetchone()["id"]
            )
            write = ArticleIndexWrite(
                account_id=account_id,
                article_title=detail.article_title,
                published_article_time=detail.published_article_time,
                article_link=detail.article_link,
                archive_dir="测试公众号/贴图测试记录",
                resource_manifest=ResourceManifest(),
                collected_time="2026-08-25 18:00:00",
            )
            repository = ArticleRepository(connection)

            first_id = repository.upsert(write)
            second_id = repository.upsert(write)

            self.assertEqual(first_id, second_id)
            self.assertEqual(
                connection.execute(
                    "SELECT COUNT(*) FROM awa_public_articles WHERE article_link = ?",
                    (link,),
                ).fetchone()[0],
                1,
            )
        finally:
            connection.close()

    def test_unrelated_or_malformed_paths_remain_rejected(self) -> None:
        matcher = WechatRequestMatcher(listen_started_at=10.0)
        rejected = (
            "https://example.com/s/o11oBVMCW8UKq14hIvS3_g",
            "https://mp.weixin.qq.com/not-an-article/o11oBVMCW8UKq14hIvS3_g",
            "https://mp.weixin.qq.com/s/too-short",
            "https://mp.weixin.qq.com/s/token/extra",
        )

        for link in rejected:
            with self.subTest(link=link):
                self.assertIsNone(matcher.match_reference(link, observed_at=11.0))


if __name__ == "__main__":
    unittest.main()
