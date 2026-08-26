from __future__ import annotations

import unittest

from src.modules.window.uia_window_test_reader import (
    _find_image_post_grid,
    _scan_image_post_grid,
)


class FakeControl:
    def __init__(
        self,
        control_type: str,
        name: str = "",
        rect: tuple[int, int, int, int] = (0, 0, 0, 0),
        children: list["FakeControl"] | None = None,
    ) -> None:
        self.ControlTypeName = control_type
        self.Name = name
        self.BoundingRectangle = rect
        self.IsOffscreen = False
        self._children = children or []

    def GetChildren(self):
        return list(self._children)


def image_card(title: str, metric: str, rect: tuple[int, int, int, int]):
    left, top, right, bottom = rect
    return FakeControl(
        "GroupControl",
        rect=rect,
        children=[
            FakeControl(
                "TextControl",
                title,
                (left + 8, bottom - 48, right - 8, bottom - 26),
            ),
            FakeControl(
                "TextControl",
                metric,
                (left + 8, bottom - 24, right - 8, bottom - 6),
            ),
        ],
    )


class WechatImagePostGridTest(unittest.TestCase):
    def test_finds_and_reads_image_post_grid_without_date_groups(self) -> None:
        first = image_card("处暑｜车行山野，静赏人间新秋", "赞 7", (20, 120, 220, 440))
        second = image_card(
            "车圈好运里程，居然被你们刷到了！",
            "赞 13 1个朋友转发",
            (230, 120, 430, 440),
        )
        grid = FakeControl(
            "GroupControl",
            rect=(10, 100, 440, 900),
            children=[first, second],
        )
        document = FakeControl(
            "DocumentControl",
            rect=(0, 0, 500, 1000),
            children=[grid],
        )

        self.assertIs(_find_image_post_grid(document, max_depth=4, max_nodes=100), grid)
        cards, _node_count, loading = _scan_image_post_grid(
            grid,
            content_viewport=(0, 80, 500, 950),
            min_visible_height=10,
            max_depth=10,
            max_nodes=500,
        )

        self.assertFalse(loading)
        self.assertEqual(
            [card.raw_title for card in cards],
            [
                "处暑|车行山野,静赏人间新秋",
                "车圈好运里程,居然被你们刷到了!",
            ],
        )
        self.assertTrue(all(card.published_date for card in cards))
        self.assertEqual(
            [card.click_point for card in cards],
            [(120, 280), (330, 280)],
        )

    def test_does_not_mistake_regular_article_cards_for_image_posts(self) -> None:
        regular = image_card("普通文章", "阅读 1029 赞 40", (20, 120, 220, 440))
        root = FakeControl(
            "DocumentControl",
            rect=(0, 0, 500, 1000),
            children=[regular],
        )

        self.assertIsNone(_find_image_post_grid(root, max_depth=4, max_nodes=100))

    def test_finds_image_post_grid_when_only_one_card_is_available(self) -> None:
        only_card = image_card("只剩一条贴图", "赞 1", (20, 120, 220, 440))
        grid = FakeControl(
            "GroupControl",
            rect=(10, 100, 440, 900),
            children=[only_card],
        )
        document = FakeControl(
            "DocumentControl",
            rect=(0, 0, 500, 1000),
            children=[grid],
        )

        self.assertIs(_find_image_post_grid(document, max_depth=4, max_nodes=100), grid)
