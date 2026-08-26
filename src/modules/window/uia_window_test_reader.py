from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from datetime import date
import re
from typing import Any

from src.modules.window.article_date_filter import normalize_home_date_text
from src.modules.window.uia_article_group_parser import (
    is_uia_date_text,
    is_uia_metric_text,
    normalize_uia_text,
)
from src.modules.window.wechat_document_reader import find_wechat_document_control
from src.modules.window.wechat_home_window_finder import rect_to_tuple
from src.modules.window.window_models import WindowInfo


Rect = tuple[int, int, int, int]
Marker = tuple[str, str]

_NAVIGATION_TEXT = frozenset({"全部", "贴图", "文章", "视频号"})
_NON_TITLE_TEXT = frozenset(
    {
        *_NAVIGATION_TEXT,
        "公众号",
        "服务号",
        "订阅号",
        "微信",
        "展开",
        "收起",
        "置顶",
        "加载更多",
        "正在加载...",
        "发消息",
        "已关注",
    }
)

_MORE_ARTICLES_PATTERN = re.compile(r"^余下\s*(\d+)\s*篇$")

# 这些限制只控制单次 UIA 结构扫描范围，不属于用户可调的业务参数。
_DATE_LIST_TAIL_PROBE_LIMIT = 8
_DATE_GROUP_SCAN_LIMIT = 256
_TOP_BOUNDARY_GROUP_COUNT = 3
_DATE_ROW_CHILD_SCAN_LIMIT = 8
_ARTICLE_CARD_GROUP_LIMIT = 32
_ARTICLE_CARD_NODE_LIMIT = 512


@dataclass(frozen=True, slots=True)
class UiaWindowTestArticleCard:
    """窗口测试使用的单篇卡片快照，不持有可变 UIA 控件引用。"""

    date_text: str
    published_date: str
    raw_title: str
    title: str
    date_rect: Rect | None
    title_rect: Rect | None
    card_rect: Rect
    visible_rect: Rect | None
    visible_height: int
    click_point: tuple[int, int] | None

    @property
    def marker(self) -> Marker:
        return (normalize_uia_text(self.date_text), normalize_uia_text(self.raw_title))

    @property
    def is_more_trigger(self) -> bool:
        """当前卡片是否为“余下 xx 篇”的文章展开入口。"""

        return _MORE_ARTICLES_PATTERN.fullmatch(
            normalize_uia_text(self.raw_title)
        ) is not None

    @property
    def remaining_count(self) -> int:
        """返回展开入口标注的剩余文章数量，普通文章返回 0。"""

        matched = _MORE_ARTICLES_PATTERN.fullmatch(
            normalize_uia_text(self.raw_title)
        )
        return int(matched.group(1)) if matched is not None else 0


@dataclass(frozen=True, slots=True)
class UiaWindowTestDateGroup:
    date_text: str
    published_date: str
    date_rect: Rect | None
    group_rect: Rect
    cards: tuple[UiaWindowTestArticleCard, ...]


@dataclass(frozen=True, slots=True)
class UiaWindowTestDateGroupHeader:
    """日期定位阶段使用的轻量日期组，不解析组内文章卡片。"""

    date_text: str
    published_date: str
    date_rect: Rect | None
    group_rect: Rect
    visible_rect: Rect | None
    runtime_id: tuple[int, ...] = ()


@dataclass(frozen=True, slots=True)
class UiaWindowTestDateSnapshot:
    groups: tuple[UiaWindowTestDateGroupHeader, ...]
    visible_groups: tuple[UiaWindowTestDateGroupHeader, ...]
    content_viewport: Rect
    node_count: int = 0
    loading: bool = False


@dataclass(frozen=True, slots=True)
class UiaWindowTestSnapshot:
    groups: tuple[UiaWindowTestDateGroup, ...]
    all_cards: tuple[UiaWindowTestArticleCard, ...]
    visible_cards: tuple[UiaWindowTestArticleCard, ...]
    content_viewport: Rect
    node_count: int = 0
    loading: bool = False


@dataclass(slots=True)
class _Node:
    control_type: str
    name: str
    rect: Rect
    children: list[int]


class UiaWindowTestReader:
    """为窗口测试生成一次不可变的主页 UIA 日期组/文章卡片快照。"""

    def __init__(
        self,
        *,
        max_depth: int = 14,
        max_nodes: int = 5000,
        min_visible_height: int = 10,
    ) -> None:
        self._max_depth = max(1, int(max_depth))
        self._max_nodes = max(1, int(max_nodes))
        self._min_visible_height = max(1, int(min_visible_height))
        self._date_list_control: Any | None = None
        self._image_post_grid_control: Any | None = None

    def read(self, home_window: WindowInfo) -> UiaWindowTestSnapshot:
        try:
            article_list, content_viewport = self._resolve_article_list(home_window)
        except RuntimeError as exc:
            if str(exc) != "公众号主页没有找到日期组列表":
                raise
            image_grid, content_viewport = self._resolve_image_post_grid(home_window)
            if image_grid is None:
                raise
            cards, node_count, loading = _scan_image_post_grid(
                image_grid,
                content_viewport=content_viewport,
                min_visible_height=self._min_visible_height,
                max_depth=self._max_depth,
                max_nodes=self._max_nodes,
            )
            group = UiaWindowTestDateGroup(
                date_text="今天",
                published_date=date.today().isoformat(),
                date_rect=None,
                group_rect=rect_to_tuple(
                    _safe_get(image_grid, "BoundingRectangle", None)
                ),
                cards=cards,
            )
            return UiaWindowTestSnapshot(
                groups=(group,) if cards else (),
                all_cards=cards,
                visible_cards=tuple(
                    card for card in cards if card.visible_rect is not None
                ),
                content_viewport=content_viewport,
                node_count=node_count,
                loading=loading,
            )
        if article_list is None:
            return UiaWindowTestSnapshot((), (), (), content_viewport)
        groups, node_count, loading = _scan_article_groups_from_tail(
            article_list,
            content_viewport=content_viewport,
            min_visible_height=self._min_visible_height,
            max_depth=self._max_depth,
            max_nodes=self._max_nodes,
        )
        all_cards = tuple(card for group in groups for card in group.cards)
        visible_cards = tuple(
            card for card in all_cards if card.visible_rect is not None
        )
        return UiaWindowTestSnapshot(
            groups=groups,
            all_cards=all_cards,
            visible_cards=visible_cards,
            content_viewport=content_viewport,
            node_count=node_count,
            loading=loading,
        )

    def read_date_groups(self, home_window: WindowInfo) -> UiaWindowTestDateSnapshot:
        """从文章列表末尾倒序读取当前视口附近的日期组。"""

        article_list, content_viewport = self._resolve_article_list(home_window)
        if article_list is None:
            return UiaWindowTestDateSnapshot((), (), content_viewport)
        groups, scanned_count, loading = _scan_date_groups_from_tail(
            article_list,
            content_viewport=content_viewport,
            min_visible_height=self._min_visible_height,
            max_nodes=self._max_nodes,
        )
        return UiaWindowTestDateSnapshot(
            groups=groups,
            visible_groups=tuple(
                group for group in groups if group.visible_rect is not None
            ),
            content_viewport=content_viewport,
            node_count=scanned_count,
            loading=loading,
        )

    def _resolve_article_list(self, home_window: WindowInfo) -> tuple[Any | None, Rect]:
        if home_window.control is None or not home_window.has_valid_rect:
            return None, (0, 0, 0, 0)
        document = find_wechat_document_control(home_window.control)
        if document is None:
            raise RuntimeError("公众号主页没有可读取的 UIA DocumentControl")

        document_rect = rect_to_tuple(getattr(document, "BoundingRectangle", None))
        content_viewport = _rect_intersection(document_rect, home_window.rect)
        if not _valid_rect(content_viewport):
            raise RuntimeError("公众号主页 UIA 内容区域坐标无效")

        article_list = self._date_list_control
        tail = _find_tail_date_group(article_list, content_viewport)
        if tail is None:
            article_list, tail = _find_date_group_list(
                document,
                content_viewport=content_viewport,
                max_depth=min(self._max_depth, 10),
                max_nodes=min(self._max_nodes, 256),
            )
            self._date_list_control = article_list
        if article_list is None or tail is None:
            raise RuntimeError("公众号主页没有找到日期组列表")

        content_top = _lightweight_content_top(
            document,
            article_list=article_list,
            viewport=content_viewport,
            max_depth=min(self._max_depth, 10),
            max_nodes=min(self._max_nodes, 256),
        )
        content_viewport = (
            content_viewport[0],
            min(max(content_viewport[1], content_top), content_viewport[3]),
            content_viewport[2],
            content_viewport[3],
        )
        return article_list, content_viewport

    def _resolve_image_post_grid(
        self,
        home_window: WindowInfo,
    ) -> tuple[Any | None, Rect]:
        """定位“贴图”分类的无日期双列网格。"""

        if home_window.control is None or not home_window.has_valid_rect:
            return None, (0, 0, 0, 0)
        document = find_wechat_document_control(home_window.control)
        if document is None:
            raise RuntimeError("公众号主页没有可读取的 UIA DocumentControl")
        document_rect = rect_to_tuple(getattr(document, "BoundingRectangle", None))
        content_viewport = _rect_intersection(document_rect, home_window.rect)
        if not _valid_rect(content_viewport):
            raise RuntimeError("公众号主页 UIA 内容区域坐标无效")

        grid = self._image_post_grid_control
        if not _control_has_image_post_cards(grid, minimum=1):
            grid = _find_image_post_grid(
                document,
                max_depth=min(self._max_depth, 10),
                max_nodes=min(self._max_nodes, 512),
            )
            self._image_post_grid_control = grid
        if grid is None:
            return None, content_viewport

        content_top = _lightweight_content_top(
            document,
            article_list=grid,
            viewport=content_viewport,
            max_depth=min(self._max_depth, 10),
            max_nodes=min(self._max_nodes, 256),
        )
        return grid, (
            content_viewport[0],
            min(max(content_viewport[1], content_top), content_viewport[3]),
            content_viewport[2],
            content_viewport[3],
        )


def cards_after_marker(
    snapshot: UiaWindowTestSnapshot,
    marker: Marker | None,
) -> tuple[UiaWindowTestArticleCard, ...]:
    """返回标记之后当前可见的卡片；首次读取从第一张可见卡片开始。"""

    if marker is None:
        return snapshot.visible_cards
    normalized_marker = (
        normalize_uia_text(marker[0]),
        normalize_uia_text(marker[1]),
    )
    marker_index = next(
        (
            index
            for index, card in enumerate(snapshot.all_cards)
            if card.marker == normalized_marker
        ),
        None,
    )
    if marker_index is None:
        return ()
    following = set(id(card) for card in snapshot.all_cards[marker_index + 1 :])
    return tuple(card for card in snapshot.visible_cards if id(card) in following)


def snapshot_contains_marker(
    snapshot: UiaWindowTestSnapshot,
    marker: Marker | None,
) -> bool:
    if marker is None:
        return True
    normalized_marker = (
        normalize_uia_text(marker[0]),
        normalize_uia_text(marker[1]),
    )
    return any(card.marker == normalized_marker for card in snapshot.all_cards)


def _article_card(
    card_index: int,
    *,
    date_text: str,
    published_date: str,
    date_rect: Rect | None,
    nodes: list[_Node],
    leaves: list[tuple[int, ...]],
    content_viewport: Rect,
    min_visible_height: int,
) -> UiaWindowTestArticleCard:
    title_index = _title_leaf(card_index, nodes=nodes, leaves=leaves)
    title_node = nodes[title_index] if title_index is not None else None
    raw_title = title_node.name if title_node is not None else ""
    display_title = raw_title if any(character.isalnum() for character in raw_title) else "xxx"
    card_rect = nodes[card_index].rect
    intersection = _rect_intersection(card_rect, content_viewport)
    visible_height = max(0, intersection[3] - intersection[1])
    visible_rect = intersection if visible_height >= min_visible_height else None
    click_point = _rect_center(visible_rect) if visible_rect is not None else None
    return UiaWindowTestArticleCard(
        date_text=date_text,
        published_date=published_date,
        raw_title=raw_title,
        title=display_title,
        date_rect=date_rect,
        title_rect=(
            title_node.rect
            if title_node is not None and _valid_rect(title_node.rect)
            else None
        ),
        card_rect=card_rect,
        visible_rect=visible_rect,
        visible_height=visible_height,
        click_point=click_point,
    )


def _title_leaf(
    card_index: int,
    *,
    nodes: list[_Node],
    leaves: list[tuple[int, ...]],
) -> int | None:
    candidates = [
        leaf_index
        for leaf_index in leaves[card_index]
        if _is_title_text(nodes[leaf_index].name)
    ]
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda index: (
            nodes[index].rect[1] if _valid_rect(nodes[index].rect) else 10**9,
            -len(nodes[index].name),
            index,
        ),
    )


def _is_title_text(text: str) -> bool:
    value = normalize_uia_text(text)
    if not value or len(value) > 500:
        return False
    if value.lower() in _NON_TITLE_TEXT or value in _NON_TITLE_TEXT:
        return False
    return (
        not is_uia_date_text(value)
        and not is_uia_metric_text(value)
        and not _is_image_post_metric_text(value)
    )


def _is_image_post_metric_text(text: str) -> bool:
    """贴图卡片只显示点赞/转发统计，不包含普通文章的“阅读”。"""

    value = normalize_uia_text(text)
    return bool(
        re.search(
            r"^赞\s*[\d.]+(?:万)?\+?(?:.*(?:朋友|转发|赞))?",
            value,
        )
    )


def _find_image_post_grid(
    root: Any,
    *,
    max_depth: int,
    max_nodes: int,
) -> Any | None:
    """查找至少包含一张贴图卡片的公共网格父节点。"""

    queue: deque[tuple[Any, int]] = deque([(root, 0)])
    best_match: Any | None = None
    best_depth = -1
    inspected = 0
    while queue and inspected < max_nodes:
        control, depth = queue.popleft()
        inspected += 1
        if _control_has_image_post_cards(control, minimum=1):
            # 单卡场景下外层 Document 也可能把整个网格误看成一张卡；
            # 继续向下寻找，取最深的直接卡片父节点。
            if depth > best_depth:
                best_match = control
                best_depth = depth
        if depth >= max_depth:
            continue
        queue.extend((child, depth + 1) for child in _safe_children(control))
    return best_match


def _control_has_image_post_cards(control: Any, *, minimum: int) -> bool:
    if control is None:
        return False
    found = 0
    for child in _safe_children(control):
        if _image_post_card_texts(child) is None:
            continue
        found += 1
        if found >= max(1, int(minimum)):
            return True
    return False


def _image_post_card_texts(control: Any) -> tuple[str, str] | None:
    if (
        str(_safe_get(control, "ControlTypeName", "") or "").lower()
        != "groupcontrol"
    ):
        return None
    if not _valid_rect(rect_to_tuple(_safe_get(control, "BoundingRectangle", None))):
        return None
    nodes = _snapshot_tree(
        control,
        max_depth=8,
        max_nodes=96,
    )
    if not nodes:
        return None
    leaves = _descendant_text_leaves(nodes)
    names = [nodes[index].name for index in leaves[0]]
    metric = next((name for name in names if _is_image_post_metric_text(name)), "")
    title_index = _title_leaf(0, nodes=nodes, leaves=leaves)
    title = nodes[title_index].name if title_index is not None else ""
    if not title or not metric:
        return None
    return title, metric


def _scan_image_post_grid(
    grid: Any,
    *,
    content_viewport: Rect,
    min_visible_height: int,
    max_depth: int,
    max_nodes: int,
) -> tuple[tuple[UiaWindowTestArticleCard, ...], int, bool]:
    """把贴图双列网格转换为主流程可点击的卡片快照。"""

    cards: list[UiaWindowTestArticleCard] = []
    node_count = 0
    loading = False
    published_date = date.today().isoformat()
    for child in _safe_children(grid):
        loading = loading or _control_is_loading(
            child,
            content_viewport=content_viewport,
        )
        if len(cards) >= _ARTICLE_CARD_GROUP_LIMIT:
            break
        if _image_post_card_texts(child) is None:
            continue
        nodes = _snapshot_tree(
            child,
            max_depth=max_depth,
            max_nodes=min(max_nodes, _ARTICLE_CARD_NODE_LIMIT),
        )
        node_count += len(nodes)
        if not nodes:
            continue
        leaves = _descendant_text_leaves(nodes)
        cards.append(
            _article_card(
                0,
                date_text="今天",
                published_date=published_date,
                date_rect=None,
                nodes=nodes,
                leaves=leaves,
                content_viewport=content_viewport,
                min_visible_height=min_visible_height,
            )
        )
    return (
        tuple(
            sorted(
                cards,
                key=lambda card: (
                    card.card_rect[1],
                    card.card_rect[0],
                    card.raw_title,
                ),
            )
        ),
        node_count,
        loading,
    )


def _find_date_group_list(
    root: Any,
    *,
    content_viewport: Rect,
    max_depth: int,
    max_nodes: int,
) -> tuple[Any | None, Any | None]:
    """浅层查找日期组公共父节点，不展开日期组内部文章卡片。"""

    queue: deque[tuple[Any, int]] = deque([(root, 0)])
    inspected = 0
    while queue and inspected < max_nodes:
        control, depth = queue.popleft()
        inspected += 1
        tail = _find_tail_date_group(control, content_viewport)
        if tail is not None:
            return control, tail
        if depth >= max_depth:
            continue
        queue.extend((child, depth + 1) for child in _safe_children(control))
    return None, None


def _find_tail_date_group(control: Any, content_viewport: Rect) -> Any | None:
    """跳过列表末尾的加载提示，从尾部找到第一个日期组。"""

    current = _safe_related_control(control, "GetLastChildControl")
    for _ in range(_DATE_LIST_TAIL_PROBE_LIMIT):
        if current is None:
            return None
        if _date_group_header_from_control(
            current,
            content_viewport=content_viewport,
            min_visible_height=1,
        ) is not None:
            return current
        current = _safe_related_control(current, "GetPreviousSiblingControl")
    return None


def _scan_date_groups_from_tail(
    article_list: Any,
    *,
    content_viewport: Rect,
    min_visible_height: int,
    max_nodes: int,
) -> tuple[tuple[UiaWindowTestDateGroupHeader, ...], int, bool]:
    """从列表尾部向前读取视口附近日期组，保留一个顶部边界组。"""

    controls, scanned_count, loading = _scan_date_group_controls_from_tail(
        article_list,
        content_viewport=content_viewport,
        min_visible_height=min_visible_height,
        max_nodes=max_nodes,
    )
    return tuple(header for _, header in controls), scanned_count, loading


def _scan_article_groups_from_tail(
    article_list: Any,
    *,
    content_viewport: Rect,
    min_visible_height: int,
    max_depth: int,
    max_nodes: int,
) -> tuple[tuple[UiaWindowTestDateGroup, ...], int, bool]:
    """只解析尾部视口批次中的文章卡片子树。"""

    controls, scanned_count, loading = _scan_date_group_controls_from_tail(
        article_list,
        content_viewport=content_viewport,
        min_visible_height=min_visible_height,
        max_nodes=max_nodes,
    )
    groups: list[UiaWindowTestDateGroup] = []
    node_count = scanned_count
    for control, header in controls:
        cards, card_node_count = _article_cards_from_date_group(
            control,
            header=header,
            content_viewport=content_viewport,
            min_visible_height=min_visible_height,
            max_depth=max_depth,
            max_nodes=max_nodes,
        )
        node_count += card_node_count
        if not cards:
            continue
        groups.append(
            UiaWindowTestDateGroup(
                date_text=header.date_text,
                published_date=header.published_date,
                date_rect=header.date_rect,
                group_rect=header.group_rect,
                cards=cards,
            )
        )
    return tuple(groups), node_count, loading


def _scan_date_group_controls_from_tail(
    article_list: Any,
    *,
    content_viewport: Rect,
    min_visible_height: int,
    max_nodes: int,
) -> tuple[tuple[tuple[Any, UiaWindowTestDateGroupHeader], ...], int, bool]:
    reverse_groups: list[tuple[Any, UiaWindowTestDateGroupHeader]] = []
    scanned_count = 0
    loading = False
    above_viewport_count = 0
    current = _safe_related_control(article_list, "GetLastChildControl")
    scan_limit = min(max(1, int(max_nodes)), _DATE_GROUP_SCAN_LIMIT)
    while current is not None and scanned_count < scan_limit:
        scanned_count += 1
        loading = loading or _control_is_loading(
            current,
            content_viewport=content_viewport,
        )
        header = _date_group_header_from_control(
            current,
            content_viewport=content_viewport,
            min_visible_height=min_visible_height,
        )
        if header is not None:
            reverse_groups.append((current, header))
            if header.group_rect[3] <= content_viewport[1]:
                above_viewport_count += 1
                if above_viewport_count >= _TOP_BOUNDARY_GROUP_COUNT:
                    break
        current = _safe_related_control(current, "GetPreviousSiblingControl")

    return tuple(reversed(reverse_groups)), scanned_count, loading


def _article_cards_from_date_group(
    control: Any,
    *,
    header: UiaWindowTestDateGroupHeader,
    content_viewport: Rect,
    min_visible_height: int,
    max_depth: int,
    max_nodes: int,
) -> tuple[tuple[UiaWindowTestArticleCard, ...], int]:
    date_row = _safe_related_control(control, "GetFirstChildControl")
    current = _safe_related_control(date_row, "GetNextSiblingControl")
    cards: list[UiaWindowTestArticleCard] = []
    node_count = 0
    while current is not None and len(cards) < _ARTICLE_CARD_GROUP_LIMIT:
        if (
            str(_safe_get(current, "ControlTypeName", "") or "").lower()
            == "groupcontrol"
            and _valid_rect(
                rect_to_tuple(_safe_get(current, "BoundingRectangle", None))
            )
        ):
            nodes = _snapshot_tree(
                current,
                max_depth=max_depth,
                max_nodes=min(max_nodes, _ARTICLE_CARD_NODE_LIMIT),
            )
            node_count += len(nodes)
            if nodes:
                leaves = _descendant_text_leaves(nodes)
                cards.append(
                    _article_card(
                        0,
                        date_text=header.date_text,
                        published_date=header.published_date,
                        date_rect=header.date_rect,
                        nodes=nodes,
                        leaves=leaves,
                        content_viewport=content_viewport,
                        min_visible_height=min_visible_height,
                    )
                )
        current = _safe_related_control(current, "GetNextSiblingControl")
    return tuple(cards), node_count


def _lightweight_content_top(
    root: Any,
    *,
    article_list: Any,
    viewport: Rect,
    max_depth: int,
    max_nodes: int,
) -> int:
    """只遍历文章列表外的浅层导航节点，确定正文可视区上边界。"""

    article_list_runtime_id = _safe_runtime_id(article_list)
    if _same_control(root, article_list, article_list_runtime_id):
        return _content_top_before_first_date_group(
            article_list,
            viewport=viewport,
            max_depth=max_depth,
            max_nodes=max_nodes,
        )
    navigation_bottoms: list[int] = []
    queue: deque[tuple[Any, int]] = deque([(root, 0)])
    inspected = 0
    while queue and inspected < max_nodes:
        control, depth = queue.popleft()
        inspected += 1
        if _same_control(control, article_list, article_list_runtime_id):
            continue
        name = normalize_uia_text(str(_safe_get(control, "Name", "") or ""))
        rect = rect_to_tuple(_safe_get(control, "BoundingRectangle", None))
        if name in _NAVIGATION_TEXT and _rect_intersects(rect, viewport):
            navigation_bottoms.append(rect[3])
        if depth >= max_depth:
            continue
        queue.extend((child, depth + 1) for child in _safe_children(control))
    return max(navigation_bottoms, default=viewport[1])


def _content_top_before_first_date_group(
    article_list: Any,
    *,
    viewport: Rect,
    max_depth: int,
    max_nodes: int,
) -> int:
    """兼容测试和简化树：只检查第一个日期组之前的导航兄弟节点。"""

    navigation_bottoms: list[int] = []
    current = _safe_related_control(article_list, "GetFirstChildControl")
    inspected = 0
    while current is not None and inspected < max_nodes:
        if _date_group_header_from_control(
            current,
            content_viewport=viewport,
            min_visible_height=1,
        ) is not None:
            break
        queue: deque[tuple[Any, int]] = deque([(current, 0)])
        while queue and inspected < max_nodes:
            control, depth = queue.popleft()
            inspected += 1
            name = normalize_uia_text(str(_safe_get(control, "Name", "") or ""))
            rect = rect_to_tuple(_safe_get(control, "BoundingRectangle", None))
            if name in _NAVIGATION_TEXT and _rect_intersects(rect, viewport):
                navigation_bottoms.append(rect[3])
            if depth < max_depth:
                queue.extend((child, depth + 1) for child in _safe_children(control))
        current = _safe_related_control(current, "GetNextSiblingControl")
    return max(navigation_bottoms, default=viewport[1])


def _same_control(
    control: Any,
    expected: Any,
    expected_runtime_id: tuple[int, ...],
) -> bool:
    if control is expected:
        return True
    return bool(
        expected_runtime_id and _safe_runtime_id(control) == expected_runtime_id
    )


def _date_group_header_from_control(
    control: Any,
    *,
    content_viewport: Rect,
    min_visible_height: int,
) -> UiaWindowTestDateGroupHeader | None:
    if str(_safe_get(control, "ControlTypeName", "") or "").lower() != "groupcontrol":
        return None
    group_rect = rect_to_tuple(_safe_get(control, "BoundingRectangle", None))
    if not _valid_rect(group_rect):
        return None

    date_row = _safe_related_control(control, "GetFirstChildControl")
    date_leaf = _direct_date_text_child(date_row)
    if date_leaf is None:
        return None
    article_card = _safe_related_control(date_row, "GetNextSiblingControl")
    if (
        article_card is None
        or str(_safe_get(article_card, "ControlTypeName", "") or "").lower()
        != "groupcontrol"
        or not _valid_rect(
            rect_to_tuple(_safe_get(article_card, "BoundingRectangle", None))
        )
    ):
        return None

    date_text = normalize_uia_text(str(_safe_get(date_leaf, "Name", "") or ""))
    parsed_date = normalize_home_date_text(date_text)
    intersection = _rect_intersection(group_rect, content_viewport)
    visible_height = max(0, intersection[3] - intersection[1])
    return UiaWindowTestDateGroupHeader(
        date_text=date_text,
        published_date=parsed_date.isoformat() if parsed_date is not None else "",
        date_rect=rect_to_tuple(_safe_get(date_leaf, "BoundingRectangle", None)),
        group_rect=group_rect,
        visible_rect=(
            intersection if visible_height >= max(1, int(min_visible_height)) else None
        ),
        runtime_id=_safe_runtime_id(control),
    )


def _direct_date_text_child(date_row: Any) -> Any | None:
    """日期行首层只读取文本子节点，不进入后续文章卡片子树。"""

    current = _safe_related_control(date_row, "GetFirstChildControl")
    for _ in range(_DATE_ROW_CHILD_SCAN_LIMIT):
        if current is None:
            return None
        name = normalize_uia_text(str(_safe_get(current, "Name", "") or ""))
        if is_uia_date_text(name):
            return current
        current = _safe_related_control(current, "GetNextSiblingControl")
    return None


def _safe_related_control(control: Any, method_name: str) -> Any | None:
    if control is None:
        return None
    try:
        method = getattr(control, method_name)
        return method()
    except Exception:
        return None


def _safe_runtime_id(control: Any) -> tuple[int, ...]:
    try:
        return tuple(int(value) for value in control.GetRuntimeId() or [])
    except Exception:
        return ()


def _control_is_loading(control: Any, *, content_viewport: Rect) -> bool:
    """只把当前可视区内的加载提示视为活动加载状态。"""

    name = normalize_uia_text(str(_safe_get(control, "Name", "") or ""))
    if name not in {"正在加载...", "加载中"}:
        return False
    if bool(_safe_get(control, "IsOffscreen", False)):
        return False
    rect = rect_to_tuple(_safe_get(control, "BoundingRectangle", None))
    return _rect_intersects(rect, content_viewport)


def _snapshot_tree(root: Any, *, max_depth: int, max_nodes: int) -> list[_Node]:
    result: list[_Node] = []
    queue: deque[tuple[Any, int, int | None]] = deque([(root, 0, None)])
    while queue and len(result) < max_nodes:
        control, depth, parent_index = queue.popleft()
        index = len(result)
        result.append(
            _Node(
                control_type=str(_safe_get(control, "ControlTypeName", "") or "").lower(),
                name=normalize_uia_text(str(_safe_get(control, "Name", "") or "")),
                rect=rect_to_tuple(_safe_get(control, "BoundingRectangle", None)),
                children=[],
            )
        )
        if parent_index is not None:
            result[parent_index].children.append(index)
        if depth >= max_depth:
            continue
        for child in _safe_children(control):
            queue.append((child, depth + 1, index))
    return result


def _descendant_text_leaves(nodes: list[_Node]) -> list[tuple[int, ...]]:
    result: list[tuple[int, ...]] = [() for _ in nodes]
    for index in range(len(nodes) - 1, -1, -1):
        node = nodes[index]
        has_text_child = any(
            nodes[child].control_type == "textcontrol" for child in node.children
        )
        if node.control_type == "textcontrol" and not has_text_child and node.name:
            result[index] = (index,)
            continue
        result[index] = tuple(
            leaf
            for child in node.children
            for leaf in result[child]
        )
    return result


def _safe_children(control: Any) -> list[Any]:
    try:
        return list(control.GetChildren() or [])
    except Exception:
        return []


def _safe_get(value: Any, name: str, default: Any) -> Any:
    try:
        return getattr(value, name)
    except Exception:
        return default


def _valid_rect(rect: Rect) -> bool:
    return rect[2] > rect[0] and rect[3] > rect[1]


def _rect_intersects(left: Rect, right: Rect) -> bool:
    return (
        _valid_rect(left)
        and _valid_rect(right)
        and left[0] < right[2]
        and left[2] > right[0]
        and left[1] < right[3]
        and left[3] > right[1]
    )


def _rect_intersection(left: Rect, right: Rect) -> Rect:
    if not _rect_intersects(left, right):
        return (0, 0, 0, 0)
    return (
        max(left[0], right[0]),
        max(left[1], right[1]),
        min(left[2], right[2]),
        min(left[3], right[3]),
    )


def _rect_center(rect: Rect) -> tuple[int, int]:
    return ((rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)


__all__ = [
    "Marker",
    "UiaWindowTestArticleCard",
    "UiaWindowTestDateGroupHeader",
    "UiaWindowTestDateGroup",
    "UiaWindowTestDateSnapshot",
    "UiaWindowTestReader",
    "UiaWindowTestSnapshot",
    "cards_after_marker",
    "snapshot_contains_marker",
]
