"""文档结构 → SlidePlan 的两条来源：分页与 .txt 列表解析。

分页（每页最多 6 条要点，超出出「（续）」页）必须对规则切分与模型分段同时成立；
.txt 的 ASCII 编号列表是列表标记，不是章节编号。
"""

from __future__ import annotations

import json
from pathlib import Path

from paiban import deck

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "fixtures" / "sample-longdoc.md"
GOLDEN = json.loads((ROOT / "fixtures" / "golden" / "expected-props.json").read_text("utf-8"))


class _Message:
    def __init__(self, content: str) -> None:
        self.content = content


class _Choice:
    def __init__(self, content: str) -> None:
        self.message = _Message(content)


class _Response:
    def __init__(self, content: str) -> None:
        self.choices = [_Choice(content)]


class _Completions:
    def __init__(self, content: str) -> None:
        self._content = content

    def create(self, **_kwargs) -> _Response:
        return _Response(self._content)


class FakeClient:
    """只实现 plan_slides 用到的那一小块 OpenAI 客户端接口。"""

    def __init__(self, content: str) -> None:
        self.chat = type("Chat", (), {"completions": _Completions(content)})()


def outline_with_heading() -> deck.DocOutline:
    return deck.DocOutline(
        title="封面标题",
        blocks=(deck.DocBlock(deck.BLOCK_HEADING, "第一讲 · 问题", 1),),
    )


# --------------------------------------------------------------------------- #
# 分页
# --------------------------------------------------------------------------- #


def test_model_plan_paginates_at_six_bullets() -> None:
    """10 条要点必须分成 6 + 4 两页，第二页标题带「（续）」。"""
    payload = json.dumps(
        {
            "slides": [
                {
                    "title": "十个要点",
                    "bullets": [f"要点{index}" for index in range(1, 11)],
                    "notes": "讲稿",
                }
            ]
        }
    )
    plan = deck.plan_slides(outline_with_heading(), client=FakeClient(payload))

    assert plan.source == deck.PLAN_SOURCE_DEEPSEEK
    assert [len(slide.bullets) for slide in plan.slides] == [6, 4]
    assert plan.slides[0].title == "十个要点"
    assert plan.slides[1].title == f"十个要点{deck.CONTINUATION_SUFFIX}"
    assert plan.slides[0].notes == "讲稿"
    assert plan.slides[1].notes == ""


def test_rules_plan_pagination_matches_golden() -> None:
    """规则路径的分页结果仍与 golden 记录的页标题一致（重构不许改输出）。"""
    plan = deck.plan_slides(deck.read_document(SOURCE))

    assert plan.source == deck.PLAN_SOURCE_RULES
    assert [slide.title for slide in plan.slides] == GOLDEN["slide_titles"]
    assert len(plan.slides) == GOLDEN["totals"]["content_slides"]


# --------------------------------------------------------------------------- #
# .txt 编号列表
# --------------------------------------------------------------------------- #


def test_numbered_list_in_txt_stays_bullets(tmp_path: Path) -> None:
    """.txt 里的 ASCII 编号行是列表项，不是章节标题。"""
    source = tmp_path / "讲稿.txt"
    source.write_text(
        "内训讲稿\n\n这一节讲清楚三件事。\n\n1. 先问清楚要回答什么问题\n2. 数据从哪来\n3. 把表格读成结构\n",
        encoding="utf-8",
    )
    outline = deck.read_document(source)

    bullets = [block for block in outline.blocks if block.kind == deck.BLOCK_BULLET]
    assert [block.text for block in bullets] == [
        "先问清楚要回答什么问题",
        "数据从哪来",
        "把表格读成结构",
    ]
    assert all(block.level == 0 for block in bullets)
    assert not [block for block in outline.blocks if block.kind == deck.BLOCK_HEADING]

    plan = deck.plan_slides(outline)
    assert [len(slide.bullets) for slide in plan.slides] == [3]


def test_chinese_numbered_lines_are_still_headings(tmp_path: Path) -> None:
    """中文编号（一、 / 1、）仍是章节标题。"""
    source = tmp_path / "讲稿.txt"
    source.write_text(
        "内训讲稿\n\n一、先问清楚\n正文一段。\n\n1、数据从哪来\n正文两段。\n",
        encoding="utf-8",
    )
    outline = deck.read_document(source)

    headings = [block for block in outline.blocks if block.kind == deck.BLOCK_HEADING]
    assert [(block.level, block.text) for block in headings] == [
        (2, "一、先问清楚"),
        (2, "1、数据从哪来"),
    ]
