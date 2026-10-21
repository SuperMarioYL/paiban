"""CJKLayoutPass 的单元测试：字体别名表、标点挤压、段落版式、确定性。

这一组只测版式引擎本身（typography.py + ooxml.py 的属性原语），不涉及长文档解析；
端到端的属性回归在 test_golden_pptx.py。
"""

from __future__ import annotations

import pytest
from pptx.oxml.ns import qn

from paiban import ooxml, typography

# 段落属性断言用的固定字号：缩进按它换算，避免落到 FALLBACK_FONT_SIZE_PT。
BODY_SIZE_PT = 20.0
TITLE_SIZE_PT = 32.0

_A_RPR = qn("a:rPr")


def tiny_deck(*bullets: str, title: str = "章节页") -> ooxml.Pptx:
    """一份最小 deck：封面 + 一页内容页（真实占位符）。"""
    return ooxml.emit(
        ooxml.DeckIR(
            title="封面标题",
            subtitle="副标题",
            slides=(
                ooxml.Slide(
                    title=title,
                    bullets=tuple(ooxml.Bullet(text) for text in bullets),
                ),
            ),
        )
    )


def layout_pass() -> typography.CJKLayoutPass:
    return typography.CJKLayoutPass(
        body_style=typography.ParagraphStyle(font_size_pt=BODY_SIZE_PT),
        title_style=typography.ParagraphStyle(
            font_size_pt=TITLE_SIZE_PT,
            first_line_indent_chars=None,
            alignment=None,
        ),
    )


def paragraphs(deck: ooxml.Pptx) -> list[tuple[str, object]]:
    """(role, paragraph) 列表，按文档顺序。"""
    return [
        (body.role, paragraph)
        for body in deck.iter_text_bodies()
        for paragraph in body.paragraphs
    ]


# --------------------------------------------------------------------------- #
# 字体别名表
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("alias", "expected"),
    [
        ("微软雅黑", "Microsoft YaHei"),
        ("宋体", "SimSun"),
        ("新宋体", "SimSun"),
        ("黑体", "SimHei"),
        ("楷体", "KaiTi"),
        ("仿宋", "FangSong"),
        ("等线", "DengXian"),
        ("思源黑体", "Source Han Sans SC"),
        ("思源宋体", "Source Han Serif SC"),
        ("Microsoft YaHei", "Microsoft YaHei"),
        ("microsoft yahei", "Microsoft YaHei"),
        ("  Microsoft   YaHei  ", "Microsoft YaHei"),
        ("Arial", "Arial"),
        ("Times New Roman", "Times New Roman"),
    ],
)
def test_resolve_font_aliases(alias: str, expected: str) -> None:
    family = typography.resolve_font(alias)
    assert family is not None
    assert family.name == expected


@pytest.mark.parametrize(
    "name", ["", None, "思源黑体 CN", "Comic Sans MS", "微软雅黑 Light", "宋 体"]
)
def test_resolve_font_rejects_unknown(name: str | None) -> None:
    assert typography.resolve_font(name) is None


def test_alias_table_normalizes_whitespace_and_case() -> None:
    assert typography.normalize_font_name("  Microsoft   YaHei ") == "microsoft yahei"
    assert typography.resolve_font("思源黑体") is typography.FONT_ALIASES["思源黑体"]


def test_chinese_families_share_one_typeface_across_slots() -> None:
    """中文族三槽同字体；拉丁族给东亚槽配中文字体，中文才不会落到方块。"""
    assert typography.resolve_font("微软雅黑").slots() == ooxml.FontSlots(
        latin="Microsoft YaHei", ea="Microsoft YaHei", cs="Microsoft YaHei"
    )
    assert typography.resolve_font("宋体").slots() == ooxml.FontSlots(
        latin="Times New Roman", ea="SimSun", cs="Times New Roman"
    )
    assert typography.resolve_font("Arial").slots() == ooxml.FontSlots(
        latin="Arial", ea="Microsoft YaHei", cs="Arial"
    )


# --------------------------------------------------------------------------- #
# 标点挤压判定
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("，。", [True, True]),
        ("！？", [True, True]),
        ("……", [True, True]),
        ("——", [True, True]),
        ("《论语》（注）", [False, False, False, True, True, False, False]),
        ("他说「好」。", [False, False, False, False, True, True]),
        ("你好，世界", [False, False, False, False, False]),
        ("", []),
    ],
)
def test_squeeze_flags(text: str, expected: list[bool]) -> None:
    assert typography.squeeze_flags(text) == expected


def test_every_full_width_punct_is_squeezable() -> None:
    """规则表里的标点都能被 squeeze_flags 判定，加字符时不会漏掉。"""
    for char in typography.FULL_WIDTH_PUNCT:
        assert typography.squeeze_flags(f"甲{char}{char}乙") == [False, True, True, False]


# --------------------------------------------------------------------------- #
# 段落版式参数
# --------------------------------------------------------------------------- #


def test_paragraph_style_defaults_match_chinese_norms() -> None:
    style = typography.ParagraphStyle()
    assert style.font == "微软雅黑"
    assert style.first_line_indent_chars == 2.0
    assert style.line_spacing == 1.5
    assert style.alignment == "just"


def test_title_style_does_not_indent_or_force_alignment() -> None:
    assert typography.TITLE_STYLE.first_line_indent_chars is None
    assert typography.TITLE_STYLE.alignment is None


def test_points_to_emu_is_two_character_width() -> None:
    """首行缩进 2 字符 = 两个字符宽，DrawingML 里就是 2 × 字号。"""
    assert ooxml.points_to_emu(2 * BODY_SIZE_PT) == 508000


# --------------------------------------------------------------------------- #
# pass 本体
# --------------------------------------------------------------------------- #


def test_pass_writes_paragraph_layout_attributes() -> None:
    deck = tiny_deck("第一条要点", "第二条要点")
    report = layout_pass().apply(deck)

    assert report.paragraphs_styled == report.paragraphs
    for role, paragraph in paragraphs(deck):
        layout = ooxml.read_paragraph_layout(paragraph.pPr)
        assert layout.line_spacing == 1.5
        assert layout.ea_line_break is True
        assert layout.hanging_punct is True
        assert layout.space_before_pt == 0.0
        assert layout.space_after_pt == 6.0
        if role == ooxml.TEXT_ROLE_TITLE:
            assert layout.indent_emu is None
            assert layout.alignment is None
        else:
            assert layout.indent_emu == ooxml.points_to_emu(2 * BODY_SIZE_PT)
            assert layout.alignment == "just"


def test_pass_writes_three_font_slots_on_every_run() -> None:
    deck = tiny_deck("中文要点，。", "Mixed 混排")
    layout_pass().apply(deck)

    expected = ooxml.FontSlots(
        latin="Microsoft YaHei", ea="Microsoft YaHei", cs="Microsoft YaHei"
    )
    runs = 0
    for _, paragraph in paragraphs(deck):
        for run in ooxml.paragraph_runs(paragraph):
            assert ooxml.read_font_slots(run.find(_A_RPR)) == expected
            assert ooxml.run_font_size_pt(run) in (BODY_SIZE_PT, TITLE_SIZE_PT)
            runs += 1
    assert runs > 0


def test_pass_squeezes_consecutive_full_width_punct() -> None:
    deck = tiny_deck("《论语》（注）", "结论——原因")
    report = layout_pass().apply(deck)

    assert report.punct_squeezed == 4  # 》（ 两个 + —— 两个
    spacings = [
        int(run.find(_A_RPR).get("spc"))
        for _, paragraph in paragraphs(deck)
        for run in ooxml.paragraph_runs(paragraph)
        if run.find(_A_RPR) is not None and run.find(_A_RPR).get("spc") is not None
    ]
    assert spacings == [int(-BODY_SIZE_PT / 2 * 100)] * 2


def test_pass_leaves_isolated_full_width_punct_alone() -> None:
    """单个全角标点不挤压：行首行尾半宽交给 hangingPunct，不做字距补偿。"""
    deck = tiny_deck("你好，世界。")
    report = layout_pass().apply(deck)

    assert report.punct_squeezed == 0
    for _, paragraph in paragraphs(deck):
        for run in ooxml.paragraph_runs(paragraph):
            rPr = run.find(_A_RPR)
            assert rPr is None or rPr.get("spc") is None


def test_pass_keeps_latin_slot_for_fonts_outside_the_alias_table() -> None:
    """别名表外的字体（符号字体、客户自有字体）只补东亚槽，不覆盖拉丁槽。"""
    deck = tiny_deck("中文要点")
    run = first_run(deck)
    ooxml.write_font_slots(ooxml.run_properties(run), ooxml.FontSlots(ea="Symbol"))

    layout_pass().apply(deck)

    assert ooxml.read_font_slots(run.find(_A_RPR)) == ooxml.FontSlots(
        latin="Symbol", ea="Microsoft YaHei", cs="Symbol"
    )


def test_pass_resolves_the_alias_of_the_source_font() -> None:
    """源字体在别名表里时按别名族重写三槽，而不是一律写默认族。"""
    deck = tiny_deck("中文要点")
    run = first_run(deck)
    ooxml.write_font_slots(ooxml.run_properties(run), ooxml.FontSlots(ea="宋体"))

    layout_pass().apply(deck)

    assert ooxml.read_font_slots(run.find(_A_RPR)) == ooxml.FontSlots(
        latin="Times New Roman", ea="SimSun", cs="Times New Roman"
    )


def test_pass_does_not_touch_text_content() -> None:
    """挤压只拆 run，段落文本一个字符都不改。"""
    deck = tiny_deck("《指标口径表》（第 2 版）……", "分组后直接对比！？")
    before = paragraph_texts(deck)

    layout_pass().apply(deck)

    assert paragraph_texts(deck) == before


def test_pass_is_deterministic() -> None:
    """同一份输入两次运行逐属性一致（plan §4 的 golden 前提）。"""
    first = tiny_deck("《指标口径表》（第 2 版）", "分组后直接对比！？")
    second = tiny_deck("《指标口径表》（第 2 版）", "分组后直接对比！？")
    layout_pass().apply(first)
    layout_pass().apply(second)

    assert snapshot(first) == snapshot(second)


def test_pass_rejects_style_font_outside_alias_table() -> None:
    deck = tiny_deck("中文要点")
    layout = typography.CJKLayoutPass(
        body_style=typography.ParagraphStyle(font="思源黑体 CN")
    )
    with pytest.raises(ValueError, match="别名表"):
        layout.apply(deck)


def test_indent_falls_back_to_default_size_only_when_nothing_declares_one() -> None:
    """段落与模板都没声明字号时，首行缩进按 FALLBACK_FONT_SIZE_PT 换算。"""
    deck = tiny_deck("没有显式字号的段落")
    typography.CJKLayoutPass().apply(deck)

    _, paragraph = next(item for item in paragraphs(deck) if item[0] == ooxml.TEXT_ROLE_BODY)
    layout = ooxml.read_paragraph_layout(paragraph.pPr)
    assert layout.indent_emu == ooxml.points_to_emu(
        2 * typography.FALLBACK_FONT_SIZE_PT
    )


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #


def first_run(deck: ooxml.Pptx):
    for _, paragraph in paragraphs(deck):
        runs = ooxml.paragraph_runs(paragraph)
        if runs:
            return runs[0]
    raise AssertionError("deck 里没有 run")


def run_texts(deck: ooxml.Pptx) -> list[str]:
    return [
        ooxml.run_text(run)
        for _, paragraph in paragraphs(deck)
        for run in ooxml.paragraph_runs(paragraph)
    ]


def paragraph_texts(deck: ooxml.Pptx) -> list[str]:
    return [
        "".join(ooxml.run_text(run) for run in ooxml.paragraph_runs(paragraph))
        for _, paragraph in paragraphs(deck)
    ]


def snapshot(deck: ooxml.Pptx) -> list[tuple[str, object]]:
    """段落属性 + run 属性的快照，用于比较两次运行的结果。"""
    snapshot: list[tuple[str, object]] = []
    for role, paragraph in paragraphs(deck):
        snapshot.append((role, ooxml.read_paragraph_layout(paragraph.pPr)))
        for run in ooxml.paragraph_runs(paragraph):
            rPr = run.find(_A_RPR)
            snapshot.append(
                (
                    ooxml.run_text(run),
                    (
                        ooxml.run_font_size_pt(run),
                        ooxml.read_font_slots(rPr),
                        None if rPr is None else rPr.get("spc"),
                    ),
                )
            )
    return snapshot
