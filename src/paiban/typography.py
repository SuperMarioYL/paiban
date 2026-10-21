"""中文版式引擎：CJKLayoutPass。

规则（mvp_plan.md §2 的 CJKLayoutPass，全部落到 OOXML 属性上）：

font_alias
    字体别名 → ``<a:latin>`` / ``<a:ea>`` / ``<a:cs>`` 三字体槽。以 run 的 ``ea``
    槽（其次 ``latin`` 槽）作为源字体名查别名表；命中就写该族三槽，别名表外的
    字体保留拉丁槽、只补东亚槽（不把符号字体覆盖掉），没有显式字体则写角色默认
    族三槽。
punct_compress
    全角标点挤压。连续标点（相邻两侧任一是全角标点）按半宽挤压，做法是把 run 拆
    成多段、给被挤压的段写 ``<a:rPr spc>`` 负字距；行首 / 行尾半宽交给
    ``<a:pPr hangingPunct="1">`` —— 断行位置由渲染器决定，本层不猜行边界，也不
    改动文本内容。
kinsoku
    避头尾。显式打开 ``<a:pPr eaLnBrk="1">``：这是 DrawingML 里控制东亚换行与
    禁则处理的开关（规范默认 true，本层保证它不会被继承成 off）。
paragraph
    首行缩进 2 字符（``<a:pPr indent>``，按段落字号换算，DrawingML 没有“字符”
    单位）、行距 1.5（``<a:lnSpc><a:spcPct val="150000"/>``）、段前后间距
    （``<a:spcBef>`` / ``<a:spcAft>``）、两端对齐（``algn="just"``）。

确定性：同一份输入两次运行必须逐属性一致（plan §4），所以本模块不做任何依赖
时间、随机数或渲染器的决策。

诚实边界：这些属性是否被 WPS 按其版本正确渲染，本层不做断言（plan 的
``schema_unverified: true``）——这里只保证写进 OOXML 的属性本身正确、可被
``paiban verify`` 复核。
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from . import ooxml

if TYPE_CHECKING:
    from pptx.oxml.text import CT_RegularTextRun, CT_TextParagraph

# 全角标点：连续出现时按半宽挤压。行首 / 行尾半宽由 hangingPunct 交给渲染器。
FULL_WIDTH_PUNCT = frozenset(
    "，。、；：？！"  # 点号
    "“”‘’"  # 引号
    "（）〔〕［］｛｝【】《》〈〉「」『』"  # 括号
    "…—～·"  # 其他
)

# 段落没有显式字号、模板也没声明字号时的回退值，只用来换算首行缩进 2 字符。
# 模板族在“版式参数”里给出 font_size_pt 时，缩进就是精确的 2 字符。
FALLBACK_FONT_SIZE_PT = 32.0


@dataclass(frozen=True)
class FontFamily:
    """一个字体族在三字体槽里的落点，以及它认得的中文别名。

    中文族让三个槽指向同一个字体；拉丁族给 ``ea`` 槽配一个中文字体，中文才不会
    落到不可用字体上（在 WPS 里就是方块）。``name`` 是规范名，也进别名表。
    """

    name: str
    latin: str
    ea: str
    cs: str
    aliases: tuple[str, ...] = ()

    def slots(self) -> ooxml.FontSlots:
        return ooxml.FontSlots(latin=self.latin, ea=self.ea, cs=self.cs)


# 别名表本体：这是本产品自己的知识，也是可以单独授权的那部分资产。
FONT_FAMILIES: tuple[FontFamily, ...] = (
    FontFamily(
        name="Microsoft YaHei",
        latin="Microsoft YaHei",
        ea="Microsoft YaHei",
        cs="Microsoft YaHei",
        aliases=("微软雅黑",),
    ),
    FontFamily(
        name="SimSun",
        latin="Times New Roman",
        ea="SimSun",
        cs="Times New Roman",
        aliases=("宋体", "新宋体"),
    ),
    FontFamily(
        name="SimHei",
        latin="Arial",
        ea="SimHei",
        cs="Arial",
        aliases=("黑体",),
    ),
    FontFamily(
        name="KaiTi",
        latin="Times New Roman",
        ea="KaiTi",
        cs="Times New Roman",
        aliases=("楷体",),
    ),
    FontFamily(
        name="FangSong",
        latin="Times New Roman",
        ea="FangSong",
        cs="Times New Roman",
        aliases=("仿宋",),
    ),
    FontFamily(
        name="DengXian",
        latin="DengXian",
        ea="DengXian",
        cs="DengXian",
        aliases=("等线",),
    ),
    FontFamily(
        name="Source Han Sans SC",
        latin="Source Han Sans SC",
        ea="Source Han Sans SC",
        cs="Source Han Sans SC",
        aliases=("思源黑体",),
    ),
    FontFamily(
        name="Source Han Serif SC",
        latin="Source Han Serif SC",
        ea="Source Han Serif SC",
        cs="Source Han Serif SC",
        aliases=("思源宋体",),
    ),
    FontFamily(
        name="Arial",
        latin="Arial",
        ea="Microsoft YaHei",
        cs="Arial",
    ),
    FontFamily(
        name="Calibri",
        latin="Calibri",
        ea="Microsoft YaHei",
        cs="Calibri",
    ),
    FontFamily(
        name="Times New Roman",
        latin="Times New Roman",
        ea="SimSun",
        cs="Times New Roman",
    ),
    FontFamily(
        name="Consolas",
        latin="Consolas",
        ea="Microsoft YaHei",
        cs="Consolas",
    ),
)


def normalize_font_name(name: str) -> str:
    """规范化字体名：合并空白 + 大小写折叠，用于查表。"""
    return " ".join(name.split()).casefold()


FONT_ALIASES: Mapping[str, FontFamily] = {
    key: family
    for family in FONT_FAMILIES
    for key in (
        normalize_font_name(family.name),
        *(normalize_font_name(alias) for alias in family.aliases),
    )
}


def resolve_font(name: str | None) -> FontFamily | None:
    """字体名 → 字体族；别名表里没有返回 None。"""
    if not name:
        return None
    return FONT_ALIASES.get(normalize_font_name(name))


@dataclass(frozen=True)
class ParagraphStyle:
    """正文 / 标题的版式参数，对应模板族 JSON 里的“版式参数”。

    字段为 None 表示不写该属性、保留母版继承，所以模板只声明自己关心的部分。
    ``first_line_indent_chars`` 为 0 表示显式取消首行缩进。
    """

    font: str = "微软雅黑"
    font_size_pt: float | None = None
    first_line_indent_chars: float | None = 2.0
    line_spacing: float | None = 1.5
    space_before_pt: float | None = 0.0
    space_after_pt: float | None = 6.0
    alignment: str | None = "just"


# 正文：首行缩进 2 字符、行距 1.5、两端对齐。标题：不缩进、不强制对齐，只统一
# 行距与段后间距，标题的居中 / 左对齐交给模板母版。
BODY_STYLE = ParagraphStyle()
TITLE_STYLE = ParagraphStyle(first_line_indent_chars=None, alignment=None)


@dataclass
class LayoutReport:
    """一次 pass 的改写统计，供 ``paiban fix`` / ``build`` 输出。"""

    text_bodies: int = 0
    paragraphs: int = 0
    runs: int = 0
    fonts_resolved: int = 0
    fonts_defaulted: int = 0
    punct_squeezed: int = 0
    paragraphs_styled: int = 0


@dataclass(frozen=True)
class CJKLayoutPass:
    """中文版式层：一次确定性的 OOXML 文本属性改写。"""

    body_style: ParagraphStyle = BODY_STYLE
    title_style: ParagraphStyle = TITLE_STYLE

    def apply(self, deck: ooxml.Pptx) -> LayoutReport:
        """对整份 deck 应用中文版式，返回改写统计。"""
        report = LayoutReport()
        for body in deck.iter_text_bodies():
            report.text_bodies += 1
            style = self.title_style if body.role == ooxml.TEXT_ROLE_TITLE else self.body_style
            family = resolve_font(style.font)
            if family is None:
                raise ValueError(f"模板字体 {style.font!r} 不在别名表里，请先补进 FONT_FAMILIES")
            for paragraph in body.paragraphs:
                report.paragraphs += 1
                self._apply_paragraph(paragraph, style, family, report)
        return report

    def _apply_paragraph(
        self,
        paragraph: CT_TextParagraph,
        style: ParagraphStyle,
        family: FontFamily,
        report: LayoutReport,
    ) -> None:
        size_pt = (
            ooxml.paragraph_font_size_pt(paragraph)
            or style.font_size_pt
            or FALLBACK_FONT_SIZE_PT
        )
        indent_emu = (
            None
            if style.first_line_indent_chars is None
            else ooxml.points_to_emu(style.first_line_indent_chars * size_pt)
        )
        ooxml.write_paragraph_layout(
            paragraph.get_or_add_pPr(),
            indent_emu=indent_emu,
            line_spacing=style.line_spacing,
            space_before_pt=style.space_before_pt,
            space_after_pt=style.space_after_pt,
            alignment=style.alignment,
            hanging_punct=True,
            ea_line_break=True,
        )
        report.paragraphs_styled += 1
        for run in ooxml.paragraph_runs(paragraph):
            report.runs += 1
            self._apply_run(run, style, family, report)
        report.punct_squeezed += self._squeeze_punctuation(paragraph, size_pt)

    def _apply_run(
        self,
        run: CT_RegularTextRun,
        style: ParagraphStyle,
        family: FontFamily,
        report: LayoutReport,
    ) -> None:
        rPr = ooxml.run_properties(run)
        source = ooxml.read_font_slots(rPr)
        # 中文用哪个字体以 ea 槽为准：它才是决定中文渲染的槽。
        source_name = None if source is None else (source.ea or source.latin)
        resolved = resolve_font(source_name)
        if resolved is not None:
            ooxml.write_font_slots(rPr, resolved.slots())
            report.fonts_resolved += 1
        elif source_name:
            # 别名表外的字体（符号字体、客户自有字体）：保留拉丁槽，只补东亚槽，
            # 中文才不会落到不可用字体上。
            ooxml.write_font_slots(
                rPr, ooxml.FontSlots(latin=source_name, ea=family.ea, cs=source_name)
            )
            report.fonts_defaulted += 1
        else:
            ooxml.write_font_slots(rPr, family.slots())
            report.fonts_defaulted += 1
        if style.font_size_pt is not None and ooxml.run_font_size_pt(run) is None:
            ooxml.write_run_font_size(rPr, style.font_size_pt)

    def _squeeze_punctuation(self, paragraph: CT_TextParagraph, size_pt: float) -> int:
        """连续全角标点按半宽挤压，返回被挤压的字符数。"""
        paragraph_text = ooxml.paragraph_text_with_runs(paragraph)
        if not paragraph_text.spans:
            return 0
        squeeze = squeeze_flags(paragraph_text.text)
        squeezed = 0
        for span in paragraph_text.spans:
            flags = squeeze[span.start : span.end]
            if not any(flags):
                continue
            run = span.run
            segments = _split_segments(ooxml.run_text(run), flags)
            run_size_pt = ooxml.run_font_size_pt(run) or size_pt
            new_runs = ooxml.split_run(run, [segment.text for segment in segments])
            for segment, new_run in zip(segments, new_runs):
                if segment.squeeze:
                    ooxml.write_character_spacing(
                        ooxml.run_properties(new_run), -run_size_pt / 2
                    )
                    squeezed += len(segment.text)
        return squeezed


def squeeze_flags(text: str) -> list[bool]:
    """逐字符标记是否需要挤压：全角标点且左右至少一侧也是全角标点。"""
    flags = [False] * len(text)
    for index, char in enumerate(text):
        if char not in FULL_WIDTH_PUNCT:
            continue
        left = index > 0 and text[index - 1] in FULL_WIDTH_PUNCT
        right = index + 1 < len(text) and text[index + 1] in FULL_WIDTH_PUNCT
        flags[index] = left or right
    return flags


@dataclass(frozen=True)
class _Segment:
    """run 文本里一段连续、挤压状态相同的字符。"""

    text: str
    squeeze: bool


def _split_segments(text: str, flags: Sequence[bool]) -> list[_Segment]:
    """按 ``flags`` 把文本切成连续段；``flags`` 与 ``text`` 一一对应。"""
    segments: list[_Segment] = []
    buffer: list[str] = []
    current: bool | None = None
    for char, flag in zip(text, flags):
        if current is not None and flag != current:
            segments.append(_Segment("".join(buffer), current))
            buffer = []
        buffer.append(char)
        current = flag
    if current is not None:
        segments.append(_Segment("".join(buffer), current))
    return segments
