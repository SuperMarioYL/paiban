"""verify 的 punct_compress 检查：只认真正承载连续全角标点的 run 有没有负字距。"""

from __future__ import annotations

from pptx.oxml.ns import qn

from paiban import deck, ooxml, typography

_A_RPR = qn("a:rPr")


def squeezed_deck() -> ooxml.Pptx:
    """一段含连续全角标点的正文页，已跑过版式 pass。"""
    source = ooxml.emit(
        ooxml.DeckIR(
            title="封面",
            slides=(
                ooxml.Slide(title="一页", bullets=(ooxml.Bullet("他说「好」呀。（注）"),)),
            ),
        )
    )
    typography.CJKLayoutPass(
        body_style=typography.ParagraphStyle(font_size_pt=20.0)
    ).apply(source)
    return source


def body_paragraph(pptx: ooxml.Pptx):
    for body in pptx.iter_text_bodies():
        if body.role == ooxml.TEXT_ROLE_BODY:
            return body.paragraphs[0]
    raise AssertionError("deck 里没有正文段落")


def punct_finding(report: deck.LintReport) -> deck.LintFinding:
    return next(finding for finding in report.findings if finding.check == "punct_compress")


def test_applied_pass_still_passes_punct_check() -> None:
    """正常跑过 pass 的段落仍是 punct_compress ok。"""
    report = deck.lint(squeezed_deck())

    assert punct_finding(report).level == deck.LINT_OK
    assert report.ok


def test_unrelated_negative_spacing_does_not_pass_punct_check() -> None:
    """标点没被挤压、只是段落里别处有负字距时，不能算已应用。"""
    pptx = squeezed_deck()
    paragraph = body_paragraph(pptx)
    runs = ooxml.paragraph_runs(paragraph)
    for run in runs:
        rPr = run.find(_A_RPR)
        if rPr is not None and rPr.get("spc") is not None:
            del rPr.attrib["spc"]
    ooxml.write_character_spacing(ooxml.run_properties(runs[0]), -1.0)

    report = deck.lint(pptx)

    assert punct_finding(report).level == deck.LINT_WARN
    assert not report.ok
