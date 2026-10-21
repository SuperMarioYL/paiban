"""paiban 命令行入口：build / verify / fix（mvp_plan.md §3 的三条命令）。

* ``build``  长文档 → 中文模板族 deck（默认 DeepSeek 语义分段，无 key 降级规则切分）；
* ``verify`` 打开 .pptx 出版式 lint 报告；
* ``fix``    对已有 .pptx 跑同一套中文版式 pass。

输出全是 typer.echo 的纯文本，不带颜色与富文本格式：终端录制、CI 日志和 README 里贴的
应当是同一份可复制的文本。退出码约定见各命令的 docstring。
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, NoReturn, Optional

import typer

from . import deck, ooxml, typography

app = typer.Typer(
    name="paiban",
    help="中文长文档 → WPS 原生可编辑 .pptx 的版式引擎。",
    no_args_is_help=True,
    add_completion=False,
)

# 标签对齐用：中文字符按两格宽计算。
_LABEL_WIDTH = 10
_EXIT_USER_ERROR = 2
_EXIT_LINT_WARN = 1

SourceArg = Annotated[
    Path,
    typer.Argument(
        help="长文档：.md / .markdown / .txt / .docx",
        exists=True,
        file_okay=True,
        dir_okay=False,
        readable=True,
    ),
]
DeckArg = Annotated[
    Path,
    typer.Argument(help="要复核的 .pptx", exists=True, file_okay=True, dir_okay=False, readable=True),
]
TemplateOption = Annotated[
    str,
    typer.Option(
        "--template",
        "-t",
        help=f"模板族：{' / '.join(deck.BUILTIN_TEMPLATES)}",
    ),
]
VerifyTemplateOption = Annotated[
    Optional[str],
    typer.Option(
        "--template",
        "-t",
        help="按该模板族的版式参数判期望值；缺省用内置中文版式规范",
    ),
]
OutOption = Annotated[
    Optional[Path],
    typer.Option("--out", "-o", help="输出 .pptx 路径；缺省与输入同名换后缀"),
]


@app.command("build")
def build_command(
    source: SourceArg,
    out: OutOption = None,
    template: TemplateOption = deck.DEFAULT_TEMPLATE_NAME,
    model: Annotated[
        str,
        typer.Option("--model", help=f"DeepSeek 模型名（默认 {deck.DEFAULT_MODEL}）"),
    ] = deck.DEFAULT_MODEL,
) -> None:
    """长文档 → 中文模板族 deck。

    没有 ``DEEPSEEK_API_KEY`` 时不报错，改用文档结构规则切分，只损失语义分段质量；
    页序来源会打印出来，规则切分不会被说成模型分段。生成物总是写出，退出码只反映
    输入 / 模板是否可用（0 正常，2 用户错误）；版式是否达标由 ``verify`` 判定。
    """
    result = _build(source, out=out, template=template, model=model)
    _print_build(result, source=source)


@app.command("verify")
def verify_command(
    path: DeckArg,
    template: VerifyTemplateOption = None,
) -> None:
    """复核 .pptx 里的中文版式属性，出 lint 报告。

    只复核 OOXML 属性，不渲染、不驱动 WPS 进程。所有 check 都是 ``ok`` 时退出码 0；
    出现 ``warn`` 时退出码 1，方便接进 CI 或交付前的闸门。
    """
    try:
        report = deck.verify(path, template=template)
    except ValueError as error:
        _fail(error)
    _print_verify(report)
    if not report.ok:
        raise typer.Exit(code=_EXIT_LINT_WARN)


@app.command("fix")
def fix_command(
    path: DeckArg,
    out: OutOption = None,
    template: TemplateOption = deck.DEFAULT_TEMPLATE_NAME,
) -> None:
    """对已有 .pptx 跑一遍中文版式 pass，就地改写属性（-o 可另存）。

    适合拿别人的模板或别的工具出的稿子做二次交付：字体别名、标点挤压、避头尾、
    缩进行距全部按模板族的版式参数重写。文本内容一个字符都不改。
    """
    deck_template = _load_template(template)
    pptx = _open_deck(path)
    layout = typography.CJKLayoutPass(
        body_style=deck_template.body_style,
        title_style=deck_template.title_style,
    ).apply(pptx)
    target = Path(out) if out is not None else path
    pptx.save(target)
    report = deck.lint(
        pptx,
        path=target,
        body_style=deck_template.body_style,
        title_style=deck_template.title_style,
    )
    typer.echo(_row("文件", str(path)))
    typer.echo(_row("模板族", deck_template.spec.name))
    _print_layout(layout)
    _print_lint(report, show_all=False)
    typer.echo(_row("输出", f"{target}（就地改写）" if out is None else str(target)))


# --------------------------------------------------------------------------- #
# 步骤包装：把用户错误转成退出码 2，而不是 traceback
# --------------------------------------------------------------------------- #


def _build(
    source: Path,
    *,
    out: Optional[Path],
    template: str,
    model: str,
) -> deck.BuildResult:
    try:
        return deck.build_deck(source, template=template, out=out, model=model)
    except ValueError as error:
        _fail(error)


def _load_template(name: str) -> deck.DeckTemplate:
    try:
        return deck.load_template(name)
    except ValueError as error:
        _fail(error)


def _open_deck(path: Path) -> ooxml.Pptx:
    try:
        return ooxml.Pptx.open(path)
    except (ValueError, KeyError) as error:
        _fail(f"打不开 {path}：{error}")


def _fail(error: Exception) -> NoReturn:
    typer.echo(f"错误：{error}", err=True)
    raise typer.Exit(code=_EXIT_USER_ERROR)


# --------------------------------------------------------------------------- #
# 输出
# --------------------------------------------------------------------------- #


def _print_build(result: deck.BuildResult, *, source: Path) -> None:
    plan = result.plan
    source_label = plan.source + (f"（{plan.model}）" if plan.model else "")
    if plan.fallback_reason:
        source_label += f"（{plan.fallback_reason}）"
    typer.echo(_row("输入", str(source)))
    typer.echo(_row("模板族", result.template))
    typer.echo(_row("页序来源", source_label))
    typer.echo(
        _row(
            "页数",
            f"{result.slide_count}（封面 1 + 内容 {len(plan.slides)}）",
        )
    )
    _print_layout(result.layout)
    _print_lint(result.lint, show_all=False)
    typer.echo(_row("输出", str(result.path)))


def _print_verify(report: deck.LintReport) -> None:
    typer.echo(_row("文件", str(report.path) if report.path is not None else "-"))
    typer.echo(_row("复核范围", deck.VERIFY_SCOPE))
    typer.echo(
        _row(
            "规模",
            f"{report.slides} 页 / {report.text_bodies} 个文本容器 / "
            f"{report.paragraphs} 段 / {report.runs} 个 run",
        )
    )
    _print_lint(report, show_all=True)
    typer.echo(_row("WPS 目标版本", deck.WPS_TARGET_VERSION_RANGE))
    typer.echo(_row("未自动验证", "以下结论只能在 WPS 里人工确认："))
    for index, item in enumerate(deck.WPS_UNVERIFIED, start=1):
        typer.echo(_row("", f"{index}. {item}"))


def _print_lint(report: deck.LintReport, *, show_all: bool) -> None:
    warnings = [finding for finding in report.findings if finding.level != deck.LINT_OK]
    summary = f"{len(report.findings)} 项检查" + (f"，{len(warnings)} 项告警" if warnings else "全过")
    typer.echo(_row("lint", f"{deck.LINT_OK if report.ok else deck.LINT_WARN} — {summary}"))
    for finding in report.findings if show_all else warnings:
        typer.echo(_row("", f"{finding.check} [{finding.level}] {finding.detail}"))


def _print_layout(layout: typography.LayoutReport) -> None:
    typer.echo(
        _row(
            "改写",
            f"段落 {layout.paragraphs} / run {layout.runs} / 标点挤压 {layout.punct_squeezed} 字",
        )
    )
    typer.echo(
        _row(
            "字体",
            f"别名表命中 {layout.fonts_resolved} / 写入默认族 {layout.fonts_defaulted}",
        )
    )


def _row(label: str, value: str) -> str:
    return f"{label}{' ' * max(2, _LABEL_WIDTH - _display_width(label))}{value}"


def _display_width(text: str) -> int:
    return sum(2 if ord(char) > 0x2E7F else 1 for char in text)
