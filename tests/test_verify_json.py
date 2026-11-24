"""paiban verify --json：stdout 只有可解析的 JSON，退出码语义不变。"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest
from typer.testing import CliRunner

from paiban import cli, deck, ooxml, typography

_ANSI = re.compile(r"\x1b\[[0-9;?]*[ -/]*[@-~]")


def plain(result) -> str:
    return _ANSI.sub("", result.output or "")


@pytest.fixture()
def clean_deck(tmp_path: Path) -> Path:
    path = tmp_path / "clean.pptx"
    pptx = ooxml.emit(
        ooxml.DeckIR(
            title="封面",
            slides=(ooxml.Slide(title="一页", bullets=(ooxml.Bullet("中文要点，。"),)),),
        )
    )
    typography.CJKLayoutPass(
        body_style=typography.ParagraphStyle(font_size_pt=20.0)
    ).apply(pptx)
    pptx.save(path)
    return path


@pytest.fixture()
def raw_deck_path(tmp_path: Path) -> Path:
    path = tmp_path / "raw.pptx"
    ooxml.emit(
        ooxml.DeckIR(
            title="未加工的 deck",
            slides=(ooxml.Slide(title="一页", bullets=(ooxml.Bullet("中文要点，。"),)),),
        )
    ).save(path)
    return path


def test_verify_json_prints_only_parseable_json(clean_deck: Path) -> None:
    result = CliRunner().invoke(cli.app, ["verify", str(clean_deck), "--json"])

    assert result.exit_code == 0, plain(result)
    payload = json.loads(plain(result))
    assert payload["ok"] is True
    assert payload["findings"]
    assert {"check", "level", "detail"} <= set(payload["findings"][0])
    assert payload["wps_unverified"]
    assert "lint " not in plain(result)  # 人类可读的表格行不该混进 stdout


def test_verify_json_keeps_exit_code_semantics(clean_deck: Path, raw_deck_path: Path) -> None:
    runner = CliRunner()

    clean = runner.invoke(cli.app, ["verify", str(clean_deck), "--json"])
    assert clean.exit_code == 0
    assert json.loads(plain(clean))["ok"] is True

    raw = runner.invoke(cli.app, ["verify", str(raw_deck_path), "--json"])
    assert raw.exit_code == cli._EXIT_LINT_WARN
    payload = json.loads(plain(raw))
    assert payload["ok"] is False
    assert any(finding["level"] == deck.LINT_WARN for finding in payload["findings"])
