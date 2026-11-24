"""版本面锁步：pyproject.toml 是唯一真源，录制产物与站点 meta 必须跟它一致。"""

from __future__ import annotations

import json
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def package_version() -> str:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text("utf-8"))
    return data["project"]["version"]


def test_demo_results_version_matches_pyproject() -> None:
    demo = json.loads((ROOT / "docs" / "demo-results.json").read_text("utf-8"))

    assert demo["version"] == package_version()


def test_site_meta_versions_match_pyproject() -> None:
    site = json.loads((ROOT / "web" / "site.json").read_text("utf-8"))
    expected = f"v{package_version()}"

    assert site["meta"]["content_version"] == expected
    assert site["meta"]["source_version"] == expected
