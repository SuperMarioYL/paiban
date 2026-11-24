# Changelog

All notable changes to Paiban are recorded here. The version in `pyproject.toml` is the
single source of truth; `docs/demo-results.json` and `web/site.json` are asserted against it
by `tests/test_version_lockstep.py`.

## 0.2.0 — 2026-11-25

Fixes found by reading the shipped v0.1.0 source, plus the one roadmap item that could be
delivered and verified in a single bump.

### Fixed

- **The six-bullets-per-slide cap now applies to the model path too.** `_parse_slides` put
  every bullet the model returned on one slide; only the rules path went through `_chunk`.
  Both sources now share `_paginate`, so a page with more than six bullets becomes a page
  plus a 「（续）」 page. This is what README.md and the site already promised.
- **A numbered list in a `.txt` is no longer parsed as section headings.** `_PLAIN_SECTION`
  accepted the ASCII list marker `1.`, and heading detection ran before bullet detection, so
  every numbered item became a level-2 heading (one near-empty section page per item). ASCII
  `1.` / `1)` now belongs to `_BULLET`; `一、` and `1、` are still chapter headings.
- **`verify`'s `punct_compress` check now checks the compression it audits.** It used to pass
  whenever any run in the paragraph had a negative `spc`, even if the consecutive full-width
  punctuation itself carried none. It now requires every run that holds squeezable punctuation
  to carry negative spacing.

### Added

- **`paiban verify --json`** — the same lint report as machine-readable JSON on stdout
  (`path` / `scope` / counts / `ok` / `findings` / `wps_unverified`). Exit codes are unchanged
  (0 when every check is `ok`, 1 on any `warn`), so `paiban verify deck.pptx --json | jq -e .ok`
  works as a delivery gate.

### Changed

- `docs/demo-results.json`, `web/site.json` (`meta.content_version`, `meta.source_version`,
  `generated_at`) and the READMEs are kept in lockstep with `pyproject.toml`; a regression test
  fails if they drift.

## 0.1.0 — 2026-10-21

First release.

- Chinese typography engine written to OOXML attributes: font alias fallback across the
  `latin` / `ea` / `cs` slots, full-width punctuation compression via negative `spc`, kinsoku
  via `eaLnBrk` / `hangingPunct`, two-character first-line indent and 1.5 line spacing.
- Long document → deck (`build`), layout lint (`verify`), and the same pass over an existing
  deck (`fix`). DeepSeek does semantic segmentation only; layout is always code.
- Built-in `report` and `courseware` template families, a bilingual README, an Agent Skill,
  and a golden attribute regression over `fixtures/sample-longdoc.md`.
