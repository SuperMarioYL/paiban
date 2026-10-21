"""Paiban — 中文版式引擎：长文档 → WPS 原生可编辑 .pptx。

模块边界（见 mvp_plan.md §4）：
  cli.py        可执行入口（build / fix / verify 子命令）
  deck.py       长文档 → SlidePlan（DeepSeek 语义分段，无 key 时规则降级）
  typography.py 中文版式引擎（字体别名、标点挤压与避头尾、缩进行距）
  ooxml.py      DeckIR → OOXML（python-pptx，真实形状与占位符）
  templates/    内置中文模板族（汇报 / 课件）
"""

from __future__ import annotations

__all__: list[str] = []
