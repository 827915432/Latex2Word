#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tex_scan_common.py

TeX 文本扫描的轻量公共函数。
这些函数均为纯函数，便于在 precheck/normalize/postcheck 复用。
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional


def strip_latex_comments(text: str) -> str:
    """
    逐行去除 LaTeX 注释，保留被转义的 `%`。
    """
    stripped_lines: list[str] = []
    for line in text.splitlines():
        cut_index = None
        escaped = False
        for idx, ch in enumerate(line):
            if ch == "\\":
                escaped = not escaped
                continue
            if ch == "%" and not escaped:
                cut_index = idx
                break
            escaped = False
        if cut_index is None:
            stripped_lines.append(line)
        else:
            stripped_lines.append(line[:cut_index])
    return "\n".join(stripped_lines)


def line_number_from_offset(text: str, offset: int) -> int:
    """
    根据字符偏移估算行号（从 1 开始）。
    """
    return text.count("\n", 0, max(0, offset)) + 1


def skip_whitespace(text: str, pos: int) -> int:
    """
    跳过空白字符，返回新的位置索引。
    """
    while pos < len(text) and text[pos].isspace():
        pos += 1
    return pos


def parse_balanced_group(text: str, start: int, open_char: str, close_char: str) -> Optional[int]:
    """
    解析配对分组，返回分组结束后一位索引；解析失败返回 None。
    """
    if start >= len(text) or text[start] != open_char:
        return None

    depth = 0
    idx = start
    while idx < len(text):
        ch = text[idx]
        if ch == "\\":
            idx += 2
            continue
        if ch == open_char:
            depth += 1
        elif ch == close_char:
            depth -= 1
            if depth == 0:
                return idx + 1
        idx += 1
    return None


def resolve_path_with_extensions(base_dir: Path, target: str, extensions: Iterable[str]) -> Optional[Path]:
    """
    解析路径并按扩展名候选补全：
    - 若 target 自带后缀，直接校验；
    - 否则尝试逐个后缀。
    """
    target = (target or "").strip()
    if not target:
        return None

    candidate = Path(target)
    if candidate.suffix:
        resolved = (base_dir / candidate).resolve()
        return resolved if resolved.exists() else None

    for ext in extensions:
        resolved = (base_dir / f"{target}{ext}").resolve()
        if resolved.exists():
            return resolved
    resolved = (base_dir / candidate).resolve()
    return resolved if resolved.exists() else None
