#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_postprocess.py

在 Pandoc 产物 docx 上执行轻量结构后处理，目标是提高可维护性：
1. 为图/表题注补充 Word SEQ 字段编号（图 N / 表 N）；
2. 为带标签的显示公式补充 SEQ 字段编号（(N)）；
3. 按 TeX 标签顺序补齐缺失书签，尽量修复内部超链接跳转。
4. 将文内图/表/公式引用从内部超链接升级为 Word REF 字段（可随编号更新）。

该模块是“best effort”设计：
- 不修改源 LaTeX；
- 不尝试重排内容结构；
- 若出现局部无法匹配，仅记录告警并继续处理其他对象。
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"

NS = {"w": W_NS, "m": M_NS}


def wqn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def mqn(local: str) -> str:
    return f"{{{M_NS}}}{local}"


def xmlqn(local: str) -> str:
    return f"{{{XML_NS}}}{local}"


FIGURE_ENVS = {"figure", "figure*"}
TABLE_ENVS = {"table", "table*", "longtable"}
EQUATION_ENVS = {
    "equation",
    "equation*",
    "align",
    "align*",
    "alignat",
    "alignat*",
    "flalign",
    "flalign*",
    "gather",
    "gather*",
    "multline",
    "multline*",
    "eqnarray",
    "eqnarray*",
    "split",
    "cases",
}

# 仅用于“显示公式段落槽位”对齐的环境集合。
# 注意：split/cases 常作为 equation/align 的内部子环境，不对应独立显示段落。
DISPLAY_EQUATION_ENVS = {
    "equation",
    "equation*",
    "align",
    "align*",
    "alignat",
    "alignat*",
    "flalign",
    "flalign*",
    "gather",
    "gather*",
    "multline",
    "multline*",
    "eqnarray",
    "eqnarray*",
}

TOKEN_PATTERN = re.compile(r"\\begin\{([^}]+)\}|\\end\{([^}]+)\}|\\label\{([^}]+)\}")
LABEL_CMD_PATTERN = re.compile(r"\\label\s*\{([^}]+)\}")
ALGORITHM_TITLE_PATTERN = re.compile(r"^\s*(?:algorithm|算法)\s*:", re.IGNORECASE)
ALGORITHM_IO_PATTERN = re.compile(r"^\s*(?:input|output|输入|输出)\s*:", re.IGNORECASE)

EQUATION_SEQ_NAME = "MTEqn"
EQUATION_DISPLAY_SEQ_SWITCHES = r"\c \* Arabic \* MERGEFORMAT"
EQUATION_HINT_SEQ_SWITCHES = r"\h \* MERGEFORMAT"
LEGACY_EQUATION_SEQ_NAMES = ("Eq", "EQ")
MATHTYPE_EQUATION_BOOKMARK_PREFIX = "ZEqnNum"


@dataclass
class LabelInventory:
    figure_labels: list[str] = field(default_factory=list)
    figure_all_labels: list[str] = field(default_factory=list)
    figure_label_alias_to_primary: dict[str, str] = field(default_factory=dict)
    table_labels: list[str] = field(default_factory=list)
    equation_labels: list[str] = field(default_factory=list)


@dataclass
class DocxPostprocessResult:
    modified: bool
    warnings: list[str] = field(default_factory=list)
    metrics: dict[str, int] = field(default_factory=dict)
    details: dict[str, object] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "modified": self.modified,
            "warnings": list(self.warnings),
            "metrics": dict(self.metrics),
            "details": dict(self.details),
        }


def _strip_comments(text: str) -> str:
    """
    逐行移除 TeX 注释（保留 \%）。
    """
    cleaned_lines: list[str] = []
    for raw_line in text.splitlines():
        line = raw_line
        cut = None
        idx = 0
        while idx < len(line):
            if line[idx] == "%":
                backslashes = 0
                j = idx - 1
                while j >= 0 and line[j] == "\\":
                    backslashes += 1
                    j -= 1
                if backslashes % 2 == 0:
                    cut = idx
                    break
            idx += 1
        if cut is not None:
            line = line[:cut]
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def _dedup_keep_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        key = value.strip()
        if not key:
            continue
        if key in seen:
            continue
        seen.add(key)
        result.append(key)
    return result


def _skip_whitespace(text: str, pos: int) -> int:
    while pos < len(text) and text[pos].isspace():
        pos += 1
    return pos


def _parse_balanced_group(text: str, start: int, open_char: str, close_char: str) -> int | None:
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


def _find_command_spans(text: str, command_name: str) -> list[tuple[int, int]]:
    pattern = re.compile(rf"\\{re.escape(command_name)}(?![A-Za-z@*])")
    spans: list[tuple[int, int]] = []
    cursor = 0
    while True:
        match = pattern.search(text, cursor)
        if not match:
            break

        pos = _skip_whitespace(text, match.end())
        if pos < len(text) and text[pos] == "[":
            opt_end = _parse_balanced_group(text, pos, "[", "]")
            if opt_end is None:
                cursor = match.end()
                continue
            pos = _skip_whitespace(text, opt_end)

        if pos >= len(text) or text[pos] != "{":
            cursor = match.end()
            continue

        body_end = _parse_balanced_group(text, pos, "{", "}")
        if body_end is None:
            cursor = match.end()
            continue

        spans.append((match.start(), body_end))
        cursor = body_end

    return spans


def _extract_environment_bodies(text: str, target_envs: set[str]) -> list[str]:
    bodies: list[str] = []
    stack: list[tuple[str, int]] = []

    for match in TOKEN_PATTERN.finditer(text):
        begin_env = match.group(1)
        end_env = match.group(2)
        if begin_env:
            stack.append((begin_env.strip().lower(), match.end()))
            continue

        if not end_env:
            continue
        env = end_env.strip().lower()
        if not stack:
            continue

        for pos in range(len(stack) - 1, -1, -1):
            stack_env, body_start = stack[pos]
            if stack_env != env:
                continue
            if env in target_envs and body_start <= match.start():
                bodies.append(text[body_start:match.start()])
            stack = stack[:pos]
            break

    return bodies


def _bind_alias(alias_to_primary: dict[str, str], *, alias: str, primary: str) -> None:
    alias_key = alias.strip()
    primary_key = primary.strip()
    if not alias_key or not primary_key:
        return
    if alias_key not in alias_to_primary:
        alias_to_primary[alias_key] = primary_key


def _resolve_figure_labels_in_body(body_text: str) -> tuple[list[str], dict[str, str], list[str]]:
    label_entries: list[tuple[int, str]] = []
    for match in LABEL_CMD_PATTERN.finditer(body_text):
        label = (match.group(1) or "").strip()
        if label:
            label_entries.append((match.start(), label))

    if not label_entries:
        return [], {}, []

    all_labels = _dedup_keep_order(label for _, label in label_entries)
    caption_spans = _find_command_spans(body_text, "caption")
    if not caption_spans:
        primary_labels = list(all_labels)
        alias_map = {label: label for label in primary_labels}
        return primary_labels, alias_map, all_labels

    primary_labels: list[str] = []
    alias_to_primary: dict[str, str] = {}
    consumed_label_indexes: set[int] = set()
    caption_primary_labels: list[str | None] = [None for _ in caption_spans]

    for cap_idx, (caption_start, caption_end) in enumerate(caption_spans):
        prev_caption_end = caption_spans[cap_idx - 1][1] if cap_idx > 0 else 0
        next_caption_start = caption_spans[cap_idx + 1][0] if cap_idx + 1 < len(caption_spans) else len(body_text)

        after_region = [
            idx
            for idx, (pos, _label) in enumerate(label_entries)
            if idx not in consumed_label_indexes and caption_end <= pos < next_caption_start
        ]
        before_region = [
            idx
            for idx, (pos, _label) in enumerate(label_entries)
            if idx not in consumed_label_indexes and prev_caption_end <= pos < caption_start
        ]

        region = after_region if after_region else before_region
        if not region:
            continue

        primary_idx = region[0] if after_region else region[-1]
        primary_label = label_entries[primary_idx][1]
        caption_primary_labels[cap_idx] = primary_label
        if primary_label not in primary_labels:
            primary_labels.append(primary_label)

        for idx in region:
            _bind_alias(alias_to_primary, alias=label_entries[idx][1], primary=primary_label)
            consumed_label_indexes.add(idx)

    if not any(caption_primary_labels):
        fallback_label = label_entries[-1][1]
        if fallback_label not in primary_labels:
            primary_labels.append(fallback_label)
        caption_primary_labels[-1] = fallback_label
        _bind_alias(alias_to_primary, alias=fallback_label, primary=fallback_label)

    for idx, (pos, label) in enumerate(label_entries):
        if idx in consumed_label_indexes:
            continue

        best_primary = None
        best_distance = None
        for cap_idx, (caption_start, _caption_end) in enumerate(caption_spans):
            primary = caption_primary_labels[cap_idx]
            if not primary:
                continue
            distance = abs(pos - caption_start)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_primary = primary

        if best_primary is None:
            if label not in primary_labels:
                primary_labels.append(label)
            best_primary = label

        _bind_alias(alias_to_primary, alias=label, primary=best_primary)

    caption_primaries = _dedup_keep_order(label for label in caption_primary_labels if label)
    if len(caption_spans) > 1 and caption_primaries:
        canonical_primary = caption_primaries[-1]
        collapsed_alias: dict[str, str] = {}
        for _, label in label_entries:
            _bind_alias(collapsed_alias, alias=label, primary=canonical_primary)
        for alias in alias_to_primary.keys():
            _bind_alias(collapsed_alias, alias=alias, primary=canonical_primary)
        alias_to_primary = collapsed_alias
        primary_labels = [canonical_primary]

    return _dedup_keep_order(primary_labels), alias_to_primary, all_labels


def extract_label_inventory(tex_files: list[Path]) -> LabelInventory:
    figure_primary_labels: list[str] = []
    figure_all_labels: list[str] = []
    figure_label_alias_to_primary: dict[str, str] = {}
    figure_fallback_labels: list[str] = []
    table_labels: list[str] = []
    equation_labels: list[str] = []

    for tex_file in tex_files:
        try:
            raw = tex_file.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            try:
                raw = tex_file.read_text(encoding="gbk")
            except Exception:
                continue
        except Exception:
            continue

        text = _strip_comments(raw)

        for figure_body in _extract_environment_bodies(text, FIGURE_ENVS):
            primary_labels, alias_map, all_labels = _resolve_figure_labels_in_body(figure_body)
            figure_primary_labels.extend(primary_labels)
            figure_all_labels.extend(all_labels)
            for alias, primary in alias_map.items():
                _bind_alias(
                    figure_label_alias_to_primary,
                    alias=alias,
                    primary=primary,
                )

        stack: list[str] = []

        for match in TOKEN_PATTERN.finditer(text):
            begin_env = match.group(1)
            end_env = match.group(2)
            label = match.group(3)

            if begin_env:
                stack.append(begin_env.strip().lower())
                continue

            if end_env:
                env = end_env.strip().lower()
                if not stack:
                    continue
                for pos in range(len(stack) - 1, -1, -1):
                    if stack[pos] == env:
                        stack = stack[:pos]
                        break
                continue

            if not label:
                continue
            label_key = label.strip()
            if not label_key:
                continue

            env_set = set(stack)
            if env_set & FIGURE_ENVS:
                figure_fallback_labels.append(label_key)
            elif env_set & TABLE_ENVS:
                table_labels.append(label_key)
            elif env_set & EQUATION_ENVS:
                equation_labels.append(label_key)

    figure_primary_labels = _dedup_keep_order(figure_primary_labels)
    figure_all_labels = _dedup_keep_order(figure_all_labels + figure_fallback_labels)

    if not figure_primary_labels:
        figure_primary_labels = _dedup_keep_order(figure_fallback_labels)

    normalized_alias: dict[str, str] = {}
    for alias, primary in figure_label_alias_to_primary.items():
        alias_key = alias.strip()
        primary_key = primary.strip()
        if not alias_key or not primary_key:
            continue
        if alias_key in normalized_alias:
            continue
        normalized_alias[alias_key] = primary_key

    for label in figure_primary_labels:
        if label not in normalized_alias:
            normalized_alias[label] = label

    return LabelInventory(
        figure_labels=figure_primary_labels,
        figure_all_labels=figure_all_labels,
        figure_label_alias_to_primary=normalized_alias,
        table_labels=_dedup_keep_order(table_labels),
        equation_labels=_dedup_keep_order(equation_labels),
    )


def _extract_labels_from_body_text(body_text: str) -> list[str]:
    labels: list[str] = []
    for match in LABEL_CMD_PATTERN.finditer(body_text):
        label = (match.group(1) or "").strip()
        if label:
            labels.append(label)
    return _dedup_keep_order(labels)


def extract_float_slots(tex_files: list[Path]) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    """
    提取图/表槽位（按 TeX 环境出现顺序）。

    返回：
    - figure_slots: [{"primary_labels": [...], "all_labels": [...], "alias_to_primary": {...}}, ...]
    - table_slots: [{"labels": [...]}, ...]
    """
    figure_slots: list[dict[str, object]] = []
    table_slots: list[dict[str, object]] = []

    for tex_file in tex_files:
        try:
            raw = tex_file.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            try:
                raw = tex_file.read_text(encoding="gbk")
            except Exception:
                continue
        except Exception:
            continue

        text = _strip_comments(raw)

        for figure_body in _extract_environment_bodies(text, FIGURE_ENVS):
            primary_labels, alias_map, all_labels = _resolve_figure_labels_in_body(figure_body)
            all_labels = _dedup_keep_order(all_labels or _extract_labels_from_body_text(figure_body))
            primary_labels = _dedup_keep_order(primary_labels)

            # 若未从 caption 语义中确定 primary，则退化为首个标签。
            if not primary_labels and all_labels:
                primary_labels = [all_labels[0]]

            normalized_alias: dict[str, str] = {}
            for alias, primary in alias_map.items():
                alias_key = alias.strip()
                primary_key = primary.strip()
                if not alias_key or not primary_key:
                    continue
                if alias_key in normalized_alias:
                    continue
                normalized_alias[alias_key] = primary_key
            for primary_label in primary_labels:
                if primary_label not in normalized_alias:
                    normalized_alias[primary_label] = primary_label

            figure_slots.append(
                {
                    "primary_labels": primary_labels,
                    "all_labels": all_labels,
                    "alias_to_primary": normalized_alias,
                }
            )

        for table_body in _extract_environment_bodies(text, TABLE_ENVS):
            table_slots.append({"labels": _extract_labels_from_body_text(table_body)})

    return figure_slots, table_slots


def extract_equation_display_slots(tex_files: list[Path]) -> list[dict[str, object]]:
    """
    提取“显示级公式槽位”列表，并保留每个槽位中的标签集合（可为空）。

    设计目的：
    - 槽位数应尽量对应 Pandoc 产物中的显示公式段落数；
    - 槽位内标签用于把编号绑定到正确公式，而不是简单按“标签列表顺序”压到前 N 个公式。
    """
    slots: list[dict[str, object]] = []

    for tex_file in tex_files:
        try:
            raw = tex_file.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            try:
                raw = tex_file.read_text(encoding="gbk")
            except Exception:
                continue
        except Exception:
            continue

        text = _strip_comments(raw)
        stack: list[dict[str, object]] = []

        for match in TOKEN_PATTERN.finditer(text):
            begin_env = match.group(1)
            end_env = match.group(2)
            label = match.group(3)

            if begin_env:
                stack.append(
                    {
                        "env": begin_env.strip().lower(),
                        "labels": [],
                    }
                )
                continue

            if end_env:
                env = end_env.strip().lower()
                if not stack:
                    continue
                for pos in range(len(stack) - 1, -1, -1):
                    item = stack[pos]
                    if item.get("env") != env:
                        continue
                    stack = stack[:pos]
                    if env in DISPLAY_EQUATION_ENVS:
                        labels = item.get("labels", [])
                        if not isinstance(labels, list):
                            labels = []
                        slots.append(
                            {
                                "env": env,
                                "labels": _dedup_keep_order(str(v) for v in labels),
                            }
                        )
                    break
                continue

            if not label:
                continue
            label_key = label.strip()
            if not label_key:
                continue

            for pos in range(len(stack) - 1, -1, -1):
                env = str(stack[pos].get("env", ""))
                if env not in DISPLAY_EQUATION_ENVS:
                    continue
                labels = stack[pos].get("labels")
                if not isinstance(labels, list):
                    labels = []
                    stack[pos]["labels"] = labels
                labels.append(label_key)
                break

    return slots


def _paragraph_style_id(paragraph: ET.Element) -> str:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return ""
    style = ppr.find("./w:pStyle", NS)
    if style is None:
        return ""
    return style.get(wqn("val"), "").strip()


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join((node.text or "") for node in paragraph.findall(".//w:t", NS))


def _paragraph_contains_drawing(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None


def _paragraph_has_numbering(paragraph: ET.Element) -> bool:
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        return False
    return ppr.find("./w:numPr", NS) is not None


def _select_algorithm_table_style_id(styles_root: ET.Element) -> str:
    """
    选择算法表格使用的样式（若存在）。
    """
    table_style_ids: set[str] = set()
    for style in styles_root.findall("./w:style", NS):
        if style.get(wqn("type"), "") != "table":
            continue
        style_id = (style.get(wqn("styleId"), "") or "").strip()
        if style_id:
            table_style_ids.add(style_id)

    for candidate in ("AlgorithmTable", "Table", "TableGrid", "a3"):
        if candidate in table_style_ids:
            return candidate

    return ""


def _find_preferred_table_style_id(styles_root: ET.Element) -> str:
    """
    在 reference.docx 的表格样式中优先选择“三线表”风格。

    匹配优先级：
    1) styleId 或样式名精确命中常见三线表关键词；
    2) 样式名包含“三线表 / threeline / booktabs”等语义词。
    """
    preferred_tokens = {
        "三线表",
        "threelinetable",
        "threeline",
        "booktabstable",
        "booktabs",
    }
    fuzzy_hits: list[str] = []

    for style in styles_root.findall("./w:style", NS):
        if style.get(wqn("type"), "") != "table":
            continue
        style_id = (style.get(wqn("styleId"), "") or "").strip()
        if not style_id:
            continue

        name_node = style.find("./w:name", NS)
        style_name = ""
        if name_node is not None:
            style_name = (name_node.get(wqn("val"), "") or "").strip()

        normalized_id = _normalize_style_name(style_id)
        normalized_name = _normalize_style_name(style_name)

        if normalized_id in preferred_tokens or normalized_name in preferred_tokens:
            return style_id

        joined = f"{normalized_id} {normalized_name}".strip()
        if any(token in joined for token in ("三线表", "threeline", "booktabs")):
            fuzzy_hits.append(style_id)

    if fuzzy_hits:
        return fuzzy_hits[0]
    return ""


def _table_style_id(table: ET.Element) -> str:
    tbl_pr = table.find("./w:tblPr", NS)
    if tbl_pr is None:
        return ""
    tbl_style = tbl_pr.find("./w:tblStyle", NS)
    if tbl_style is None:
        return ""
    return (tbl_style.get(wqn("val"), "") or "").strip()


def _set_table_style_id(
    table: ET.Element,
    style_id: str,
    *,
    replace_default_only: bool = True,
) -> bool:
    if not style_id:
        return False

    tbl_pr = table.find("./w:tblPr", NS)
    if tbl_pr is None:
        tbl_pr = ET.Element(wqn("tblPr"))
        table.insert(0, tbl_pr)

    tbl_style = tbl_pr.find("./w:tblStyle", NS)
    current_style = ""
    if tbl_style is not None:
        current_style = (tbl_style.get(wqn("val"), "") or "").strip()

    if current_style == style_id:
        return False

    if replace_default_only and current_style:
        normalized_current = _normalize_style_name(current_style)
        default_style_tokens = {"table", "tablegrid", "normaltable"}
        if normalized_current not in default_style_tokens:
            return False

    if tbl_style is None:
        tbl_style = ET.SubElement(tbl_pr, wqn("tblStyle"))
    tbl_style.set(wqn("val"), style_id)
    return True


def _table_looks_like_algorithm_wrapper(table: ET.Element) -> bool:
    paragraphs = table.findall(".//w:p", NS)
    if not paragraphs:
        return False

    title_text = _paragraph_text(paragraphs[0]).strip()
    if not ALGORITHM_TITLE_PATTERN.match(title_text):
        return False

    return any(_paragraph_has_numbering(paragraph) for paragraph in paragraphs[1:])


def _apply_preferred_table_style(
    document_root: ET.Element,
    *,
    preferred_style_id: str,
) -> tuple[int, int, int]:
    """
    将 docx 中普通表格样式切换为三线表样式。

    返回：
    - applied_count
    - candidate_count
    - skipped_algorithm_count
    """
    if not preferred_style_id:
        return 0, 0, 0

    applied_count = 0
    candidate_count = 0
    skipped_algorithm_count = 0

    for table in document_root.findall(".//w:tbl", NS):
        candidate_count += 1
        if _table_looks_like_algorithm_wrapper(table):
            skipped_algorithm_count += 1
            continue
        if _set_table_style_id(table, preferred_style_id, replace_default_only=True):
            applied_count += 1

    return applied_count, candidate_count, skipped_algorithm_count


def _build_algorithm_table(paragraphs: list[ET.Element], table_style_id: str) -> ET.Element:
    """
    将若干段落封装为 1x1 的算法表格。
    """
    tbl = ET.Element(wqn("tbl"))
    tbl_pr = ET.SubElement(tbl, wqn("tblPr"))

    if table_style_id:
        tbl_style = ET.SubElement(tbl_pr, wqn("tblStyle"))
        tbl_style.set(wqn("val"), table_style_id)

    tbl_w = ET.SubElement(tbl_pr, wqn("tblW"))
    tbl_w.set(wqn("w"), "0")
    tbl_w.set(wqn("type"), "auto")

    tbl_borders = ET.SubElement(tbl_pr, wqn("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = ET.SubElement(tbl_borders, wqn(side))
        border.set(wqn("val"), "single")
        border.set(wqn("sz"), "4")
        border.set(wqn("space"), "0")
        border.set(wqn("color"), "auto")

    tbl_cell_mar = ET.SubElement(tbl_pr, wqn("tblCellMar"))
    for side in ("top", "left", "bottom", "right"):
        margin = ET.SubElement(tbl_cell_mar, wqn(side))
        margin.set(wqn("w"), "80")
        margin.set(wqn("type"), "dxa")

    tbl_grid = ET.SubElement(tbl, wqn("tblGrid"))
    grid_col = ET.SubElement(tbl_grid, wqn("gridCol"))
    grid_col.set(wqn("w"), "9000")

    tr = ET.SubElement(tbl, wqn("tr"))
    tc = ET.SubElement(tr, wqn("tc"))
    tc_pr = ET.SubElement(tc, wqn("tcPr"))
    tc_w = ET.SubElement(tc_pr, wqn("tcW"))
    tc_w.set(wqn("w"), "9000")
    tc_w.set(wqn("type"), "dxa")
    tc_valign = ET.SubElement(tc_pr, wqn("vAlign"))
    tc_valign.set(wqn("val"), "top")

    for paragraph in paragraphs:
        tc.append(paragraph)

    if not paragraphs:
        tc.append(ET.Element(wqn("p")))

    return tbl


def _wrap_algorithm_blocks_as_tables(
    *,
    body: ET.Element,
    styles_root: ET.Element,
) -> tuple[int, int, int, int, str, list[dict[str, object]]]:
    """
    识别并将算法块封装为表格。

    识别规则（稳定）：
    - 起始段落文本以 "Algorithm:"（或“算法:”）开头；
    - 后续可包含 Input/Output 段落；
    - 必须包含至少 1 个带编号属性（numPr）的步骤段落；
    - 在遇到不属于上述集合的段落时结束块。
    """
    table_style_id = _select_algorithm_table_style_id(styles_root)
    wrapped_block_count = 0
    wrapped_paragraph_count = 0
    detected_block_count = 0
    skipped_block_count = 0
    block_samples: list[dict[str, object]] = []

    while True:
        children = list(body)
        changed = False
        index = 0

        while index < len(children):
            element = children[index]
            if element.tag != wqn("p"):
                index += 1
                continue

            title_text = _paragraph_text(element).strip()
            if not ALGORITHM_TITLE_PATTERN.match(title_text):
                index += 1
                continue

            detected_block_count += 1
            block: list[ET.Element] = [element]
            step_count = 0
            cursor = index + 1

            while cursor < len(children):
                follower = children[cursor]
                if follower.tag != wqn("p"):
                    break

                follower_text = _paragraph_text(follower).strip()
                if ALGORITHM_IO_PATTERN.match(follower_text):
                    block.append(follower)
                    cursor += 1
                    continue

                if _paragraph_has_numbering(follower):
                    block.append(follower)
                    step_count += 1
                    cursor += 1
                    continue

                break

            if step_count <= 0:
                skipped_block_count += 1
                index += 1
                continue

            for paragraph in block:
                body.remove(paragraph)

            table_element = _build_algorithm_table(block, table_style_id)
            body.insert(index, table_element)

            wrapped_block_count += 1
            wrapped_paragraph_count += len(block)
            block_samples.append(
                {
                    "start_index": index,
                    "paragraph_count": len(block),
                    "step_count": step_count,
                    "title_text": title_text[:120],
                }
            )
            changed = True
            break

        if not changed:
            break

    return (
        wrapped_block_count,
        wrapped_paragraph_count,
        detected_block_count,
        skipped_block_count,
        table_style_id,
        block_samples,
    )


def _paragraph_has_seq_field(paragraph: ET.Element, seq_keyword: str | None = None) -> bool:
    keyword = (seq_keyword or "").strip().upper()

    for fld in paragraph.findall(".//w:fldSimple", NS):
        instr = (fld.get(wqn("instr"), "") or "").upper()
        if "SEQ " not in instr:
            continue
        if not keyword or f"SEQ {keyword}" in instr:
            return True

    for node in paragraph.findall(".//w:instrText", NS):
        instr_text = (node.text or "").upper()
        if "SEQ " not in instr_text:
            continue
        if not keyword or f"SEQ {keyword}" in instr_text:
            return True
    return False


def _make_run(text: str) -> ET.Element:
    run = ET.Element(wqn("r"))
    text_node = ET.SubElement(run, wqn("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(xmlqn("space"), "preserve")
    text_node.text = text
    return run


def _make_tab_run() -> ET.Element:
    run = ET.Element(wqn("r"))
    ET.SubElement(run, wqn("tab"))
    return run


def _make_seq_field(
    seq_name: str,
    display_number: int | None = None,
    *,
    field_switches: str = r"\* ARABIC",
) -> ET.Element:
    fld = ET.Element(wqn("fldSimple"))
    instr = f"SEQ {seq_name}"
    switches = (field_switches or "").strip()
    if switches:
        instr = f"{instr} {switches}"
    fld.set(wqn("instr"), instr)
    run = ET.SubElement(fld, wqn("r"))
    text_node = ET.SubElement(run, wqn("t"))
    text_node.text = str(display_number) if display_number is not None else ""
    return fld


def _make_ref_field(bookmark_name: str, display_text: str) -> ET.Element:
    fld = ET.Element(wqn("fldSimple"))
    fld.set(wqn("instr"), f" REF {bookmark_name} \\h ")
    run = ET.SubElement(fld, wqn("r"))
    text_node = ET.SubElement(run, wqn("t"))
    text_node.text = display_text if display_text else "?"
    return fld


def _make_field_char_run(field_char_type: str) -> ET.Element:
    run = ET.Element(wqn("r"))
    fld_char = ET.SubElement(run, wqn("fldChar"))
    fld_char.set(wqn("fldCharType"), field_char_type)
    return run


def _make_instr_text_run(instr_text: str) -> ET.Element:
    run = ET.Element(wqn("r"))
    instr_node = ET.SubElement(run, wqn("instrText"))
    instr_node.set(xmlqn("space"), "preserve")
    instr_node.text = f" {instr_text.strip()} "
    return run


def _build_simple_ref_nodes(bookmark_name: str, display_text: str) -> list[ET.Element]:
    return [_make_ref_field(bookmark_name, display_text)]


def _build_mathtype_equation_ref_nodes(bookmark_name: str, display_text: str) -> list[ET.Element]:
    text = display_text if display_text else "(?)"
    return [
        _make_field_char_run("begin"),
        _make_instr_text_run(f"GOTOBUTTON {bookmark_name}  \\* MERGEFORMAT"),
        _make_field_char_run("begin"),
        _make_instr_text_run(f"REF {bookmark_name} \\* Charformat \\! \\* MERGEFORMAT"),
        _make_field_char_run("separate"),
        _make_run(text),
        _make_field_char_run("end"),
        _make_field_char_run("end"),
    ]


def _insert_nodes_after_ppr(paragraph: ET.Element, nodes: list[ET.Element]) -> None:
    index = 0
    children = list(paragraph)
    if children and children[0].tag == wqn("pPr"):
        index = 1
    for node in nodes:
        paragraph.insert(index, node)
        index += 1


def _run_text(run: ET.Element) -> str:
    if run.tag != wqn("r"):
        return ""
    return "".join((node.text or "") for node in run.findall("./w:t", NS))


def _run_text_matches_paren(run: ET.Element, paren: str) -> bool:
    text = _run_text(run)
    if not text and run.tag == wqn("r"):
        text = "".join((node.text or "") for node in run.findall("./w:instrText", NS))
    if not text:
        return False
    compact = re.sub(r"\s+", "", text)
    return compact == paren


def _set_run_text(run: ET.Element, text: str) -> bool:
    if run.tag != wqn("r"):
        return False
    text_nodes = run.findall("./w:t", NS)
    target_nodes = text_nodes if text_nodes else run.findall("./w:instrText", NS)
    if not target_nodes:
        return False
    changed = False
    for idx, node in enumerate(target_nodes):
        target = text if idx == 0 else ""
        if (node.text or "") != target:
            node.text = target
            changed = True
        if node.tag == wqn("t") and xmlqn("space") in node.attrib:
            del node.attrib[xmlqn("space")]
    return changed


def _is_tab_run(run: ET.Element) -> bool:
    return run.tag == wqn("r") and run.find("./w:tab", NS) is not None


def _find_first_math_child_index(paragraph: ET.Element) -> int:
    for idx, child in enumerate(list(paragraph)):
        if child.tag in {mqn("oMathPara"), mqn("oMath")}:
            return idx
    return -1


def _find_eq_seq_child_index(paragraph: ET.Element) -> int:
    seq_candidates = {f"SEQ {EQUATION_SEQ_NAME.upper()}"}
    seq_candidates.update(f"SEQ {name.upper()}" for name in LEGACY_EQUATION_SEQ_NAMES)
    for idx, child in enumerate(list(paragraph)):
        if child.tag != wqn("fldSimple"):
            continue
        instr = (child.get(wqn("instr"), "") or "").upper()
        if any(candidate in instr for candidate in seq_candidates):
            return idx
    return -1


def _paragraph_has_equation_seq_field(paragraph: ET.Element) -> bool:
    if _paragraph_has_seq_field(paragraph, EQUATION_SEQ_NAME):
        return True
    for legacy_name in LEGACY_EQUATION_SEQ_NAMES:
        if _paragraph_has_seq_field(paragraph, legacy_name):
            return True
    return False


def _normalize_equation_display_seq_fields(paragraph: ET.Element) -> bool:
    """
    将公式“显示编号”字段规范化为：
    `SEQ MTEqn \\c \\* Arabic \\* MERGEFORMAT`
    """
    changed = False
    seq_tokens = [EQUATION_SEQ_NAME, *LEGACY_EQUATION_SEQ_NAMES]
    seq_tokens_upper = tuple(f"SEQ {name.upper()}" for name in seq_tokens)
    target_instr = f"SEQ {EQUATION_SEQ_NAME} {EQUATION_DISPLAY_SEQ_SWITCHES}"

    for child in list(paragraph):
        if child.tag != wqn("fldSimple"):
            continue
        raw_instr = child.get(wqn("instr"), "") or ""
        instr_upper = raw_instr.upper()
        if not any(token in instr_upper for token in seq_tokens_upper):
            continue
        # 跳过 hint 域（理论上 hint 使用复杂域 instrText，不应出现在 fldSimple）
        if "\\H" in instr_upper and "\\C" not in instr_upper:
            continue
        if raw_instr.strip() == target_instr:
            continue
        child.set(wqn("instr"), target_instr)
        changed = True
    return changed


def _paragraph_has_mathtype_number_macrobutton(paragraph: ET.Element) -> bool:
    for node in paragraph.findall(".//w:instrText", NS):
        instr = (node.text or "").upper()
        if "MACROBUTTON MTPLACEREF" in instr:
            return True
    return False


def _ensure_equation_number_macrobutton_wrapper(paragraph: ET.Element) -> bool:
    """
    确保公式编号区使用 MathType 外层域：
    { MACROBUTTON MTPlaceRef \\* MERGEFORMAT (SEQ MTEqn) }
    """
    if _paragraph_has_mathtype_number_macrobutton(paragraph):
        return False

    seq_idx = _find_eq_seq_child_index(paragraph)
    if seq_idx < 0:
        return False

    children = list(paragraph)
    left_idx = seq_idx - 1 if seq_idx > 0 and _run_text_matches_paren(children[seq_idx - 1], "(") else seq_idx
    right_idx = seq_idx + 1 if seq_idx + 1 < len(children) and _run_text_matches_paren(children[seq_idx + 1], ")") else seq_idx

    prefix_nodes = [
        _make_field_char_run("begin"),
        _make_instr_text_run("MACROBUTTON MTPlaceRef \\* MERGEFORMAT"),
        _make_field_char_run("begin"),
        _make_instr_text_run(f"SEQ {EQUATION_SEQ_NAME} {EQUATION_HINT_SEQ_SWITCHES}"),
        _make_field_char_run("end"),
    ]

    for offset, node in enumerate(prefix_nodes):
        paragraph.insert(left_idx + offset, node)

    suffix_insert_idx = right_idx + len(prefix_nodes) + 1
    paragraph.insert(suffix_insert_idx, _make_field_char_run("end"))
    return True


def _ensure_equation_prefix_tab(paragraph: ET.Element) -> bool:
    """
    确保显示公式前有一个制表符（tab）。
    """
    math_idx = _find_first_math_child_index(paragraph)
    if math_idx < 0:
        return False

    children = list(paragraph)
    if math_idx > 0 and _is_tab_run(children[math_idx - 1]):
        return False

    paragraph.insert(math_idx, _make_tab_run())
    return True


def _ensure_equation_number_tab_layout(paragraph: ET.Element) -> bool:
    """
    确保编号区布局为：TAB (SEQ MTEqn)
    即：tab equ tab (num)
    """
    changed = False
    seq_idx = _find_eq_seq_child_index(paragraph)
    if seq_idx < 0:
        return False

    children = list(paragraph)

    # 1) 确保在 SEQ 前有 "(" 运行。
    paren_idx = seq_idx
    if seq_idx > 0 and _run_text_matches_paren(children[seq_idx - 1], "("):
        paren_idx = seq_idx - 1
        changed = _set_run_text(children[paren_idx], "(") or changed
    else:
        paragraph.insert(seq_idx, _make_run("("))
        changed = True
        paren_idx = seq_idx

    # 2) 确保 "(" 前有 tab。
    children = list(paragraph)
    if paren_idx == 0 or not _is_tab_run(children[paren_idx - 1]):
        paragraph.insert(paren_idx, _make_tab_run())
        changed = True

    # 3) 确保 SEQ 后有 ")"。
    seq_idx = _find_eq_seq_child_index(paragraph)
    children = list(paragraph)
    if seq_idx < 0:
        return changed
    if seq_idx + 1 < len(children) and _run_text_matches_paren(children[seq_idx + 1], ")"):
        changed = _set_run_text(children[seq_idx + 1], ")") or changed
    else:
        paragraph.insert(seq_idx + 1, _make_run(")"))
        changed = True

    return changed


def _prepend_caption_seq(
    paragraph: ET.Element,
    *,
    prefix: str,
    seq_name: str,
    sequence_index: int,
) -> bool:
    if _paragraph_has_seq_field(paragraph):
        return False
    nodes = [_make_run(prefix), _make_seq_field(seq_name, sequence_index), _make_run(" ")]
    _insert_nodes_after_ppr(paragraph, nodes)
    return True


def _append_equation_seq(
    paragraph: ET.Element,
    *,
    seq_name: str = EQUATION_SEQ_NAME,
    sequence_index: int,
) -> bool:
    changed = _normalize_equation_display_seq_fields(paragraph)
    if _paragraph_has_equation_seq_field(paragraph):
        return changed
    text = _paragraph_text(paragraph).strip()
    if re.search(r"\(\d+\)$", text):
        return changed
    paragraph.append(_make_tab_run())
    paragraph.append(_make_run("("))
    paragraph.append(
        _make_seq_field(
            seq_name,
            sequence_index,
            field_switches=EQUATION_DISPLAY_SEQ_SWITCHES,
        )
    )
    paragraph.append(_make_run(")"))
    return True


def _is_display_equation_paragraph(paragraph: ET.Element) -> bool:
    if paragraph.find(".//m:oMathPara", NS) is not None:
        return True
    if paragraph.find(".//m:oMath", NS) is None:
        return False
    if _paragraph_contains_drawing(paragraph):
        return False
    # 仅在“几乎纯数学段”场景下，把 m:oMath 视作显示公式候选。
    if _paragraph_text(paragraph).strip():
        return False
    return True


def _parse_caption_style_ids(styles_root: ET.Element) -> tuple[set[str], set[str], set[str]]:
    figure_style_ids: set[str] = {"ImageCaption"}
    table_style_ids: set[str] = {"TableCaption"}
    generic_caption_style_ids: set[str] = set()

    figure_names = {"image caption", "figure caption", "图题", "图注", "图片题注"}
    table_names = {"table caption", "表题", "表注", "表格题注"}
    generic_names = {"caption", "题注"}

    for style in styles_root.findall(".//w:style[@w:type='paragraph']", NS):
        style_id = (style.get(wqn("styleId"), "") or "").strip()
        if not style_id:
            continue
        name_node = style.find("./w:name", NS)
        name_value = ""
        if name_node is not None:
            name_value = (name_node.get(wqn("val"), "") or "").strip().lower()
        sid_lower = style_id.lower()
        if sid_lower == "imagecaption" or name_value in figure_names:
            figure_style_ids.add(style_id)
        if sid_lower == "tablecaption" or name_value in table_names:
            table_style_ids.add(style_id)
        if sid_lower in {"caption"} or name_value in generic_names:
            generic_caption_style_ids.add(style_id)

    return figure_style_ids, table_style_ids, generic_caption_style_ids


def _normalize_style_name(value: str) -> str:
    return re.sub(r"\s+", "", (value or "").strip().lower())


def _find_mt_display_equation_style_id(styles_root: ET.Element) -> str:
    """
    在 styles.xml 中查找 MathType 显示公式段落样式 ID（MTDisplayEquation）。
    匹配规则：
    - 样式名或 styleId 规范化后命中 mtdisplayequation / mathtypedisplayequation。
    """
    preferred_targets = {"mtdisplayequation", "mathtypedisplayequation"}
    preferred_id_candidates: list[str] = []
    preferred_name_candidates: list[str] = []

    for style in styles_root.findall(".//w:style[@w:type='paragraph']", NS):
        style_id = (style.get(wqn("styleId"), "") or "").strip()
        if not style_id:
            continue

        normalized_style_id = _normalize_style_name(style_id)
        if normalized_style_id in preferred_targets:
            preferred_id_candidates.append(style_id)

        name_node = style.find("./w:name", NS)
        style_name = ""
        if name_node is not None:
            style_name = (name_node.get(wqn("val"), "") or "").strip()
        normalized_name = _normalize_style_name(style_name)
        if normalized_name in preferred_targets:
            preferred_name_candidates.append(style_id)

    # 按优先级返回：preferred(name) > preferred(id)
    if preferred_name_candidates:
        return preferred_name_candidates[0]
    if preferred_id_candidates:
        return preferred_id_candidates[0]
    return ""


def _set_paragraph_style_id(paragraph: ET.Element, style_id: str) -> bool:
    if not style_id:
        return False
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(wqn("pPr"))
        paragraph.insert(0, ppr)

    style_node = ppr.find("./w:pStyle", NS)
    if style_node is None:
        style_node = ET.SubElement(ppr, wqn("pStyle"))

    current = (style_node.get(wqn("val"), "") or "").strip()
    if current == style_id:
        return False
    style_node.set(wqn("val"), style_id)
    return True


def _paragraph_has_visible_payload(paragraph: ET.Element) -> bool:
    if _paragraph_text(paragraph).strip():
        return True
    if _paragraph_contains_drawing(paragraph):
        return True
    if paragraph.find(".//m:oMath", NS) is not None:
        return True
    return False


def _infer_caption_kind(
    elements: list[ET.Element],
    para_index: int,
    paragraph: ET.Element,
) -> str:
    text = _paragraph_text(paragraph).strip().lower()
    if text.startswith(("表", "table")):
        return "table"
    if text.startswith(("图", "figure", "fig.")):
        return "figure"

    for pos in range(para_index - 1, -1, -1):
        node = elements[pos]
        if node.tag == wqn("tbl"):
            return "table"
        if node.tag == wqn("p"):
            if _paragraph_contains_drawing(node):
                return "figure"
            if _paragraph_has_visible_payload(node):
                break

    for pos in range(para_index + 1, len(elements)):
        node = elements[pos]
        if node.tag == wqn("tbl"):
            return "table"
        if node.tag == wqn("p"):
            if _paragraph_contains_drawing(node):
                return "figure"
            if _paragraph_has_visible_payload(node):
                break

    return "figure"


def _pandoc_hashed_anchor(label: str) -> str:
    digest = hashlib.sha1(label.encode("utf-8")).hexdigest()
    return "X" + digest[1:]


def _candidate_anchor_names(label: str) -> list[str]:
    label_key = label.strip()
    if not label_key:
        return []
    names = [label_key]
    hashed = _pandoc_hashed_anchor(label_key)
    if hashed != label_key:
        names.append(hashed)
    return names


def _number_bookmark_name(object_kind: str, label: str) -> str:
    if object_kind == "equation":
        return _mathtype_equation_number_bookmark_name(label)

    digest = hashlib.sha1(label.strip().encode("utf-8")).hexdigest()[:32]
    prefix = {
        "figure": "figNum",
        "table": "tabNum",
    }.get(object_kind, "")
    if not prefix:
        return ""
    return f"{prefix}_{digest}"


def _mathtype_equation_number_bookmark_name(label: str) -> str:
    """
    MathType 风格公式号书签：`ZEqnNum##########`（确定性生成）。
    """
    digest = hashlib.sha1(label.strip().encode("utf-8")).hexdigest()
    numeric = int(digest[:12], 16) % 10_000_000_000
    return f"{MATHTYPE_EQUATION_BOOKMARK_PREFIX}{numeric:010d}"


def _subfigure_marker_bookmark_name(label: str) -> str:
    digest = hashlib.sha1(label.strip().encode("utf-8")).hexdigest()[:32]
    return f"figSub_{digest}"


def _normalize_for_similarity(text: str) -> str:
    return re.sub(r"\s+", "", text or "").strip()


def _char_jaccard(lhs: str, rhs: str) -> float:
    left = _normalize_for_similarity(lhs)
    right = _normalize_for_similarity(rhs)
    if not left or not right:
        return 0.0
    left_set = set(left)
    right_set = set(right)
    union = left_set | right_set
    if not union:
        return 0.0
    return len(left_set & right_set) / len(union)


def _label_text_for_match(label: str) -> str:
    token = (label or "").strip()
    if ":" in token:
        token = token.split(":", 1)[1]
    return _normalize_for_similarity(token.replace("_", ""))


def _caption_text_for_match(text: str) -> str:
    compact = _normalize_for_similarity(text)
    # 兼容“图 1 xxx / 表 1 xxx / Figure 1 xxx / Table 1 xxx”等前缀。
    compact = re.sub(r"^(图|表|figure|table|fig\.?)[\d\.\-_:：]*", "", compact, flags=re.IGNORECASE)
    return compact


def _label_caption_similarity(label: str, caption_text: str) -> float:
    label_text = _label_text_for_match(label)
    caption_text_normalized = _caption_text_for_match(caption_text)
    if not label_text or not caption_text_normalized:
        return 0.0
    if label_text in caption_text_normalized:
        return 1.0
    return _char_jaccard(label_text, caption_text_normalized)


def _align_label_pairs_to_captions(
    caption_paragraphs: list[ET.Element],
    labels: list[str],
) -> tuple[list[tuple[ET.Element, str]], list[str], dict[str, str]]:
    """
    对 (caption, label) 做单调对齐：
    - 默认按顺序一一配对；
    - 当标签数量多于题注数量且“下一个标签”明显更像当前题注时，跳过当前标签；
    - 记录跳过标签的别名建议（alias -> next_label），用于 figure REF 映射补偿。
    """
    pairs: list[tuple[ET.Element, str]] = []
    skipped_labels: list[str] = []
    alias_suggestions: dict[str, str] = {}

    if not caption_paragraphs or not labels:
        return pairs, skipped_labels, alias_suggestions

    cap_texts = [_paragraph_text(paragraph) for paragraph in caption_paragraphs]
    label_index = 0
    caption_index = 0

    while label_index < len(labels) and caption_index < len(caption_paragraphs):
        label = labels[label_index]
        caption_text = cap_texts[caption_index]
        score_current = _label_caption_similarity(label, caption_text)

        score_next = -1.0
        has_next_label = label_index + 1 < len(labels)
        if has_next_label:
            score_next = _label_caption_similarity(labels[label_index + 1], caption_text)

        remaining_labels = len(labels) - label_index
        remaining_captions = len(caption_paragraphs) - caption_index
        can_skip_label = remaining_labels > remaining_captions
        should_skip = False

        if can_skip_label and has_next_label:
            if score_next >= 0.33 and score_next >= score_current + 0.08:
                should_skip = True
            elif score_current < 0.2 and score_next > score_current + 0.03:
                should_skip = True

        if should_skip:
            skipped = labels[label_index]
            skipped_labels.append(skipped)
            alias_suggestions.setdefault(skipped, labels[label_index + 1])
            label_index += 1
            continue

        pairs.append((caption_paragraphs[caption_index], label))
        label_index += 1
        caption_index += 1

    if label_index < len(labels):
        skipped_labels.extend(labels[label_index:])

    return pairs, _dedup_keep_order(skipped_labels), alias_suggestions


def _slot_labels_for_pairing(
    slot: dict[str, object],
    *,
    object_kind: str,
) -> tuple[list[str], list[str], dict[str, str]]:
    """
    统一读取槽位中的 primary/all/alias 信息。
    """
    if object_kind == "figure":
        raw_primary = slot.get("primary_labels", [])
        raw_all = slot.get("all_labels", [])
        raw_alias = slot.get("alias_to_primary", {})

        primary_labels = _dedup_keep_order(str(v) for v in raw_primary) if isinstance(raw_primary, list) else []
        all_labels = _dedup_keep_order(str(v) for v in raw_all) if isinstance(raw_all, list) else []
        if not all_labels:
            all_labels = list(primary_labels)
        if not primary_labels and all_labels:
            primary_labels = [all_labels[0]]

        alias_to_primary: dict[str, str] = {}
        if isinstance(raw_alias, dict):
            for alias, primary in raw_alias.items():
                alias_key = str(alias).strip()
                primary_key = str(primary).strip()
                if not alias_key or not primary_key:
                    continue
                if alias_key in alias_to_primary:
                    continue
                alias_to_primary[alias_key] = primary_key
        return primary_labels, all_labels, alias_to_primary

    raw_labels = slot.get("labels", [])
    labels = _dedup_keep_order(str(v) for v in raw_labels) if isinstance(raw_labels, list) else []
    return labels, labels, {}


def _build_caption_label_pairs_by_slots(
    *,
    object_kind: str,
    caption_paragraphs: list[ET.Element],
    slots: list[dict[str, object]],
    fallback_labels: list[str],
) -> tuple[list[tuple[ET.Element, str]], list[str], dict[str, str], str, list[str]]:
    """
    基于 TeX 槽位驱动图/表 (caption, label) 配对。

    返回：
    - label_pairs
    - skipped_labels（槽位配对中被跳过的标签）
    - alias_suggestions（alias -> primary）
    - pairing_strategy: slots / slots_mismatch_resync / fallback_label_alignment / no_caption
    - pairing_warnings
    """
    warnings: list[str] = []
    pairs: list[tuple[ET.Element, str]] = []
    skipped_labels: list[str] = []
    alias_suggestions: dict[str, str] = {}

    if not caption_paragraphs:
        return pairs, skipped_labels, alias_suggestions, "no_caption", warnings

    if not slots:
        warnings.append(
            f"未提取到 {object_kind} TeX 槽位，已回退到标签-题注顺序对齐（fallback_label_alignment）。"
        )
        fallback_pairs, fallback_skipped, fallback_alias = _align_label_pairs_to_captions(
            caption_paragraphs,
            fallback_labels,
        )
        return fallback_pairs, fallback_skipped, fallback_alias, "fallback_label_alignment", warnings

    cap_texts = [_paragraph_text(paragraph) for paragraph in caption_paragraphs]

    slot_primaries: list[list[str]] = []
    slot_all_labels: list[list[str]] = []
    slot_alias_maps: list[dict[str, str]] = []
    for slot in slots:
        primary_labels, all_labels, alias_map = _slot_labels_for_pairing(slot, object_kind=object_kind)
        slot_primaries.append(primary_labels)
        slot_all_labels.append(all_labels)
        slot_alias_maps.append(alias_map)

    labeled_slot_suffix: list[int] = [0] * (len(slots) + 1)
    for idx in range(len(slots) - 1, -1, -1):
        labeled_slot_suffix[idx] = labeled_slot_suffix[idx + 1] + (1 if slot_primaries[idx] else 0)

    slot_index = 0
    caption_index = 0
    while slot_index < len(slots) and caption_index < len(caption_paragraphs):
        primary_labels = slot_primaries[slot_index]
        all_labels = slot_all_labels[slot_index]
        alias_map = slot_alias_maps[slot_index]

        if not primary_labels:
            # 无标签槽位只用于占位，保持与题注位置同步推进。
            slot_index += 1
            caption_index += 1
            continue

        primary = primary_labels[0]
        caption_text = cap_texts[caption_index]
        score_current = _label_caption_similarity(primary, caption_text)

        next_labeled_slot = -1
        for pos in range(slot_index + 1, len(slots)):
            if slot_primaries[pos]:
                next_labeled_slot = pos
                break

        score_next = -1.0
        next_primary = ""
        if next_labeled_slot >= 0:
            next_primary = slot_primaries[next_labeled_slot][0]
            score_next = _label_caption_similarity(next_primary, caption_text)

        remaining_labeled_slots = labeled_slot_suffix[slot_index]
        remaining_captions = len(caption_paragraphs) - caption_index
        can_skip_slot = remaining_labeled_slots > remaining_captions
        should_skip = False

        if can_skip_slot and next_labeled_slot >= 0:
            if score_next >= 0.33 and score_next >= score_current + 0.08:
                should_skip = True
            elif score_current < 0.2 and score_next > score_current + 0.03:
                should_skip = True

        if should_skip:
            skipped_labels.extend(all_labels)
            for alias_label in all_labels:
                if not alias_label or alias_label == next_primary:
                    continue
                alias_suggestions.setdefault(alias_label, next_primary)
            slot_index += 1
            continue

        pairs.append((caption_paragraphs[caption_index], primary))
        for alias_label in all_labels:
            if alias_label and alias_label != primary:
                alias_suggestions.setdefault(alias_label, primary)
        for alias_label, target_label in alias_map.items():
            if not alias_label or not target_label:
                continue
            if target_label == primary:
                alias_suggestions.setdefault(alias_label, primary)

        slot_index += 1
        caption_index += 1

    if slot_index < len(slots):
        for pos in range(slot_index, len(slots)):
            skipped_labels.extend(slot_all_labels[pos])

    strategy = "slots" if len(slots) == len(caption_paragraphs) else "slots_mismatch_resync"
    if strategy == "slots_mismatch_resync":
        warnings.append(
            f"{object_kind} 槽位数与 docx 题注段落数不一致，已采用槽位驱动重同步配对（slots_mismatch_resync）。"
        )

    return (
        pairs,
        _dedup_keep_order(skipped_labels),
        alias_suggestions,
        strategy,
        warnings,
    )


def _add_bookmark(
    paragraph: ET.Element,
    *,
    bookmark_name: str,
    bookmark_id: int,
) -> None:
    start = ET.Element(wqn("bookmarkStart"))
    start.set(wqn("id"), str(bookmark_id))
    start.set(wqn("name"), bookmark_name)
    _insert_nodes_after_ppr(paragraph, [start])

    end = ET.Element(wqn("bookmarkEnd"))
    end.set(wqn("id"), str(bookmark_id))
    paragraph.append(end)


def _find_direct_seq_field(paragraph: ET.Element, seq_name: str) -> ET.Element | None:
    keyword = f"SEQ {seq_name.upper()}"
    for child in list(paragraph):
        if child.tag != wqn("fldSimple"):
            continue
        instr = (child.get(wqn("instr"), "") or "").upper()
        if keyword in instr:
            return child
    return None


def _insert_bookmark_around_child(
    paragraph: ET.Element,
    child: ET.Element,
    *,
    bookmark_name: str,
    bookmark_id: int,
) -> bool:
    children = list(paragraph)
    try:
        idx = children.index(child)
    except ValueError:
        return False

    start = ET.Element(wqn("bookmarkStart"))
    start.set(wqn("id"), str(bookmark_id))
    start.set(wqn("name"), bookmark_name)
    end = ET.Element(wqn("bookmarkEnd"))
    end.set(wqn("id"), str(bookmark_id))

    paragraph.insert(idx, start)
    paragraph.insert(idx + 2, end)
    return True


def _insert_bookmark_around_child_range(
    paragraph: ET.Element,
    start_child: ET.Element,
    end_child: ET.Element,
    *,
    bookmark_name: str,
    bookmark_id: int,
) -> bool:
    children = list(paragraph)
    try:
        start_idx = children.index(start_child)
        end_idx = children.index(end_child)
    except ValueError:
        return False

    if end_idx < start_idx:
        start_idx, end_idx = end_idx, start_idx

    start = ET.Element(wqn("bookmarkStart"))
    start.set(wqn("id"), str(bookmark_id))
    start.set(wqn("name"), bookmark_name)
    end = ET.Element(wqn("bookmarkEnd"))
    end.set(wqn("id"), str(bookmark_id))

    paragraph.insert(start_idx, start)
    paragraph.insert(end_idx + 2, end)
    return True


def _insert_equation_number_bookmark(
    paragraph: ET.Element,
    *,
    seq_name: str,
    bookmark_name: str,
    bookmark_id: int,
) -> bool:
    """
    将书签包裹在公式编号整体 `(SEQ)` 上，符合 MathType 引用期望。
    """
    seq_field = _find_direct_seq_field(paragraph, seq_name)
    if seq_field is None:
        return False

    children = list(paragraph)
    try:
        seq_idx = children.index(seq_field)
    except ValueError:
        return False

    start_child = seq_field
    end_child = seq_field
    if seq_idx > 0 and _run_text_matches_paren(children[seq_idx - 1], "("):
        start_child = children[seq_idx - 1]
    if seq_idx + 1 < len(children) and _run_text_matches_paren(children[seq_idx + 1], ")"):
        end_child = children[seq_idx + 1]

    return _insert_bookmark_around_child_range(
        paragraph,
        start_child,
        end_child,
        bookmark_name=bookmark_name,
        bookmark_id=bookmark_id,
    )


def _find_subcaption_marker_run(paragraph: ET.Element) -> tuple[ET.Element | None, str]:
    marker_pattern = re.compile(r"^[\s\u00A0]*([\(（]\s*[A-Za-z0-9]+\s*[\)）])")
    for child in list(paragraph):
        if child.tag != wqn("r"):
            continue
        text = _run_text(child)
        if not text:
            continue
        match = marker_pattern.match(text)
        if not match:
            continue
        marker_text = match.group(1)
        if marker_text.startswith("（") and marker_text.endswith("）"):
            marker_text = "(" + marker_text[1:-1].strip() + ")"
        else:
            marker_text = marker_text.replace("（", "(").replace("）", ")")
        marker_text = re.sub(r"\s+", "", marker_text)
        return child, marker_text
    return None, ""


def _build_subfigure_ref_mapping(
    *,
    document_root: ET.Element,
    labels: list[str],
    existing_bookmarks: set[str],
    bookmark_id_seed: int,
) -> tuple[dict[str, str], dict[str, str], int, int, list[str], list[str]]:
    """
    为子图标签构建 REF 映射：
    - 在子图单元格题注行 `(a)/(b)` 处补充书签；
    - 将子图标签 anchor 映射到该书签，用于 \subref 的 REF 字段升级。
    """
    next_id = bookmark_id_seed
    bookmark_added = 0
    labels_bound = 0
    labels_missing_marker: list[str] = []
    if not labels:
        return {}, {}, bookmark_added, labels_bound, labels_missing_marker, []

    dedup_labels = _dedup_keep_order(labels)
    anchor_to_label: dict[str, str] = {}
    for label in dedup_labels:
        for anchor_name in _candidate_anchor_names(label):
            anchor_to_label.setdefault(anchor_name, label)

    label_to_marker_bookmark: dict[str, str] = {}
    label_to_marker_text: dict[str, str] = {}
    encountered_labels: set[str] = set()

    for table_cell in document_root.findall(".//w:tc", NS):
        paragraphs = table_cell.findall("./w:p", NS)
        if not paragraphs:
            continue

        for index, paragraph in enumerate(paragraphs):
            bookmark_names = [
                (node.get(wqn("name"), "") or "").strip()
                for node in paragraph.findall(".//w:bookmarkStart", NS)
            ]
            candidate_labels = _dedup_keep_order(
                anchor_to_label[name]
                for name in bookmark_names
                if name in anchor_to_label
            )
            if not candidate_labels:
                continue
            encountered_labels.update(candidate_labels)

            marker_paragraph = None
            marker_run = None
            marker_text = ""
            for probe_index in range(index, min(index + 3, len(paragraphs))):
                run, marker = _find_subcaption_marker_run(paragraphs[probe_index])
                if run is None:
                    continue
                marker_paragraph = paragraphs[probe_index]
                marker_run = run
                marker_text = marker
                break

            if marker_paragraph is None or marker_run is None:
                continue

            for label in candidate_labels:
                if label in label_to_marker_bookmark:
                    continue
                marker_bookmark = _subfigure_marker_bookmark_name(label)
                if marker_bookmark not in existing_bookmarks:
                    wrapped = _insert_bookmark_around_child(
                        marker_paragraph,
                        marker_run,
                        bookmark_name=marker_bookmark,
                        bookmark_id=next_id,
                    )
                    if wrapped:
                        existing_bookmarks.add(marker_bookmark)
                        next_id += 1
                        bookmark_added += 1
                if marker_bookmark in existing_bookmarks:
                    label_to_marker_bookmark[label] = marker_bookmark
                    label_to_marker_text[label] = marker_text

    mapping: dict[str, str] = {}
    display_text: dict[str, str] = {}
    detected_labels = _dedup_keep_order(label for label in dedup_labels if label in encountered_labels)
    for label in detected_labels:
        marker_bookmark = label_to_marker_bookmark.get(label)
        if not marker_bookmark:
            labels_missing_marker.append(label)
            continue
        labels_bound += 1
        marker_text = label_to_marker_text.get(label, "")
        for anchor_name in _candidate_anchor_names(label):
            mapping.setdefault(anchor_name, marker_bookmark)
            if marker_text:
                display_text.setdefault(anchor_name, marker_text)

    return (
        mapping,
        display_text,
        bookmark_added,
        labels_bound,
        _dedup_keep_order(labels_missing_marker),
        detected_labels,
    )


def _repair_label_bookmarks(
    *,
    label_pairs: list[tuple[ET.Element, str]],
    missing_anchors: set[str],
    existing_bookmarks: set[str],
    bookmark_id_seed: int,
) -> tuple[int, int]:
    next_id = bookmark_id_seed
    added = 0
    touched_labels = 0

    for paragraph, label in label_pairs:
        label_added = False
        for anchor_name in _candidate_anchor_names(label):
            if anchor_name not in missing_anchors:
                continue
            if anchor_name in existing_bookmarks:
                continue
            _add_bookmark(paragraph, bookmark_name=anchor_name, bookmark_id=next_id)
            next_id += 1
            added += 1
            label_added = True
            existing_bookmarks.add(anchor_name)
            missing_anchors.discard(anchor_name)
        if label_added:
            touched_labels += 1
    return added, touched_labels


def _group_aliases_by_primary(
    alias_to_primary: dict[str, str],
    primary_labels: list[str],
) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = {}
    for primary in primary_labels:
        key = primary.strip()
        if not key or key in grouped:
            continue
        grouped[key] = [key]

    for alias, primary in alias_to_primary.items():
        alias_key = alias.strip()
        primary_key = primary.strip()
        if not alias_key or not primary_key:
            continue
        bucket = grouped.get(primary_key)
        if bucket is None:
            continue
        if alias_key not in bucket:
            bucket.append(alias_key)

    return grouped


def _build_equation_label_pairs_by_slots(
    *,
    display_equation_paragraphs: list[ET.Element],
    equation_slots: list[dict[str, object]],
    fallback_equation_labels: list[str],
) -> tuple[list[tuple[ET.Element, str]], dict[str, list[str]], list[ET.Element], str, list[str], int]:
    """
    基于 TeX 槽位对齐显示公式段落，生成 equation (paragraph, label) 配对。

    返回：
    - label_pairs
    - label_aliases_by_primary
    - number_targets（需要补 SEQ MTEqn 的公式段落）
    - pairing_strategy: slots / fallback_zip
    - pairing_warnings
    - unlabeled_equation_numbered_count（由 `equation` 环境触发的无标签编号数量）
    """
    warnings: list[str] = []

    if display_equation_paragraphs and equation_slots and len(equation_slots) == len(display_equation_paragraphs):
        pairs: list[tuple[ET.Element, str]] = []
        alias_to_primary: dict[str, str] = {}
        primary_labels: list[str] = []
        number_targets: list[ET.Element] = []
        unlabeled_equation_numbered_count = 0

        for paragraph, slot in zip(display_equation_paragraphs, equation_slots):
            env = str(slot.get("env", "")).strip().lower()
            raw_labels = slot.get("labels", [])
            if isinstance(raw_labels, list):
                dedup_labels = _dedup_keep_order(str(v) for v in raw_labels)
            else:
                dedup_labels = []

            # 规则：
            # 1) 任意带标签的显示公式都编号；
            # 2) 额外对 \begin{equation}...\end{equation} 的无标签独立公式补编号。
            should_number = bool(dedup_labels) or env == "equation"
            if should_number:
                number_targets.append(paragraph)
            if env == "equation" and not dedup_labels:
                unlabeled_equation_numbered_count += 1

            if not dedup_labels:
                continue

            primary = dedup_labels[0]
            if primary not in primary_labels:
                primary_labels.append(primary)
            pairs.append((paragraph, primary))

            for alias in dedup_labels:
                _bind_alias(alias_to_primary, alias=alias, primary=primary)

        grouped_aliases = _group_aliases_by_primary(alias_to_primary, primary_labels)
        return (
            pairs,
            grouped_aliases,
            number_targets,
            "slots",
            warnings,
            unlabeled_equation_numbered_count,
        )

    warnings.append(
        "公式槽位数与 docx 显示公式段落数不一致，已回退到顺序配对（fallback_zip）。"
    )
    fallback_pairs = list(zip(display_equation_paragraphs, fallback_equation_labels))
    return (
        fallback_pairs,
        {},
        [paragraph for paragraph, _label in fallback_pairs],
        "fallback_zip",
        warnings,
        0,
    )


def _build_caption_ref_bookmark_mapping(
    *,
    label_pairs: list[tuple[ET.Element, str]],
    label_aliases_by_primary: dict[str, list[str]] | None,
    existing_bookmarks: set[str],
    bookmark_id_seed: int,
    seq_name: str,
    object_kind: str,
) -> tuple[dict[str, str], int, int, list[str]]:
    """
    为指定对象（figure/table/equation）标签创建“编号专用书签”，并输出 anchor -> 书签名映射。
    """
    next_id = bookmark_id_seed
    added = 0
    labels_bound = 0
    labels_missing_seq: list[str] = []
    mapping: dict[str, str] = {}

    for paragraph, label in label_pairs:
        label_key = label.strip()
        if not label_key:
            continue

        number_bookmark = _number_bookmark_name(object_kind, label_key)
        if not number_bookmark:
            labels_missing_seq.append(label_key)
            continue
        if number_bookmark not in existing_bookmarks:
            if object_kind == "equation":
                wrapped = _insert_equation_number_bookmark(
                    paragraph,
                    seq_name=seq_name,
                    bookmark_name=number_bookmark,
                    bookmark_id=next_id,
                )
            else:
                seq_field = _find_direct_seq_field(paragraph, seq_name)
                if seq_field is None:
                    labels_missing_seq.append(label_key)
                    continue
                wrapped = _insert_bookmark_around_child(
                    paragraph,
                    seq_field,
                    bookmark_name=number_bookmark,
                    bookmark_id=next_id,
                )
            if not wrapped:
                labels_missing_seq.append(label_key)
                continue
            existing_bookmarks.add(number_bookmark)
            next_id += 1
            added += 1

        labels_bound += 1
        alias_labels = [label_key]
        if label_aliases_by_primary:
            alias_labels = label_aliases_by_primary.get(label_key, alias_labels)
        for alias_label in alias_labels:
            alias_key = alias_label.strip()
            if not alias_key:
                continue
            for anchor_name in _candidate_anchor_names(alias_key):
                mapping.setdefault(anchor_name, number_bookmark)

    return mapping, added, labels_bound, labels_missing_seq


def _convert_hyperlinks_to_ref_fields(
    document_root: ET.Element,
    *,
    anchor_to_number_bookmark: dict[str, str],
    anchor_display_text: dict[str, str] | None = None,
    ref_node_builder: Callable[[str, str], list[ET.Element]] | None = None,
) -> tuple[int, int, int]:
    """
    将文内 hyperlink(anchor) 转换为 REF 字段（仅处理传入映射中的锚点）。
    返回：
    - converted_link_count
    - candidate_link_count
    - candidate_anchor_name_count
    """
    converted = 0
    candidate = 0
    candidate_anchor_names: set[str] = set()
    builder = ref_node_builder or _build_simple_ref_nodes

    for paragraph in document_root.findall(".//w:p", NS):
        children = list(paragraph)
        for child in children:
            if child.tag != wqn("hyperlink"):
                continue
            anchor_name = child.get(wqn("anchor"), "").strip()
            if not anchor_name:
                continue
            bookmark_name = anchor_to_number_bookmark.get(anchor_name)
            if not bookmark_name:
                continue

            candidate += 1
            candidate_anchor_names.add(anchor_name)
            preferred_text = ""
            if anchor_display_text:
                preferred_text = anchor_display_text.get(anchor_name, "")
            display_text = preferred_text or "".join((node.text or "") for node in child.findall(".//w:t", NS))
            replacement_nodes = builder(bookmark_name, display_text)
            if not replacement_nodes:
                continue

            insertion_index = list(paragraph).index(child)
            paragraph.remove(child)
            for offset, node in enumerate(replacement_nodes):
                paragraph.insert(insertion_index + offset, node)
            converted += 1

    return converted, candidate, len(candidate_anchor_names)


def _repair_remaining_anchors_by_similarity(
    *,
    missing_anchors: set[str],
    candidate_paragraphs: list[tuple[ET.Element, str]],
    existing_bookmarks: set[str],
    bookmark_id_seed: int,
) -> tuple[int, list[dict[str, object]]]:
    """
    对仍缺失的锚点做近似修复（仅用于图表等文本标签）：
    - 以字符集合 Jaccard 评估 anchor 与候选题注文本相似度；
    - 要求 best >= 0.45 且领先 second 至少 0.05，避免误绑定。
    """
    next_id = bookmark_id_seed
    added = 0
    mapping_records: list[dict[str, object]] = []

    if not candidate_paragraphs:
        return added, mapping_records

    for anchor_name in sorted(missing_anchors):
        if anchor_name in existing_bookmarks:
            continue
        if anchor_name.startswith("X") and len(anchor_name) == 40:
            # 哈希锚点通常来源于公式标签，留给精确映射，不做模糊绑定。
            continue

        scored: list[tuple[float, ET.Element, str]] = []
        for paragraph, text in candidate_paragraphs:
            score = _char_jaccard(anchor_name, text)
            if score > 0:
                scored.append((score, paragraph, text))
        if not scored:
            continue

        scored.sort(key=lambda item: item[0], reverse=True)
        best_score, best_paragraph, best_text = scored[0]
        second_score = scored[1][0] if len(scored) > 1 else 0.0
        if best_score < 0.45:
            continue
        if best_score - second_score < 0.05:
            continue

        _add_bookmark(best_paragraph, bookmark_name=anchor_name, bookmark_id=next_id)
        next_id += 1
        added += 1
        existing_bookmarks.add(anchor_name)
        missing_anchors.discard(anchor_name)
        mapping_records.append(
            {
                "anchor": anchor_name,
                "matched_caption_text": best_text,
                "similarity": round(best_score, 4),
            }
        )

    return added, mapping_records


def run_docx_postprocess(docx_path: Path, tex_files: list[Path]) -> DocxPostprocessResult:
    if not docx_path.exists():
        raise RuntimeError(f"DOCX 不存在：{docx_path}")

    inventory = extract_label_inventory(tex_files)
    warnings: list[str] = []

    with zipfile.ZipFile(docx_path, "r") as zin:
        members = zin.namelist()
        payload = {name: zin.read(name) for name in members}

    if "word/document.xml" not in payload:
        raise RuntimeError("DOCX 缺少 word/document.xml。")
    if "word/styles.xml" not in payload:
        raise RuntimeError("DOCX 缺少 word/styles.xml。")

    ET.register_namespace("w", W_NS)
    ET.register_namespace("m", M_NS)

    document_root = ET.fromstring(payload["word/document.xml"])
    styles_root = ET.fromstring(payload["word/styles.xml"])

    figure_style_ids, table_style_ids, generic_caption_style_ids = _parse_caption_style_ids(styles_root)
    mt_display_equation_style_id = _find_mt_display_equation_style_id(styles_root)
    preferred_table_style_id = _find_preferred_table_style_id(styles_root)

    body = document_root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("document.xml 缺少 w:body。")

    (
        algorithm_table_wrapped_count,
        algorithm_table_wrapped_paragraph_count,
        algorithm_block_detected_count,
        algorithm_block_skipped_count,
        algorithm_table_style_id,
        algorithm_block_samples,
    ) = _wrap_algorithm_blocks_as_tables(
        body=body,
        styles_root=styles_root,
    )
    (
        table_style_applied_count,
        table_style_candidate_count,
        table_style_skipped_algorithm_count,
    ) = _apply_preferred_table_style(
        document_root=document_root,
        preferred_style_id=preferred_table_style_id,
    )

    elements = list(body)
    figure_caption_paragraphs: list[ET.Element] = []
    table_caption_paragraphs: list[ET.Element] = []
    display_equation_paragraphs: list[ET.Element] = []

    for index, element in enumerate(elements):
        if element.tag != wqn("p"):
            continue
        style_id = _paragraph_style_id(element)
        if style_id in figure_style_ids:
            figure_caption_paragraphs.append(element)
        elif style_id in table_style_ids:
            table_caption_paragraphs.append(element)
        elif style_id in generic_caption_style_ids:
            kind = _infer_caption_kind(elements, index, element)
            if kind == "table":
                table_caption_paragraphs.append(element)
            else:
                figure_caption_paragraphs.append(element)

        if _is_display_equation_paragraph(element):
            display_equation_paragraphs.append(element)

    anchors = {
        node.get(wqn("anchor"), "")
        for node in document_root.findall(".//w:hyperlink[@w:anchor]", NS)
        if node.get(wqn("anchor"), "")
    }
    bookmark_nodes = document_root.findall(".//w:bookmarkStart", NS)
    existing_bookmarks = {
        node.get(wqn("name"), "")
        for node in bookmark_nodes
        if node.get(wqn("name"), "")
    }
    missing_anchors = set(anchor for anchor in anchors if anchor not in existing_bookmarks)

    max_bookmark_id = 0
    for node in bookmark_nodes:
        raw = node.get(wqn("id"), "")
        if raw.isdigit():
            max_bookmark_id = max(max_bookmark_id, int(raw))
    next_bookmark_id = max_bookmark_id + 1

    figure_seq_added = 0
    for index, paragraph in enumerate(figure_caption_paragraphs, start=1):
        if _prepend_caption_seq(
            paragraph,
            prefix="图 ",
            seq_name="Figure",
            sequence_index=index,
        ):
            figure_seq_added += 1

    table_seq_added = 0
    for index, paragraph in enumerate(table_caption_paragraphs, start=1):
        if _prepend_caption_seq(
            paragraph,
            prefix="表 ",
            seq_name="Table",
            sequence_index=index,
        ):
            table_seq_added += 1

    figure_slots, table_slots = extract_float_slots(tex_files)
    (
        figure_label_pairs,
        figure_labels_skipped_by_pairing,
        figure_alias_suggestions,
        figure_pairing_strategy,
        figure_pairing_warnings,
    ) = _build_caption_label_pairs_by_slots(
        object_kind="figure",
        caption_paragraphs=figure_caption_paragraphs,
        slots=figure_slots,
        fallback_labels=inventory.figure_labels,
    )
    (
        table_label_pairs,
        table_labels_skipped_by_pairing,
        _table_alias_suggestions,
        table_pairing_strategy,
        table_pairing_warnings,
    ) = _build_caption_label_pairs_by_slots(
        object_kind="table",
        caption_paragraphs=table_caption_paragraphs,
        slots=table_slots,
        fallback_labels=inventory.table_labels,
    )
    warnings.extend(figure_pairing_warnings)
    warnings.extend(table_pairing_warnings)
    equation_slots = extract_equation_display_slots(tex_files)
    (
        equation_label_pairs,
        equation_aliases_by_primary,
        equation_number_targets,
        equation_pairing_strategy,
        equation_pairing_warnings,
        equation_unlabeled_equation_numbered_count,
    ) = _build_equation_label_pairs_by_slots(
        display_equation_paragraphs=display_equation_paragraphs,
        equation_slots=equation_slots,
        fallback_equation_labels=inventory.equation_labels,
    )
    warnings.extend(equation_pairing_warnings)
    figure_aliases_by_primary = _group_aliases_by_primary(
        inventory.figure_label_alias_to_primary,
        inventory.figure_labels,
    )
    paired_figure_labels = {label for _paragraph, label in figure_label_pairs}
    for alias_label, primary_label in figure_alias_suggestions.items():
        if primary_label not in paired_figure_labels:
            continue
        bucket = figure_aliases_by_primary.get(primary_label)
        if bucket is None:
            bucket = [primary_label]
            figure_aliases_by_primary[primary_label] = bucket
        if alias_label not in bucket:
            bucket.append(alias_label)

    bookmark_added_figure, figure_label_repaired = _repair_label_bookmarks(
        label_pairs=figure_label_pairs,
        missing_anchors=missing_anchors,
        existing_bookmarks=existing_bookmarks,
        bookmark_id_seed=next_bookmark_id,
    )
    next_bookmark_id += bookmark_added_figure

    bookmark_added_table, table_label_repaired = _repair_label_bookmarks(
        label_pairs=table_label_pairs,
        missing_anchors=missing_anchors,
        existing_bookmarks=existing_bookmarks,
        bookmark_id_seed=next_bookmark_id,
    )
    next_bookmark_id += bookmark_added_table

    equation_prefix_tab_added = 0
    for paragraph in display_equation_paragraphs:
        if _ensure_equation_prefix_tab(paragraph):
            equation_prefix_tab_added += 1

    equation_seq_added = 0
    for index, paragraph in enumerate(equation_number_targets, start=1):
        if _append_equation_seq(paragraph, seq_name=EQUATION_SEQ_NAME, sequence_index=index):
            equation_seq_added += 1

    equation_number_tab_layout_fixed = 0
    for paragraph in display_equation_paragraphs:
        if _ensure_equation_number_tab_layout(paragraph):
            equation_number_tab_layout_fixed += 1

    equation_macrobutton_wrapped = 0
    for paragraph in equation_number_targets:
        if _ensure_equation_number_macrobutton_wrapper(paragraph):
            equation_macrobutton_wrapped += 1

    equation_style_applied = 0
    if mt_display_equation_style_id:
        for paragraph in display_equation_paragraphs:
            if _set_paragraph_style_id(paragraph, mt_display_equation_style_id):
                equation_style_applied += 1

    bookmark_added_equation, equation_label_repaired = _repair_label_bookmarks(
        label_pairs=equation_label_pairs,
        missing_anchors=missing_anchors,
        existing_bookmarks=existing_bookmarks,
        bookmark_id_seed=next_bookmark_id,
    )
    next_bookmark_id += bookmark_added_equation

    figure_anchor_labels = inventory.figure_all_labels or inventory.figure_labels
    subfigure_candidate_labels = [
        label
        for label in figure_anchor_labels
        if inventory.figure_label_alias_to_primary.get(label, label) != label
    ]
    (
        subfigure_ref_mapping,
        subfigure_ref_display_text,
        subfigure_marker_bookmark_added,
        subfigure_labels_bound,
        subfigure_labels_missing_marker,
        subfigure_detected_labels,
    ) = _build_subfigure_ref_mapping(
        document_root=document_root,
        labels=subfigure_candidate_labels,
        existing_bookmarks=existing_bookmarks,
        bookmark_id_seed=next_bookmark_id,
    )
    next_bookmark_id += subfigure_marker_bookmark_added

    figure_anchor_name_all = {
        anchor_name
        for label in figure_anchor_labels
        for anchor_name in _candidate_anchor_names(label)
    }
    table_anchor_name_all = {
        anchor_name
        for label in inventory.table_labels
        for anchor_name in _candidate_anchor_names(label)
    }
    equation_anchor_name_all = {
        anchor_name
        for label in inventory.equation_labels
        for anchor_name in _candidate_anchor_names(label)
    }
    figure_ref_mapping, figure_num_bookmark_added, figure_labels_bound, figure_labels_missing_num = (
        _build_caption_ref_bookmark_mapping(
            label_pairs=figure_label_pairs,
            label_aliases_by_primary=figure_aliases_by_primary,
            existing_bookmarks=existing_bookmarks,
            bookmark_id_seed=next_bookmark_id,
            seq_name="Figure",
            object_kind="figure",
        )
    )
    next_bookmark_id += figure_num_bookmark_added
    table_ref_mapping, table_num_bookmark_added, table_labels_bound, table_labels_missing_num = (
        _build_caption_ref_bookmark_mapping(
            label_pairs=table_label_pairs,
            label_aliases_by_primary=None,
            existing_bookmarks=existing_bookmarks,
            bookmark_id_seed=next_bookmark_id,
            seq_name="Table",
            object_kind="table",
        )
    )
    next_bookmark_id += table_num_bookmark_added
    equation_ref_mapping, equation_num_bookmark_added, equation_labels_bound, equation_labels_missing_num = (
        _build_caption_ref_bookmark_mapping(
            label_pairs=equation_label_pairs,
            label_aliases_by_primary=equation_aliases_by_primary or None,
            existing_bookmarks=existing_bookmarks,
            bookmark_id_seed=next_bookmark_id,
            seq_name=EQUATION_SEQ_NAME,
            object_kind="equation",
        )
    )
    next_bookmark_id += equation_num_bookmark_added

    subfigure_anchor_name_all = {
        anchor_name
        for label in subfigure_detected_labels
        for anchor_name in _candidate_anchor_names(label)
    }
    subfigure_hyperlink_anchors_in_doc = anchors & subfigure_anchor_name_all
    figure_hyperlink_anchors_in_doc = anchors & figure_anchor_name_all
    table_hyperlink_anchors_in_doc = anchors & table_anchor_name_all
    equation_hyperlink_anchors_in_doc = anchors & equation_anchor_name_all
    subfigure_anchor_unmapped = sorted(
        anchor_name
        for anchor_name in subfigure_hyperlink_anchors_in_doc
        if anchor_name not in subfigure_ref_mapping
    )
    figure_anchor_unmapped = sorted(
        anchor_name for anchor_name in figure_hyperlink_anchors_in_doc if anchor_name not in figure_ref_mapping
    )
    table_anchor_unmapped = sorted(
        anchor_name for anchor_name in table_hyperlink_anchors_in_doc if anchor_name not in table_ref_mapping
    )
    equation_anchor_unmapped = sorted(
        anchor_name for anchor_name in equation_hyperlink_anchors_in_doc if anchor_name not in equation_ref_mapping
    )
    subfigure_ref_converted, subfigure_ref_candidate_links, subfigure_ref_candidate_anchor_count = (
        _convert_hyperlinks_to_ref_fields(
            document_root,
            anchor_to_number_bookmark=subfigure_ref_mapping,
            anchor_display_text=subfigure_ref_display_text,
        )
    )
    figure_ref_converted, figure_ref_candidate_links, figure_ref_candidate_anchor_count = (
        _convert_hyperlinks_to_ref_fields(
            document_root,
            anchor_to_number_bookmark=figure_ref_mapping,
        )
    )
    table_ref_converted, table_ref_candidate_links, table_ref_candidate_anchor_count = (
        _convert_hyperlinks_to_ref_fields(
            document_root,
            anchor_to_number_bookmark=table_ref_mapping,
        )
    )
    equation_ref_converted, equation_ref_candidate_links, equation_ref_candidate_anchor_count = (
        _convert_hyperlinks_to_ref_fields(
            document_root,
            anchor_to_number_bookmark=equation_ref_mapping,
            ref_node_builder=_build_mathtype_equation_ref_nodes,
        )
    )

    resolved_anchor_names = (
        set(subfigure_ref_mapping)
        | set(figure_ref_mapping)
        | set(table_ref_mapping)
        | set(equation_ref_mapping)
    )
    missing_anchors.difference_update(resolved_anchor_names)

    caption_candidates = [
        (paragraph, _paragraph_text(paragraph).strip())
        for paragraph in (figure_caption_paragraphs + table_caption_paragraphs)
        if _paragraph_text(paragraph).strip()
    ]
    bookmark_added_fuzzy, fuzzy_records = _repair_remaining_anchors_by_similarity(
        missing_anchors=missing_anchors,
        candidate_paragraphs=caption_candidates,
        existing_bookmarks=existing_bookmarks,
        bookmark_id_seed=next_bookmark_id,
    )
    next_bookmark_id += bookmark_added_fuzzy

    bookmark_added_total = (
        subfigure_marker_bookmark_added
        + bookmark_added_figure
        + bookmark_added_table
        + bookmark_added_equation
        + bookmark_added_fuzzy
    )
    modified = any(
        value > 0
        for value in (
            algorithm_table_wrapped_count,
            table_style_applied_count,
            figure_seq_added,
            table_seq_added,
            equation_seq_added,
            equation_prefix_tab_added,
            equation_number_tab_layout_fixed,
            equation_macrobutton_wrapped,
            bookmark_added_total,
            subfigure_marker_bookmark_added,
            figure_num_bookmark_added,
            table_num_bookmark_added,
            equation_num_bookmark_added,
            subfigure_ref_converted,
            figure_ref_converted,
            table_ref_converted,
            equation_ref_converted,
        )
    )

    metrics = {
        "algorithm_block_detected_count": algorithm_block_detected_count,
        "algorithm_block_wrapped_count": algorithm_table_wrapped_count,
        "algorithm_block_skipped_count": algorithm_block_skipped_count,
        "algorithm_paragraph_wrapped_count": algorithm_table_wrapped_paragraph_count,
        "table_style_candidate_count": table_style_candidate_count,
        "table_style_applied_count": table_style_applied_count,
        "table_style_skipped_algorithm_count": table_style_skipped_algorithm_count,
        "figure_caption_para_count": len(figure_caption_paragraphs),
        "table_caption_para_count": len(table_caption_paragraphs),
        "display_equation_para_count": len(display_equation_paragraphs),
        "figure_slot_count_from_tex": len(figure_slots),
        "table_slot_count_from_tex": len(table_slots),
        "figure_label_count_from_tex": len(inventory.figure_all_labels),
        "figure_label_primary_count_from_tex": len(inventory.figure_labels),
        "table_label_count_from_tex": len(inventory.table_labels),
        "equation_label_count_from_tex": len(inventory.equation_labels),
        "equation_slot_count_from_tex": len(equation_slots),
        "equation_label_pair_count": len(equation_label_pairs),
        "equation_number_target_count": len(equation_number_targets),
        "equation_unlabeled_equation_numbered_count": equation_unlabeled_equation_numbered_count,
        "figure_label_pair_count": len(figure_label_pairs),
        "table_label_pair_count": len(table_label_pairs),
        "figure_labels_skipped_by_pairing_count": len(figure_labels_skipped_by_pairing),
        "table_labels_skipped_by_pairing_count": len(table_labels_skipped_by_pairing),
        "figure_labels_skipped_by_alignment_count": len(figure_labels_skipped_by_pairing),
        "table_labels_skipped_by_alignment_count": len(table_labels_skipped_by_pairing),
        "figure_pairing_warning_count": len(figure_pairing_warnings),
        "table_pairing_warning_count": len(table_pairing_warnings),
        "equation_pairing_warning_count": len(equation_pairing_warnings),
        "figure_caption_seq_added": figure_seq_added,
        "table_caption_seq_added": table_seq_added,
        "equation_seq_added": equation_seq_added,
        "equation_prefix_tab_added": equation_prefix_tab_added,
        "equation_number_tab_layout_fixed": equation_number_tab_layout_fixed,
        "equation_macrobutton_wrapped": equation_macrobutton_wrapped,
        "equation_style_applied_count": equation_style_applied,
        "equation_style_candidate_count": len(display_equation_paragraphs),
        "bookmark_added_total": bookmark_added_total,
        "subfigure_marker_bookmark_added": subfigure_marker_bookmark_added,
        "bookmark_added_figure": bookmark_added_figure,
        "bookmark_added_table": bookmark_added_table,
        "bookmark_added_equation": bookmark_added_equation,
        "bookmark_added_fuzzy": bookmark_added_fuzzy,
        "figure_num_bookmark_added": figure_num_bookmark_added,
        "table_num_bookmark_added": table_num_bookmark_added,
        "equation_num_bookmark_added": equation_num_bookmark_added,
        "figure_ref_mapping_anchor_count": len(figure_ref_mapping),
        "figure_ref_candidate_link_count": figure_ref_candidate_links,
        "figure_ref_candidate_anchor_count": figure_ref_candidate_anchor_count,
        "figure_ref_converted_link_count": figure_ref_converted,
        "figure_ref_unmapped_anchor_count": len(figure_anchor_unmapped),
        "subfigure_ref_mapping_anchor_count": len(subfigure_ref_mapping),
        "subfigure_detected_label_count": len(subfigure_detected_labels),
        "subfigure_ref_candidate_link_count": subfigure_ref_candidate_links,
        "subfigure_ref_candidate_anchor_count": subfigure_ref_candidate_anchor_count,
        "subfigure_ref_converted_link_count": subfigure_ref_converted,
        "subfigure_ref_unmapped_anchor_count": len(subfigure_anchor_unmapped),
        "table_ref_mapping_anchor_count": len(table_ref_mapping),
        "table_ref_candidate_link_count": table_ref_candidate_links,
        "table_ref_candidate_anchor_count": table_ref_candidate_anchor_count,
        "table_ref_converted_link_count": table_ref_converted,
        "table_ref_unmapped_anchor_count": len(table_anchor_unmapped),
        "equation_ref_mapping_anchor_count": len(equation_ref_mapping),
        "equation_ref_candidate_link_count": equation_ref_candidate_links,
        "equation_ref_candidate_anchor_count": equation_ref_candidate_anchor_count,
        "equation_ref_converted_link_count": equation_ref_converted,
        "equation_ref_unmapped_anchor_count": len(equation_anchor_unmapped),
        "missing_anchor_count_after": len(missing_anchors),
    }

    details = {
        "algorithm_table_style_id": algorithm_table_style_id,
        "preferred_table_style_id": preferred_table_style_id,
        "algorithm_block_samples": algorithm_block_samples[:20],
        "figure_labels_repaired": figure_label_repaired,
        "table_labels_repaired": table_label_repaired,
        "equation_labels_repaired": equation_label_repaired,
        "equation_display_style_id": mt_display_equation_style_id,
        "figure_labels_bound_to_ref_bookmark": figure_labels_bound,
        "figure_labels_missing_number_bookmark": figure_labels_missing_num,
        "subfigure_labels_bound_to_marker_bookmark": subfigure_labels_bound,
        "subfigure_labels_missing_marker_bookmark": subfigure_labels_missing_marker[:50],
        "subfigure_detected_labels": subfigure_detected_labels[:50],
        "figure_label_alias_count": len(inventory.figure_label_alias_to_primary),
        "figure_primary_with_alias_count": sum(
            1 for labels in figure_aliases_by_primary.values() if len(labels) > 1
        ),
        "figure_pairing_strategy": figure_pairing_strategy,
        "table_pairing_strategy": table_pairing_strategy,
        "figure_pairing_warnings": figure_pairing_warnings[:20],
        "table_pairing_warnings": table_pairing_warnings[:20],
        "figure_labels_skipped_by_pairing": figure_labels_skipped_by_pairing[:50],
        "table_labels_skipped_by_pairing": table_labels_skipped_by_pairing[:50],
        "figure_labels_skipped_by_alignment": figure_labels_skipped_by_pairing[:50],
        "table_labels_skipped_by_alignment": table_labels_skipped_by_pairing[:50],
        "figure_alias_suggestions_applied": {
            alias: primary
            for alias, primary in figure_alias_suggestions.items()
            if primary in paired_figure_labels
        },
        "table_labels_bound_to_ref_bookmark": table_labels_bound,
        "table_labels_missing_number_bookmark": table_labels_missing_num,
        "equation_labels_bound_to_ref_bookmark": equation_labels_bound,
        "equation_labels_missing_number_bookmark": equation_labels_missing_num,
        "equation_pairing_strategy": equation_pairing_strategy,
        "equation_pairing_warnings": equation_pairing_warnings[:20],
        "equation_alias_group_count": len(equation_aliases_by_primary),
        "equation_slot_unlabeled_count": sum(
            1
            for slot in equation_slots
            if not isinstance(slot.get("labels"), list) or not slot.get("labels")
        ),
        "equation_slot_unlabeled_equation_count": sum(
            1
            for slot in equation_slots
            if str(slot.get("env", "")).strip().lower() == "equation"
            and (not isinstance(slot.get("labels"), list) or not slot.get("labels"))
        ),
        "figure_ref_unmapped_anchors": figure_anchor_unmapped[:50],
        "subfigure_ref_unmapped_anchors": subfigure_anchor_unmapped[:50],
        "table_ref_unmapped_anchors": table_anchor_unmapped[:50],
        "equation_ref_unmapped_anchors": equation_anchor_unmapped[:50],
        "fuzzy_anchor_repairs": fuzzy_records,
        "missing_anchors_sample_after": sorted(missing_anchors)[:50],
    }

    if len(inventory.figure_labels) > len(figure_caption_paragraphs):
        warnings.append(
            "figure 标签数量大于 docx 可识别图题段落数量，部分 figure 书签无法自动修复。"
        )
    if table_style_candidate_count > 0 and not preferred_table_style_id:
        warnings.append(
            "未在 reference.docx 中检测到三线表样式；表格保持 Pandoc 产出的原样式。"
        )
    if algorithm_block_skipped_count > 0:
        warnings.append(
            f"检测到 {algorithm_block_skipped_count} 个算法标题块未匹配到编号步骤，已保持原段落结构。"
        )
    if figure_labels_skipped_by_pairing:
        warnings.append(
            f"figure 槽位配对中跳过 {len(figure_labels_skipped_by_pairing)} 个标签（疑似子图或被 Pandoc 合并的子题注）。"
        )
    if len(inventory.table_labels) > len(table_caption_paragraphs):
        warnings.append(
            "table 标签数量大于 docx 可识别表题段落数量，部分 table 书签无法自动修复。"
        )
    if table_labels_skipped_by_pairing:
        warnings.append(
            f"table 槽位配对中跳过 {len(table_labels_skipped_by_pairing)} 个标签（疑似被 Pandoc 合并或丢失题注）。"
        )
    if len(inventory.equation_labels) > len(display_equation_paragraphs):
        warnings.append(
            "equation 标签数量大于 docx 可识别显示公式段落数量，部分公式书签无法自动修复。"
        )
    if display_equation_paragraphs and not mt_display_equation_style_id:
        warnings.append(
            "未在 reference.docx 中检测到 MTDisplayEquation 样式；显示公式段落保持原样式。"
        )
    if figure_labels_missing_num:
        warnings.append(
            f"{len(figure_labels_missing_num)} 个 figure 标签未能绑定图号书签，相关图引用无法升级为 REF 字段。"
        )
    if subfigure_labels_missing_marker:
        warnings.append(
            f"{len(subfigure_labels_missing_marker)} 个子图标签未能绑定子图标记书签，相关 \\subref 可能无法升级为子图级 REF。"
        )
    if table_labels_missing_num:
        warnings.append(
            f"{len(table_labels_missing_num)} 个 table 标签未能绑定表号书签，相关表引用无法升级为 REF 字段。"
        )
    if equation_labels_missing_num:
        warnings.append(
            f"{len(equation_labels_missing_num)} 个 equation 标签未能绑定公式号书签，相关公式引用无法升级为 REF 字段。"
        )
    if figure_ref_candidate_links > 0 and figure_ref_converted < figure_ref_candidate_links:
        warnings.append(
            "部分 figure 超链接未成功转换为 REF 字段，请人工核对图引用。"
        )
    if subfigure_ref_candidate_links > 0 and subfigure_ref_converted < subfigure_ref_candidate_links:
        warnings.append(
            "部分子图超链接未成功转换为子图级 REF 字段，请人工核对子图引用。"
        )
    if table_ref_candidate_links > 0 and table_ref_converted < table_ref_candidate_links:
        warnings.append(
            "部分 table 超链接未成功转换为 REF 字段，请人工核对表引用。"
        )
    if equation_ref_candidate_links > 0 and equation_ref_converted < equation_ref_candidate_links:
        warnings.append(
            "部分 equation 超链接未成功转换为 REF 字段，请人工核对公式引用。"
        )
    if figure_anchor_unmapped:
        warnings.append(
            f"检测到 {len(figure_anchor_unmapped)} 个 figure 风格锚点未映射到图题，已保留原超链接。"
        )
    if subfigure_anchor_unmapped:
        warnings.append(
            f"检测到 {len(subfigure_anchor_unmapped)} 个子图锚点未映射到子图标记，已保留原超链接。"
        )
    if table_anchor_unmapped:
        warnings.append(
            f"检测到 {len(table_anchor_unmapped)} 个 table 风格锚点未映射到表题，已保留原超链接。"
        )
    if equation_anchor_unmapped:
        warnings.append(
            f"检测到 {len(equation_anchor_unmapped)} 个 equation 风格锚点未映射到公式编号，已保留原超链接。"
        )
    if missing_anchors:
        warnings.append(
            f"仍有 {len(missing_anchors)} 个内部锚点缺失，需人工校核剩余交叉引用。"
        )
    if bookmark_added_fuzzy > 0:
        warnings.append(
            f"已通过近似匹配修复 {bookmark_added_fuzzy} 个锚点，请人工复核这些跳转目标是否符合预期。"
        )

    if modified:
        payload["word/document.xml"] = ET.tostring(
            document_root,
            encoding="utf-8",
            xml_declaration=True,
        )
        temp_output = docx_path.with_name(docx_path.name + ".tmp")
        with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in members:
                zout.writestr(name, payload[name])
        temp_output.replace(docx_path)

    return DocxPostprocessResult(
        modified=modified,
        warnings=warnings,
        metrics=metrics,
        details=details,
    )
