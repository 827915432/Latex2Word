#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_postprocess.py

在 Pandoc 产物 docx 上执行轻量结构后处理，目标是提高可维护性：
1. 为图/表题注补充 Word SEQ 字段编号（图 N / 表 N）；
2. 为带标签的显示公式补充 SEQ 字段编号（(N)）；
3. 按 TeX 标签顺序补齐缺失书签，尽量修复内部超链接跳转。
4. 将文内图引用从内部超链接升级为 Word REF 字段（可随编号更新）。

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
from typing import Iterable
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

TOKEN_PATTERN = re.compile(r"\\begin\{([^}]+)\}|\\end\{([^}]+)\}|\\label\{([^}]+)\}")


@dataclass
class LabelInventory:
    figure_labels: list[str] = field(default_factory=list)
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


def extract_label_inventory(tex_files: list[Path]) -> LabelInventory:
    figure_labels: list[str] = []
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
                figure_labels.append(label_key)
            elif env_set & TABLE_ENVS:
                table_labels.append(label_key)
            elif env_set & EQUATION_ENVS:
                equation_labels.append(label_key)

    return LabelInventory(
        figure_labels=_dedup_keep_order(figure_labels),
        table_labels=_dedup_keep_order(table_labels),
        equation_labels=_dedup_keep_order(equation_labels),
    )


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


def _make_seq_field(seq_name: str, display_number: int | None = None) -> ET.Element:
    fld = ET.Element(wqn("fldSimple"))
    fld.set(wqn("instr"), f"SEQ {seq_name} \\* ARABIC")
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


def _insert_nodes_after_ppr(paragraph: ET.Element, nodes: list[ET.Element]) -> None:
    index = 0
    children = list(paragraph)
    if children and children[0].tag == wqn("pPr"):
        index = 1
    for node in nodes:
        paragraph.insert(index, node)
        index += 1


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
    seq_name: str = "Eq",
    sequence_index: int,
) -> bool:
    if _paragraph_has_seq_field(paragraph, seq_name):
        return False
    text = _paragraph_text(paragraph).strip()
    if re.search(r"\(\d+\)$", text):
        return False
    paragraph.append(_make_run(" ("))
    paragraph.append(_make_seq_field(seq_name, sequence_index))
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


def _figure_num_bookmark_name(label: str) -> str:
    digest = hashlib.sha1(label.strip().encode("utf-8")).hexdigest()[:32]
    return f"figNum_{digest}"


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


def _build_figure_ref_bookmark_mapping(
    *,
    figure_label_pairs: list[tuple[ET.Element, str]],
    existing_bookmarks: set[str],
    bookmark_id_seed: int,
) -> tuple[dict[str, str], int, int, list[str]]:
    """
    为 figure 标签创建“图号专用书签”，并输出 anchor -> 书签名映射。
    """
    next_id = bookmark_id_seed
    added = 0
    labels_bound = 0
    labels_missing_seq: list[str] = []
    mapping: dict[str, str] = {}

    for paragraph, label in figure_label_pairs:
        label_key = label.strip()
        if not label_key:
            continue

        number_bookmark = _figure_num_bookmark_name(label_key)
        if number_bookmark not in existing_bookmarks:
            seq_field = _find_direct_seq_field(paragraph, "Figure")
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
        for anchor_name in _candidate_anchor_names(label_key):
            mapping[anchor_name] = number_bookmark

    return mapping, added, labels_bound, labels_missing_seq


def _convert_figure_hyperlinks_to_ref_fields(
    document_root: ET.Element,
    *,
    anchor_to_number_bookmark: dict[str, str],
) -> tuple[int, int, int]:
    """
    将文内 figure 引用从 hyperlink(anchor) 转换为 REF 字段。
    返回：
    - converted_link_count
    - candidate_link_count
    - candidate_anchor_name_count
    """
    converted = 0
    candidate = 0
    candidate_anchor_names: set[str] = set()

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
            display_text = "".join((node.text or "") for node in child.findall(".//w:t", NS))
            ref_field = _make_ref_field(bookmark_name, display_text)

            insertion_index = list(paragraph).index(child)
            paragraph.remove(child)
            paragraph.insert(insertion_index, ref_field)
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

    body = document_root.find("./w:body", NS)
    if body is None:
        raise RuntimeError("document.xml 缺少 w:body。")

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

    figure_label_pairs = list(zip(figure_caption_paragraphs, inventory.figure_labels))
    table_label_pairs = list(zip(table_caption_paragraphs, inventory.table_labels))
    equation_label_pairs = list(zip(display_equation_paragraphs, inventory.equation_labels))

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

    equation_seq_added = 0
    for index, (paragraph, _label) in enumerate(equation_label_pairs, start=1):
        if _append_equation_seq(paragraph, seq_name="Eq", sequence_index=index):
            equation_seq_added += 1

    bookmark_added_equation, equation_label_repaired = _repair_label_bookmarks(
        label_pairs=equation_label_pairs,
        missing_anchors=missing_anchors,
        existing_bookmarks=existing_bookmarks,
        bookmark_id_seed=next_bookmark_id,
    )
    next_bookmark_id += bookmark_added_equation

    figure_anchor_name_all = {
        anchor_name
        for label in inventory.figure_labels
        for anchor_name in _candidate_anchor_names(label)
    }
    figure_ref_mapping, figure_num_bookmark_added, figure_labels_bound, figure_labels_missing_num = (
        _build_figure_ref_bookmark_mapping(
            figure_label_pairs=figure_label_pairs,
            existing_bookmarks=existing_bookmarks,
            bookmark_id_seed=next_bookmark_id,
        )
    )
    next_bookmark_id += figure_num_bookmark_added

    figure_hyperlink_anchors_in_doc = anchors & figure_anchor_name_all
    figure_anchor_unmapped = sorted(
        anchor_name for anchor_name in figure_hyperlink_anchors_in_doc if anchor_name not in figure_ref_mapping
    )
    figure_ref_converted, figure_ref_candidate_links, figure_ref_candidate_anchor_count = (
        _convert_figure_hyperlinks_to_ref_fields(
            document_root,
            anchor_to_number_bookmark=figure_ref_mapping,
        )
    )

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
        bookmark_added_figure
        + bookmark_added_table
        + bookmark_added_equation
        + bookmark_added_fuzzy
    )
    modified = any(
        value > 0
        for value in (
            figure_seq_added,
            table_seq_added,
            equation_seq_added,
            bookmark_added_total,
        )
    )

    metrics = {
        "figure_caption_para_count": len(figure_caption_paragraphs),
        "table_caption_para_count": len(table_caption_paragraphs),
        "display_equation_para_count": len(display_equation_paragraphs),
        "figure_label_count_from_tex": len(inventory.figure_labels),
        "table_label_count_from_tex": len(inventory.table_labels),
        "equation_label_count_from_tex": len(inventory.equation_labels),
        "figure_caption_seq_added": figure_seq_added,
        "table_caption_seq_added": table_seq_added,
        "equation_seq_added": equation_seq_added,
        "bookmark_added_total": bookmark_added_total,
        "bookmark_added_figure": bookmark_added_figure,
        "bookmark_added_table": bookmark_added_table,
        "bookmark_added_equation": bookmark_added_equation,
        "bookmark_added_fuzzy": bookmark_added_fuzzy,
        "figure_num_bookmark_added": figure_num_bookmark_added,
        "figure_ref_mapping_anchor_count": len(figure_ref_mapping),
        "figure_ref_candidate_link_count": figure_ref_candidate_links,
        "figure_ref_candidate_anchor_count": figure_ref_candidate_anchor_count,
        "figure_ref_converted_link_count": figure_ref_converted,
        "figure_ref_unmapped_anchor_count": len(figure_anchor_unmapped),
        "missing_anchor_count_after": len(missing_anchors),
    }

    details = {
        "figure_labels_repaired": figure_label_repaired,
        "table_labels_repaired": table_label_repaired,
        "equation_labels_repaired": equation_label_repaired,
        "figure_labels_bound_to_ref_bookmark": figure_labels_bound,
        "figure_labels_missing_number_bookmark": figure_labels_missing_num,
        "figure_ref_unmapped_anchors": figure_anchor_unmapped[:50],
        "fuzzy_anchor_repairs": fuzzy_records,
        "missing_anchors_sample_after": sorted(missing_anchors)[:50],
    }

    if len(inventory.figure_labels) > len(figure_caption_paragraphs):
        warnings.append(
            "figure 标签数量大于 docx 可识别图题段落数量，部分 figure 书签无法自动修复。"
        )
    if len(inventory.table_labels) > len(table_caption_paragraphs):
        warnings.append(
            "table 标签数量大于 docx 可识别表题段落数量，部分 table 书签无法自动修复。"
        )
    if len(inventory.equation_labels) > len(display_equation_paragraphs):
        warnings.append(
            "equation 标签数量大于 docx 可识别显示公式段落数量，部分公式书签无法自动修复。"
        )
    if figure_labels_missing_num:
        warnings.append(
            f"{len(figure_labels_missing_num)} 个 figure 标签未能绑定图号书签，相关图引用无法升级为 REF 字段。"
        )
    if figure_ref_candidate_links > 0 and figure_ref_converted < figure_ref_candidate_links:
        warnings.append(
            "部分 figure 超链接未成功转换为 REF 字段，请人工核对图引用。"
        )
    if figure_anchor_unmapped:
        warnings.append(
            f"检测到 {len(figure_anchor_unmapped)} 个 figure 风格锚点未映射到图题，已保留原超链接。"
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
