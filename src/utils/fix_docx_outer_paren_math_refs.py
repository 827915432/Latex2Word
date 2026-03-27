#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
fix_docx_outer_paren_math_refs_v3.py

修复 DOCX 中公式引用出现双重括号的问题，例如：
    ((11))
其真实结构通常是：
    w:t "("
    { GOTOBUTTON ZEqnNum...  ...  [fldSimple/complex REF result already equals "(11)"] }
    w:t ")"

本脚本只删除字段外层那一对普通文本括号，保留中间字段不变。
不会把引用改成纯文本。

适配：
- 复杂域起点为 w:r/w:fldChar begin
- 字段代码包含 GOTOBUTTON 与 ZEqnNum
- 字段结果可能在：
    - 内部 fldSimple 的 w:t 中，如 "(11)"
    - 内部复杂域的 w:instrText / w:t 中
- 右括号前允许夹杂 bookmarkStart / bookmarkEnd

用法：
    python fix_docx_outer_paren_math_refs_v3.py input.docx -o output.docx
    python fix_docx_outer_paren_math_refs_v3.py input.docx --in-place
    python fix_docx_outer_paren_math_refs_v3.py input.docx --dry-run
"""

from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def W(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def run_text(run: etree._Element) -> str:
    return "".join(node.text or "" for node in run.findall("./w:t", namespaces=NS))


def instr_text(run: etree._Element) -> str:
    return "".join(node.text or "" for node in run.findall("./w:instrText", namespaces=NS))


def set_run_text(run: etree._Element, text: str) -> bool:
    text_nodes = run.findall("./w:t", namespaces=NS)
    if not text_nodes:
        return False
    changed = False
    for i, node in enumerate(text_nodes):
        target = text if i == 0 else ""
        if (node.text or "") != target:
            node.text = target
            changed = True
    return changed


def remove_or_trim_paren_run(run: etree._Element, side: str) -> bool:
    """
    side == "left": 删除 run 文本里最后一个 "("
    side == "right": 删除 run 文本里第一个 ")"
    """
    text = run_text(run)
    if not text:
        return False

    if side == "left":
        idx = text.rfind("(")
        if idx < 0:
            return False
        new_text = text[:idx] + text[idx + 1 :]
    elif side == "right":
        idx = text.find(")")
        if idx < 0:
            return False
        new_text = text[:idx] + text[idx + 1 :]
    else:
        return False

    parent = run.getparent()
    if new_text == "":
        if parent is not None:
            parent.remove(run)
            return True
        return False

    return set_run_text(run, new_text)


def get_field_char_type(run: etree._Element) -> str:
    fld = run.find("./w:fldChar", namespaces=NS)
    if fld is None:
        return ""
    return fld.get(W("fldCharType"), "")


def complex_field_span(children: List[etree._Element], start_idx: int) -> Optional[int]:
    """
    给定段落直属 child 中某个 begin 位置，返回匹配 end 的索引。
    """
    if start_idx >= len(children):
        return None

    node = children[start_idx]
    if node.tag != W("r") or get_field_char_type(node) != "begin":
        return None

    depth = 0
    for i in range(start_idx, len(children)):
        cur = children[i]
        if cur.tag != W("r"):
            continue

        fld_type = get_field_char_type(cur)
        if fld_type == "begin":
            depth += 1
        elif fld_type == "end":
            depth -= 1
            if depth == 0:
                return i
    return None


def collect_all_visible_text(node: etree._Element) -> str:
    texts = []
    for t in node.findall(".//w:t", namespaces=NS):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def collect_all_instr_text(node: etree._Element) -> str:
    texts = []
    for t in node.findall(".//w:instrText", namespaces=NS):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def field_block_looks_like_math_ref(
    children: List[etree._Element], start_idx: int, end_idx: int
) -> Tuple[bool, str]:
    """
    判断该复杂域是否是公式引用块：
        GOTOBUTTON + ZEqnNum
    并收集字段内部所有文本（包括子 fldSimple）。
    """
    instr_parts: List[str] = []
    visible_parts: List[str] = []

    for i in range(start_idx, end_idx + 1):
        cur = children[i]

        # 当前节点自身与其内部所有后代
        itxt = collect_all_instr_text(cur)
        if itxt:
            instr_parts.append(itxt)

        ttxt = collect_all_visible_text(cur)
        if ttxt:
            visible_parts.append(ttxt)

    instr_joined = " ".join(instr_parts)
    is_math_ref = (
        "GOTOBUTTON" in instr_joined and
        "ZEqnNum" in instr_joined
    )

    field_result = "".join(instr_parts + visible_parts)
    return is_math_ref, field_result


def field_result_has_inner_parentheses(
    children: List[etree._Element], start_idx: int, end_idx: int
) -> bool:
    """
    判断字段内部结果是否已经自带括号。
    注意：这里会扫描整个字段块内所有后代节点，
    包括嵌套的 fldSimple / w:t / w:instrText。
    """
    has_l = False
    has_r = False

    for i in range(start_idx, end_idx + 1):
        cur = children[i]
        txt = collect_all_instr_text(cur) + collect_all_visible_text(cur)
        if "(" in txt:
            has_l = True
        if ")" in txt:
            has_r = True

    return has_l and has_r


def find_prev_left_paren_run(children: List[etree._Element], before_idx: int) -> Optional[int]:
    """
    向左找最近的可见左括号 run。
    允许跳过 bookmarkStart / bookmarkEnd。
    遇到普通可见文本且不是括号，就停止。
    """
    for i in range(before_idx - 1, -1, -1):
        cur = children[i]

        if cur.tag in {W("bookmarkStart"), W("bookmarkEnd")}:
            continue

        if cur.tag != W("r"):
            continue

        txt = run_text(cur)
        if not txt:
            continue

        if "(" in txt:
            return i
        return None

    return None


def find_next_right_paren_run(children: List[etree._Element], after_idx: int) -> Optional[int]:
    """
    向右找最近的可见右括号 run。
    允许跳过 bookmarkStart / bookmarkEnd。
    遇到普通可见文本且不是括号，就停止。
    """
    for i in range(after_idx + 1, len(children)):
        cur = children[i]

        if cur.tag in {W("bookmarkStart"), W("bookmarkEnd")}:
            continue

        if cur.tag != W("r"):
            continue

        txt = run_text(cur)
        if not txt:
            continue

        if ")" in txt:
            return i
        return None

    return None


def fix_paragraph(paragraph: etree._Element) -> int:
    """
    找到：
        "(" + [GOTOBUTTON/ZEqnNum 字段块，且字段内部已有括号] + ")"
    然后删掉外层这一对括号。
    """
    fixes = 0

    while True:
        children = list(paragraph)
        changed = False

        for i, node in enumerate(children):
            if node.tag != W("r") or get_field_char_type(node) != "begin":
                continue

            end_idx = complex_field_span(children, i)
            if end_idx is None:
                continue

            is_math_ref, _field_result = field_block_looks_like_math_ref(children, i, end_idx)
            if not is_math_ref:
                continue

            if not field_result_has_inner_parentheses(children, i, end_idx):
                continue

            left_idx = find_prev_left_paren_run(children, i)
            right_idx = find_next_right_paren_run(children, end_idx)

            if left_idx is None or right_idx is None:
                continue

            left_ok = remove_or_trim_paren_run(children[left_idx], "left")

            # 结构可能变化，重新找右括号
            children2 = list(paragraph)

            right_node = None
            for j in range(max(0, right_idx - 3), len(children2)):
                cur = children2[j]
                if cur.tag == W("r") and ")" in run_text(cur):
                    right_node = cur
                    break

            if right_node is None:
                for cur in children2:
                    if cur.tag == W("r") and ")" in run_text(cur):
                        right_node = cur
                        break

            right_ok = False
            if right_node is not None:
                right_ok = remove_or_trim_paren_run(right_node, "right")

            if left_ok or right_ok:
                fixes += 1
                changed = True
                break

        if not changed:
            break

    return fixes


def process_xml(xml_bytes: bytes) -> Tuple[bytes, int]:
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    root = etree.fromstring(xml_bytes, parser=parser)

    fixed = 0
    for p in root.findall(".//w:p", namespaces=NS):
        fixed += fix_paragraph(p)

    out = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=None,
    )
    return out, fixed


def process_docx(input_path: Path, output_path: Path, dry_run: bool = False) -> dict:
    report = {
        "input": str(input_path),
        "output": str(output_path),
        "files_touched": [],
        "paragraphs_fixed": 0,
    }

    with zipfile.ZipFile(input_path, "r") as zin:
        members = zin.namelist()
        payload = {name: zin.read(name) for name in members}

    for name in members:
        if not (name.startswith("word/") and name.endswith(".xml")):
            continue

        try:
            new_bytes, count = process_xml(payload[name])
        except Exception:
            continue

        if count:
            report["files_touched"].append({
                "name": name,
                "paragraphs_fixed": count,
            })
            report["paragraphs_fixed"] += count
            if not dry_run:
                payload[name] = new_bytes

    if not dry_run:
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in members:
                zout.writestr(name, payload[name])

    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Fix outer parentheses around MathType equation reference fields in DOCX."
    )
    parser.add_argument("input_docx", help="Input .docx path")
    parser.add_argument("-o", "--output", default=None, help="Output .docx path")
    parser.add_argument("--in-place", action="store_true", help="Overwrite input file in place")
    parser.add_argument("--dry-run", action="store_true", help="Scan only; do not write output")
    parser.add_argument("--report-json", default=None, help="Optional JSON report path")
    args = parser.parse_args()

    input_path = Path(args.input_docx).resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise SystemExit("Input must be a .docx file")

    if args.in_place and args.output:
        raise SystemExit("Use either --in-place or -o/--output, not both.")

    if args.in_place:
        temp_output = input_path.with_name(input_path.stem + ".outerparen_tmp.docx")
        output_path = temp_output
        final_path = input_path
    else:
        output_path = (
            Path(args.output).resolve()
            if args.output
            else input_path.with_name(input_path.stem + "_outerparen_fixed.docx")
        )
        final_path = output_path

    report = process_docx(input_path, output_path, dry_run=args.dry_run)

    if args.in_place and not args.dry_run and output_path.exists():
        output_path.replace(final_path)

    if args.report_json:
        Path(args.report_json).resolve().write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
