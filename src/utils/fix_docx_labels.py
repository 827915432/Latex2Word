#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_docx_dotted_labels.py

用途：
- 修复已经生成的 .docx 中，由标签名包含 "." 等不稳定字符引起的内部锚点问题。
- 将 Word 侧对应的 bookmark / hyperlink anchor / field target 名称统一改成下划线风格。

它修复的对象不是 LaTeX 源码本身，而是 docx 内部这些“标签等价物”：
- w:bookmarkStart @w:name
- w:hyperlink @w:anchor
- 域代码中的 REF / PAGEREF / NOTEREF / GOTOBUTTON / HYPERLINK \\l 目标名

典型示例：
- SEC7.1   -> SEC7_1
- EQ3.1    -> EQ3_1
- fig3.2.1a -> fig3_2_1a

用法：
    python fix_docx_dotted_labels.py input.docx -o output_fixed.docx
    python fix_docx_dotted_labels.py input.docx --in-place
    python fix_docx_dotted_labels.py input.docx --dry-run
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import OrderedDict
from pathlib import Path
from typing import Dict, Iterable, Tuple

BOOKMARK_NAME_RE = re.compile(r'w:bookmarkStart\b[^>]*\bw:name="([^"]+)"')
ANCHOR_ATTR_RE_TEMPLATE = r'(\bw:anchor="){}(")'
BOOKMARK_ATTR_RE_TEMPLATE = r'(\bw:name="){}(")'

# Word 字段里常见的内部目标
FIELD_PATTERNS = [
    r'(\bREF\s+){}(?=(?:\s|\\|"))',
    r'(\bPAGEREF\s+){}(?=(?:\s|\\|"))',
    r'(\bNOTEREF\s+){}(?=(?:\s|\\|"))',
    r'(\bGOTOBUTTON\s+){}(?=(?:\s|\\|"))',
    r'(HYPERLINK\s+\\l\s+"){}(?=")',
]

XML_FILE_RE = re.compile(r'.+\.(xml|rels)$', re.IGNORECASE)


def sanitize_bookmark_name(name: str) -> str:
    """
    将不稳定的书签/锚点名规范化：
    - 非 [A-Za-z0-9_] 全部替换为 _
    - 合并连续下划线
    - 若首字符不是字母，则补前缀 bk_
    """
    new = re.sub(r'[^A-Za-z0-9_]', '_', name)
    new = re.sub(r'_+', '_', new).strip('_')
    if not new:
        new = 'bk'
    if not re.match(r'^[A-Za-z]', new):
        new = f'bk_{new}'
    return new


def uniquify_mapping(names: Iterable[str]) -> Dict[str, str]:
    """
    对所有待修复名字构造 old -> new 映射，并处理重名冲突。
    """
    mapping: Dict[str, str] = OrderedDict()
    used = set()

    for old in names:
        base = sanitize_bookmark_name(old)
        new = base
        idx = 1
        while new in used:
            idx += 1
            new = f"{base}_{idx}"
        mapping[old] = new
        used.add(new)

    return mapping


def collect_dotted_bookmarks(zip_path: Path) -> Dict[str, str]:
    found = OrderedDict()
    with zipfile.ZipFile(zip_path, 'r') as zin:
        for name in zin.namelist():
            if not XML_FILE_RE.fullmatch(name):
                continue
            try:
                text = zin.read(name).decode('utf-8')
            except UnicodeDecodeError:
                continue
            for bm in BOOKMARK_NAME_RE.findall(text):
                if '.' in bm or re.search(r'[^A-Za-z0-9_]', bm):
                    found.setdefault(bm, None)

    return uniquify_mapping(found.keys())


def apply_mapping_to_xml_text(text: str, mapping: Dict[str, str]) -> Tuple[str, int]:
    """
    只做“定点替换”，避免重写整个 XML 结构。
    """
    count = 0
    out = text

    for old, new in mapping.items():
        if old == new:
            continue
        old_esc = re.escape(old)

        # 1) bookmarkStart @w:name
        out, n = re.subn(BOOKMARK_ATTR_RE_TEMPLATE.format(old_esc), rf'\1{new}\2', out)
        count += n

        # 2) hyperlink @w:anchor
        out, n = re.subn(ANCHOR_ATTR_RE_TEMPLATE.format(old_esc), rf'\1{new}\2', out)
        count += n

        # 3) 域代码中的目标名
        for patt in FIELD_PATTERNS:
            out, n = re.subn(patt.format(old_esc), rf'\1{new}', out)
            count += n

    return out, count


def process_docx(input_path: Path, output_path: Path, dry_run: bool = False) -> dict:
    mapping = collect_dotted_bookmarks(input_path)

    report = {
        "input": str(input_path),
        "output": str(output_path),
        "mapping": mapping,
        "files_touched": [],
        "total_replacements": 0,
    }

    if not mapping:
        return report

    if dry_run:
        return report

    with zipfile.ZipFile(input_path, 'r') as zin:
        members = zin.namelist()
        payload = {name: zin.read(name) for name in members}

    for name in members:
        if not XML_FILE_RE.fullmatch(name):
            continue
        try:
            text = payload[name].decode('utf-8')
        except UnicodeDecodeError:
            continue

        new_text, n = apply_mapping_to_xml_text(text, mapping)
        if n > 0:
            payload[name] = new_text.encode('utf-8')
            report["files_touched"].append({"name": name, "replacements": n})
            report["total_replacements"] += n

    with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name in members:
            zout.writestr(name, payload[name])

    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Fix dotted labels/bookmarks inside a DOCX.")
    parser.add_argument("input_docx", help="Input .docx path")
    parser.add_argument("-o", "--output", default=None, help="Output .docx path")
    parser.add_argument("--in-place", action="store_true", help="Overwrite input file in place")
    parser.add_argument("--dry-run", action="store_true", help="Only print planned replacements")
    parser.add_argument("--report-json", default=None, help="Optional report JSON path")
    args = parser.parse_args()

    input_path = Path(args.input_docx).resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise SystemExit("Input must be a .docx file")

    if args.in_place and args.output:
        raise SystemExit("Use either --in-place or -o/--output, not both.")

    if args.in_place:
        output_path = input_path
        work_input = input_path
        temp_output = input_path.with_name(input_path.stem + ".fixed_tmp.docx")
    else:
        output_path = Path(args.output).resolve() if args.output else input_path.with_name(input_path.stem + "_labels_fixed.docx")
        work_input = input_path
        temp_output = output_path

    report = process_docx(work_input, temp_output, dry_run=args.dry_run)

    if args.in_place and not args.dry_run and temp_output.exists():
        temp_output.replace(output_path)

    if args.report_json:
        report_path = Path(args.report_json).resolve()
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
