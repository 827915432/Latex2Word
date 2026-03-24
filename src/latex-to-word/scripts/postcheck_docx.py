#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
postcheck_docx.py

功能概述
--------
本脚本用于在 Pandoc 主转换完成之后，对生成的 Word 文档（.docx）执行“结构级后检查”。

它解决的问题包括：
1. 检查 .docx 文件是否存在、是否可打开、是否包含基本的 Word XML 结构；
2. 检查图片、表格、标题、公式、书签、内部链接、字段等结构是否存在；
3. 结合前面阶段的报告（precheck / normalization / pandoc conversion），
   对“期望有但实际缺失”的对象给出明确告警或错误；
4. 为后续 build_manual_fix_list.py 提供结构化输入；
5. 给出 PASS / PASS_WITH_WARNINGS / FAIL 的明确状态。

设计边界
--------
本脚本只做“检查”，不做以下事情：
- 不修改 .docx；
- 不修改源工程；
- 不执行任何 Word 自动化；
- 不替代人工校对；
- 不生成最终人工修复清单（那是 build_manual_fix_list.py 的职责）。

依赖
----
- Python 3.9+
- 仅使用标准库，不依赖第三方包

典型用法
--------
python scripts/postcheck_docx.py --work-root D:/work/my-paper__latex_to_word_work

或显式指定 docx：
python scripts/postcheck_docx.py \
  --work-root D:/work/my-paper__latex_to_word_work \
  --docx D:/work/my-paper__latex_to_word_work/stage_convert/output.docx

输出
----
默认会在 work-root 的 `stage_postcheck/` 下生成：
- postcheck-report.json
- postcheck-report.md

退出码
------
- 0: PASS 或 PASS_WITH_WARNINGS
- 1: FAIL
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from pipeline_common import (
    load_json_if_exists,
    locate_skill_root,
    read_text_file,
    safe_relative,
)
from pipeline_constants import (
    REQUIRED_RULE_FILES,
    SEVERITY_ERROR,
    SEVERITY_INFO,
    SEVERITY_ORDER,
    SEVERITY_WARN,
    STATUS_FAIL,
    STATUS_PASS,
    STATUS_PASS_WITH_WARNINGS,
)
from pipeline_layout import (
    STAGE_CONVERT,
    STAGE_NORMALIZE,
    STAGE_POSTCHECK,
    STAGE_PRECHECK,
    resolve_explicit_or_stage_input,
    resolve_explicit_or_stage_output,
    stage_dir,
)
from stage_reporting import persist_stage_report
from tex_scan_common import strip_latex_comments


# -----------------------------------------------------------------------------
# 常量定义
# -----------------------------------------------------------------------------

# WordprocessingML / OMML / DrawingML 相关命名空间。
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
}

# -----------------------------------------------------------------------------
# 数据结构定义
# -----------------------------------------------------------------------------

@dataclass
class Finding:
    """
    结构化问题记录。

    字段说明
    --------
    severity:
        INFO / WARN / ERROR

    code:
        稳定、机器可读的问题代码，用于后续脚本筛选和汇总。

    message:
        面向人的问题描述。

    location:
        问题位置。对 docx 来说通常不是“文件:行号”，而是：
        - 文档级位置，例如 "word/document.xml"
        - 或逻辑位置，例如 "document body"

    details:
        结构化附加信息。
    """
    severity: str
    code: str
    message: str
    location: Optional[str] = None
    details: dict = field(default_factory=dict)


@dataclass
class PostcheckReport:
    """
    docx 后检查报告对象。

    设计目标：
    - 既可 JSON 结构化消费，也适合 Markdown 人读；
    - 明确区分 source 侧期望值与 docx 侧实际值；
    - 明确列出状态、问题和下一步建议。
    """
    status: str
    can_continue: bool
    work_root: str
    source_project_root: Optional[str]
    input_docx: str
    used_conversion_report: bool
    used_normalization_report: bool
    used_precheck_report: bool
    findings: list[dict]
    source_inventory: dict
    docx_inventory: dict
    metrics: dict
    summary: dict
    recommendations: list[str]


# -----------------------------------------------------------------------------
# 参数与通用工具函数
# -----------------------------------------------------------------------------

def build_argument_parser() -> argparse.ArgumentParser:
    """
    构造命令行参数解析器。

    参数设计原则：
    - work-root 为必需，因为这一步默认面向“规范化工作目录”运转；
    - docx / conversion / normalization / precheck 报告允许显式指定；
    - 若不指定，则按 skill 流水线默认文件名推断。
    """
    parser = argparse.ArgumentParser(
        description="Perform structural postcheck for a generated DOCX file."
    )
    parser.add_argument(
        "--work-root",
        required=True,
        help="规范化工作目录。",
    )
    parser.add_argument(
        "--docx",
        default=None,
        help="待检查的 docx 文件路径；默认优先 <work-root>/stage_convert/output.docx，再回退 <work-root>/output.docx",
    )
    parser.add_argument(
        "--conversion-json",
        default=None,
        help="Pandoc 主转换报告 JSON 路径；默认优先 <work-root>/stage_convert/pandoc-conversion-report.json，再回退 <work-root>/pandoc-conversion-report.json",
    )
    parser.add_argument(
        "--normalization-json",
        default=None,
        help="规范化报告 JSON 路径；默认优先 <work-root>/stage_normalize/normalization-report.json，再回退 <work-root>/normalization-report.json",
    )
    parser.add_argument(
        "--precheck-json",
        default=None,
        help="预检查报告 JSON 路径；默认优先 <work-root>/stage_precheck/precheck-report.json，再回退 <work-root>/precheck-report.json，若不存在则尝试从 normalization-report.json 中的 project_root 寻找。",
    )
    parser.add_argument(
        "--json-out",
        default=None,
        help="后检查 JSON 报告输出路径；默认 <work-root>/stage_postcheck/postcheck-report.json",
    )
    parser.add_argument(
        "--md-out",
        default=None,
        help="后检查 Markdown 报告输出路径；默认 <work-root>/stage_postcheck/postcheck-report.md",
    )
    return parser


def check_rule_files(skill_root: Path) -> list[Finding]:
    """
    检查规则文件是否存在。

    postcheck 阶段不会解析规则语义，但若 skill 目录缺失规则文件，应尽早通过报告暴露。
    """
    findings: list[Finding] = []
    for relative in REQUIRED_RULE_FILES:
        candidate = (skill_root / relative).resolve()
        if not candidate.exists():
            findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="RULE_FILE_MISSING",
                    message="缺少规则文件；后续状态判断的可解释性会下降。",
                    location=safe_relative(candidate, skill_root),
                    details={"required_file": relative},
                )
            )
    return findings


# -----------------------------------------------------------------------------
# 源侧 inventory 统计
# -----------------------------------------------------------------------------

def collect_source_inventory(work_root: Path, normalization_report: Optional[dict]) -> dict:
    """
    从“规范化后的 TeX 文件”中统计一个近似的 source inventory。

    这样做的目的：
    - postcheck 不仅要看 docx 里“有什么”，还要看“原本预期应该有什么”；
    - 规范化工作副本比原始工程更适合作为后续 तुलना基准；
    - 若 normalization-report.json 存在，则优先用其中的 tex_files_processed 作为统计范围。

    统计内容（近似值）：
    - includegraphics 数量
    - figure 环境数量
    - table / longtable 数量
    - caption 数量
    - heading 命令数量
    - equation 类环境数量
    - label / ref / cite 数量
    """
    inventory = {
        "tex_file_count": 0,
        "image_command_count": 0,
        "figure_env_count": 0,
        "table_env_count": 0,
        "caption_command_count": 0,
        "heading_command_count": 0,
        "equation_env_count": 0,
        "label_count": 0,
        "ref_count": 0,
        "cite_count": 0,
    }

    normalized_source_root = work_root.resolve()
    if normalization_report and isinstance(normalization_report.get("summary"), dict):
        value = normalization_report["summary"].get("normalized_source_root")
        if isinstance(value, str) and value.strip():
            candidate_root = Path(value).resolve()
            if candidate_root.exists() and candidate_root.is_dir():
                normalized_source_root = candidate_root

    tex_files: list[Path] = []
    if normalization_report and isinstance(normalization_report.get("tex_files_processed"), list):
        for relative in normalization_report["tex_files_processed"]:
            rel_path = Path(str(relative))
            candidates = []
            if rel_path.is_absolute():
                candidates.append(rel_path.resolve())
            else:
                candidates.append((normalized_source_root / rel_path).resolve())
                candidates.append((work_root / rel_path).resolve())
            for candidate in candidates:
                if candidate.exists() and candidate.suffix.lower() == ".tex":
                    tex_files.append(candidate)
                    break

    if not tex_files:
        tex_files = sorted(normalized_source_root.rglob("*.tex"))
    if not tex_files:
        tex_files = sorted(work_root.rglob("*.tex"))

    inventory["tex_file_count"] = len(tex_files)

    if not tex_files:
        return inventory

    # 源侧统计用正则集合。
    includegraphics_pattern = re.compile(r"\\includegraphics(?:\[[^\]]*\])?\{[^}]+\}")
    figure_env_pattern = re.compile(r"\\begin\{figure\*?\}")
    table_env_pattern = re.compile(r"\\begin\{(?:table\*?|longtable)\}")
    caption_pattern = re.compile(r"\\caption(?:\[[^\]]*\])?\{")
    heading_pattern = re.compile(r"\\(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\b")
    equation_env_pattern = re.compile(
        r"\\begin\{(?:equation\*?|align\*?|gather\*?|multline\*?|split|cases|matrix|pmatrix|bmatrix|vmatrix|Vmatrix)\}"
    )
    display_math_bracket_pattern = re.compile(r"\\\[")
    label_pattern = re.compile(r"\\label\{[^}]+\}")
    ref_pattern = re.compile(r"\\(?:ref|eqref|pageref|autoref|cref|Cref|nameref)\{[^}]+\}")
    cite_pattern = re.compile(
        r"\\(?:cite|citep|citet|parencite|textcite|autocite|footcite|supercite)\*?(?:\[[^\]]*\]){0,2}\{[^}]+\}"
    )

    for tex_file in tex_files:
        try:
            text = read_text_file(tex_file)
        except Exception:
            continue
        text = strip_latex_comments(text)

        inventory["image_command_count"] += len(includegraphics_pattern.findall(text))
        inventory["figure_env_count"] += len(figure_env_pattern.findall(text))
        inventory["table_env_count"] += len(table_env_pattern.findall(text))
        inventory["caption_command_count"] += len(caption_pattern.findall(text))
        inventory["heading_command_count"] += len(heading_pattern.findall(text))
        inventory["equation_env_count"] += len(equation_env_pattern.findall(text))
        inventory["equation_env_count"] += len(display_math_bracket_pattern.findall(text))
        inventory["label_count"] += len(label_pattern.findall(text))
        inventory["ref_count"] += len(ref_pattern.findall(text))
        inventory["cite_count"] += len(cite_pattern.findall(text))

    return inventory


# -----------------------------------------------------------------------------
# docx XML 检查工具
# -----------------------------------------------------------------------------

def parse_xml_from_zip(docx_zip: zipfile.ZipFile, member_name: str) -> Optional[ET.Element]:
    """
    从 docx zip 中读取并解析一个 XML 成员。

    若成员不存在或 XML 不可解析，则返回 None。
    """
    try:
        raw = docx_zip.read(member_name)
    except KeyError:
        return None

    try:
        return ET.fromstring(raw)
    except ET.ParseError:
        return None


def get_style_maps(styles_root: Optional[ET.Element]) -> tuple[dict[str, str], set[str], set[str]]:
    """
    从 word/styles.xml 构造样式映射与关键样式集合。

    返回：
    1. style_id_to_name:
       例如 "Heading1" -> "heading 1" 或自定义名称
    2. heading_style_ids:
       被判定为标题样式的 styleId 集合
    3. caption_style_ids:
       被判定为图题/表题样式的 styleId 集合

    说明：
    - Word 内建样式的 styleId 通常是英语形式，例如 Heading1 / Caption；
    - 样式名称 name 可能因模板而不同，因此同时检查 styleId 和 name；
    - 这里不做更复杂的样式继承解析，因为后检查只需要结构级判断。
    """
    style_id_to_name: dict[str, str] = {}
    heading_style_ids: set[str] = set()
    caption_style_ids: set[str] = set()

    if styles_root is None:
        return style_id_to_name, heading_style_ids, caption_style_ids

    for style in styles_root.findall(".//w:style", NS):
        style_id = style.attrib.get(f"{{{NS['w']}}}styleId", "")
        name_elem = style.find("w:name", NS)
        style_name = name_elem.attrib.get(f"{{{NS['w']}}}val", "") if name_elem is not None else ""
        style_id_to_name[style_id] = style_name

        normalized_name = style_name.strip().lower()
        normalized_id = style_id.strip().lower()

        # 标题样式判定：
        # - styleId 以 heading 开头；
        # - 或样式名称以 heading / 标题 开头。
        if normalized_id.startswith("heading") or normalized_name.startswith("heading") or normalized_name.startswith("标题"):
            heading_style_ids.add(style_id)

        # Caption 样式判定：
        # - styleId 为 caption
        # - 或样式名称为 caption / 题注
        if normalized_id == "caption" or normalized_name == "caption" or normalized_name == "题注":
            caption_style_ids.add(style_id)

    return style_id_to_name, heading_style_ids, caption_style_ids


def get_paragraph_style_id(paragraph: ET.Element) -> Optional[str]:
    """
    获取段落的 pStyle styleId。

    若段落未显式设置样式，则返回 None。
    """
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return None
    p_style = p_pr.find("w:pStyle", NS)
    if p_style is None:
        return None
    return p_style.attrib.get(f"{{{NS['w']}}}val")


def get_paragraph_text(paragraph: ET.Element) -> str:
    """
    提取段落纯文本内容。

    只拼接 w:t 节点文本，不加入复杂格式语义。
    这对标题、图题、表题、参考文献段落的后检查足够。
    """
    texts: list[str] = []
    for text_node in paragraph.findall(".//w:t", NS):
        if text_node.text:
            texts.append(text_node.text)
    return "".join(texts).strip()


def is_heading_paragraph(paragraph: ET.Element, heading_style_ids: set[str], style_id_to_name: dict[str, str]) -> bool:
    """
    判断一个段落是否应视为标题段落。

    判定策略：
    1. pStyle 命中 heading_style_ids；
    2. pPr 中存在 outlineLvl；
    3. styleId / styleName 仍表现出 heading 语义。

    这比只看 styleId 更稳，因为 reference.docx 可能做了样式定制。
    """
    style_id = get_paragraph_style_id(paragraph)
    if style_id and style_id in heading_style_ids:
        return True

    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is not None and p_pr.find("w:outlineLvl", NS) is not None:
        return True

    if style_id:
        style_name = style_id_to_name.get(style_id, "").strip().lower()
        if style_id.lower().startswith("heading") or style_name.startswith("heading") or style_name.startswith("标题"):
            return True

    return False


def is_caption_paragraph(
    paragraph: ET.Element,
    caption_style_ids: set[str],
    style_id_to_name: dict[str, str],
) -> tuple[bool, str]:
    """
    判断一个段落是否应视为图题/表题候选。

    返回：
    - is_caption_candidate: bool
    - caption_kind: "figure" / "table" / "unknown" / "none"

    判定策略：
    1. 若段落样式为 Caption，则直接作为 caption 候选；
    2. 否则使用文本启发式：
       - Figure / Fig. / Table
       - 图 / 表
    """
    style_id = get_paragraph_style_id(paragraph)
    text = get_paragraph_text(paragraph)

    # 样式判定优先。
    if style_id and style_id in caption_style_ids:
        lowered = text.lower()
        if re.match(r"^\s*(table)\b", lowered) or re.match(r"^\s*表\s*\d*", text):
            return True, "table"
        if re.match(r"^\s*(figure|fig\.?)\b", lowered) or re.match(r"^\s*图\s*\d*", text):
            return True, "figure"
        return True, "unknown"

    # 文本启发式判定。
    lowered = text.lower()
    if re.match(r"^\s*(figure|fig\.?)\s*\d+", lowered) or re.match(r"^\s*图\s*\d+", text):
        return True, "figure"
    if re.match(r"^\s*(table)\s*\d+", lowered) or re.match(r"^\s*表\s*\d+", text):
        return True, "table"

    return False, "none"


def paragraph_contains_image(paragraph: ET.Element) -> bool:
    """
    判断段落中是否包含图片对象。

    这里只做结构级判定：
    - DrawingML: a:blip[@r:embed]
    - VML: v:imagedata[@r:id]
    """
    return bool(paragraph.findall(".//a:blip[@r:embed]", NS) or paragraph.findall(".//v:imagedata[@r:id]", NS))


def is_likely_proximity_caption_text(text: str) -> bool:
    """
    判断“图/表后紧邻段落文本”是否像题注。

    该判定用于补强 caption 检测，减少仅靠样式/前缀匹配带来的漏检。
    """
    candidate = text.strip()
    if not candidate:
        return False

    # 太长的段落通常不是题注。
    if len(candidate) > 120:
        return False

    # 标题样式文本、编号章节标题不应被视为题注。
    if re.match(r"^\s*(\d+(\.\d+)*)\s+", candidate):
        return False

    # 带明显句末标点的完整句子更可能是正文。
    if candidate.endswith(("。", "；", ";", "？", "?", "！", "!")):
        return False

    # 避免将明显参考文献条目误判为题注。
    if is_likely_bibliography_entry(candidate):
        return False

    return True


def is_likely_bibliography_entry(text: str) -> bool:
    """
    判断段落是否像“参考文献条目”。

    该判定用于“无明确标题但有条目”的推断，不追求严格格式学判断，
    目标是降低后检查阶段的误报。
    """
    candidate = text.strip()
    if not candidate:
        return False

    if re.match(r"^\s*\[\d+\]", candidate):
        return True

    if re.match(r"^\s*\d+\.\s+", candidate):
        return True

    if re.search(r"\b(19|20)\d{2}\b", candidate):
        # 年份 + 常见文献分隔符，是较强信号。
        if "," in candidate or "." in candidate or "“" in candidate or "\"" in candidate:
            return True

    if re.search(r"\bdoi\b", candidate, re.IGNORECASE):
        return True

    return False


def inspect_docx(docx_path: Path) -> tuple[dict, list[Finding]]:
    """
    检查 docx 文件结构并输出 inventory。

    检查内容包括：
    - .docx 是否可作为 zip 打开；
    - 是否存在 word/document.xml；
    - 图片嵌入文件数量；
    - 正文中的图片引用数量；
    - 表格数量；
    - 标题段落数量；
    - 图题/表题候选数量；
    - Word 数学对象（OMML）数量；
    - 书签、内部超链接、字段数量；
    - 参考文献章节候选情况。

    说明：
    - 这是一种“结构级检查”，不是视觉级渲染检查；
    - 它不能替代人工打开 Word 肉眼核对，但能显著缩小人工排查范围。
    """
    findings: list[Finding] = []

    inventory = {
        "docx_openable": False,
        "docx_file_size_bytes": docx_path.stat().st_size if docx_path.exists() else 0,
        "package_member_count": 0,
        "has_document_xml": False,
        "media_file_count": 0,
        "image_reference_count_in_body": 0,
        "table_count": 0,
        "paragraph_count": 0,
        "heading_count": 0,
        "title_like_count": 0,
        "caption_count_total": 0,
        "figure_caption_count": 0,
        "table_caption_count": 0,
        "unknown_caption_count": 0,
        "proximity_caption_count_total": 0,
        "proximity_figure_caption_count": 0,
        "proximity_table_caption_count": 0,
        "effective_caption_count_total": 0,
        "math_object_count": 0,
        "bookmark_count": 0,
        "internal_hyperlink_count": 0,
        "field_ref_count": 0,
        "field_pageref_count": 0,
        "field_seq_count": 0,
        "field_toc_count": 0,
        "bibliography_section_found": False,
        "bibliography_heading_text": None,
        "bibliography_following_paragraph_count": 0,
        "bibliography_entry_like_count": 0,
        "bibliography_inferred_by_entries": False,
        "bibliography_detection_mode": "none",
        "external_hyperlink_relation_count": 0,
        "styles_xml_found": False,
    }

    if not docx_path.exists():
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="DOCX_NOT_FOUND",
                message="待检查的 docx 文件不存在。",
                location=str(docx_path),
            )
        )
        return inventory, findings

    try:
        with zipfile.ZipFile(docx_path, "r") as docx_zip:
            inventory["docx_openable"] = True
            members = docx_zip.namelist()
            inventory["package_member_count"] = len(members)

            if "word/document.xml" not in members:
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="MISSING_DOCUMENT_XML",
                        message=".docx 中缺少 word/document.xml，文档结构不完整。",
                        location="word/document.xml",
                    )
                )
                return inventory, findings

            inventory["has_document_xml"] = True

            # 统计 media 文件数量。
            inventory["media_file_count"] = sum(1 for name in members if name.startswith("word/media/") and not name.endswith("/"))

            # styles.xml 与 relationships.xml 可选读取。
            styles_root = parse_xml_from_zip(docx_zip, "word/styles.xml")
            if styles_root is not None:
                inventory["styles_xml_found"] = True
            style_id_to_name, heading_style_ids, caption_style_ids = get_style_maps(styles_root)

            rels_root = parse_xml_from_zip(docx_zip, "word/_rels/document.xml.rels")
            if rels_root is not None:
                relationship_type_attr = "{http://schemas.openxmlformats.org/package/2006/relationships}Type"
                for rel in rels_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    rel_type = rel.attrib.get(relationship_type_attr, "")
                    if rel_type.endswith("/hyperlink"):
                        inventory["external_hyperlink_relation_count"] += 1

            document_root = parse_xml_from_zip(docx_zip, "word/document.xml")
            if document_root is None:
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="DOCUMENT_XML_PARSE_FAILED",
                        message="word/document.xml 无法解析为 XML。",
                        location="word/document.xml",
                    )
                )
                return inventory, findings

            # 统计图片引用：
            # - DrawingML 图片: a:blip[@r:embed]
            # - 旧式 VML 图片: v:imagedata[@r:id]
            inventory["image_reference_count_in_body"] = len(
                document_root.findall(".//a:blip[@r:embed]", NS)
            ) + len(
                document_root.findall(".//v:imagedata[@r:id]", NS)
            )

            # 统计表格与数学对象。
            inventory["table_count"] = len(document_root.findall(".//w:tbl", NS))
            inventory["math_object_count"] = len(document_root.findall(".//m:oMath", NS)) + len(
                document_root.findall(".//m:oMathPara", NS)
            )

            # 统计书签、内部链接、字段。
            inventory["bookmark_count"] = len(document_root.findall(".//w:bookmarkStart", NS))
            inventory["internal_hyperlink_count"] = len(document_root.findall(".//w:hyperlink[@w:anchor]", NS))

            instr_text_nodes = document_root.findall(".//w:instrText", NS)
            for node in instr_text_nodes:
                raw = (node.text or "").upper()
                if " REF " in f" {raw} " or raw.strip().startswith("REF "):
                    inventory["field_ref_count"] += 1
                if " PAGEREF " in f" {raw} " or raw.strip().startswith("PAGEREF "):
                    inventory["field_pageref_count"] += 1
                if " SEQ " in f" {raw} " or raw.strip().startswith("SEQ "):
                    inventory["field_seq_count"] += 1
                if " TOC " in f" {raw} " or raw.strip().startswith("TOC "):
                    inventory["field_toc_count"] += 1

            # 补充统计 simple field（w:fldSimple）中的字段指令：
            # Pandoc 与部分后处理器会直接写入 w:fldSimple@w:instr，
            # 不一定生成 w:instrText。
            fld_simple_nodes = document_root.findall(".//w:fldSimple", NS)
            instr_attr_name = f"{{{NS['w']}}}instr"
            for node in fld_simple_nodes:
                raw = (node.get(instr_attr_name, "") or "").upper()
                if " REF " in f" {raw} " or raw.strip().startswith("REF "):
                    inventory["field_ref_count"] += 1
                if " PAGEREF " in f" {raw} " or raw.strip().startswith("PAGEREF "):
                    inventory["field_pageref_count"] += 1
                if " SEQ " in f" {raw} " or raw.strip().startswith("SEQ "):
                    inventory["field_seq_count"] += 1
                if " TOC " in f" {raw} " or raw.strip().startswith("TOC "):
                    inventory["field_toc_count"] += 1

            # 段落级统计。
            body = document_root.find(".//w:body", NS)
            if body is None:
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="DOCX_MISSING_BODY_XML",
                        message="word/document.xml 中未找到 w:body。",
                        location="word/document.xml",
                    )
                )
                return inventory, findings

            body_children = list(body)
            paragraphs = [child for child in body_children if child.tag == f"{{{NS['w']}}}p"]
            inventory["paragraph_count"] = len(paragraphs)

            paragraph_texts: list[str] = []
            heading_flags: list[bool] = []
            child_idx_to_para_idx: dict[int, int] = {}
            para_idx_to_child_idx: dict[int, int] = {}
            para_idx_to_heading: dict[int, bool] = {}
            para_idx_to_explicit_caption: dict[int, tuple[bool, str]] = {}

            for child_idx, child in enumerate(body_children):
                if child.tag != f"{{{NS['w']}}}p":
                    continue

                paragraph = child
                para_idx = len(paragraph_texts)
                child_idx_to_para_idx[child_idx] = para_idx
                para_idx_to_child_idx[para_idx] = child_idx

                text = get_paragraph_text(paragraph)
                paragraph_texts.append(text)

                is_heading = is_heading_paragraph(paragraph, heading_style_ids, style_id_to_name)
                heading_flags.append(is_heading)
                para_idx_to_heading[para_idx] = is_heading
                if is_heading:
                    inventory["heading_count"] += 1

                # Title 样式单独粗略计数，便于诊断“只有标题没有章节”的情况。
                style_id = get_paragraph_style_id(paragraph)
                if style_id:
                    style_name = style_id_to_name.get(style_id, "").strip().lower()
                    if style_id.lower() == "title" or style_name == "title":
                        inventory["title_like_count"] += 1

                is_caption, caption_kind = is_caption_paragraph(paragraph, caption_style_ids, style_id_to_name)
                para_idx_to_explicit_caption[para_idx] = (is_caption, caption_kind)
                if is_caption:
                    inventory["caption_count_total"] += 1
                    if caption_kind == "figure":
                        inventory["figure_caption_count"] += 1
                    elif caption_kind == "table":
                        inventory["table_caption_count"] += 1
                    else:
                        inventory["unknown_caption_count"] += 1

            # 结构邻近检测（图/表后紧邻段落）：
            # 当段落样式和前缀都没命中 caption 时，邻近结构可以补充识别。
            proximity_para_indices: set[int] = set()
            for child_idx, child in enumerate(body_children):
                object_kind: Optional[str] = None
                if child.tag == f"{{{NS['w']}}}tbl":
                    object_kind = "table"
                elif child.tag == f"{{{NS['w']}}}p" and paragraph_contains_image(child):
                    object_kind = "figure"

                if object_kind is None:
                    continue

                candidate_para_idx: Optional[int] = None
                for next_idx in range(child_idx + 1, len(body_children)):
                    next_child = body_children[next_idx]
                    if next_child.tag != f"{{{NS['w']}}}p":
                        continue
                    next_para_idx = child_idx_to_para_idx.get(next_idx)
                    if next_para_idx is None:
                        continue
                    next_text = paragraph_texts[next_para_idx].strip()
                    if not next_text:
                        continue
                    candidate_para_idx = next_para_idx
                    break

                if candidate_para_idx is None:
                    continue

                if para_idx_to_heading.get(candidate_para_idx, False):
                    continue

                explicit_caption, _ = para_idx_to_explicit_caption.get(candidate_para_idx, (False, "none"))
                if explicit_caption:
                    continue

                candidate_text = paragraph_texts[candidate_para_idx]
                if not is_likely_proximity_caption_text(candidate_text):
                    continue

                if candidate_para_idx in proximity_para_indices:
                    continue

                proximity_para_indices.add(candidate_para_idx)
                inventory["proximity_caption_count_total"] += 1
                if object_kind == "figure":
                    inventory["proximity_figure_caption_count"] += 1
                elif object_kind == "table":
                    inventory["proximity_table_caption_count"] += 1

            inventory["effective_caption_count_total"] = (
                inventory["caption_count_total"] + inventory["proximity_caption_count_total"]
            )

            # 参考文献章节启发式检查。
            bibliography_heading_patterns = [
                re.compile(r"^\s*references\s*$", re.IGNORECASE),
                re.compile(r"^\s*bibliography\s*$", re.IGNORECASE),
                re.compile(r"^\s*参考文献\s*$"),
            ]

            bibliography_heading_index: Optional[int] = None
            bibliography_heading_text: Optional[str] = None

            for index, text in enumerate(paragraph_texts):
                if not text:
                    continue
                if any(pattern.match(text) for pattern in bibliography_heading_patterns):
                    bibliography_heading_index = index
                    bibliography_heading_text = text
                    break

            if bibliography_heading_index is not None:
                inventory["bibliography_section_found"] = True
                inventory["bibliography_detection_mode"] = "heading"
                inventory["bibliography_heading_text"] = bibliography_heading_text

                count = 0
                entry_like_count = 0
                for idx in range(bibliography_heading_index + 1, len(paragraph_texts)):
                    # 遇到下一条标题则停止，避免把后续章节算进 bibliography。
                    if heading_flags[idx]:
                        break
                    current_text = paragraph_texts[idx].strip()
                    if not current_text:
                        continue
                    count += 1
                    if is_likely_bibliography_entry(current_text):
                        entry_like_count += 1

                inventory["bibliography_following_paragraph_count"] = count
                inventory["bibliography_entry_like_count"] = entry_like_count
            else:
                # 无明确标题时，尝试根据文末条目结构推断参考文献区。
                last_heading_index = -1
                for idx, flag in enumerate(heading_flags):
                    if flag:
                        last_heading_index = idx

                start_idx = max(last_heading_index + 1, int(len(paragraph_texts) * 0.6))
                entry_like_count = 0
                for idx in range(start_idx, len(paragraph_texts)):
                    if heading_flags[idx]:
                        continue
                    current_text = paragraph_texts[idx].strip()
                    if not current_text:
                        continue
                    if is_likely_bibliography_entry(current_text):
                        entry_like_count += 1

                inventory["bibliography_entry_like_count"] = entry_like_count
                if entry_like_count >= 2:
                    inventory["bibliography_inferred_by_entries"] = True
                    inventory["bibliography_detection_mode"] = "entry_inference"

    except zipfile.BadZipFile:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="DOCX_BAD_ZIP",
                message=".docx 文件无法作为有效 ZIP 包打开，文档已损坏或格式错误。",
                location=str(docx_path),
            )
        )
        return inventory, findings

    except Exception as exc:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="DOCX_INSPECTION_EXCEPTION",
                message="检查 .docx 时发生未预期异常。",
                location=str(docx_path),
                details={"error": str(exc)},
            )
        )
        return inventory, findings

    return inventory, findings


# -----------------------------------------------------------------------------
# 报告比对与状态判定
# -----------------------------------------------------------------------------

def analyze_results(
    work_root: Path,
    docx_path: Path,
    source_project_root: Optional[Path],
    precheck_report: Optional[dict],
    normalization_report: Optional[dict],
    conversion_report: Optional[dict],
    source_inventory: dict,
    docx_inventory: dict,
    initial_findings: list[Finding],
) -> PostcheckReport:
    """
    结合源侧期望值与 docx 实际值，生成最终 postcheck 报告。

    设计原则：
    - 报错要尽量面向“缺失关键对象”；
    - 对于本质上只能部分自动验证的对象（如 caption、交叉引用），
      尽量用 WARN 而不是武断 FAIL；
    - 状态判定遵循 acceptance_criteria 的最低可用标准思想。
    """
    findings: list[Finding] = list(initial_findings)

    # -------------------------------------------------------------------------
    # 预备信息
    # -------------------------------------------------------------------------
    expected_image_count = int(source_inventory.get("image_command_count", 0))
    expected_table_count = int(source_inventory.get("table_env_count", 0))
    expected_heading_count = int(source_inventory.get("heading_command_count", 0))
    expected_equation_count = int(source_inventory.get("equation_env_count", 0))
    expected_caption_count = int(source_inventory.get("caption_command_count", 0))

    expected_ref_count = 0
    expected_cite_count = 0
    if precheck_report and isinstance(precheck_report.get("metrics"), dict):
        expected_ref_count = int(precheck_report["metrics"].get("ref_count", 0))
        expected_cite_count = int(precheck_report["metrics"].get("cite_count", 0))
    else:
        expected_ref_count = int(source_inventory.get("ref_count", 0))
        expected_cite_count = int(source_inventory.get("cite_count", 0))

    # -------------------------------------------------------------------------
    # 基础 docx 可用性检查
    # -------------------------------------------------------------------------
    if not docx_inventory.get("docx_openable", False):
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="DOCX_NOT_OPENABLE",
                message="生成的 docx 无法打开，不满足最低可用标准。",
                location=str(docx_path),
            )
        )

    if not docx_inventory.get("has_document_xml", False):
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="DOCX_MISSING_BODY_XML",
                message="docx 缺少主体文档 XML，不满足最低可用标准。",
                location="word/document.xml",
            )
        )

    if docx_inventory.get("docx_file_size_bytes", 0) < 1024 and docx_inventory.get("docx_openable", False):
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="DOCX_FILE_TOO_SMALL",
                message="docx 文件尺寸非常小，请确认输出内容是否完整。",
                location=str(docx_path),
                details={"file_size_bytes": docx_inventory.get("docx_file_size_bytes", 0)},
            )
        )

    # -------------------------------------------------------------------------
    # 标题结构检查
    # -------------------------------------------------------------------------
    if expected_heading_count > 0 and docx_inventory.get("heading_count", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="HEADINGS_MISSING",
                message="源文档存在标题结构，但 docx 中未检测到任何标题段落。",
                location="document body",
                details={
                    "expected_heading_count": expected_heading_count,
                    "actual_heading_count": docx_inventory.get("heading_count", 0),
                },
            )
        )
    elif expected_heading_count > 0 and docx_inventory.get("heading_count", 0) < max(1, expected_heading_count // 3):
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="HEADINGS_POSSIBLY_INCOMPLETE",
                message="docx 中检测到的标题段落明显少于源文档预期，请重点检查目录与章节层级。",
                location="document body",
                details={
                    "expected_heading_count": expected_heading_count,
                    "actual_heading_count": docx_inventory.get("heading_count", 0),
                },
            )
        )

    # -------------------------------------------------------------------------
    # 图片检查
    # -------------------------------------------------------------------------
    if expected_image_count > 0 and docx_inventory.get("image_reference_count_in_body", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="IMAGES_MISSING_IN_BODY",
                message="源文档存在图片命令，但 docx 正文中未检测到任何图片引用。",
                location="document body",
                details={
                    "expected_image_count": expected_image_count,
                    "actual_image_reference_count_in_body": docx_inventory.get("image_reference_count_in_body", 0),
                },
            )
        )
    elif expected_image_count > 0 and docx_inventory.get("image_reference_count_in_body", 0) < expected_image_count:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="IMAGE_REFERENCE_COUNT_LOWER_THAN_EXPECTED",
                message="docx 中的图片引用数少于源文档预期，请检查是否有图片缺失或未嵌入。",
                location="document body",
                details={
                    "expected_image_count": expected_image_count,
                    "actual_image_reference_count_in_body": docx_inventory.get("image_reference_count_in_body", 0),
                },
            )
        )

    if docx_inventory.get("media_file_count", 0) == 0 and expected_image_count > 0:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="NO_MEDIA_FILES_FOUND",
                message="预期存在图片，但 docx 包中未发现 word/media 文件，请重点检查图片嵌入情况。",
                location="word/media",
            )
        )

    # -------------------------------------------------------------------------
    # 表格检查
    # -------------------------------------------------------------------------
    if expected_table_count > 0 and docx_inventory.get("table_count", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="TABLES_MISSING",
                message="源文档存在表格环境，但 docx 中未检测到任何表格对象。",
                location="document body",
                details={
                    "expected_table_count": expected_table_count,
                    "actual_table_count": docx_inventory.get("table_count", 0),
                },
            )
        )
    elif expected_table_count > 0 and docx_inventory.get("table_count", 0) < expected_table_count:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="TABLE_COUNT_LOWER_THAN_EXPECTED",
                message="docx 中的表格数量少于源文档预期，请重点检查复杂表格是否损坏或丢失。",
                location="document body",
                details={
                    "expected_table_count": expected_table_count,
                    "actual_table_count": docx_inventory.get("table_count", 0),
                },
            )
        )

    # -------------------------------------------------------------------------
    # 图题 / 表题检查
    # -------------------------------------------------------------------------
    actual_caption_count_total = int(
        docx_inventory.get("effective_caption_count_total", docx_inventory.get("caption_count_total", 0))
    )
    explicit_caption_count_total = int(docx_inventory.get("caption_count_total", 0))
    proximity_caption_count_total = int(docx_inventory.get("proximity_caption_count_total", 0))

    if expected_caption_count > 0 and actual_caption_count_total == 0:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="NO_CAPTION_CANDIDATES_DETECTED",
                message="源文档存在 caption 命令，但 docx 中未检测到明确的图题/表题候选段落。",
                location="document body",
                details={
                    "expected_caption_count": expected_caption_count,
                    "actual_caption_count_total": actual_caption_count_total,
                    "explicit_caption_count_total": explicit_caption_count_total,
                    "proximity_caption_count_total": proximity_caption_count_total,
                },
            )
        )
    elif expected_caption_count > 0 and actual_caption_count_total < expected_caption_count:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="CAPTION_COUNT_LOWER_THAN_EXPECTED",
                message="docx 中检测到的图题/表题候选数量少于源文档预期，请重点检查图题表题完整性。",
                location="document body",
                details={
                    "expected_caption_count": expected_caption_count,
                    "actual_caption_count_total": actual_caption_count_total,
                    "explicit_caption_count_total": explicit_caption_count_total,
                    "proximity_caption_count_total": proximity_caption_count_total,
                },
            )
        )

    # -------------------------------------------------------------------------
    # 数学对象检查
    # -------------------------------------------------------------------------
    if expected_equation_count > 0 and docx_inventory.get("math_object_count", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="MATH_OBJECTS_MISSING",
                message="源文档存在公式环境，但 docx 中未检测到任何 Word 数学对象（OMML）。",
                location="document body",
                details={
                    "expected_equation_count": expected_equation_count,
                    "actual_math_object_count": docx_inventory.get("math_object_count", 0),
                },
            )
        )
    elif expected_equation_count > 0 and docx_inventory.get("math_object_count", 0) < max(1, expected_equation_count // 2):
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="MATH_OBJECT_COUNT_LOWER_THAN_EXPECTED",
                message="docx 中的数学对象数量明显少于源文档预期，请重点检查复杂公式是否被降级或遗漏。",
                location="document body",
                details={
                    "expected_equation_count": expected_equation_count,
                    "actual_math_object_count": docx_inventory.get("math_object_count", 0),
                },
            )
        )

    # -------------------------------------------------------------------------
    # 引用与内部链接检查
    # -------------------------------------------------------------------------
    internal_reference_structures = (
        docx_inventory.get("internal_hyperlink_count", 0)
        + docx_inventory.get("field_ref_count", 0)
        + docx_inventory.get("field_pageref_count", 0)
        + docx_inventory.get("bookmark_count", 0)
    )

    if expected_ref_count > 0 and internal_reference_structures == 0:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="NO_INTERNAL_REFERENCE_STRUCTURES_DETECTED",
                message="源文档存在内部引用，但 docx 中未检测到书签、内部跳转或 REF/PAGEREF 字段，请重点检查交叉引用。",
                location="document body",
                details={
                    "expected_ref_count": expected_ref_count,
                    "internal_hyperlink_count": docx_inventory.get("internal_hyperlink_count", 0),
                    "field_ref_count": docx_inventory.get("field_ref_count", 0),
                    "field_pageref_count": docx_inventory.get("field_pageref_count", 0),
                    "bookmark_count": docx_inventory.get("bookmark_count", 0),
                },
            )
        )
    elif expected_ref_count > 0 and docx_inventory.get("bookmark_count", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="BOOKMARKS_NOT_DETECTED",
                message="源文档存在内部引用，但 docx 中未检测到书签，部分交叉引用恢复能力可能有限。",
                location="document body",
                details={"expected_ref_count": expected_ref_count},
            )
        )

    # -------------------------------------------------------------------------
    # 目录与字段检查
    # -------------------------------------------------------------------------
    if docx_inventory.get("heading_count", 0) > 0 and docx_inventory.get("field_toc_count", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="TOC_FIELD_NOT_DETECTED",
                message="检测到标题结构，但未检测到 TOC 字段；目录可能需要在 Word 中手动插入或更新。",
                location="document body",
            )
        )

    # SEQ 字段是 Word caption / 编号体系的重要线索，但 Pandoc 不一定总生成。
    if expected_caption_count > 0 and actual_caption_count_total > 0 and docx_inventory.get("field_seq_count", 0) == 0:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="SEQ_FIELD_NOT_DETECTED",
                message="源文档存在 caption，但 docx 中未检测到 SEQ 字段；图表编号的 Word 原生可维护性可能有限。",
                location="document body",
            )
        )

    # -------------------------------------------------------------------------
    # 参考文献检查
    # -------------------------------------------------------------------------
    bibliography_detected = bool(
        docx_inventory.get("bibliography_section_found", False) or docx_inventory.get("bibliography_inferred_by_entries", False)
    )
    bibliography_entry_like_count = int(docx_inventory.get("bibliography_entry_like_count", 0))

    if expected_cite_count > 0 and not bibliography_detected:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="BIBLIOGRAPHY_SECTION_NOT_DETECTED",
                message="源文档存在文内引用，但 docx 中未检测到明确的参考文献章节或条目结构，请重点检查文末参考文献。",
                location="document body",
                details={
                    "expected_cite_count": expected_cite_count,
                    "bibliography_entry_like_count": bibliography_entry_like_count,
                    "bibliography_detection_mode": docx_inventory.get("bibliography_detection_mode", "none"),
                },
            )
        )
    elif (
        expected_cite_count > 0
        and docx_inventory.get("bibliography_section_found", False)
        and docx_inventory.get("bibliography_following_paragraph_count", 0) == 0
        and bibliography_entry_like_count == 0
    ):
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="BIBLIOGRAPHY_SECTION_EMPTY_OR_UNCLEAR",
                message="检测到了参考文献章节标题，但其后的参考文献段落数为 0，请检查参考文献是否真正生成。",
                location="document body",
                details={
                    "bibliography_heading_text": docx_inventory.get("bibliography_heading_text"),
                    "expected_cite_count": expected_cite_count,
                    "bibliography_entry_like_count": bibliography_entry_like_count,
                },
            )
        )

    # -------------------------------------------------------------------------
    # conversion / normalization / precheck 报告一致性检查
    # -------------------------------------------------------------------------
    if conversion_report and conversion_report.get("status") == STATUS_FAIL:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="UPSTREAM_CONVERSION_FAILED",
                message="Pandoc 主转换报告显示失败，当前 docx 结果不应视为可靠可交付产物。",
                location="pandoc-conversion-report.json",
                details={"upstream_status": conversion_report.get("status")},
            )
        )

    if normalization_report and normalization_report.get("status") == STATUS_FAIL:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="UPSTREAM_NORMALIZATION_FAILED",
                message="规范化报告显示失败，当前 docx 结果不应视为可靠可交付产物。",
                location="normalization-report.json",
                details={"upstream_status": normalization_report.get("status")},
            )
        )

    if precheck_report and precheck_report.get("status") == STATUS_FAIL:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="UPSTREAM_PRECHECK_FAILED",
                message="预检查报告显示失败；即使生成了 docx，仍需高度怀疑其完整性。",
                location="precheck-report.json",
                details={"upstream_status": precheck_report.get("status")},
            )
        )

    # -------------------------------------------------------------------------
    # 最终状态判定
    #
    # 判定原则：
    # - ERROR > 0 => FAIL
    # - 否则 WARN > 0 => PASS_WITH_WARNINGS
    # - 否则 PASS
    #
    # 这样做与前面脚本保持一致，也契合 acceptance_criteria 中的三态要求。
    # -------------------------------------------------------------------------
    error_count = sum(1 for finding in findings if finding.severity == SEVERITY_ERROR)
    warn_count = sum(1 for finding in findings if finding.severity == SEVERITY_WARN)
    info_count = sum(1 for finding in findings if finding.severity == SEVERITY_INFO)

    if error_count > 0:
        status = STATUS_FAIL
        can_continue = False
    elif warn_count > 0:
        status = STATUS_PASS_WITH_WARNINGS
        can_continue = True
    else:
        status = STATUS_PASS
        can_continue = True

    recommendations: list[str] = []

    if status == STATUS_FAIL:
        recommendations.append("当前 docx 未达到最低可用标准；请先修复 ERROR 级问题，再重新执行转换或检查。")
    else:
        recommendations.append("可进入 build_manual_fix_list.py 阶段，基于当前 postcheck 结果生成人工修复清单。")

    if warn_count > 0:
        recommendations.append("打开 Word 后优先执行全选并更新所有字段，再检查目录、图表编号和交叉引用。")

    if expected_ref_count > 0:
        recommendations.append("重点核对图、表、公式的交叉引用；必要时在 Word 中重建关键交叉引用。")

    if expected_table_count > 0:
        recommendations.append("重点核对复杂表格的单元格合并、跨页与版式。")

    if expected_equation_count > 0:
        recommendations.append("重点核对复杂公式是否仍为 Word 数学对象，以及公式编号是否正确。")

    if expected_cite_count > 0:
        recommendations.append("重点核对文末参考文献是否完整，以及文内引用是否能跳转到对应条目。")

    summary = {
        "status": status,
        "can_continue": can_continue,
        "docx_openable": docx_inventory.get("docx_openable", False),
        "minimum_usable_standard_likely_met": status != STATUS_FAIL,
        "used_conversion_report": conversion_report is not None,
        "used_normalization_report": normalization_report is not None,
        "used_precheck_report": precheck_report is not None,
        "errors": error_count,
        "warnings": warn_count,
        "infos": info_count,
    }

    metrics = {
        "expected_image_count": expected_image_count,
        "expected_table_count": expected_table_count,
        "expected_heading_count": expected_heading_count,
        "expected_equation_count": expected_equation_count,
        "expected_caption_count": expected_caption_count,
        "expected_ref_count": expected_ref_count,
        "expected_cite_count": expected_cite_count,
        "actual_image_reference_count_in_body": docx_inventory.get("image_reference_count_in_body", 0),
        "actual_table_count": docx_inventory.get("table_count", 0),
        "actual_heading_count": docx_inventory.get("heading_count", 0),
        "actual_math_object_count": docx_inventory.get("math_object_count", 0),
        "actual_caption_count_total": docx_inventory.get("caption_count_total", 0),
        "actual_effective_caption_count_total": actual_caption_count_total,
        "actual_proximity_caption_count_total": proximity_caption_count_total,
        "bibliography_entry_like_count": bibliography_entry_like_count,
        "bibliography_detected_by_entries": bool(docx_inventory.get("bibliography_inferred_by_entries", False)),
        "actual_internal_reference_structures": internal_reference_structures,
        "finding_error_count": error_count,
        "finding_warn_count": warn_count,
        "finding_info_count": info_count,
    }

    report = PostcheckReport(
        status=status,
        can_continue=can_continue,
        work_root=str(work_root.resolve()),
        source_project_root=str(source_project_root.resolve()) if source_project_root else None,
        input_docx=str(docx_path.resolve()),
        used_conversion_report=conversion_report is not None,
        used_normalization_report=normalization_report is not None,
        used_precheck_report=precheck_report is not None,
        findings=[asdict(finding) for finding in findings],
        source_inventory=source_inventory,
        docx_inventory=docx_inventory,
        metrics=metrics,
        summary=summary,
        recommendations=recommendations,
    )
    return report


# -----------------------------------------------------------------------------
# Markdown 报告渲染
# -----------------------------------------------------------------------------

def render_markdown_report(report: PostcheckReport) -> str:
    """
    将 postcheck 结果渲染为 Markdown 报告。

    报告结构重点包括：
    - 总体状态
    - source 侧期望 inventory
    - docx 侧实际 inventory
    - findings
    - recommendations

    这样既方便开发者排查，也方便人工审校者快速进入修复阶段。
    """
    lines: list[str] = []

    lines.append("# DOCX Postcheck Report")
    lines.append("")
    lines.append(f"- Status: **{report.status}**")
    lines.append(f"- Can continue: **{report.can_continue}**")
    lines.append(f"- Work root: `{report.work_root}`")
    lines.append(f"- Source project root: `{report.source_project_root or 'N/A'}`")
    lines.append(f"- Input DOCX: `{report.input_docx}`")
    lines.append(f"- Used conversion report: **{report.used_conversion_report}**")
    lines.append(f"- Used normalization report: **{report.used_normalization_report}**")
    lines.append(f"- Used precheck report: **{report.used_precheck_report}**")
    lines.append("")

    lines.append("## Summary")
    lines.append("")
    for key, value in report.summary.items():
        lines.append(f"- {key}: **{value}**")
    lines.append("")

    lines.append("## Metrics")
    lines.append("")
    for key, value in report.metrics.items():
        lines.append(f"- {key}: **{value}**")
    lines.append("")

    lines.append("## Source Inventory")
    lines.append("")
    for key, value in report.source_inventory.items():
        lines.append(f"- {key}: **{value}**")
    lines.append("")

    lines.append("## DOCX Inventory")
    lines.append("")
    for key, value in report.docx_inventory.items():
        lines.append(f"- {key}: **{value}**")
    lines.append("")

    lines.append("## Findings")
    lines.append("")
    if not report.findings:
        lines.append("- No findings.")
        lines.append("")
    else:
        grouped: dict[str, list[dict]] = defaultdict(list)
        for finding in report.findings:
            grouped[finding["severity"]].append(finding)

        for severity in SEVERITY_ORDER:
            if severity not in grouped:
                continue
            lines.append(f"### {severity}")
            lines.append("")
            for item in grouped[severity]:
                location = f" ({item['location']})" if item.get("location") else ""
                lines.append(f"- **[{item['code']}]** {item['message']}{location}")
            lines.append("")

    lines.append("## Recommendations")
    lines.append("")
    if report.recommendations:
        for recommendation in report.recommendations:
            lines.append(f"- {recommendation}")
    else:
        lines.append("- No immediate action required.")
    lines.append("")

    return "\n".join(lines)


# -----------------------------------------------------------------------------
# 主入口
# -----------------------------------------------------------------------------

def main() -> int:
    """
    主入口函数。

    固定执行顺序：
    1. 解析参数；
    2. 解析 work-root 与默认输入/输出路径；
    3. 读取前面阶段的报告；
    4. 构造 source inventory；
    5. 检查 docx 结构；
    6. 综合判断结果；
    7. 写出 JSON / Markdown 报告；
    8. 打印控制台摘要；
    9. 返回退出码。
    """
    parser = build_argument_parser()
    args = parser.parse_args()

    work_root = Path(args.work_root).resolve()
    if not work_root.exists() or not work_root.is_dir():
        print(f"[ERROR] 无效的工作目录: {work_root}", file=sys.stderr)
        return 1

    skill_root = locate_skill_root()
    postcheck_stage_dir = stage_dir(work_root, STAGE_POSTCHECK)
    postcheck_stage_dir.mkdir(parents=True, exist_ok=True)

    docx_path = resolve_explicit_or_stage_input(
        args.docx,
        work_root,
        STAGE_CONVERT,
        "output.docx",
        legacy_filename="output.docx",
    )
    conversion_json_path = resolve_explicit_or_stage_input(
        args.conversion_json,
        work_root,
        STAGE_CONVERT,
        "pandoc-conversion-report.json",
        legacy_filename="pandoc-conversion-report.json",
    )
    normalization_json_path = resolve_explicit_or_stage_input(
        args.normalization_json,
        work_root,
        STAGE_NORMALIZE,
        "normalization-report.json",
        legacy_filename="normalization-report.json",
    )
    json_out = resolve_explicit_or_stage_output(
        args.json_out,
        work_root,
        STAGE_POSTCHECK,
        "postcheck-report.json",
    )
    md_out = resolve_explicit_or_stage_output(
        args.md_out,
        work_root,
        STAGE_POSTCHECK,
        "postcheck-report.md",
    )

    initial_findings: list[Finding] = []
    initial_findings.extend(check_rule_files(skill_root))

    conversion_report = load_json_if_exists(conversion_json_path)
    if conversion_report is None:
        initial_findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="CONVERSION_REPORT_MISSING",
                message="未找到 pandoc-conversion-report.json；postcheck 将仅基于 docx 和可用的其他输入执行。",
                location=str(conversion_json_path),
            )
        )

    normalization_report = load_json_if_exists(normalization_json_path)
    if normalization_report is None:
        initial_findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="NORMALIZATION_REPORT_MISSING",
                message="未找到 normalization-report.json；source inventory 统计可能不完整。",
                location=str(normalization_json_path),
            )
        )

    # precheck-report.json 的查找策略：
    # 1. 若用户显式指定，则直接使用；
    # 2. 否则优先尝试 stage_precheck/precheck-report.json；
    # 3. 再回退 work_root/precheck-report.json；
    # 4. 若 normalization-report.json 中存在 project_root，则再尝试 source project_root/precheck-report.json。
    precheck_report: Optional[dict] = None
    source_project_root: Optional[Path] = None

    if normalization_report and normalization_report.get("project_root"):
        source_project_root = Path(normalization_report["project_root"]).resolve()

    if args.precheck_json:
        precheck_json_path = Path(args.precheck_json).resolve()
        precheck_report = load_json_if_exists(precheck_json_path)
        if precheck_report is None:
            initial_findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="EXPLICIT_PRECHECK_REPORT_MISSING",
                    message="显式指定的 precheck-report.json 不存在或无法读取。",
                    location=str(precheck_json_path),
                )
            )
    else:
        candidate_in_work_root = resolve_explicit_or_stage_input(
            None,
            work_root,
            STAGE_PRECHECK,
            "precheck-report.json",
            legacy_filename="precheck-report.json",
        )
        precheck_report = load_json_if_exists(candidate_in_work_root)

        if precheck_report is None and source_project_root is not None:
            candidate_in_source_root = (source_project_root / "precheck-report.json").resolve()
            precheck_report = load_json_if_exists(candidate_in_source_root)

    # 收集源侧 inventory。
    source_inventory = collect_source_inventory(work_root, normalization_report)

    # docx 检查。
    docx_inventory, docx_findings = inspect_docx(docx_path)
    initial_findings.extend(docx_findings)

    # 综合分析。
    report = analyze_results(
        work_root=work_root,
        docx_path=docx_path,
        source_project_root=source_project_root,
        precheck_report=precheck_report,
        normalization_report=normalization_report,
        conversion_report=conversion_report,
        source_inventory=source_inventory,
        docx_inventory=docx_inventory,
        initial_findings=initial_findings,
    )

    # 写出报告。
    persist_stage_report(
        work_root=work_root,
        stage=STAGE_POSTCHECK,
        report_obj=report,
        markdown_text=render_markdown_report(report),
        report_json_path=json_out,
        report_md_path=md_out,
        status=report.status,
        can_continue=report.can_continue,
        artifacts={
            "postcheck_report_json": json_out,
            "postcheck_report_md": md_out,
            "input_docx": docx_path,
            "conversion_report_json": conversion_json_path,
            "normalization_report_json": normalization_json_path,
        },
        summary=report.summary,
        metrics=report.metrics,
        top_level_artifacts={
            "reports": {
                "postcheck_report_json": json_out,
                "postcheck_report_md": md_out,
            }
        },
    )

    # 控制台摘要。
    print(f"[{report.status}] DOCX postcheck completed.")
    print(f"Input DOCX: {report.input_docx}")
    print(f"Errors: {report.metrics.get('finding_error_count', 0)}")
    print(f"Warnings: {report.metrics.get('finding_warn_count', 0)}")
    print(f"Infos: {report.metrics.get('finding_info_count', 0)}")
    print(f"JSON report: {json_out}")
    print(f"Markdown report: {md_out}")

    return 1 if report.status == STATUS_FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
