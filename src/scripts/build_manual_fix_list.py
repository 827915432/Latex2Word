#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_manual_fix_list.py

功能概述
--------
本脚本用于在 LaTeX -> Word 流水线的最后阶段，聚合前面各阶段的报告，
生成一份“真正可执行”的人工修复清单。

它解决的问题包括：
1. 汇总 precheck / normalization / pandoc conversion / postcheck 各阶段的剩余问题；
2. 将技术性问题翻译成用户可执行的修复任务；
3. 为每个待修项给出：
   - 优先级
   - 问题类别
   - 问题位置
   - 为什么重要
   - 推荐修法
   - 是否影响交付
4. 输出 JSON + Markdown 两份清单，供用户在 Word 中按顺序修复；
5. 将“需要从哪里开始修”固定成标准顺序，避免用户通篇乱查。

设计边界
--------
本脚本只负责“清单生成”，不做以下事情：
- 不再次运行 Pandoc；
- 不再次检查或修改 .docx；
- 不修改源工程；
- 不自动执行 Word 修复；
- 不替代人工最终验收。

依赖
----
- Python 3.9+
- 仅使用标准库，不依赖第三方包

典型用法
--------
python scripts/build_manual_fix_list.py --work-root D:/work/my-paper__latex_to_word_work

输出
----
默认会在 work-root 下生成两类结果：

1) 阶段化产物（`stage_checklist/`）：
- manual-fix-checklist.json
- manual-fix-checklist.md

2) 面向用户阅读的目录视图：
- deliverables/
- reports/
- logs/
- debug/
- README_RUN.md

退出码
------
- 0: PASS 或 PASS_WITH_WARNINGS
- 1: FAIL
"""

from __future__ import annotations

import argparse
import shutil
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Optional

from pipeline_common import (
    load_json_if_exists,
    locate_skill_root,
    safe_relative,
    write_json,
    write_markdown,
)
from pipeline_constants import (
    REQUIRED_RULE_FILES,
    SEVERITY_ERROR,
    SEVERITY_WARN,
    STATUS_FAIL,
    STATUS_PASS,
    STATUS_PASS_WITH_WARNINGS,
)
from pipeline_layout import (
    STAGE_CHECKLIST,
    STAGE_CONVERT,
    STAGE_NORMALIZE,
    STAGE_POSTCHECK,
    STAGE_PRECHECK,
    resolve_explicit_or_stage_input,
    resolve_explicit_or_stage_output,
    stage_default_or_legacy,
    stage_dir,
)
from stage_reporting import persist_stage_report


# -----------------------------------------------------------------------------
# 常量定义
# -----------------------------------------------------------------------------

DELIVERY_NONE = "none"
DELIVERY_LOW = "low"
DELIVERY_MEDIUM = "medium"
DELIVERY_HIGH = "high"


# -----------------------------------------------------------------------------
# 数据结构定义
# -----------------------------------------------------------------------------

@dataclass
class ChecklistItem:
    """
    人工修复清单中的单个任务项。

    字段说明
    --------
    item_id:
        稳定的人类可读编号，例如 G01、P03、N02、C01、D05。
        其中前缀含义：
        - G: 全局通用步骤（Global）
        - P: precheck 来源
        - N: normalization 来源
        - C: conversion 来源
        - D: postcheck 来源

    priority:
        数值越小表示越优先。建议用户按 priority 升序处理。

    category:
        问题类别，例如：
        - fields
        - toc
        - cross_references
        - images
        - tables
        - equations
        - bibliography
        - headings
        - custom_macros
        - conditional_build

    title:
        简洁标题。

    source_stage:
        来源阶段：
        - global
        - precheck
        - normalization
        - conversion
        - postcheck

    source_code:
        原始 finding code 或 action_type；若为全局固定项，则可写固定值。

    location:
        问题位置。允许是文件路径、docx 逻辑位置、标签名、章节说明等。

    issue_summary:
        对问题的简明描述。

    why_it_matters:
        说明为什么这个问题值得优先处理。

    recommended_actions:
        推荐用户执行的操作步骤列表。
        设计成列表而非单字符串，是为了让后续 UI 或脚本消费更稳定。

    affects_delivery:
        是否影响最终交付质量：
        - none / low / medium / high

    details:
        保留上游报告中的结构化上下文，便于后续脚本或人工深挖。
    """
    item_id: str
    priority: int
    category: str
    title: str
    source_stage: str
    source_code: str
    location: Optional[str]
    issue_summary: str
    why_it_matters: str
    recommended_actions: list[str]
    affects_delivery: str
    details: dict = field(default_factory=dict)


@dataclass
class ManualFixChecklistReport:
    """
    最终人工修复清单报告对象。

    设计目标：
    - 一份 JSON，便于后续脚本消费；
    - 一份 Markdown，便于用户直接照着执行；
    - 汇总工作目录、输入报告状态、总项数、优先级分布等信息。
    """
    status: str
    can_continue: bool
    work_root: str
    source_project_root: Optional[str]
    used_precheck_report: bool
    used_normalization_report: bool
    used_conversion_report: bool
    used_postcheck_report: bool
    items: list[dict]
    metrics: dict
    summary: dict
    recommendations: list[str]
    user_view_generated: bool
    user_view_root: Optional[str]
    published_file_count: int


# -----------------------------------------------------------------------------
# 参数与通用工具函数
# -----------------------------------------------------------------------------

def build_argument_parser() -> argparse.ArgumentParser:
    """
    构造命令行参数解析器。

    参数设计原则：
    - 以 work-root 为中心；
    - 默认约定与前面各阶段保持一致；
    - 允许用户显式覆盖报告路径，但不暴露多余配置。
    """
    parser = argparse.ArgumentParser(
        description="Build a user-facing manual fix checklist from all pipeline reports."
    )
    parser.add_argument(
        "--work-root",
        required=True,
        help="规范化工作目录。",
    )
    parser.add_argument(
        "--precheck-json",
        default=None,
        help="预检查报告 JSON 路径；默认优先 <work-root>/stage_precheck/precheck-report.json，若不存在则尝试源工程目录。",
    )
    parser.add_argument(
        "--normalization-json",
        default=None,
        help="规范化报告 JSON 路径；默认优先 <work-root>/stage_normalize/normalization-report.json",
    )
    parser.add_argument(
        "--conversion-json",
        default=None,
        help="Pandoc 主转换报告 JSON 路径；默认优先 <work-root>/stage_convert/pandoc-conversion-report.json",
    )
    parser.add_argument(
        "--postcheck-json",
        default=None,
        help="docx 后检查报告 JSON 路径；默认优先 <work-root>/stage_postcheck/postcheck-report.json",
    )
    parser.add_argument(
        "--json-out",
        default=None,
        help="清单 JSON 输出路径；默认 <work-root>/stage_checklist/manual-fix-checklist.json",
    )
    parser.add_argument(
        "--md-out",
        default=None,
        help="清单 Markdown 输出路径；默认 <work-root>/stage_checklist/manual-fix-checklist.md",
    )
    parser.add_argument(
        "--no-user-view",
        action="store_true",
        help="跳过用户视图目录（deliverables/reports/logs/debug）与 README_RUN.md 的生成。",
    )
    return parser


def check_rule_files(skill_root: Path) -> list[ChecklistItem]:
    """
    检查规则文件是否存在。

    这一步不是为了阻止清单生成，而是为了在 skill 目录损坏时给出明确告警。
    """
    items: list[ChecklistItem] = []
    seq = 1
    for relative in REQUIRED_RULE_FILES:
        candidate = (skill_root / relative).resolve()
        if not candidate.exists():
            items.append(
                ChecklistItem(
                    item_id=f"GX{seq:02d}",
                    priority=950,
                    category="tooling",
                    title="补齐缺失的规则文件",
                    source_stage="global",
                    source_code="RULE_FILE_MISSING",
                    location=safe_relative(candidate, skill_root),
                    issue_summary="Skill 目录中缺少规则文件，后续报告解释性可能下降。",
                    why_it_matters="规则文件缺失不会直接破坏当前 docx，但会影响流程一致性和问题解释能力。",
                    recommended_actions=[
                        f"确认文件 `{relative}` 是否存在于 skill 根目录下。",
                        "若文件被误删，请恢复该规则文件后重新执行相关阶段。",
                    ],
                    affects_delivery=DELIVERY_LOW,
                    details={"required_file": relative},
                )
            )
            seq += 1
    return items


# -----------------------------------------------------------------------------
# 默认路径解析
# -----------------------------------------------------------------------------

def resolve_report_paths(
    work_root: Path,
    precheck_arg: Optional[str],
    normalization_arg: Optional[str],
    conversion_arg: Optional[str],
    postcheck_arg: Optional[str],
) -> tuple[Optional[Path], Optional[Path], Optional[Path], Optional[Path], Optional[Path]]:
    """
    解析四类报告路径，并尽量从 normalization-report.json 中恢复 source project root。

    返回：
    - precheck_json_path
    - normalization_json_path
    - conversion_json_path
    - postcheck_json_path
    - source_project_root

    说明：
    - precheck-report.json 的默认查找优先级与 postcheck 阶段一致；
    - source_project_root 用于补充“源工程目录”信息，方便报告展示。
    """
    normalization_json_path = resolve_explicit_or_stage_input(
        normalization_arg,
        work_root,
        STAGE_NORMALIZE,
        "normalization-report.json",
        legacy_filename="normalization-report.json",
    )
    normalization_report = load_json_if_exists(normalization_json_path)

    source_project_root: Optional[Path] = None
    if normalization_report and normalization_report.get("project_root"):
        source_project_root = Path(normalization_report["project_root"]).resolve()

    conversion_json_path = resolve_explicit_or_stage_input(
        conversion_arg,
        work_root,
        STAGE_CONVERT,
        "pandoc-conversion-report.json",
        legacy_filename="pandoc-conversion-report.json",
    )
    postcheck_json_path = resolve_explicit_or_stage_input(
        postcheck_arg,
        work_root,
        STAGE_POSTCHECK,
        "postcheck-report.json",
        legacy_filename="postcheck-report.json",
    )

    candidate_in_work_root = resolve_explicit_or_stage_input(
        precheck_arg,
        work_root,
        STAGE_PRECHECK,
        "precheck-report.json",
        legacy_filename="precheck-report.json",
    )
    if candidate_in_work_root.exists() or precheck_arg:
        precheck_json_path = candidate_in_work_root
    elif source_project_root is not None:
        candidate_in_source_root = (source_project_root / "precheck-report.json").resolve()
        precheck_json_path = candidate_in_source_root
    else:
        precheck_json_path = candidate_in_work_root

    return (
        precheck_json_path,
        normalization_json_path,
        conversion_json_path,
        postcheck_json_path,
        source_project_root,
    )


def _resolve_path_like(path_like: Optional[str], base_dir: Path) -> Optional[Path]:
    """
    将报告中的路径字段解析为绝对路径。

    说明：
    - 若 path_like 为空，返回 None；
    - 若是相对路径，则按 base_dir 解析；
    - 仅做路径解析，不检查是否存在。
    """
    if not path_like or not isinstance(path_like, str):
        return None
    candidate = Path(path_like)
    if not candidate.is_absolute():
        return (base_dir / candidate).resolve()
    return candidate.resolve()


def _paired_markdown_for_json(json_path: Optional[Path]) -> Optional[Path]:
    """
    根据 json 报告路径推断对应的 markdown 报告路径。
    """
    if json_path is None:
        return None
    return json_path.with_suffix(".md")


def publish_user_view(
    *,
    work_root: Path,
    precheck_json_path: Optional[Path],
    normalization_json_path: Optional[Path],
    conversion_json_path: Optional[Path],
    postcheck_json_path: Optional[Path],
    checklist_json_path: Path,
    checklist_md_path: Path,
    conversion_report: Optional[dict],
) -> tuple[bool, Path, int, list[str]]:
    """
    生成面向用户阅读的目录视图。

    设计原则：
    - 不移动原文件，只复制，保持历史兼容；
    - 目录结构固定：deliverables / reports / logs / debug；
    - 缺失非关键文件仅记告警，不阻塞主流程。
    """
    user_view_root = work_root.resolve()
    deliverables_dir = user_view_root / "deliverables"
    reports_dir = user_view_root / "reports"
    logs_dir = user_view_root / "logs"
    debug_dir = user_view_root / "debug"

    for folder in [deliverables_dir, reports_dir, logs_dir, debug_dir]:
        folder.mkdir(parents=True, exist_ok=True)

    warnings: list[str] = []
    published_file_count = 0

    def publish_file(src: Optional[Path], dst: Path, label: str, required: bool) -> None:
        nonlocal published_file_count, warnings
        if src is None:
            if required:
                warnings.append(f"{label}: 源路径不可用。")
            return

        resolved_src = src.resolve()
        if not resolved_src.exists() or not resolved_src.is_file():
            if required:
                warnings.append(f"{label}: 未找到文件 `{resolved_src}`。")
            return

        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            same_file = False
            if dst.exists():
                try:
                    same_file = resolved_src.samefile(dst)
                except Exception:
                    same_file = False
            if not same_file:
                shutil.copy2(str(resolved_src), str(dst))
            published_file_count += 1
        except Exception as exc:
            if required:
                warnings.append(f"{label}: 复制失败（{exc}）。")

    # deliverables
    output_docx_path = None
    pandoc_log_path = None
    if conversion_report and isinstance(conversion_report, dict):
        output_docx_path = _resolve_path_like(conversion_report.get("output_docx"), work_root)
        pandoc_log_path = _resolve_path_like(conversion_report.get("pandoc_log"), work_root)

    if output_docx_path is None:
        output_docx_path = stage_default_or_legacy(
            work_root,
            STAGE_CONVERT,
            "output.docx",
            legacy_filename="output.docx",
        )
    if pandoc_log_path is None:
        pandoc_log_path = stage_default_or_legacy(
            work_root,
            STAGE_CONVERT,
            "pandoc-conversion.log",
            legacy_filename="pandoc-conversion.log",
        )

    publish_file(output_docx_path, deliverables_dir / "output.docx", "交付文档 output.docx", required=True)
    publish_file(
        checklist_md_path.resolve(),
        deliverables_dir / "manual-fix-checklist.md",
        "人工修复清单 manual-fix-checklist.md",
        required=True,
    )

    # reports
    precheck_md_path = _paired_markdown_for_json(precheck_json_path)
    normalization_md_path = _paired_markdown_for_json(normalization_json_path)
    conversion_md_path = _paired_markdown_for_json(conversion_json_path)
    postcheck_md_path = _paired_markdown_for_json(postcheck_json_path)

    publish_file(precheck_json_path, reports_dir / "precheck-report.json", "预检查报告 JSON", required=False)
    publish_file(precheck_md_path, reports_dir / "precheck-report.md", "预检查报告 Markdown", required=False)

    publish_file(
        normalization_json_path,
        reports_dir / "normalization-report.json",
        "规范化报告 JSON",
        required=False,
    )
    publish_file(
        normalization_md_path,
        reports_dir / "normalization-report.md",
        "规范化报告 Markdown",
        required=False,
    )

    publish_file(
        conversion_json_path,
        reports_dir / "pandoc-conversion-report.json",
        "Pandoc 转换报告 JSON",
        required=False,
    )
    publish_file(
        conversion_md_path,
        reports_dir / "pandoc-conversion-report.md",
        "Pandoc 转换报告 Markdown",
        required=False,
    )

    publish_file(postcheck_json_path, reports_dir / "postcheck-report.json", "后检查报告 JSON", required=False)
    publish_file(postcheck_md_path, reports_dir / "postcheck-report.md", "后检查报告 Markdown", required=False)

    publish_file(
        checklist_json_path.resolve(),
        reports_dir / "manual-fix-checklist.json",
        "人工修复清单 JSON",
        required=True,
    )

    # logs
    publish_file(pandoc_log_path, logs_dir / "pandoc-conversion.log", "Pandoc 转换日志", required=False)

    # debug（统一去掉前导点，避免影响可见性）
    metadata_src = stage_default_or_legacy(
        work_root,
        STAGE_CONVERT,
        ".pandoc_metadata.json",
        legacy_filename=".pandoc_metadata.json",
    )
    resource_dirs_src = stage_default_or_legacy(
        work_root,
        STAGE_CONVERT,
        ".pandoc_resource_dirs.txt",
        legacy_filename=".pandoc_resource_dirs.txt",
    )
    bibs_src = stage_default_or_legacy(
        work_root,
        STAGE_CONVERT,
        ".pandoc_bibliographies.txt",
        legacy_filename=".pandoc_bibliographies.txt",
    )
    publish_file(
        metadata_src,
        debug_dir / "pandoc_metadata.json",
        "Pandoc 元数据",
        required=False,
    )
    publish_file(
        resource_dirs_src,
        debug_dir / "pandoc_resource_dirs.txt",
        "Pandoc 资源目录清单",
        required=False,
    )
    publish_file(
        bibs_src,
        debug_dir / "pandoc_bibliographies.txt",
        "Pandoc 参考文献清单",
        required=False,
    )

    return published_file_count > 0, user_view_root, published_file_count, warnings


def write_run_readme(
    *,
    work_root: Path,
    checklist_status: str,
    precheck_status: str,
    normalization_status: str,
    conversion_status: str,
    postcheck_status: str,
    user_view_generated: bool,
    published_file_count: int,
    user_view_warnings: list[str],
) -> Path:
    """
    生成用户入口说明 README_RUN.md。
    """
    readme_path = (work_root / "README_RUN.md").resolve()
    lines: list[str] = []

    lines.append("# 转换结果阅读入口")
    lines.append("")
    lines.append("## 总体状态")
    lines.append("")
    lines.append(f"- 人工修复清单状态：**{checklist_status}**")
    lines.append(f"- precheck：**{precheck_status}**")
    lines.append(f"- normalize：**{normalization_status}**")
    lines.append(f"- pandoc convert：**{conversion_status}**")
    lines.append(f"- postcheck：**{postcheck_status}**")
    lines.append(f"- 用户视图目录已生成：**{user_view_generated}**")
    lines.append(f"- 已发布文件数：**{published_file_count}**")
    lines.append("")
    lines.append("## 先看这些文件")
    lines.append("")
    lines.append("1. `deliverables/output.docx`")
    lines.append("2. `deliverables/manual-fix-checklist.md`")
    lines.append("")
    lines.append("## 报告阅读顺序")
    lines.append("")
    lines.append("1. `reports/precheck-report.md`")
    lines.append("2. `reports/normalization-report.md`")
    lines.append("3. `reports/pandoc-conversion-report.md`")
    lines.append("4. `reports/postcheck-report.md`")
    lines.append("5. `reports/manual-fix-checklist.json`")
    lines.append("")
    lines.append("## 其他目录")
    lines.append("")
    lines.append("- `logs/`：原始转换日志")
    lines.append("- `debug/`：Pandoc 调试辅助文件")
    lines.append("")

    if user_view_warnings:
        lines.append("## 用户视图生成告警（非阻塞）")
        lines.append("")
        for warning in user_view_warnings:
            lines.append(f"- {warning}")
        lines.append("")

    write_markdown(readme_path, "\n".join(lines))
    return readme_path


# -----------------------------------------------------------------------------
# 全局固定步骤
# -----------------------------------------------------------------------------

def build_global_items(
    source_inventory: dict,
    postcheck_report: Optional[dict],
) -> list[ChecklistItem]:
    """
    生成所有文档都建议执行的“全局固定步骤”。

    这些步骤不依赖具体错误，即使文档整体质量较好，也建议用户按顺序执行。
    它们的核心目的是：
    - 先让 Word 刷新字段；
    - 再看目录；
    - 再看图表标题与引用；
    - 最后才处理局部版式问题。

    这样可以显著降低用户在错误顺序上浪费时间。
    """
    items: list[ChecklistItem] = []

    expected_headings = int(source_inventory.get("heading_command_count", 0))
    expected_captions = int(source_inventory.get("caption_command_count", 0))
    expected_refs = int(source_inventory.get("ref_count", 0))
    expected_cites = int(source_inventory.get("cite_count", 0))
    expected_tables = int(source_inventory.get("table_env_count", 0))
    expected_images = int(source_inventory.get("image_command_count", 0))
    expected_equations = int(source_inventory.get("equation_env_count", 0))

    items.append(
        ChecklistItem(
            item_id="G01",
            priority=10,
            category="fields",
            title="先在 Word 中更新全文字段",
            source_stage="global",
            source_code="UPDATE_ALL_FIELDS",
            location="entire document",
            issue_summary="目录、图表编号、交叉引用、页码等 Word 字段可能尚未刷新。",
            why_it_matters="很多看起来像‘错误’的问题，其实在更新字段后会自动缓解或消失；先做这一步可以避免误判。",
            recommended_actions=[
                "在 Word 中打开文档后，先全选全文（Ctrl+A）。",
                "执行“更新域/更新字段”。",
                "若 Word 询问更新目录，请选择更新整个目录。",
            ],
            affects_delivery=DELIVERY_HIGH,
            details={},
        )
    )

    if expected_headings > 0:
        items.append(
            ChecklistItem(
                item_id="G02",
                priority=20,
                category="toc",
                title="检查目录与章节层级",
                source_stage="global",
                source_code="CHECK_TOC_AND_HEADINGS",
                location="table of contents / heading structure",
                issue_summary="文档存在章节层级，目录是否可更新、标题是否仍保持为 Word 标题样式，需要人工确认。",
                why_it_matters="标题层级和目录是 Word 文档后续可维护性的核心；若这里不对，后续编号和导航都会受影响。",
                recommended_actions=[
                    "更新字段后，检查目录是否生成或能否正确更新。",
                    "随机抽查若干章节标题，确认其仍是 Word 标题样式而非普通正文。",
                    "若目录缺失但标题层级正确，可在 Word 中手动插入目录。",
                ],
                affects_delivery=DELIVERY_HIGH,
                details={"expected_heading_count": expected_headings},
            )
        )

    if expected_captions > 0 and (expected_images > 0 or expected_tables > 0):
        items.append(
            ChecklistItem(
                item_id="G03",
                priority=30,
                category="captions",
                title="检查图题、表题以及图目录/表目录基础",
                source_stage="global",
                source_code="CHECK_CAPTIONS_AND_LISTS",
                location="captions / list of figures / list of tables",
                issue_summary="文档存在图题或表题，应人工确认其完整性以及是否可支撑图目录、表目录生成。",
                why_it_matters="图题表题是图表引用、图目录和表目录的基础；若标题丢失或不规范，后续交叉引用维护会很困难。",
                recommended_actions=[
                    "抽查若干图片和表格，确认图题表题完整存在。",
                    "若项目要求图目录或表目录，请在 Word 中尝试插入或更新。",
                    "若个别图题/表题未被识别为可维护对象，可在 Word 中重建对应 caption。",
                ],
                affects_delivery=DELIVERY_HIGH,
                details={
                    "expected_caption_count": expected_captions,
                    "expected_image_count": expected_images,
                    "expected_table_count": expected_tables,
                },
            )
        )

    if expected_refs > 0:
        items.append(
            ChecklistItem(
                item_id="G04",
                priority=40,
                category="cross_references",
                title="检查关键交叉引用（图、表、公式、附录）",
                source_stage="global",
                source_code="CHECK_KEY_CROSS_REFERENCES",
                location="cross-references",
                issue_summary="源文档存在内部引用，需人工确认关键图表公式引用是否仍可点击、是否指向正确对象。",
                why_it_matters="交叉引用是 LaTeX 转 Word 最容易受损的部分；即使文档看起来正常，也可能隐藏断链或错链。",
                recommended_actions=[
                    "优先抽查正文中最关键的图引用、表引用和公式引用。",
                    "若引用文本存在但跳转不对，可在 Word 中重建关键交叉引用。",
                    "若附录引用或页码引用存在，也应一并核对。",
                ],
                affects_delivery=DELIVERY_HIGH,
                details={"expected_ref_count": expected_refs},
            )
        )

    if expected_cites > 0:
        items.append(
            ChecklistItem(
                item_id="G05",
                priority=50,
                category="bibliography",
                title="检查文内引用与文末参考文献",
                source_stage="global",
                source_code="CHECK_BIBLIOGRAPHY",
                location="citations / bibliography",
                issue_summary="源文档存在文内引用，需要人工确认文末参考文献是否完整，以及文内引用是否能定位到对应条目。",
                why_it_matters="参考文献系统是论文和技术报告的核心交付要求之一；若这里错配，整篇文档可信度会明显下降。",
                recommended_actions=[
                    "抽查几个文内引用，确认其格式与样式符合要求。",
                    "滚动到文末，确认参考文献章节存在且条目不为空。",
                    "若文内引用无法跳到文末条目，至少要保证文末条目完整且编号/作者年制正确。",
                ],
                affects_delivery=DELIVERY_HIGH,
                details={"expected_cite_count": expected_cites},
            )
        )

    if expected_equations > 0:
        items.append(
            ChecklistItem(
                item_id="G06",
                priority=60,
                category="equations",
                title="检查复杂公式与公式编号",
                source_stage="global",
                source_code="CHECK_EQUATIONS",
                location="equations",
                issue_summary="源文档存在公式环境，需人工确认关键公式是否仍为 Word 数学对象，公式编号是否正确。",
                why_it_matters="公式本体和编号是技术文档的关键语义对象；复杂公式一旦损坏，往往无法通过简单格式微调修复。",
                recommended_actions=[
                    "抽查正文中的复杂公式、矩阵、cases 和多行公式。",
                    "确认关键公式不是整体图片，仍可在 Word 中编辑。",
                    "若公式编号与引用不一致，优先修复关键公式的编号与交叉引用。",
                ],
                affects_delivery=DELIVERY_HIGH,
                details={"expected_equation_count": expected_equations},
            )
        )

    # 若已有 postcheck，提示用户优先结合后检查报告处理。
    if postcheck_report is not None:
        items.append(
            ChecklistItem(
                item_id="G07",
                priority=70,
                category="review_strategy",
                title="按 postcheck 报告优先修复高风险对象",
                source_stage="global",
                source_code="FOLLOW_POSTCHECK_PRIORITY",
                location="postcheck-report.json / postcheck-report.md",
                issue_summary="postcheck 已对结构级问题进行了归类，建议用户不要通篇盲查，而应按报告顺序处理。",
                why_it_matters="先修结构性问题，再修局部版式，通常比从头逐页浏览更高效，也更不容易漏掉关键对象。",
                recommended_actions=[
                    "先处理 ERROR 级问题，再处理 WARN 级问题。",
                    "对 postcheck 已标记的图表、交叉引用和参考文献问题进行定点修复。",
                    "修完一轮后，再执行一次全字段更新并做终审。",
                ],
                affects_delivery=DELIVERY_MEDIUM,
                details={},
            )
        )

    return items


# -----------------------------------------------------------------------------
# 单项映射与生成辅助函数
# -----------------------------------------------------------------------------

def make_item(
    *,
    item_id: str,
    priority: int,
    category: str,
    title: str,
    source_stage: str,
    source_code: str,
    location: Optional[str],
    issue_summary: str,
    why_it_matters: str,
    recommended_actions: list[str],
    affects_delivery: str,
    details: Optional[dict] = None,
) -> ChecklistItem:
    """
    创建 ChecklistItem 的便捷函数。

    这样可以让后续映射逻辑更紧凑，同时保持每个字段都被显式填写。
    """
    return ChecklistItem(
        item_id=item_id,
        priority=priority,
        category=category,
        title=title,
        source_stage=source_stage,
        source_code=source_code,
        location=location,
        issue_summary=issue_summary,
        why_it_matters=why_it_matters,
        recommended_actions=recommended_actions,
        affects_delivery=affects_delivery,
        details=details or {},
    )


def item_dedup_key(item: ChecklistItem) -> tuple[str, str, str, str]:
    """
    定义清单项去重键。

    去重原则：
    - 同一来源阶段
    - 同一 source_code
    - 同一 category
    - 同一 location

    这样可以避免上游多个报告重复生成同一性质任务。
    """
    return (
        item.source_stage,
        item.source_code,
        item.category,
        item.location or "",
    )


def append_dedup(items: list[ChecklistItem], seen: set[tuple[str, str, str, str]], item: ChecklistItem) -> None:
    """
    仅当未重复时，才把 item 加入清单。

    这样可显著减少由多阶段重复告警带来的清单噪声。
    """
    key = item_dedup_key(item)
    if key not in seen:
        items.append(item)
        seen.add(key)


def shorten_for_line(text: str, limit: int = 80) -> str:
    """
    截断单行文本，避免清单项过长。
    """
    value = (text or "").strip()
    if len(value) <= limit:
        return value
    if limit <= 3:
        return value[:limit]
    return value[: limit - 3].rstrip() + "..."


def format_unrenderable_image_examples(details: dict, limit: int = 8) -> list[str]:
    """
    从 finding details 中提取“不可渲染图片”对象摘要。
    """
    raw_examples = details.get("unsupported_image_examples", [])
    if not isinstance(raw_examples, list):
        return []

    summaries: list[str] = []
    for raw in raw_examples:
        if not isinstance(raw, dict):
            continue
        source_hint = shorten_for_line(str(raw.get("source_hint", "")).strip(), 70)
        target = shorten_for_line(str(raw.get("target", "")).strip(), 70)
        paragraph_text = shorten_for_line(str(raw.get("paragraph_text", "")).strip(), 45)
        next_paragraph_text = shorten_for_line(str(raw.get("next_paragraph_text", "")).strip(), 45)
        rid = str(raw.get("rid", "")).strip()
        paragraph_index = raw.get("paragraph_index")

        generic_hints = {"picture", "image", "图", "图片"}
        source_lower = source_hint.lower()
        source_is_generic = source_lower in generic_hints or source_lower.startswith("picture ")

        context_hint = next_paragraph_text or paragraph_text
        meaningful_source = "" if source_is_generic else source_hint
        anchor = meaningful_source or context_hint or target or rid
        if not anchor:
            continue

        location_parts: list[str] = []
        if rid:
            location_parts.append(f"rId={rid}")
        if paragraph_index:
            location_parts.append(f"段落#{paragraph_index}")
        location_text = f" ({', '.join(location_parts)})" if location_parts else ""

        if target and target != anchor:
            summaries.append(f"{anchor} -> {target}{location_text}")
        else:
            summaries.append(f"{anchor}{location_text}")
        if len(summaries) >= limit:
            break

    return summaries


# -----------------------------------------------------------------------------
# 从各阶段报告映射为人工修复项
# -----------------------------------------------------------------------------

def items_from_precheck(precheck_report: Optional[dict]) -> list[ChecklistItem]:
    """
    将 precheck 阶段的 finding 转为人工修复项。

    设计原则：
    - 只转换对“人工收口”有意义的问题；
    - 明确把结构性高风险对象转成用户可执行语言；
    - 已在后续阶段被更具体地捕获的问题，允许后续用更高优先级版本覆盖。
    """
    items: list[ChecklistItem] = []
    if not precheck_report:
        return items

    findings = precheck_report.get("findings", [])
    seq = 1
    for finding in findings:
        severity = finding.get("severity", "")
        code = finding.get("code", "")
        location = finding.get("file") or None
        details = finding.get("details", {}) or {}

        # precheck 只将确实值得人工关注的对象转入清单。
        if code == "MISSING_IMAGE":
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=120,
                    category="images",
                    title="补查缺失图片资源",
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary="预检查阶段检测到图片资源缺失。",
                    why_it_matters="图片资源缺失会直接导致 Word 中的图片为空白或丢失，属于明显影响交付的问题。",
                    recommended_actions=[
                        "确认原始 LaTeX 工程中该图片是否确实存在。",
                        "若图片路径写错，请在源工程中修正后重新执行流水线。",
                        "若该图片允许降级为其他格式，请先补齐可用资源再重转。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code == "COMPLEX_TABLES_DETECTED":
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=220,
                    category="tables",
                    title="重点人工核对复杂表格",
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary="预检查阶段已识别出复杂表格文件，后续 Word 表格极可能需要人工修正。",
                    why_it_matters="复杂表格是 LaTeX 转 Word 的高风险对象，若不优先检查，常会在最终交付前才暴露大面积错位问题。",
                    recommended_actions=[
                        "打开 postcheck 报告，优先定位复杂表格相关问题。",
                        "在 Word 中重点检查跨页、跨行、跨列、表头重复和列宽。",
                        "若表格已近似转换成功，优先人工修补版式而不是整体重做。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code == "HIGH_RISK_ENVIRONMENTS":
            envs = details.get("environments", {})
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=260,
                    category="high_risk_objects",
                    title="核对高风险环境的转换结果",
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary="预检查阶段发现高风险环境，转换结果需要人工确认。",
                    why_it_matters="子图、算法、代码环境、TikZ 等对象即使主体内容保住，也经常在 Word 中出现版式退化或结构弱化。",
                    recommended_actions=[
                        "优先检查子图、算法、代码块和 TikZ 图。",
                        "若对象已降级，请确认降级结果仍可读、可交付。",
                        "对关键结果图或关键算法说明，应优先进行版式修整。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details={"environments": envs},
                )
            )
            seq += 1

        elif code == "DUPLICATE_LABEL":
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=180,
                    category="cross_references",
                    title="修复重复标签引起的交叉引用风险",
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary="预检查阶段发现重复标签，图表或公式引用可能指向不稳定。",
                    why_it_matters="重复标签会让 LaTeX 源侧和 Word 侧的引用恢复都变得不可靠，属于后续交叉引用问题的根源之一。",
                    recommended_actions=[
                        "回到源工程中修复重复 label，并重新执行流水线。",
                        "若当前只做应急交付，请至少在 Word 中核对所有相关图表公式引用是否指向正确对象。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"UNDEFINED_LABEL_REFERENCE", "UNDEFINED_CITATION_KEY"}:
            category = "cross_references" if code == "UNDEFINED_LABEL_REFERENCE" else "bibliography"
            title = "处理未定义的内部引用" if code == "UNDEFINED_LABEL_REFERENCE" else "处理未定义的文献引用键"
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=170,
                    category=category,
                    title=title,
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary=f"预检查阶段发现 `{code}`，说明源工程本身存在引用完整性问题。",
                    why_it_matters="若源工程引用系统本身不完整，即使 Word 文档生成成功，也很难保证交叉引用或文献引用最终正确。",
                    recommended_actions=[
                        "优先修复源工程中的未定义引用问题。",
                        "修复后重新运行至少 precheck -> normalize -> convert -> postcheck 四步。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"CUSTOM_ENVIRONMENTS_DETECTED", "CUSTOM_COMMANDS_DETECTED", "UNKNOWN_ENVIRONMENTS_DETECTED"}:
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=320,
                    category="custom_macros",
                    title="核对自定义命令/环境的残余影响",
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary="预检查阶段检测到自定义命令、环境或未知环境，后续转换可能做了部分降级或跳过。",
                    why_it_matters="自定义结构往往不是 Pandoc 的稳定输入，一旦未被安全展开，最终 Word 里可能出现结构弱化或内容表现异常。",
                    recommended_actions=[
                        "结合 normalization 报告查看哪些自定义命令/环境被跳过或未安全展开。",
                        "在 Word 中重点检查相应章节或对象块是否仍然完整、可读。",
                        "若问题明显，优先回源工程做更明确的结构化改写后再重转。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details=details,
                )
            )
            seq += 1

        elif code == "CONDITIONAL_COMPILATION_DETECTED":
            items.append(
                make_item(
                    item_id=f"P{seq:02d}",
                    priority=360,
                    category="conditional_build",
                    title="确认条件编译分支是否正确展开",
                    source_stage="precheck",
                    source_code=code,
                    location=location,
                    issue_summary="工程依赖条件编译，当前转换结果可能只对应某一个展开分支。",
                    why_it_matters="若条件编译分支不对，Word 文档可能少章节、少图表或少附录，且这种缺失不一定会在版式层面立刻暴露。",
                    recommended_actions=[
                        "对照原始 PDF 或预期编译配置，确认当前 Word 版本包含正确的章节与附录。",
                        "若当前分支不对，先在源工程中固定正确分支后再重转。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details=details,
                )
            )
            seq += 1

        else:
            # 对 precheck 的其他 INFO/WARN/ERROR 不一概转入清单，避免噪声。
            # 只有真正与后续人工修复有关的项才入清单。
            _ = severity  # 保持变量已使用，便于静态阅读。
            continue

    return items


def items_from_normalization(normalization_report: Optional[dict]) -> list[ChecklistItem]:
    """
    将 normalization 阶段的 action 转为人工修复项。

    设计原则：
    - 只关注带有降级意味或“被跳过”的动作；
    - 常规安全规范化动作（如换行统一、补全扩展名）不进入人工清单；
    - 已经成功完成的规范化，不应再给用户增加无意义工作。
    """
    items: list[ChecklistItem] = []
    if not normalization_report:
        return items

    actions = normalization_report.get("actions", [])
    seq = 1
    for action in actions:
        severity = action.get("severity", "")
        action_type = action.get("action_type", "")
        location = action.get("file") or None
        details = action.get("details", {}) or {}

        if action_type in {"normalize_autoref", "normalize_cref"}:
            items.append(
                make_item(
                    item_id=f"N{seq:02d}",
                    priority=140,
                    category="cross_references",
                    title="核对由扩展引用命令降级而来的交叉引用",
                    source_stage="normalization",
                    source_code=action_type,
                    location=location,
                    issue_summary="规范化阶段已把扩展引用命令（如 autoref/cref）保守降级为基础 ref 形式。",
                    why_it_matters="这有利于主体转换，但对象类型前缀和部分复杂引用语义可能变弱，最终仍需人工确认关键引用是否可接受。",
                    recommended_actions=[
                        "重点检查图、表、公式的交叉引用文本是否仍符合期望。",
                        "若对象类型前缀缺失或表述不自然，可在 Word 中局部手工调整。",
                        "若关键引用跳转失效，请重建 Word 交叉引用。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details=details,
                )
            )
            seq += 1

        elif action_type.startswith("skip_"):
            items.append(
                make_item(
                    item_id=f"N{seq:02d}",
                    priority=300,
                    category="custom_macros",
                    title="检查被跳过的复杂宏定义或复杂结构",
                    source_stage="normalization",
                    source_code=action_type,
                    location=location,
                    issue_summary="规范化阶段为避免误改正文，跳过了某些复杂自定义命令或结构。",
                    why_it_matters="被跳过并不等于已处理；这些对象若参与正文结构或关键对象排版，可能直接影响最终 Word 表现。",
                    recommended_actions=[
                        "回看 normalization 报告中该 action 的 details，确认具体被跳过了什么。",
                        "在 Word 中定位相关内容块，检查是否存在结构弱化、丢格式或内容异常。",
                        "若问题明显，优先回源工程做更显式的重写后再重转。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details=details,
                )
            )
            seq += 1

        elif action_type == "override_safe_zero_arg_macro":
            items.append(
                make_item(
                    item_id=f"N{seq:02d}",
                    priority=340,
                    category="custom_macros",
                    title="核对被覆盖的零参数宏展开结果",
                    source_stage="normalization",
                    source_code=action_type,
                    location=location,
                    issue_summary="同名零参数宏在不同文件中发生了覆盖，最终展开结果需要人工确认。",
                    why_it_matters="同名宏覆盖可能导致不同章节出现不同替换效果，若宏承载术语或符号语义，容易形成局部不一致。",
                    recommended_actions=[
                        "定位相关章节，检查关键术语、符号或短宏展开后的文本是否正确。",
                        "若发现同一术语在不同位置表现不一致，建议回源工程统一宏定义后重新转换。",
                    ],
                    affects_delivery=DELIVERY_LOW,
                    details=details,
                )
            )
            seq += 1

        else:
            # 常规成功动作不进入人工清单。
            _ = severity
            continue

    return items


def items_from_conversion(conversion_report: Optional[dict]) -> list[ChecklistItem]:
    """
    将 Pandoc 主转换阶段的告警转为人工修复项。

    说明：
    - conversion 报告中的 warnings 通常是人类可读文本；
    - 这里不试图过度语义理解，而是将其转为用户应关注的检查点；
    - 若 conversion 阶段本身 FAIL，则更适合由 postcheck / 总体状态承接，
      这里只在仍生成了报告时补充用户动作建议。
    """
    items: list[ChecklistItem] = []
    if not conversion_report:
        return items

    warnings = conversion_report.get("warnings", [])
    seq = 1
    for warning in warnings:
        text = str(warning).strip()
        if not text:
            continue

        # 尽量按关键词给出更具体分类；否则归入 conversion。
        lowered = text.lower()
        if "bibliograph" in lowered or "引用" in text or "参考文献" in text:
            category = "bibliography"
            title = "核对参考文献主转换告警"
            priority = 160
            impact = DELIVERY_HIGH
        elif "resource-path" in lowered or "目录" in text:
            category = "resources"
            title = "核对资源搜索路径相关告警"
            priority = 500
            impact = DELIVERY_LOW
        else:
            category = "conversion"
            title = "处理 Pandoc 主转换阶段的告警"
            priority = 280
            impact = DELIVERY_MEDIUM

        items.append(
            make_item(
                item_id=f"C{seq:02d}",
                priority=priority,
                category=category,
                title=title,
                source_stage="conversion",
                source_code="conversion_warning",
                location="pandoc-conversion-report.json",
                issue_summary=text,
                why_it_matters="主转换阶段的告警往往意味着文档虽然生成成功，但某些对象的质量仍不可靠，后续需要重点核查。",
                recommended_actions=[
                    "打开 pandoc-conversion-report.md 查看该告警对应的上下文。",
                    "结合 postcheck 报告，确认相关对象是否真的在 docx 中出现质量问题。",
                    "若该告警与关键图表、公式或参考文献相关，应提高其修复优先级。",
                ],
                affects_delivery=impact,
                details={"warning_text": text},
            )
        )
        seq += 1

    return items


def items_from_postcheck(postcheck_report: Optional[dict]) -> list[ChecklistItem]:
    """
    将 postcheck 阶段的 finding 转为人工修复项。

    这是清单生成最重要的来源，因为 postcheck 已经是“结构级结果判断”。
    因此这里会更积极地把 ERROR/WARN 转成人工待办项。
    """
    items: list[ChecklistItem] = []
    if not postcheck_report:
        return items

    findings = postcheck_report.get("findings", [])
    seq = 1
    for finding in findings:
        severity = finding.get("severity", "")
        code = finding.get("code", "")
        location = finding.get("location") or None
        details = finding.get("details", {}) or {}

        if code in {"HEADINGS_MISSING", "HEADINGS_POSSIBLY_INCOMPLETE"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=80,
                    category="headings",
                    title="修复标题层级或目录基础",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示标题层级可能丢失或不完整。",
                    why_it_matters="若标题层级不对，目录、导航、章节编号和整体文档可维护性都会受到直接影响。",
                    recommended_actions=[
                        "在 Word 中抽查各级标题，确认其仍是标题样式。",
                        "若标题被转成普通段落，请将关键章节标题恢复为正确的 Word 标题样式。",
                        "恢复标题样式后重新更新目录与字段。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code == "UNRENDERABLE_IMAGE_FORMATS_DETECTED":
            unsupported_formats = details.get("unsupported_formats", [])
            if not isinstance(unsupported_formats, list):
                unsupported_formats = []
            format_text = ", ".join(str(item) for item in unsupported_formats if str(item).strip()) or "PDF/EPS"

            unsupported_ref_count = int(details.get("unsupported_reference_count_in_body", 0) or 0)
            unsupported_examples = format_unrenderable_image_examples(details, limit=8)
            example_text = "；".join(unsupported_examples)

            issue_summary = (
                f"后检查发现 {unsupported_ref_count} 处图片引用关联到 {format_text}，"
                "Word 通常不会渲染这类内嵌图片。"
            )
            if example_text:
                issue_summary += f" 已定位对象：{example_text}"

            recommended_actions = [
                "根据清单中的对象定位到源图，优先将 PDF/EPS 转为 PNG/JPG/EMF 等可渲染格式。",
                "替换后重新运行 convert -> postcheck -> checklist，确认该告警已消失。",
                "若暂时不能回源替换，请在 Word 中手工重插对应图片并复核题注与文内引用。",
            ]
            if unsupported_examples:
                recommended_actions.append(f"优先处理这些对象：{'; '.join(unsupported_examples)}")

            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=85,
                    category="images",
                    title="处理 Word 不可渲染的 PDF/EPS 图片",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary=issue_summary,
                    why_it_matters="这类图片在 Word 中往往直接空白，属于用户可见的高风险缺陷，必须人工收口。",
                    recommended_actions=recommended_actions,
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"IMAGES_MISSING_IN_BODY", "IMAGE_REFERENCE_COUNT_LOWER_THAN_EXPECTED", "NO_MEDIA_FILES_FOUND"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=90,
                    category="images",
                    title="修复缺失或数量异常的图片",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示图片嵌入数量异常，或正文中未检测到预期图片引用。",
                    why_it_matters="图片缺失会直接影响结果图、系统框图和示意图的交付完整性，通常属于必须优先处理的问题。",
                    recommended_actions=[
                        "在 Word 中逐个核对关键图片是否存在且显示正常。",
                        "若图片确实缺失，优先回源工程修正路径或补齐资源后重转。",
                        "若只是个别图片版式异常，可在 Word 中做局部修整，但不建议大面积手工重新插图。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"TABLES_MISSING", "TABLE_COUNT_LOWER_THAN_EXPECTED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=100,
                    category="tables",
                    title="修复缺失或数量异常的表格",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示表格数量低于源文档预期，或表格对象未成功进入 docx。",
                    why_it_matters="表格缺失或破坏不仅影响版式，还会直接导致数据不可用，通常需要优先处理。",
                    recommended_actions=[
                        "在 Word 中定位所有关键表格，确认是否存在、是否可读。",
                        "若普通表格缺失，优先回源工程与转换链修复后重转。",
                        "若仅复杂表格局部错位，可在 Word 中局部修复合并单元格、列宽和边框。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"NO_CAPTION_CANDIDATES_DETECTED", "CAPTION_COUNT_LOWER_THAN_EXPECTED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=110,
                    category="captions",
                    title="补查图题和表题完整性",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示图题/表题候选数量低于源文档预期。",
                    why_it_matters="图题和表题是图表编号、交叉引用以及图目录/表目录的基础；若丢失，后续维护成本很高。",
                    recommended_actions=[
                        "在 Word 中核对所有关键图和表是否都带有完整标题。",
                        "若个别标题变成普通文本，请在 Word 中重建为规范的 caption。",
                        "修复标题后，重新更新字段并检查图目录/表目录。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"MATH_OBJECTS_MISSING", "MATH_OBJECT_COUNT_LOWER_THAN_EXPECTED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=130,
                    category="equations",
                    title="修复缺失或异常的公式对象",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示 Word 数学对象数量低于源文档预期，或未检测到数学对象。",
                    why_it_matters="公式对象异常意味着部分公式可能丢失、图片化或不可编辑，这会直接影响技术文档的核心语义。",
                    recommended_actions=[
                        "优先检查文中最关键的编号公式和复杂公式。",
                        "若关键公式已损坏或图片化明显，建议回源工程后重转，而不是在 Word 中大面积重打。",
                        "若只是少量编号和排版问题，可在 Word 中局部修复编号与引用。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"NO_INTERNAL_REFERENCE_STRUCTURES_DETECTED", "BOOKMARKS_NOT_DETECTED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=150,
                    category="cross_references",
                    title="重建或核对关键交叉引用",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示内部引用结构不足，关键图表公式引用可能无法正确跳转。",
                    why_it_matters="内部交叉引用是 LaTeX 转 Word 最难完全自动恢复的部分，通常需要人工对关键引用做收口。",
                    recommended_actions=[
                        "优先核对正文中最关键的图、表、公式和附录引用。",
                        "对无法点击或指向错误的引用，在 Word 中重建交叉引用。",
                        "修完一轮后，再执行一次全字段更新并复查。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"TOC_FIELD_NOT_DETECTED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=200,
                    category="toc",
                    title="补建或更新目录字段",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查未检测到 TOC 字段，目录可能尚未建立或未正确保留。",
                    why_it_matters="目录是长文档的基本导航结构；若目录不存在，会显著降低 Word 文档的可用性与交付完整度。",
                    recommended_actions=[
                        "确认标题样式是否正确；若标题样式存在，可在 Word 中手动插入目录。",
                        "若目录已显示但无法更新，请先修复标题层级后再更新字段。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details=details,
                )
            )
            seq += 1

        elif code in {"SEQ_FIELD_NOT_DETECTED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=210,
                    category="captions",
                    title="补查图表编号的可维护性",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查未检测到 SEQ 字段，图表编号的 Word 原生可维护性可能不足。",
                    why_it_matters="即使图题文本存在，若编号体系不是可维护对象，后续插入或删除图表时编号更新会变得不可靠。",
                    recommended_actions=[
                        "优先核对关键图表的编号是否正确、是否能随字段更新。",
                        "若项目要求后续持续编辑，建议对关键图表重建 Word 原生 caption/编号对象。",
                    ],
                    affects_delivery=DELIVERY_MEDIUM,
                    details=details,
                )
            )
            seq += 1

        elif code in {"BIBLIOGRAPHY_SECTION_NOT_DETECTED", "BIBLIOGRAPHY_SECTION_EMPTY_OR_UNCLEAR"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=160,
                    category="bibliography",
                    title="补查参考文献章节与条目完整性",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示参考文献章节不存在、为空，或结构不清晰。",
                    why_it_matters="参考文献章节异常通常意味着文内引用无法完整落地，属于交付风险较高的问题。",
                    recommended_actions=[
                        "滚动到文末，确认参考文献章节是否存在且条目不为空。",
                        "若条目确实缺失，优先回源工程或 Pandoc 配置修复后重转。",
                        "若条目存在但标题不规范，可在 Word 中补正标题和章节层级。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"DOCX_NOT_FOUND", "DOCX_BAD_ZIP", "DOCX_NOT_OPENABLE", "DOCX_MISSING_BODY_XML", "DOCUMENT_XML_PARSE_FAILED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=5,
                    category="fatal_docx",
                    title="重新生成可用的 docx 文件",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查显示当前 docx 不可用或结构损坏。",
                    why_it_matters="若 docx 本体不可用，则任何人工修复都没有意义，必须先回到转换阶段重新生成可用文档。",
                    recommended_actions=[
                        "先查看 pandoc-conversion.log 和 pandoc-conversion-report.json。",
                        "修复阻塞问题后重新执行 convert_with_pandoc.py 和 postcheck_docx.py。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        elif code in {"UPSTREAM_CONVERSION_FAILED", "UPSTREAM_NORMALIZATION_FAILED"}:
            items.append(
                make_item(
                    item_id=f"D{seq:02d}",
                    priority=15,
                    category="pipeline",
                    title="回到上游阶段修复阻塞问题",
                    source_stage="postcheck",
                    source_code=code,
                    location=location,
                    issue_summary="后检查表明上游阶段本身已失败，当前 docx 不应被当作可靠结果。",
                    why_it_matters="如果上游阶段已失败，当前看到的许多问题只是表象；继续在 Word 中硬修通常性价比很低。",
                    recommended_actions=[
                        "先修复 normalization 或 conversion 阶段的 ERROR 级问题。",
                        "完成后重新运行后续阶段，再基于新的 postcheck 结果修正文档。",
                    ],
                    affects_delivery=DELIVERY_HIGH,
                    details=details,
                )
            )
            seq += 1

        else:
            # 对 postcheck 中其他 WARN/ERROR，生成通用人工检查项。
            # 这样能保证后检查阶段的信息不会静默丢失。
            if severity in {SEVERITY_WARN, SEVERITY_ERROR}:
                impact = DELIVERY_HIGH if severity == SEVERITY_ERROR else DELIVERY_MEDIUM
                priority = 400 if severity == SEVERITY_WARN else 190
                items.append(
                    make_item(
                        item_id=f"D{seq:02d}",
                        priority=priority,
                        category="postcheck_other",
                        title="处理 postcheck 阶段的未分类问题",
                        source_stage="postcheck",
                        source_code=code or "UNKNOWN_POSTCHECK_CODE",
                        location=location,
                        issue_summary=finding.get("message", "后检查报告中存在一个需要人工确认的问题。"),
                        why_it_matters="虽然该问题未被映射为特定类别，但已在后检查阶段被识别，说明它可能影响可交付性或可维护性。",
                        recommended_actions=[
                            "打开 postcheck-report.md 查看该问题的上下文。",
                            "在 Word 中定位对应对象并确认其是否影响交付。",
                            "若确认影响关键内容，应优先修复或回源工程重转。",
                        ],
                        affects_delivery=impact,
                        details=details,
                    )
                )
                seq += 1

    return items


# -----------------------------------------------------------------------------
# 清单汇总、排序与报告渲染
# -----------------------------------------------------------------------------

def build_summary_recommendations(
    status: str,
    item_count: int,
    high_impact_count: int,
    used_postcheck_report: bool,
) -> list[str]:
    """
    根据总体状态与清单规模，生成面向用户的收尾建议。

    这些建议不是具体问题项，而是对“如何使用这份清单”的元指导。
    """
    recommendations: list[str] = []

    if status == STATUS_FAIL:
        recommendations.append("当前清单仍已生成，但上游结果存在严重问题；建议优先回到上游阶段修复，而不是直接在 Word 中大面积硬修。")
    else:
        recommendations.append("按 priority 从小到大执行，不要从头逐页盲查。")

    recommendations.append("先完成 G 类全局步骤，再处理 D 类（postcheck）和其他阶段的定点问题。")

    if high_impact_count > 0:
        recommendations.append("优先处理 affects_delivery=high 的任务；这些问题最可能阻碍最终交付。")

    if used_postcheck_report:
        recommendations.append("每修复一轮后，建议在 Word 中再次全选并更新字段，然后做一次终审。")

    if item_count == 0:
        recommendations.append("当前没有生成任何人工修复项，但仍建议人工抽查目录、关键图表、关键公式和参考文献。 ")

    return recommendations


def render_markdown_report(report: ManualFixChecklistReport) -> str:
    """
    将人工修复清单渲染为 Markdown。

    设计目标：
    - 用户可以直接照着这份 Markdown 执行；
    - 清单项要有明显的优先级、来源和操作建议；
    - 不要求用户理解上游脚本的内部实现。
    """
    lines: list[str] = []

    lines.append("# Manual Fix Checklist")
    lines.append("")
    lines.append(f"- Status: **{report.status}**")
    lines.append(f"- Can continue: **{report.can_continue}**")
    lines.append(f"- Work root: `{report.work_root}`")
    lines.append(f"- Source project root: `{report.source_project_root or 'N/A'}`")
    lines.append(f"- Used precheck report: **{report.used_precheck_report}**")
    lines.append(f"- Used normalization report: **{report.used_normalization_report}**")
    lines.append(f"- Used conversion report: **{report.used_conversion_report}**")
    lines.append(f"- Used postcheck report: **{report.used_postcheck_report}**")
    lines.append(f"- User view generated: **{report.user_view_generated}**")
    lines.append(f"- User view root: `{report.user_view_root or 'N/A'}`")
    lines.append(f"- Published file count: **{report.published_file_count}**")
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

    lines.append("## Execution Order")
    lines.append("")
    lines.append("建议按以下顺序执行：")
    lines.append("")
    lines.append("1. 先处理 G 类全局步骤")
    lines.append("2. 再处理 D 类（postcheck）中的高优先级问题")
    lines.append("3. 再处理 P / N / C 类中的源侧或流程侧遗留问题")
    lines.append("4. 每修一轮后，在 Word 中再次更新全部字段并做抽查")
    lines.append("")

    lines.append("## Checklist Items")
    lines.append("")
    if not report.items:
        lines.append("- No checklist items generated.")
        lines.append("")
    else:
        for item in report.items:
            lines.append(f"### {item['item_id']} | priority={item['priority']} | {item['title']}")
            lines.append("")
            lines.append(f"- Category: **{item['category']}**")
            lines.append(f"- Source stage: **{item['source_stage']}**")
            lines.append(f"- Source code: **{item['source_code']}**")
            lines.append(f"- Location: `{item['location'] or 'N/A'}`")
            lines.append(f"- Affects delivery: **{item['affects_delivery']}**")
            lines.append(f"- Issue: {item['issue_summary']}")
            lines.append(f"- Why it matters: {item['why_it_matters']}")
            lines.append("- Recommended actions:")
            if item["recommended_actions"]:
                for action in item["recommended_actions"]:
                    lines.append(f"  - {action}")
            else:
                lines.append("  - (none)")
            lines.append("")

    lines.append("## Recommendations")
    lines.append("")
    if report.recommendations:
        for recommendation in report.recommendations:
            lines.append(f"- {recommendation}")
    else:
        lines.append("- No additional recommendations.")
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
    2. 解析四类报告路径；
    3. 读取报告；
    4. 生成全局步骤；
    5. 从 precheck / normalization / conversion / postcheck 各阶段提取人工任务；
    6. 去重、排序；
    7. 生成 JSON / Markdown 输出；
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
    checklist_stage_dir = stage_dir(work_root, STAGE_CHECKLIST)
    checklist_stage_dir.mkdir(parents=True, exist_ok=True)

    (
        precheck_json_path,
        normalization_json_path,
        conversion_json_path,
        postcheck_json_path,
        source_project_root,
    ) = resolve_report_paths(
        work_root=work_root,
        precheck_arg=args.precheck_json,
        normalization_arg=args.normalization_json,
        conversion_arg=args.conversion_json,
        postcheck_arg=args.postcheck_json,
    )

    json_out = resolve_explicit_or_stage_output(
        args.json_out,
        work_root,
        STAGE_CHECKLIST,
        "manual-fix-checklist.json",
    )
    md_out = resolve_explicit_or_stage_output(
        args.md_out,
        work_root,
        STAGE_CHECKLIST,
        "manual-fix-checklist.md",
    )

    precheck_report = load_json_if_exists(precheck_json_path) if precheck_json_path else None
    normalization_report = load_json_if_exists(normalization_json_path) if normalization_json_path else None
    conversion_report = load_json_if_exists(conversion_json_path) if conversion_json_path else None
    postcheck_report = load_json_if_exists(postcheck_json_path) if postcheck_json_path else None

    used_precheck_report = precheck_report is not None
    used_normalization_report = normalization_report is not None
    used_conversion_report = conversion_report is not None
    used_postcheck_report = postcheck_report is not None

    # 若 normalization 缺失，则无法恢复 source project root；此时仍尽量继续。
    if source_project_root is None and normalization_report and normalization_report.get("project_root"):
        source_project_root = Path(normalization_report["project_root"]).resolve()

    # source_inventory 优先来自 postcheck（它已经结合 normalization 统计过一次）；
    # 若 postcheck 不存在，再尝试从 precheck / normalization 中补出最基础数据。
    source_inventory = {}
    if postcheck_report and isinstance(postcheck_report.get("source_inventory"), dict):
        source_inventory = dict(postcheck_report["source_inventory"])
    else:
        # 最小后备 inventory，避免全局步骤生成完全失去依据。
        source_inventory = {
            "heading_command_count": 0,
            "caption_command_count": 0,
            "ref_count": 0,
            "cite_count": 0,
            "table_env_count": 0,
            "image_command_count": 0,
            "equation_env_count": 0,
        }
        if precheck_report and isinstance(precheck_report.get("metrics"), dict):
            metrics = precheck_report["metrics"]
            source_inventory["ref_count"] = int(metrics.get("ref_count", 0))
            source_inventory["cite_count"] = int(metrics.get("cite_count", 0))
            # precheck 没有 caption / heading / table / image 的完整细分，不强行猜测。
        if normalization_report and isinstance(normalization_report.get("summary"), dict):
            # normalization 并不提供这些统计，因此不额外填充。
            _ = normalization_report["summary"]

    # 清单项收集
    items: list[ChecklistItem] = []
    seen: set[tuple[str, str, str, str]] = set()

    # 先加规则文件缺失提示（若有）
    for item in check_rule_files(skill_root):
        append_dedup(items, seen, item)

    # 再加全局固定步骤
    for item in build_global_items(source_inventory, postcheck_report):
        append_dedup(items, seen, item)

    # 再加四阶段衍生项
    for item in items_from_precheck(precheck_report):
        append_dedup(items, seen, item)

    for item in items_from_normalization(normalization_report):
        append_dedup(items, seen, item)

    for item in items_from_conversion(conversion_report):
        append_dedup(items, seen, item)

    for item in items_from_postcheck(postcheck_report):
        append_dedup(items, seen, item)

    # 排序策略：
    # 1. priority 升序
    # 2. item_id 升序，保证输出稳定
    items.sort(key=lambda item: (item.priority, item.item_id))

    high_impact_count = sum(1 for item in items if item.affects_delivery == DELIVERY_HIGH)
    medium_impact_count = sum(1 for item in items if item.affects_delivery == DELIVERY_MEDIUM)
    low_impact_count = sum(1 for item in items if item.affects_delivery == DELIVERY_LOW)
    none_impact_count = sum(1 for item in items if item.affects_delivery == DELIVERY_NONE)

    # 状态判定：
    # - 若四类报告全部缺失，则 FAIL；
    # - 若 postcheck 缺失，但其余存在，则 PASS_WITH_WARNINGS；
    # - 若清单已生成，且至少有 normalization / conversion / postcheck 之一，则 PASS 或 PASS_WITH_WARNINGS。
    if not any([used_precheck_report, used_normalization_report, used_conversion_report, used_postcheck_report]):
        status = STATUS_FAIL
        can_continue = False
    elif not used_postcheck_report:
        status = STATUS_PASS_WITH_WARNINGS
        can_continue = True
    else:
        # 若 postcheck 存在，并且其状态为 FAIL，则清单仍然可生成，但不宜声称流程完全通过。
        if postcheck_report and postcheck_report.get("status") == STATUS_FAIL:
            status = STATUS_PASS_WITH_WARNINGS
            can_continue = True
        else:
            status = STATUS_PASS
            can_continue = True

    summary = {
        "status": status,
        "can_continue": can_continue,
        "item_count": len(items),
        "high_impact_count": high_impact_count,
        "medium_impact_count": medium_impact_count,
        "low_impact_count": low_impact_count,
        "none_impact_count": none_impact_count,
        "user_view_generated": False,
        "published_file_count": 0,
    }

    metrics = {
        "used_precheck_report": used_precheck_report,
        "used_normalization_report": used_normalization_report,
        "used_conversion_report": used_conversion_report,
        "used_postcheck_report": used_postcheck_report,
        "global_item_count": sum(1 for item in items if item.source_stage == "global"),
        "precheck_item_count": sum(1 for item in items if item.source_stage == "precheck"),
        "normalization_item_count": sum(1 for item in items if item.source_stage == "normalization"),
        "conversion_item_count": sum(1 for item in items if item.source_stage == "conversion"),
        "postcheck_item_count": sum(1 for item in items if item.source_stage == "postcheck"),
        "total_item_count": len(items),
        "user_view_warning_count": 0,
        "user_view_published_file_count": 0,
    }

    recommendations = build_summary_recommendations(
        status=status,
        item_count=len(items),
        high_impact_count=high_impact_count,
        used_postcheck_report=used_postcheck_report,
    )

    if not used_postcheck_report:
        recommendations.append("建议先运行 postcheck_docx.py，再重新生成一次人工修复清单，这样优先级会更准确。")

    if not used_precheck_report:
        recommendations.append("缺少 precheck 报告时，源工程侧的风险信息会不完整；若需要回源修复，建议补跑 precheck.py。")

    report = ManualFixChecklistReport(
        status=status,
        can_continue=can_continue,
        work_root=str(work_root.resolve()),
        source_project_root=str(source_project_root.resolve()) if source_project_root else None,
        used_precheck_report=used_precheck_report,
        used_normalization_report=used_normalization_report,
        used_conversion_report=used_conversion_report,
        used_postcheck_report=used_postcheck_report,
        items=[asdict(item) for item in items],
        metrics=metrics,
        summary=summary,
        recommendations=recommendations,
        user_view_generated=False,
        user_view_root=None,
        published_file_count=0,
    )

    # 先写一次清单文件，确保后续“用户视图发布”可以直接复用当前产物。
    report_dict = asdict(report)
    write_json(json_out, report_dict)
    write_markdown(md_out, render_markdown_report(report))

    user_view_warnings: list[str] = []
    readme_run_path: Optional[Path] = None
    if not args.no_user_view:
        user_view_generated, user_view_root, published_file_count, user_view_warnings = publish_user_view(
            work_root=work_root,
            precheck_json_path=precheck_json_path,
            normalization_json_path=normalization_json_path,
            conversion_json_path=conversion_json_path,
            postcheck_json_path=postcheck_json_path,
            checklist_json_path=json_out,
            checklist_md_path=md_out,
            conversion_report=conversion_report,
        )

        precheck_status = precheck_report.get("status") if precheck_report else "N/A"
        normalization_status = normalization_report.get("status") if normalization_report else "N/A"
        conversion_status = conversion_report.get("status") if conversion_report else "N/A"
        postcheck_status = postcheck_report.get("status") if postcheck_report else "N/A"

        readme_run_path = write_run_readme(
            work_root=work_root,
            checklist_status=status,
            precheck_status=str(precheck_status),
            normalization_status=str(normalization_status),
            conversion_status=str(conversion_status),
            postcheck_status=str(postcheck_status),
            user_view_generated=user_view_generated,
            published_file_count=published_file_count,
            user_view_warnings=user_view_warnings,
        )

        report.user_view_generated = user_view_generated
        report.user_view_root = str(user_view_root)
        report.published_file_count = published_file_count
        report.summary["user_view_generated"] = user_view_generated
        report.summary["published_file_count"] = published_file_count
        report.metrics["user_view_warning_count"] = len(user_view_warnings)
        report.metrics["user_view_published_file_count"] = published_file_count

        if user_view_generated:
            report.recommendations.append("已生成用户视图目录（deliverables/reports/logs/debug），可直接按 README_RUN.md 导航阅读。")
        else:
            report.recommendations.append("用户视图目录未生成有效产物；请先检查上游报告与输出文件是否存在。")

        if user_view_warnings:
            report.recommendations.append(
                f"用户视图生成时出现 {len(user_view_warnings)} 个非阻塞告警，请查看 README_RUN.md 的告警段落。"
            )
    else:
        report.recommendations.append("已按参数 --no-user-view 跳过用户视图目录生成。")

    # 回写最终版报告（包含用户视图生成结果）。
    top_level_artifacts = {
        "reports": {
            "manual_fix_checklist_json": json_out,
            "manual_fix_checklist_md": md_out,
        }
    }
    if readme_run_path is not None:
        top_level_artifacts["deliverables"] = {"readme_run": readme_run_path}

    persist_stage_report(
        work_root=work_root,
        stage=STAGE_CHECKLIST,
        report_obj=report,
        markdown_text=render_markdown_report(report),
        report_json_path=json_out,
        report_md_path=md_out,
        status=report.status,
        can_continue=report.can_continue,
        artifacts={
            "manual_fix_checklist_json": json_out,
            "manual_fix_checklist_md": md_out,
            "readme_run": readme_run_path if readme_run_path else None,
        },
        summary=report.summary,
        metrics=report.metrics,
        notes=report.recommendations,
        top_level_artifacts=top_level_artifacts,
    )

    print(f"[{report.status}] manual fix checklist generated.")
    print(f"Checklist items: {len(items)}")
    print(f"High impact items: {high_impact_count}")
    print(f"User view generated: {report.user_view_generated}")
    print(f"Published file count: {report.published_file_count}")
    if readme_run_path is not None:
        print(f"User view README: {readme_run_path}")
    print(f"JSON report: {json_out}")
    print(f"Markdown report: {md_out}")

    return 1 if report.status == STATUS_FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
