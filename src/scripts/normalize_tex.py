#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
normalize_tex.py

功能概述
--------
本脚本用于在 LaTeX -> Word 主转换之前，对 LaTeX 工程执行“保守规范化”。

它解决的问题包括：
1. 绝不覆盖原始工程，而是在独立工作目录中生成规范化副本；
2. 对确定性高、低风险的写法做标准化，降低 Pandoc 主转换的不稳定性；
3. 为后续步骤保留一份可追踪的“我改了什么、为什么改、改到了哪里”的报告。

本脚本当前实现的规范化动作
--------------------------
1. 复制原始工程到独立工作目录；
2. 统一文本文件换行风格为 LF（仅作用于工作副本中的 .tex 文件）；
3. 为 \\includegraphics 中缺失扩展名的图片路径补全实际扩展名；
4. 为 \\addbibresource / \\bibliography 中的 bibliography 资源补全 .bib 扩展名；
5. 将常见扩展交叉引用命令做保守降级：
   - \\autoref{...} -> \\ref{...}
   - \\cref{a,b} / \\Cref{a,b} -> \\ref{a}, \\ref{b}
   - \\subref{...} -> \\ref{...}
6. 在 align 系环境中，当检测到“行首 \\label”时，将其调整到该行公式尾部；
7. 在数学上下文中规范化旧式命令（如 \\rm/\\bf/\\cal/\\buildrel）；
8. 在 figure / table 环境中，当检测到“\\label 紧邻且位于 \\caption 前”时，
   将其调整为“\\caption 后紧跟 \\label”；
9. 将 ``table*`` 环境保守降级为 ``table``（保留内容与可选参数不变）；
10. 将 ``algorithm / algorithm*``（含 algorithm2e 语法）降级为 Pandoc 稳定块结构；
11. 安全展开“零参数且不含 # 占位符”的简单自定义命令：
   - 支持 \\newcommand / \\renewcommand / \\providecommand
   - 仅展开零参数、短小、无参数占位符的定义
   - 不覆盖原定义，只在正文使用位置做文本替换
12. 生成结构化 JSON 报告与 Markdown 报告。

设计边界
--------
本脚本刻意不做以下事情：
1. 不改写正文语义；
2. 不重构章节结构；
3. 不删除用户内容；
4. 不做激进宏展开；
5. 不尝试自动修复复杂 TikZ、复杂表格、复杂 theorem 系统；
6. 不直接调用 Pandoc；
7. 不修改原始工程。

依赖
----
- Python 3.9+
- 仅使用标准库，不依赖第三方包

典型用法
--------
python scripts/normalize_tex.py --project-root D:/work/my-paper --main-tex main.tex --force

或依赖 precheck-report.json 自动获取主文件：

python scripts/normalize_tex.py --project-root D:/work/my-paper --force

输出
----
默认会生成一个独立工作目录，例如：
<project-root-parent>/<project-name>__latex_to_word_work/

其中包含阶段化输出：
- `stage_normalize/source_snapshot/`（规范化后的工程副本）
- `stage_normalize/normalization-report.json`
- `stage_normalize/normalization-report.md`

退出码
------
- 0: 成功生成规范化工作副本
- 1: 失败
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Optional

from pipeline_common import (
    load_json_if_exists,
    locate_skill_root,
    read_text_file,
    safe_relative,
    split_csv_payload,
    write_json,
    write_text_file,
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
    STAGE_NORMALIZE,
    STAGE_PRECHECK,
    default_work_root_for_project,
    resolve_explicit_or_default,
    resolve_explicit_or_stage_input,
    stage_dir,
)
from stage_reporting import persist_stage_report
from tex_scan_common import (
    line_number_from_offset,
    parse_balanced_group,
    resolve_path_with_extensions,
    skip_whitespace,
)


# -----------------------------------------------------------------------------
# 常量定义
# -----------------------------------------------------------------------------

IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".pdf", ".svg", ".eps", ".bmp", ".tif", ".tiff"]
BIB_EXTENSIONS = [".bib"]
MATH_BLOCK_ENV_NAMES = (
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
)

# 默认忽略复制的目录/文件名。
# 这些内容对 LaTeX -> Word 主流程没有帮助，反而会污染工作副本。
DEFAULT_COPY_IGNORE_NAMES = {
    ".git",
    ".idea",
    ".vscode",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".DS_Store",
}

# -----------------------------------------------------------------------------
# 数据结构定义
# -----------------------------------------------------------------------------

@dataclass
class ActionRecord:
    """
    表示一次实际发生的规范化动作。

    字段说明
    --------
    action_type:
        稳定、机器可读的动作类型，例如：
        - normalize_newlines
        - add_graphics_extension
        - normalize_autoref
        - expand_zero_arg_macro

    severity:
        规范化动作的性质：
        - INFO: 常规安全规范化
        - WARN: 带有降级含义或需后续重点检查的规范化
        - ERROR: 仅用于记录导致该文件无法正常处理的情况

    file:
        目标文件，相对工程根目录或工作根目录的相对路径字符串。

    line:
        近似行号。由于规范化过程中内容可能移动，行号只要求“足够用于人工定位”。

    message:
        面向人的简短说明。

    details:
        结构化补充信息，供 JSON 报告和后续脚本消费。
    """
    action_type: str
    severity: str
    file: str
    line: Optional[int]
    message: str
    details: dict = field(default_factory=dict)


@dataclass
class FileSummary:
    """
    单个 TeX 文件的规范化摘要。
    """
    file: str
    modified: bool
    action_count: int
    actions_by_type: dict


@dataclass
class NormalizationReport:
    """
    规范化阶段的最终报告对象。

    设计目标：
    - 既方便 JSON 结构化消费，也方便 Markdown 渲染；
    - 清楚说明是否使用了 precheck 结果；
    - 清楚说明工作副本在哪里；
    - 清楚说明实际改写了哪些文件、做了哪些动作。
    """
    status: str
    can_continue: bool
    used_precheck_report: bool
    project_root: str
    work_root: str
    source_main_tex: Optional[str]
    normalized_main_tex: Optional[str]
    tex_files_processed: list[str]
    tex_files_processed_count: int
    tex_files_modified_count: int
    actions: list[dict]
    file_summaries: list[dict]
    metrics: dict
    summary: dict
    recommendations: list[str]


# -----------------------------------------------------------------------------
# 通用工具函数
# -----------------------------------------------------------------------------

def build_argument_parser() -> argparse.ArgumentParser:
    """
    构造命令行参数解析器。

    参数设计原则：
    - 只暴露规范化阶段真正需要的输入；
    - 默认行为稳定、保守；
    - 不引入与当前职责无关的参数。
    """
    parser = argparse.ArgumentParser(
        description="Create a normalized LaTeX working copy for LaTeX -> Word conversion."
    )
    parser.add_argument(
        "--project-root",
        required=True,
        help="原始 LaTeX 工程根目录。",
    )
    parser.add_argument(
        "--main-tex",
        default=None,
        help="主 TeX 文件路径；可为相对 project-root 的路径。若不提供，优先从 precheck-report.json 读取。",
    )
    parser.add_argument(
        "--precheck-json",
        default=None,
        help="预检查 JSON 报告路径；默认优先 <work-root>/stage_precheck/precheck-report.json，再回退 <project-root>/precheck-report.json。",
    )
    parser.add_argument(
        "--work-root",
        default=None,
        help="规范化工作目录；默认生成到 <project-root-parent>/<project-name>__latex_to_word_work。",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="若工作目录已存在，则先删除再重建。默认不覆盖。",
    )
    return parser


def check_rule_files(skill_root: Path) -> list[ActionRecord]:
    """
    检查规则文件是否存在。

    normalize 阶段不会解析规则文件内容，但若 skill 目录损坏，应尽早通过报告暴露。
    """
    actions: list[ActionRecord] = []
    for relative in REQUIRED_RULE_FILES:
        candidate = (skill_root / relative).resolve()
        if not candidate.exists():
            actions.append(
                ActionRecord(
                    action_type="missing_rule_file",
                    severity=SEVERITY_WARN,
                    file=safe_relative(candidate, skill_root),
                    line=None,
                    message="缺少规则文件；后续流程行为可能不完整。",
                    details={"required_file": relative},
                )
            )
    return actions


# -----------------------------------------------------------------------------
# 主文件与工作目录解析
# -----------------------------------------------------------------------------

def resolve_main_tex(
    project_root: Path,
    main_tex_arg: Optional[str],
    precheck_report: Optional[dict],
) -> tuple[Optional[Path], list[ActionRecord]]:
    """
    解析主 TeX 文件，优先级如下：
    1. 命令行参数 --main-tex
    2. precheck-report.json 中的 main_tex
    3. 常见主文件名启发式：main.tex
    4. 含 \\documentclass 与 \\begin{document} 的 .tex 文件

    返回：
    - 主文件绝对路径或 None
    - 附加动作记录（主要是告警/信息）
    """
    actions: list[ActionRecord] = []

    if main_tex_arg:
        candidate = Path(main_tex_arg)
        if not candidate.is_absolute():
            candidate = (project_root / candidate).resolve()
        if not candidate.exists():
            actions.append(
                ActionRecord(
                    action_type="resolve_main_tex_failed",
                    severity=SEVERITY_ERROR,
                    file=safe_relative(candidate, project_root),
                    line=None,
                    message="指定的主 TeX 文件不存在。",
                    details={"source": "cli_argument"},
                )
            )
            return None, actions
        return candidate.resolve(), actions

    if precheck_report and precheck_report.get("main_tex"):
        candidate = (project_root / precheck_report["main_tex"]).resolve()
        if candidate.exists():
            actions.append(
                ActionRecord(
                    action_type="resolve_main_tex_from_precheck",
                    severity=SEVERITY_INFO,
                    file=safe_relative(candidate, project_root),
                    line=None,
                    message="主 TeX 文件来自 precheck-report.json。",
                    details={"source": "precheck_report"},
                )
            )
            return candidate, actions

    main_candidate = (project_root / "main.tex").resolve()
    if main_candidate.exists():
        actions.append(
            ActionRecord(
                action_type="resolve_main_tex_by_name",
                severity=SEVERITY_WARN,
                file=safe_relative(main_candidate, project_root),
                line=None,
                message="未显式提供主文件，已按常见命名约定选择 main.tex。",
                details={"source": "heuristic_main_tex_name"},
            )
        )
        return main_candidate, actions

    tex_files = sorted(project_root.rglob("*.tex"))
    best_score = -1
    best_path: Optional[Path] = None
    for tex_file in tex_files:
        try:
            text = read_text_file(tex_file)
        except Exception:
            continue
        score = 0
        if re.search(r"\\documentclass(?:\[[^\]]*\])?\{[^}]+\}", text):
            score += 5
        if r"\begin{document}" in text:
            score += 5
        if score > best_score:
            best_score = score
            best_path = tex_file.resolve()

    if best_path is not None and best_score > 0:
        actions.append(
            ActionRecord(
                action_type="resolve_main_tex_by_content",
                severity=SEVERITY_WARN,
                file=safe_relative(best_path, project_root),
                line=None,
                message="未显式提供主文件，已按文档结构启发式自动选择主文件。",
                details={"source": "heuristic_document_structure"},
            )
        )
        return best_path, actions

    actions.append(
        ActionRecord(
            action_type="resolve_main_tex_failed",
            severity=SEVERITY_ERROR,
            file=safe_relative(project_root, project_root),
            line=None,
            message="无法确定主 TeX 文件，请显式传入 --main-tex 或先运行 precheck.py。",
            details={"source": "all_failed"},
        )
    )
    return None, actions


def default_work_root_for(project_root: Path) -> Path:
    """
    计算默认工作目录。

    采用“工程同级目录”的形式，而不是把工作副本放进工程目录内部，
    原因如下：
    1. 避免递归复制自身；
    2. 避免污染原始工程；
    3. 更容易一眼区分“源工程”和“规范化副本”。
    """
    return default_work_root_for_project(project_root)


def ensure_work_root_is_safe(project_root: Path, work_root: Path) -> Optional[str]:
    """
    检查工作目录是否安全。

    安全要求：
    - 不能与 project_root 相同；
    - 不建议位于 project_root 内部，否则复制过程会递归污染；
    - 也不能让 project_root 位于 work_root 内部导致覆盖混乱。

    返回：
    - 若安全，返回 None
    - 若不安全，返回错误消息
    """
    project_root_resolved = project_root.resolve()
    work_root_resolved = work_root.resolve()

    if project_root_resolved == work_root_resolved:
        return "工作目录不能与原始工程目录相同。"

    # work_root 不能位于 project_root 内部
    try:
        work_root_resolved.relative_to(project_root_resolved)
        return "工作目录不能位于原始工程目录内部，否则复制过程可能递归污染工程。"
    except ValueError:
        pass

    # project_root 不能位于 work_root 内部
    try:
        project_root_resolved.relative_to(work_root_resolved)
        return "原始工程目录不能位于工作目录内部。"
    except ValueError:
        pass

    return None


# -----------------------------------------------------------------------------
# 复制与目标文件选择
# -----------------------------------------------------------------------------

def copy_project_tree(project_root: Path, work_root: Path) -> None:
    """
    将原始工程复制到工作目录。

    实现说明：
    - 使用 shutil.copytree 保持目录结构；
    - 忽略常见缓存目录与开发环境目录；
    - 不在这里做任何内容修改，只负责复制。
    """
    def ignore_func(_src: str, names: list[str]) -> set[str]:
        ignored = set()
        for name in names:
            if name in DEFAULT_COPY_IGNORE_NAMES:
                ignored.add(name)
        return ignored

    shutil.copytree(project_root, work_root, ignore=ignore_func, copy_function=shutil.copy2)


def collect_target_tex_files(
    project_root: Path,
    work_root: Path,
    precheck_report: Optional[dict],
) -> list[Path]:
    """
    确定需要进行规范化处理的 TeX 文件集合。

    优先策略：
    1. 若 precheck-report.json 中存在 scanned_tex_files，则只处理主文件可达闭包；
    2. 否则处理工作副本中所有 .tex 文件。

    这样设计的原因：
    - 优先尊重 precheck 阶段已经确定的“有效输入闭包”；
    - 在没有 precheck 时，也保证脚本具备独立工作能力。
    """
    targets: list[Path] = []

    if precheck_report and isinstance(precheck_report.get("scanned_tex_files"), list):
        for relative in precheck_report["scanned_tex_files"]:
            candidate = (work_root / relative).resolve()
            if candidate.exists() and candidate.suffix.lower() == ".tex":
                targets.append(candidate)

    if targets:
        return sorted(set(targets))

    return sorted(work_root.rglob("*.tex"))


# -----------------------------------------------------------------------------
# 自定义命令收集与安全展开
# -----------------------------------------------------------------------------

@dataclass
class MacroDefinition:
    """
    表示一个可安全展开的零参数宏定义。
    """
    name: str
    replacement: str
    defined_in: str
    line: int


def collect_zero_arg_macros_from_text(text: str, file_path: Path, project_root: Path) -> tuple[dict[str, MacroDefinition], list[ActionRecord]]:
    """
    从单个 TeX 文件中收集“可安全展开”的零参数宏定义。

    支持的定义来源：
    - \\newcommand
    - \\renewcommand
    - \\providecommand

    只接受如下安全条件：
    1. 宏名规范；
    2. 参数个数为 0 或未显式给出；
    3. replacement 中不含 # 占位符；
    4. replacement 长度适中；
    5. 不含可选默认参数这类高风险结构。

    为什么如此保守：
    - 规范化阶段不应成为一个“宏编译器”；
    - 目标只是稳定主转换，不是重实现 TeX 宏系统。
    """
    macros: dict[str, MacroDefinition] = {}
    actions: list[ActionRecord] = []

    definition_head = re.compile(r"\\(?:newcommand|renewcommand|providecommand)\*?")

    pos = 0
    while True:
        match = definition_head.search(text, pos)
        if not match:
            break

        cursor = match.end()
        cursor = skip_whitespace(text, cursor)

        # 解析宏名。
        macro_name: Optional[str] = None
        if cursor < len(text) and text[cursor] == "{":
            end = parse_balanced_group(text, cursor, "{", "}")
            if end is not None:
                inner = text[cursor + 1:end - 1].strip()
                if re.fullmatch(r"\\[A-Za-z@]+", inner):
                    macro_name = inner[1:]
                cursor = end
            else:
                cursor += 1
        elif cursor < len(text) and text[cursor] == "\\":
            name_match = re.match(r"\\([A-Za-z@]+)", text[cursor:])
            if name_match:
                macro_name = name_match.group(1)
                cursor += len(name_match.group(0))

        if not macro_name:
            pos = match.end()
            continue

        cursor = skip_whitespace(text, cursor)

        # 解析可选参数个数 [n]
        arg_count = 0
        if cursor < len(text) and text[cursor] == "[":
            arg_end = parse_balanced_group(text, cursor, "[", "]")
            if arg_end is None:
                pos = match.end()
                continue
            arg_payload = text[cursor + 1:arg_end - 1].strip()
            if arg_payload.isdigit():
                arg_count = int(arg_payload)
                cursor = arg_end
            else:
                # 形如 [default] 或其他复杂写法，直接视为高风险，跳过展开。
                line_no = line_number_from_offset(text, match.start())
                actions.append(
                    ActionRecord(
                        action_type="skip_complex_macro_definition",
                        severity=SEVERITY_WARN,
                        file=safe_relative(file_path, project_root),
                        line=line_no,
                        message="跳过带复杂可选参数的自定义命令定义。",
                        details={"macro_name": macro_name},
                    )
                )
                pos = match.end()
                continue

        cursor = skip_whitespace(text, cursor)

        # 若紧跟第二个 []，说明这是“带默认参数值”的定义，风险较高，跳过。
        if cursor < len(text) and text[cursor] == "[":
            line_no = line_number_from_offset(text, match.start())
            actions.append(
                ActionRecord(
                    action_type="skip_macro_with_default_argument",
                    severity=SEVERITY_WARN,
                    file=safe_relative(file_path, project_root),
                    line=line_no,
                    message="跳过带默认参数值的自定义命令定义。",
                    details={"macro_name": macro_name},
                )
            )
            pos = match.end()
            continue

        cursor = skip_whitespace(text, cursor)

        # 解析 replacement body
        if cursor >= len(text) or text[cursor] != "{":
            pos = match.end()
            continue
        body_end = parse_balanced_group(text, cursor, "{", "}")
        if body_end is None:
            pos = match.end()
            continue

        replacement = text[cursor + 1:body_end - 1]
        line_no = line_number_from_offset(text, match.start())

        if arg_count == 0 and "#" not in replacement and len(replacement) <= 200:
            macros[macro_name] = MacroDefinition(
                name=macro_name,
                replacement=replacement,
                defined_in=safe_relative(file_path, project_root),
                line=line_no,
            )
            actions.append(
                ActionRecord(
                    action_type="collect_safe_zero_arg_macro",
                    severity=SEVERITY_INFO,
                    file=safe_relative(file_path, project_root),
                    line=line_no,
                    message="收集到可安全展开的零参数宏定义。",
                    details={"macro_name": macro_name},
                )
            )
        else:
            actions.append(
                ActionRecord(
                    action_type="skip_nonzero_or_complex_macro_definition",
                    severity=SEVERITY_WARN,
                    file=safe_relative(file_path, project_root),
                    line=line_no,
                    message="跳过非零参数或不安全的自定义命令定义。",
                    details={
                        "macro_name": macro_name,
                        "arg_count": arg_count,
                        "contains_parameter_marker": "#" in replacement,
                        "replacement_length": len(replacement),
                    },
                )
            )

        pos = body_end

    return macros, actions


def collect_zero_arg_macros(tex_files: list[Path], project_root: Path) -> tuple[dict[str, MacroDefinition], list[ActionRecord]]:
    """
    从所有目标 TeX 文件中收集可安全展开的零参数宏。

    策略：
    - 后定义覆盖先定义；
    - 若同名宏在多个文件中出现，以后者为准，并记录覆盖行为。
    """
    macros: dict[str, MacroDefinition] = {}
    actions: list[ActionRecord] = []

    for tex_file in tex_files:
        try:
            text = read_text_file(tex_file)
        except Exception as exc:
            actions.append(
                ActionRecord(
                    action_type="read_tex_failed_for_macro_scan",
                    severity=SEVERITY_WARN,
                    file=safe_relative(tex_file, project_root),
                    line=None,
                    message="宏扫描时无法读取 TeX 文件，已跳过该文件。",
                    details={"error": str(exc)},
                )
            )
            continue

        file_macros, file_actions = collect_zero_arg_macros_from_text(text, tex_file, project_root)
        actions.extend(file_actions)

        for name, definition in file_macros.items():
            if name in macros:
                actions.append(
                    ActionRecord(
                        action_type="override_safe_zero_arg_macro",
                        severity=SEVERITY_WARN,
                        file=definition.defined_in,
                        line=definition.line,
                        message="后续文件中的同名零参数宏覆盖了先前定义。",
                        details={
                            "macro_name": name,
                            "previous_defined_in": macros[name].defined_in,
                            "previous_line": macros[name].line,
                        },
                    )
                )
            macros[name] = definition

    return macros, actions


# -----------------------------------------------------------------------------
# 规范化动作实现
# -----------------------------------------------------------------------------

def normalize_newlines(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    将文本统一为 LF 换行。

    这是最保守、最稳定的一类规范化动作：
    - 不改变语义；
    - 有利于后续脚本行号与正则扫描稳定；
    - 有利于 Pandoc / Python 跨平台处理。
    """
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    actions: list[ActionRecord] = []
    if normalized != text:
        actions.append(
            ActionRecord(
                action_type="normalize_newlines",
                severity=SEVERITY_INFO,
                file=safe_relative(file_path, root),
                line=1,
                message="统一换行风格为 LF。",
            )
        )
    return normalized, actions


def add_graphics_extensions(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    为 \\includegraphics 中缺失扩展名的图片路径补全真实扩展名。

    设计理由：
    - 这是对 Pandoc 主转换帮助较大的确定性改写；
    - 只有当实际文件存在且扩展名唯一可解析时才改写；
    - 已带扩展名的路径不动。
    """
    actions: list[ActionRecord] = []

    pattern = re.compile(r"\\includegraphics(?P<opts>\[[^\]]*\])?\{(?P<target>[^}]+)\}")

    def repl(match: re.Match) -> str:
        nonlocal actions
        target = match.group("target").strip()
        opts = match.group("opts") or ""
        if Path(target).suffix:
            return match.group(0)

        resolved = resolve_path_with_extensions(file_path.parent, target, IMAGE_EXTENSIONS)
        if resolved is None:
            return match.group(0)

        relative_resolved = resolved.relative_to(file_path.parent.resolve())
        # 将 Windows 反斜杠统一为 LaTeX 更稳妥的正斜杠。
        normalized_target = relative_resolved.as_posix()

        line_no = line_number_from_offset(text, match.start())
        actions.append(
            ActionRecord(
                action_type="add_graphics_extension",
                severity=SEVERITY_INFO,
                file=safe_relative(file_path, root),
                line=line_no,
                message="为图片路径补全实际扩展名。",
                details={"old_target": target, "new_target": normalized_target},
            )
        )
        return f"\\includegraphics{opts}{{{normalized_target}}}"

    new_text = pattern.sub(repl, text)
    return new_text, actions


def add_bibliography_extensions(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    为 bibliography 资源补全 .bib 扩展名。

    支持：
    - \\addbibresource{refs}
    - \\bibliography{refs,more_refs}

    说明：
    - 只在能唯一解析到实际 .bib 文件时才补全；
    - 已带扩展名则不动；
    - 这一步是路径规范化，不改变 bibliography 语义。
    """
    actions: list[ActionRecord] = []

    addbib_pattern = re.compile(r"\\addbibresource(?P<opts>\[[^\]]*\])?\{(?P<target>[^}]+)\}")
    bibliography_pattern = re.compile(r"\\bibliography\{(?P<targets>[^}]+)\}")

    def addbib_repl(match: re.Match) -> str:
        nonlocal actions
        target = match.group("target").strip()
        opts = match.group("opts") or ""

        if Path(target).suffix:
            return match.group(0)

        resolved = resolve_path_with_extensions(file_path.parent, target, BIB_EXTENSIONS)
        if resolved is None:
            return match.group(0)

        relative_resolved = resolved.relative_to(file_path.parent.resolve())
        normalized_target = relative_resolved.as_posix()
        line_no = line_number_from_offset(text, match.start())
        actions.append(
            ActionRecord(
                action_type="add_bib_extension",
                severity=SEVERITY_INFO,
                file=safe_relative(file_path, root),
                line=line_no,
                message="为 addbibresource 路径补全 .bib 扩展名。",
                details={"old_target": target, "new_target": normalized_target},
            )
        )
        return f"\\addbibresource{opts}{{{normalized_target}}}"

    def bibliography_repl(match: re.Match) -> str:
        nonlocal actions
        targets = split_csv_payload(match.group("targets"))
        if not targets:
            return match.group(0)

        changed = False
        new_targets: list[str] = []
        line_no = line_number_from_offset(text, match.start())

        for target in targets:
            if Path(target).suffix:
                new_targets.append(target)
                continue
            resolved = resolve_path_with_extensions(file_path.parent, target, BIB_EXTENSIONS)
            if resolved is None:
                new_targets.append(target)
                continue
            relative_resolved = resolved.relative_to(file_path.parent.resolve())
            normalized_target = relative_resolved.as_posix()
            new_targets.append(normalized_target)
            changed = True
            actions.append(
                ActionRecord(
                    action_type="add_bib_extension",
                    severity=SEVERITY_INFO,
                    file=safe_relative(file_path, root),
                    line=line_no,
                    message="为 bibliography 路径补全 .bib 扩展名。",
                    details={"old_target": target, "new_target": normalized_target},
                )
            )

        if not changed:
            return match.group(0)
        return f"\\bibliography{{{','.join(new_targets)}}}"

    text = addbib_pattern.sub(addbib_repl, text)
    text = bibliography_pattern.sub(bibliography_repl, text)
    return text, actions


def normalize_extended_refs(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    对常见扩展交叉引用命令做保守降级。

    当前实现：
    - \\autoref{label} -> \\ref{label}
    - \\cref{a,b} / \\Cref{a,b} -> \\ref{a}, \\ref{b}
    - \\subref{label} / \\subref*{label} -> \\ref{label}

    为什么这样做：
    - Pandoc 对 LaTeX 内部扩展交叉引用并不总能稳定保持；
    - 规范化阶段优先保留“目标信息”和“可恢复编号关系”；
    - 将类型前缀（Figure / Table / Equation 等）弱化，属于可接受降级。
    """
    actions: list[ActionRecord] = []

    autoref_pattern = re.compile(r"\\autoref\{([^}]+)\}")
    cref_pattern = re.compile(r"\\(cref|Cref)\{([^}]+)\}")
    subref_pattern = re.compile(r"\\subref\*?(?:\[[^\]]*\])?\{([^}]+)\}")

    def autoref_repl(match: re.Match) -> str:
        label = match.group(1).strip()
        line_no = line_number_from_offset(text, match.start())
        actions.append(
            ActionRecord(
                action_type="normalize_autoref",
                severity=SEVERITY_WARN,
                file=safe_relative(file_path, root),
                line=line_no,
                message="将 \\autoref 保守降级为 \\ref。",
                details={"label": label},
            )
        )
        return f"\\ref{{{label}}}"

    def cref_repl(match: re.Match) -> str:
        labels = split_csv_payload(match.group(2))
        line_no = line_number_from_offset(text, match.start())

        if not labels:
            return match.group(0)

        actions.append(
            ActionRecord(
                action_type="normalize_cref",
                severity=SEVERITY_WARN,
                file=safe_relative(file_path, root),
                line=line_no,
                message="将 \\cref / \\Cref 保守降级为多个 \\ref。",
                details={"labels": labels, "source_command": match.group(1)},
            )
        )
        return ", ".join(f"\\ref{{{label}}}" for label in labels)

    def subref_repl(match: re.Match) -> str:
        label = match.group(1).strip()
        line_no = line_number_from_offset(text, match.start())
        actions.append(
            ActionRecord(
                action_type="normalize_subref",
                severity=SEVERITY_WARN,
                file=safe_relative(file_path, root),
                line=line_no,
                message="将 \\subref 保守降级为 \\ref（子图引用后续由 DOCX 后处理升级）。",
                details={"label": label},
            )
        )
        return f"\\ref{{{label}}}"

    text = autoref_pattern.sub(autoref_repl, text)
    text = cref_pattern.sub(cref_repl, text)
    text = subref_pattern.sub(subref_repl, text)
    return text, actions


def normalize_align_leading_labels(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    规范化 align 系环境中“行首 \\label”的位置。

    目标：
    - 将形如 ``\\label{eq:x}& a=b\\\\`` 的写法改为 ``& a=b\\label{eq:x}\\\\``；
    - 降低 Pandoc 对 align 数学块解析失败概率；
    - 保持标签文本不变，仅调整槽位位置。

    处理边界：
    - 仅处理 ``align / align* / alignat / alignat* / flalign / flalign*``；
    - 仅当 ``\\label`` 与同一行公式内容共行时移动；
    - 独立成行的 ``\\label`` 保持不动，避免错误挂接到下一行。
    """
    actions: list[ActionRecord] = []

    align_env_pattern = re.compile(
        r"\\begin\{(?P<env>align\*?|alignat\*?|flalign\*?)\}(?P<body>.*?)\\end\{(?P=env)\}",
        re.DOTALL,
    )
    leading_label_pattern = re.compile(
        r"^(?P<indent>[ \t]*)\\label\{(?P<label>[^}]+)\}(?P<tail>[^\n]*)$",
        re.MULTILINE,
    )

    def normalize_body(body: str, body_offset: int) -> str:
        def move_label(match: re.Match) -> str:
            tail = match.group("tail")
            if not tail.strip():
                return match.group(0)

            label = match.group("label").strip()
            indent = match.group("indent")

            tail_stripped = tail.rstrip()
            trailing_ws = tail[len(tail_stripped):]

            linebreak_match = re.search(r"(\\\\(?:\[[^\]]*\])?)\s*$", tail_stripped)
            if linebreak_match:
                equation_part = tail_stripped[:linebreak_match.start()].rstrip()
                linebreak_part = linebreak_match.group(1)
                if not equation_part:
                    return match.group(0)
                new_tail = f"{equation_part}\\label{{{label}}}{linebreak_part}"
            else:
                equation_part = tail_stripped
                if not equation_part:
                    return match.group(0)
                new_tail = f"{equation_part}\\label{{{label}}}"

            line_no = line_number_from_offset(text, body_offset + match.start())
            actions.append(
                ActionRecord(
                    action_type="normalize_align_leading_label",
                    severity=SEVERITY_INFO,
                    file=safe_relative(file_path, root),
                    line=line_no,
                    message="将 align 系环境中的行首 \\label 调整到公式行尾。",
                    details={"label": label},
                )
            )
            return f"{indent}{new_tail}{trailing_ws}"

        return leading_label_pattern.sub(move_label, body)

    def repl(match: re.Match) -> str:
        env_name = match.group("env")
        body = match.group("body")
        body_offset = match.start("body")
        new_body = normalize_body(body, body_offset)
        if new_body == body:
            return match.group(0)
        return f"\\begin{{{env_name}}}{new_body}\\end{{{env_name}}}"

    new_text = align_env_pattern.sub(repl, text)
    return new_text, actions


def normalize_legacy_math_commands(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    在数学上下文中规范化旧式命令，提升 Pandoc 数学解析稳定性。

    处理内容：
    - ``\\textbf{\\textit{x}}`` / ``\\textit{\\textbf{x}}`` -> ``\\boldsymbol{x}``
    - ``\\rm{...}`` / ``{\\rm ...}`` -> ``\\mathrm{...}``
    - ``\\bf{...}`` / ``{\\bf ...}`` -> ``\\mathbf{...}``
    - ``\\cal X`` / ``\\cal{X}`` -> ``\\mathcal{X}``
    - ``\\buildrel A \\over B`` -> ``\\overset{A}{B}``

    处理边界：
    - 仅作用于数学上下文（$...$、\\(...\\)、\\[...\\]、常见数学环境）；
    - 不改写普通正文中的文本样式命令。
    """
    actions: list[ActionRecord] = []
    working_text = text

    # 数学内容内部的局部替换规则。
    bold_italic_pattern = re.compile(r"\\textbf\s*\{\s*\\textit\s*\{(?P<body>[^{}]+)\}\s*\}")
    italic_bold_pattern = re.compile(r"\\textit\s*\{\s*\\textbf\s*\{(?P<body>[^{}]+)\}\s*\}")
    rm_boldtext_group_pattern = re.compile(
        r"\\rm\s*\{\s*\\textbf\s*\{(?P<body>[^{}]+)\}\s*\}"
    )
    rm_boldtext_legacy_pattern = re.compile(
        r"\{\s*\\rm\s*\{\s*\\textbf\s*\{(?P<body>[^{}]+)\}\s*\}\s*\}"
    )
    rm_group_pattern = re.compile(r"\\rm\s*\{(?P<body>[^{}]+)\}")
    rm_legacy_pattern = re.compile(r"\{\s*\\rm\s+(?P<body>[^{}]+?)\s*\}")
    bf_group_pattern = re.compile(r"\\bf\s*\{(?P<body>[^{}]+)\}")
    bf_legacy_pattern = re.compile(r"\{\s*\\bf\s+(?P<body>[^{}]+?)\s*\}")
    bf_token_pattern = re.compile(r"\\bf\s+(?P<body>[^\s\\{}(),;]+)")
    cal_group_pattern = re.compile(r"\\cal\s*\{(?P<body>[^{}]+)\}")
    cal_token_pattern = re.compile(r"\\cal\s+(?P<body>[A-Za-z])")
    buildrel_pattern = re.compile(
        r"\\buildrel\s+"
        r"(?P<top>(?:\\[A-Za-z]+|\\\S|\{[^{}]+\}|[^\\{}\s]+))"
        r"\s+\\over\s+"
        r"(?P<base>(?:\\[A-Za-z]+|\\\S|\{[^{}]+\}|[^\\{}\s]+))"
    )

    def transform_math_fragment(fragment: str, fragment_offset: int) -> str:
        nonlocal actions
        local_text = fragment

        def apply_pattern(
            content: str,
            pattern: re.Pattern,
            action_type: str,
            message: str,
            builder,
        ) -> str:
            def repl(match: re.Match) -> str:
                line_no = line_number_from_offset(working_text, fragment_offset + match.start())
                replacement = builder(match)
                actions.append(
                    ActionRecord(
                        action_type=action_type,
                        severity=SEVERITY_INFO,
                        file=safe_relative(file_path, root),
                        line=line_no,
                        message=message,
                        details={
                            "from": match.group(0),
                            "to": replacement,
                        },
                    )
                )
                return replacement

            return pattern.sub(repl, content)

        local_text = apply_pattern(
            local_text,
            bold_italic_pattern,
            "normalize_math_bold_italic",
            "将数学中的 \\textbf{\\textit{...}} 规范化为 \\boldsymbol{...}。",
            lambda match: f"\\boldsymbol{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            italic_bold_pattern,
            "normalize_math_bold_italic",
            "将数学中的 \\textit{\\textbf{...}} 规范化为 \\boldsymbol{...}。",
            lambda match: f"\\boldsymbol{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            rm_boldtext_group_pattern,
            "normalize_math_rm",
            "将数学中的 \\rm{\\textbf{...}} 规范化为 \\text{...}。",
            lambda match: f"\\text{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            rm_boldtext_legacy_pattern,
            "normalize_math_rm",
            "将数学中的 {\\rm{\\textbf{...}}} 规范化为 \\text{...}。",
            lambda match: f"\\text{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            rm_group_pattern,
            "normalize_math_rm",
            "将数学中的 \\rm{...} 规范化为 \\mathrm{...}。",
            lambda match: f"\\mathrm{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            rm_legacy_pattern,
            "normalize_math_rm",
            "将数学中的 {\\rm ...} 规范化为 \\mathrm{...}。",
            lambda match: f"\\mathrm{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            bf_group_pattern,
            "normalize_math_bf",
            "将数学中的 \\bf{...} 规范化为 \\mathbf{...}。",
            lambda match: f"\\mathbf{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            bf_legacy_pattern,
            "normalize_math_bf",
            "将数学中的 {\\bf ...} 规范化为 \\mathbf{...}。",
            lambda match: f"\\mathbf{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            bf_token_pattern,
            "normalize_math_bf",
            "将数学中的 \\bf X 规范化为 \\mathbf{X}。",
            lambda match: f"\\mathbf{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            cal_group_pattern,
            "normalize_math_cal",
            "将数学中的 \\cal{...} 规范化为 \\mathcal{...}。",
            lambda match: f"\\mathcal{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            cal_token_pattern,
            "normalize_math_cal",
            "将数学中的 \\cal X 规范化为 \\mathcal{X}。",
            lambda match: f"\\mathcal{{{match.group('body').strip()}}}",
        )
        local_text = apply_pattern(
            local_text,
            buildrel_pattern,
            "normalize_math_buildrel",
            "将数学中的 \\buildrel A \\over B 规范化为 \\overset{A}{B}。",
            lambda match: f"\\overset{{{match.group('top').strip()}}}{{{match.group('base').strip()}}}",
        )
        return local_text

    def normalize_by_pattern(
        source: str,
        pattern: re.Pattern,
        content_group: str,
        wrapper,
    ) -> str:
        cursor = 0
        parts: list[str] = []
        for match in pattern.finditer(source):
            parts.append(source[cursor:match.start()])
            original_content = match.group(content_group)
            content_offset = match.start(content_group)
            normalized_content = transform_math_fragment(original_content, content_offset)
            parts.append(wrapper(match, normalized_content))
            cursor = match.end()
        parts.append(source[cursor:])
        return "".join(parts)

    env_names = "|".join(re.escape(item) for item in MATH_BLOCK_ENV_NAMES)
    math_env_pattern = re.compile(
        rf"\\begin\{{(?P<env>{env_names})\}}(?P<body>.*?)\\end\{{(?P=env)\}}",
        re.DOTALL,
    )
    display_bracket_pattern = re.compile(r"\\\[(?P<body>.*?)\\\]", re.DOTALL)
    display_dollar_pattern = re.compile(r"(?<!\\)\$\$(?P<body>.*?)(?<!\\)\$\$", re.DOTALL)
    inline_paren_pattern = re.compile(r"\\\((?P<body>.*?)\\\)", re.DOTALL)
    inline_dollar_pattern = re.compile(r"(?<!\\)\$(?!\$)(?P<body>.*?)(?<!\\)\$(?!\$)", re.DOTALL)

    working_text = normalize_by_pattern(
        working_text,
        math_env_pattern,
        "body",
        lambda match, body: f"\\begin{{{match.group('env')}}}{body}\\end{{{match.group('env')}}}",
    )
    working_text = normalize_by_pattern(
        working_text,
        display_bracket_pattern,
        "body",
        lambda _match, body: f"\\[{body}\\]",
    )
    working_text = normalize_by_pattern(
        working_text,
        display_dollar_pattern,
        "body",
        lambda _match, body: f"$${body}$$",
    )
    working_text = normalize_by_pattern(
        working_text,
        inline_paren_pattern,
        "body",
        lambda _match, body: f"\\({body}\\)",
    )
    working_text = normalize_by_pattern(
        working_text,
        inline_dollar_pattern,
        "body",
        lambda _match, body: f"${body}$",
    )
    return working_text, actions


def downgrade_subfloat_wrappers(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    在 figure 环境内将 ``\\subfloat`` / ``\\subfigure`` 保守降级为稳定 minipage 结构。

    背景：
    - Pandoc 对 ``\\subfloat{...}`` 解析不稳定，常见现象是子图整体丢失，仅保留总图题；
    - 为提升 Word 端可读性与可维护性，需要在降级时保留子图题注信息。

    处理策略：
    - 仅在 ``figure / figure*`` 环境内识别 ``\\subfloat`` / ``\\subfigure``；
    - 解析最多两个可选参数 ``[...]`` 与必须参数 ``{...}``；
    - 将每个子图转换为：
      ``minipage(图片主体 + 子题注行)``，形成“图在上、题注在下”的稳定结构；
    - 子图 ``\\label`` 保持在主体中，不改写标签文本。
    """
    actions: list[ActionRecord] = []
    figure_pattern = re.compile(
        r"\\begin\{(?P<env>figure\*?)\}(?P<body>.*?)\\end\{(?P=env)\}",
        re.DOTALL,
    )
    subfloat_pattern = re.compile(r"\\(subfloat|subfigure)(?![A-Za-z@])")
    letters = "abcdefghijklmnopqrstuvwxyz"

    def marker_for_index(index: int) -> str:
        if index < 0:
            return "(?)"
        if index < len(letters):
            return f"({letters[index]})"
        # 26 以后使用 aa/ab/... 形式，确保索引稳定。
        value = index
        chars: list[str] = []
        while True:
            chars.append(letters[value % 26])
            value = value // 26 - 1
            if value < 0:
                break
        return f"({''.join(reversed(chars))})"

    def width_for_count(count: int) -> str:
        if count <= 1:
            return "0.96\\linewidth"
        if count == 2:
            return "0.48\\linewidth"
        if count == 3:
            return "0.31\\linewidth"
        width = max(0.18, min(0.48, 0.96 / max(count, 1)))
        return f"{width:.2f}\\linewidth"

    def choose_subcaption(optionals: list[str]) -> str:
        if len(optionals) >= 2 and optionals[1].strip():
            return optionals[1].strip()
        if optionals and optionals[0].strip():
            return optionals[0].strip()
        return ""

    def render_subfloat_minipage(
        *,
        body_content: str,
        marker: str,
        subcaption: str,
        width: str,
    ) -> str:
        lines: list[str] = [
            f"\\begin{{minipage}}[t]{{{width}}}",
            "\\centering",
            body_content.strip(),
        ]
        if subcaption:
            lines.extend(
                [
                    "\\par\\smallskip",
                    f"\\footnotesize {marker} {subcaption}",
                ]
            )
        lines.append("\\end{minipage}")
        return "\n".join(lines)

    def parse_subfloat_records(body: str) -> list[dict]:
        records: list[dict] = []
        cursor = 0
        while True:
            match = subfloat_pattern.search(body, cursor)
            if not match:
                break

            command_name = match.group(1)
            cmd_start = match.start()
            pos = skip_whitespace(body, match.end())

            optionals: list[str] = []
            parse_ok = True
            for _ in range(2):
                pos = skip_whitespace(body, pos)
                if pos < len(body) and body[pos] == "[":
                    opt_end = parse_balanced_group(body, pos, "[", "]")
                    if opt_end is None:
                        parse_ok = False
                        break
                    optionals.append(body[pos + 1 : opt_end - 1])
                    pos = opt_end
                else:
                    break

            if not parse_ok:
                cursor = match.end()
                continue

            pos = skip_whitespace(body, pos)
            if pos >= len(body) or body[pos] != "{":
                cursor = match.end()
                continue

            body_end = parse_balanced_group(body, pos, "{", "}")
            if body_end is None:
                cursor = match.end()
                continue

            records.append(
                {
                    "command": command_name,
                    "start": cmd_start,
                    "end": body_end,
                    "optionals": optionals,
                    "content": body[pos + 1 : body_end - 1],
                }
            )
            cursor = body_end
        return records

    def repl(match: re.Match) -> str:
        nonlocal actions
        env_name = match.group("env")
        body = match.group("body")
        records = parse_subfloat_records(body)
        if not records:
            return match.group(0)

        width = width_for_count(len(records))
        chunks: list[str] = []
        cursor = 0
        for index, record in enumerate(records):
            marker = marker_for_index(index)
            subcaption = choose_subcaption(record["optionals"])
            replacement = render_subfloat_minipage(
                body_content=record["content"],
                marker=marker,
                subcaption=subcaption,
                width=width,
            )

            chunks.append(body[cursor : record["start"]])
            chunks.append(replacement)
            cursor = record["end"]

            line_no = line_number_from_offset(text, match.start() + record["start"])
            actions.append(
                ActionRecord(
                    action_type="downgrade_subfloat_wrapper",
                    severity=SEVERITY_WARN,
                    file=safe_relative(file_path, root),
                    line=line_no,
                    message="将子图命令降级为 minipage（保留子图题注）以避免 Pandoc 丢图/丢子题注。",
                    details={
                        "environment": env_name,
                        "command": record["command"],
                        "subcaption_preserved": bool(subcaption),
                        "subcaption_marker": marker,
                        "subfloat_count_in_figure": len(records),
                    },
                )
            )

        chunks.append(body[cursor:])
        new_body = "".join(chunks)
        return f"\\begin{{{env_name}}}{new_body}\\end{{{env_name}}}"

    new_text = figure_pattern.sub(repl, text)
    return new_text, actions


def downgrade_table_star_environments(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    将 ``table*`` 环境保守降级为 ``table``。

    背景：
    - Pandoc 对 ``table*`` 的支持不稳定，可能出现表格保留但题注丢失；
    - 对 Word 交付而言，题注/编号/交叉引用稳定性优先于双栏跨栏语义。

    处理策略：
    - 仅替换环境名：``\\begin{table*}`` -> ``\\begin{table}``，
      ``\\end{table*}`` -> ``\\end{table}``；
    - 保留可选参数、内部内容、caption/label 次序不变。
    """
    actions: list[ActionRecord] = []
    begin_pattern = re.compile(r"\\begin\{table\*\}")
    end_pattern = re.compile(r"\\end\{table\*\}")

    begin_matches = list(begin_pattern.finditer(text))
    end_matches = list(end_pattern.finditer(text))
    if not begin_matches and not end_matches:
        return text, actions

    new_text = begin_pattern.sub(r"\\begin{table}", text)
    new_text = end_pattern.sub(r"\\end{table}", new_text)

    for match in begin_matches:
        line_no = line_number_from_offset(text, match.start())
        actions.append(
            ActionRecord(
                action_type="downgrade_table_star",
                severity=SEVERITY_WARN,
                file=safe_relative(file_path, root),
                line=line_no,
                message="将 table* 环境降级为 table，以提升 Pandoc 到 Word 的题注保真度。",
                details={"from_env": "table*", "to_env": "table"},
            )
        )

    if len(begin_matches) != len(end_matches):
        # 不中断流程，提示源文档中 table* 标记可能不成对。
        actions.append(
            ActionRecord(
                action_type="table_star_env_mismatch",
                severity=SEVERITY_WARN,
                file=safe_relative(file_path, root),
                line=None,
                message="检测到 table* begin/end 数量不一致，请人工核对该文件中的浮动体结构。",
                details={
                    "begin_count": len(begin_matches),
                    "end_count": len(end_matches),
                },
            )
        )

    return new_text, actions


def _find_command_occurrences_with_payload(
    text: str,
    command_name: str,
) -> list[tuple[int, int, str]]:
    """
    查找 ``\\command{...}`` / ``\\command[...]{...}``，并返回 payload。

    返回列表元素：
    - start: 命令开始位置
    - end: 命令结束位置（半开区间）
    - payload: 花括号内文本
    """
    pattern = re.compile(rf"\\{re.escape(command_name)}(?:\[[^\]]*\])?")
    results: list[tuple[int, int, str]] = []
    cursor = 0

    while True:
        match = pattern.search(text, cursor)
        if not match:
            break

        pos = skip_whitespace(text, match.end())
        if pos >= len(text) or text[pos] != "{":
            cursor = match.end()
            continue

        end = parse_balanced_group(text, pos, "{", "}")
        if end is None:
            cursor = match.end()
            continue

        payload = text[pos + 1 : end - 1].strip()
        results.append((match.start(), end, payload))
        cursor = end

    return results


def _remove_spans_from_text(text: str, spans: list[tuple[int, int]]) -> str:
    """
    从文本中移除若干区间（半开区间，按原文坐标）。
    """
    if not spans:
        return text

    merged: list[tuple[int, int]] = []
    for start, end in sorted(spans, key=lambda item: (item[0], item[1])):
        if start < 0 or end <= start:
            continue
        if not merged or start > merged[-1][1]:
            merged.append((start, end))
        else:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))

    parts: list[str] = []
    cursor = 0
    for start, end in merged:
        parts.append(text[cursor:start])
        cursor = end
    parts.append(text[cursor:])
    return "".join(parts)


def downgrade_algorithm_environments(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    将 ``algorithm / algorithm*``（含 algorithm2e 常见命令）降级为 Pandoc 稳定块结构。

    输出结构（示意）：
    - quote 块
    - 算法标题行（保留 caption）
    - 输入/输出行（保留 KwIn/KwOut 或 Require/Ensure）
    - enumerate 步骤列表（按 ``\\;`` 与原换行切分）
    """
    actions: list[ActionRecord] = []
    algorithm_pattern = re.compile(
        r"\\begin\{(?P<env>algorithm\*?)\}(?P<body>.*?)\\end\{(?P=env)\}",
        re.DOTALL,
    )

    def repl(match: re.Match) -> str:
        nonlocal actions
        env_name = match.group("env")
        body = match.group("body")

        caption_matches = _find_command_occurrences_with_payload(body, "caption")
        label_matches = _find_command_occurrences_with_payload(body, "label")
        kwin_matches = _find_command_occurrences_with_payload(body, "KwIn")
        kwout_matches = _find_command_occurrences_with_payload(body, "KwOut")
        require_matches = _find_command_occurrences_with_payload(body, "Require")
        ensure_matches = _find_command_occurrences_with_payload(body, "Ensure")

        caption_text = caption_matches[0][2] if caption_matches else ""
        label_text = label_matches[0][2] if label_matches else ""
        input_payloads = [item[2] for item in (kwin_matches + require_matches) if item[2]]
        output_payloads = [item[2] for item in (kwout_matches + ensure_matches) if item[2]]

        spans_to_remove: list[tuple[int, int]] = []
        for start, end, _payload in (
            caption_matches
            + label_matches
            + kwin_matches
            + kwout_matches
            + require_matches
            + ensure_matches
        ):
            spans_to_remove.append((start, end))

        working = _remove_spans_from_text(body, spans_to_remove)

        # 处理 algorithm2e 常见控制命令与行结束符。
        working = re.sub(r"\\BlankLine\b", "\n\n", working)
        working = re.sub(r"\\Indp\b", "", working)
        working = re.sub(r"\\Indm\b", "", working)
        working = re.sub(r"\\DontPrintSemicolon\b", "", working)
        working = re.sub(r"\\SetAlgoLined\b", "", working)
        working = re.sub(r"\\SetKw[A-Za-z]*\s*\{[^}]*\}(?:\s*\{[^}]*\})?", "", working)
        working = working.replace("\\;", "\n")

        step_candidates = re.split(r"\n+|\\\\", working)
        step_items: list[str] = []
        for raw in step_candidates:
            line = raw.strip()
            if not line:
                continue

            line = re.sub(r"\\emph\s*\{([^}]*)\}\s*", r"\1 ", line)
            line = re.sub(r"\s+", " ", line).strip()
            if not line:
                continue
            if line in {"\\", "\\par"}:
                continue
            step_items.append(line)

        title = caption_text if caption_text else "Algorithm"
        lines: list[str] = ["\\begin{quote}"]

        heading = f"\\noindent\\textbf{{Algorithm: {title}}}"
        if label_text:
            heading += f"\\label{{{label_text}}}"
        lines.append(heading)

        if input_payloads:
            lines.append(f"\\par\\textbf{{Input:}} {'；'.join(input_payloads)}")
        if output_payloads:
            lines.append(f"\\par\\textbf{{Output:}} {'；'.join(output_payloads)}")

        if step_items:
            lines.append("\\begin{enumerate}")
            for item in step_items:
                lines.append(f"\\item {item}")
            lines.append("\\end{enumerate}")

        lines.append("\\end{quote}")

        line_no = line_number_from_offset(text, match.start())
        actions.append(
            ActionRecord(
                action_type="downgrade_algorithm_environment",
                severity=SEVERITY_WARN,
                file=safe_relative(file_path, root),
                line=line_no,
                message="将 algorithm 环境降级为 quote+enumerate 块结构，以提升 Pandoc 转 Word 的稳定性。",
                details={
                    "environment": env_name,
                    "caption_preserved": bool(caption_text),
                    "label_preserved": bool(label_text),
                    "input_count": len(input_payloads),
                    "output_count": len(output_payloads),
                    "step_count": len(step_items),
                },
            )
        )

        return "\n".join(lines)

    new_text = algorithm_pattern.sub(repl, text)
    return new_text, actions


def find_first_command_span(text: str, command_name: str) -> Optional[tuple[int, int]]:
    """
    在一段文本中查找第一个类似 \\command[optional]{mandatory} 的命令区间。

    返回：
    - (start, end) 形式的半开区间
    - 若未找到或 mandatory 参数不完整，则返回 None

    用途：
    - 在 float 环境内部寻找首个 \\caption
    - 在 float 环境内部寻找首个 \\label
    """
    command_pattern = re.compile(rf"\\{re.escape(command_name)}(?:\[[^\]]*\])?")
    match = command_pattern.search(text)
    if not match:
        return None

    cursor = match.end()
    cursor = skip_whitespace(text, cursor)
    if cursor >= len(text) or text[cursor] != "{":
        return None

    end = parse_balanced_group(text, cursor, "{", "}")
    if end is None:
        return None

    return match.start(), end


def reorder_label_after_caption_in_floats(text: str, file_path: Path, root: Path) -> tuple[str, list[ActionRecord]]:
    """
    在 figure / table / figure* / table* 环境中，将“caption 前紧邻 label”调整为“caption 后 label”。

    为什么只做这种保守变换：
    - 这是最常见、最有助于后续交叉引用恢复的写法规范化；
    - 若 label 与 caption 之间夹杂大量其他内容，则不强行移动；
    - 避免因为过度“聪明”而改变 float 内部真正语义。
    """
    actions: list[ActionRecord] = []

    float_pattern = re.compile(
        r"\\begin\{(?P<env>figure\*?|table\*?)\}(?P<body>.*?)\\end\{(?P=env)\}",
        re.DOTALL,
    )

    def repl(match: re.Match) -> str:
        nonlocal actions
        env_name = match.group("env")
        body = match.group("body")

        label_span = find_first_command_span(body, "label")
        caption_span = find_first_command_span(body, "caption")

        if not label_span or not caption_span:
            return match.group(0)

        label_start, label_end = label_span
        caption_start, caption_end = caption_span

        # 仅在“label 明显位于 caption 前，且两者之间只有空白”时重排。
        # 这是保守且安全的情况。
        if label_start < caption_start:
            middle = body[label_end:caption_start]
            if middle.strip() == "":
                label_text = body[label_start:label_end]
                caption_text = body[caption_start:caption_end]

                new_body = (
                    body[:label_start]
                    + caption_text
                    + "\n"
                    + label_text
                    + body[caption_end:]
                )

                line_no = line_number_from_offset(text, match.start())
                actions.append(
                    ActionRecord(
                        action_type="reorder_label_after_caption",
                        severity=SEVERITY_INFO,
                        file=safe_relative(file_path, root),
                        line=line_no,
                        message="在 float 环境中将 \\label 调整到 \\caption 之后。",
                        details={"environment": env_name},
                    )
                )
                return f"\\begin{{{env_name}}}{new_body}\\end{{{env_name}}}"

        return match.group(0)

    new_text = float_pattern.sub(repl, text)
    return new_text, actions


def split_document_body(text: str) -> tuple[str, str]:
    """
    将文本拆成“前导部分”和“正文部分”。

    规则：
    - 若包含 \\begin{document}，则：
      * prefix 包含到该命令结束为止
      * body 为其后的内容
    - 若不包含，则 prefix 为空，body 为全文

    用途：
    - 避免在主文件导言区误展开宏使用；
    - 对普通章节子文件则直接视为正文整体。
    """
    marker = r"\begin{document}"
    idx = text.find(marker)
    if idx == -1:
        return "", text
    split_point = idx + len(marker)
    return text[:split_point], text[split_point:]


def expand_zero_arg_macros(
    text: str,
    file_path: Path,
    root: Path,
    macros: dict[str, MacroDefinition],
) -> tuple[str, list[ActionRecord]]:
    """
    在正文区域内展开“可安全展开”的零参数宏。

    策略：
    - 仅在正文区域做替换，尽量避免改动导言区和宏定义区；
    - 对每个宏做精确命令边界匹配；
    - 使用最多两轮替换，以支持简单的链式零参数展开；
    - 绝不删除原始定义，只替换实际使用位置。

    注意：
    - 这是“保守辅助展开”，不是完整宏求值器；
    - 若宏用法复杂，本函数不会试图处理。
    """
    if not macros:
        return text, []

    actions: list[ActionRecord] = []
    prefix, body = split_document_body(text)

    def is_macro_definition_target(content: str, start: int) -> bool:
        """
        判断当前位置是否落在“宏定义目标位”。

        例如以下写法中的 `\\foo` 不应被展开：
        - \\renewcommand\\foo{...}
        - \\renewcommand{\\foo}{...}
        - \\newcommand\\foo{...}
        - \\def\\foo{...}
        - \\let\\foo\\bar
        """
        window_start = max(0, start - 160)
        context = content[window_start:start]
        return bool(
            re.search(
                r"(?:\\(?:newcommand|renewcommand|providecommand)\*?\s*\{?\s*|\\(?:def|gdef|xdef|edef|let)\s*)$",
                context,
            )
        )

    # 为减少短命令与长命令前缀碰撞，按命令名长度逆序处理。
    ordered_macros = sorted(macros.values(), key=lambda item: len(item.name), reverse=True)

    # 最多两轮，避免出现意外递归膨胀。
    for _ in range(2):
        any_replaced = False
        for macro in ordered_macros:
            pattern = re.compile(rf"(?<!\\)\\{re.escape(macro.name)}(?![A-Za-z@])")

            def repl(match: re.Match) -> str:
                nonlocal actions, any_replaced, body
                if is_macro_definition_target(body, match.start()):
                    return match.group(0)

                any_replaced = True
                line_no = line_number_from_offset(prefix + body, len(prefix) + match.start())
                actions.append(
                    ActionRecord(
                        action_type="expand_zero_arg_macro",
                        severity=SEVERITY_INFO,
                        file=safe_relative(file_path, root),
                        line=line_no,
                        message="展开可安全展开的零参数宏。",
                        details={
                            "macro_name": macro.name,
                            "defined_in": macro.defined_in,
                            "defined_line": macro.line,
                        },
                    )
                )
                return macro.replacement

            body = pattern.sub(repl, body)

        if not any_replaced:
            break

    return prefix + body, actions


# -----------------------------------------------------------------------------
# 文件处理与报告渲染
# -----------------------------------------------------------------------------

def process_tex_file(
    file_path: Path,
    work_root: Path,
    macros: dict[str, MacroDefinition],
) -> tuple[bool, list[ActionRecord]]:
    """
    对单个 TeX 文件执行规范化动作。

    返回：
    - modified: 文件内容是否发生变化
    - actions: 动作列表

    执行顺序固定如下：
    1. 统一换行
    2. 补全图片扩展名
    3. 补全 bib 扩展名
    4. 规范化扩展交叉引用
    5. 规范化 align 系环境中的行首 label 位置
    6. 规范化数学上下文中的旧式命令
    7. 降级 subfloat / subfigure 包装
    8. 调整 float 中 caption / label 顺序
    9. 将 table* 环境降级为 table
    10. 将 algorithm 环境降级为 Pandoc 稳定块结构
    11. 展开安全零参数宏

    顺序理由：
    - 先做低风险路径规范化；
    - 再做交叉引用规范化；
    - 最后做宏展开，避免展开后的路径/引用再次复杂化。
    """
    actions: list[ActionRecord] = []

    original_text = read_text_file(file_path)
    text = original_text

    text, file_actions = normalize_newlines(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = add_graphics_extensions(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = add_bibliography_extensions(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = normalize_extended_refs(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = normalize_align_leading_labels(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = normalize_legacy_math_commands(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = downgrade_subfloat_wrappers(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = reorder_label_after_caption_in_floats(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = downgrade_table_star_environments(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = downgrade_algorithm_environments(text, file_path, work_root)
    actions.extend(file_actions)

    text, file_actions = expand_zero_arg_macros(text, file_path, work_root, macros)
    actions.extend(file_actions)

    modified = text != original_text
    if modified:
        write_text_file(file_path, text)

    return modified, actions


def render_markdown_report(report: NormalizationReport) -> str:
    """
    将规范化报告渲染为 Markdown。

    报告目标：
    - 让用户知道“工作副本在哪里”；
    - 让后续脚本知道“改了什么类型的东西”；
    - 让人工审查者知道“哪些改写带有降级含义”。
    """
    lines: list[str] = []
    lines.append("# Normalization Report")
    lines.append("")
    lines.append(f"- Status: **{report.status}**")
    lines.append(f"- Can continue: **{report.can_continue}**")
    lines.append(f"- Used precheck report: **{report.used_precheck_report}**")
    lines.append(f"- Project root: `{report.project_root}`")
    lines.append(f"- Work root: `{report.work_root}`")
    lines.append(f"- Source main TeX: `{report.source_main_tex or 'N/A'}`")
    lines.append(f"- Normalized main TeX: `{report.normalized_main_tex or 'N/A'}`")
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

    lines.append("## Processed TeX Files")
    lines.append("")
    if report.tex_files_processed:
        for item in report.tex_files_processed:
            lines.append(f"- `{item}`")
    else:
        lines.append("- (none)")
    lines.append("")

    lines.append("## File Summaries")
    lines.append("")
    if report.file_summaries:
        for item in report.file_summaries:
            lines.append(f"- `{item['file']}`")
            lines.append(f"  - modified: **{item['modified']}**")
            lines.append(f"  - action_count: **{item['action_count']}**")
            if item["actions_by_type"]:
                for key, value in item["actions_by_type"].items():
                    lines.append(f"  - {key}: **{value}**")
            else:
                lines.append("  - actions_by_type: (none)")
    else:
        lines.append("- (none)")
    lines.append("")

    lines.append("## Actions")
    lines.append("")
    if report.actions:
        grouped: dict[str, list[dict]] = defaultdict(list)
        for item in report.actions:
            grouped[item["severity"]].append(item)

        for severity in SEVERITY_ORDER:
            if severity not in grouped:
                continue
            lines.append(f"### {severity}")
            lines.append("")
            for item in grouped[severity]:
                location = f"`{item['file']}`"
                if item.get("line"):
                    location += f":{item['line']}"
                lines.append(f"- **[{item['action_type']}]** {item['message']} ({location})")
            lines.append("")
    else:
        lines.append("- No actions recorded.")
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


def build_action_metrics(action_log: list[ActionRecord]) -> dict:
    """
    统一统计 action_log 的严重级别计数，避免在失败分支重复拼装。
    """
    return {
        "action_count": len(action_log),
        "error_count": sum(1 for a in action_log if a.severity == SEVERITY_ERROR),
        "warn_count": sum(1 for a in action_log if a.severity == SEVERITY_WARN),
        "info_count": sum(1 for a in action_log if a.severity == SEVERITY_INFO),
    }


def persist_normalization_report(
    *,
    work_root: Path,
    report: NormalizationReport,
    json_out: Path,
    md_out: Path,
    normalized_source_root: Optional[Path] = None,
    normalized_main_tex: Optional[Path] = None,
) -> None:
    """
    统一写出 normalize 阶段报告，并更新 stage 级 manifest 状态。
    """
    artifacts = {
        "normalization_report_json": json_out,
        "normalization_report_md": md_out,
    }
    if normalized_source_root is not None:
        artifacts["normalized_source_root"] = normalized_source_root
    if normalized_main_tex is not None:
        artifacts["normalized_main_tex"] = normalized_main_tex

    persist_stage_report(
        work_root=work_root,
        stage=STAGE_NORMALIZE,
        report_obj=report,
        markdown_text=render_markdown_report(report),
        report_json_path=json_out,
        report_md_path=md_out,
        status=report.status,
        can_continue=report.can_continue,
        artifacts=artifacts,
        summary=report.summary,
        metrics=report.metrics,
        top_level_artifacts={
            "reports": {
                "normalization_report_json": json_out,
                "normalization_report_md": md_out,
            }
        },
    )


# -----------------------------------------------------------------------------
# 主入口
# -----------------------------------------------------------------------------

def main() -> int:
    """
    主入口函数。

    固定执行顺序：
    1. 解析参数；
    2. 解析工程根目录、skill 根目录、precheck 报告；
    3. 解析主文件；
    4. 校验工作目录安全性；
    5. 复制工程到工作目录；
    6. 收集目标 TeX 文件；
    7. 收集可安全展开的零参数宏；
    8. 逐文件规范化；
    9. 生成 JSON / Markdown 报告；
    10. 输出控制台摘要；
    11. 返回退出码。
    """
    parser = build_argument_parser()
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    if not project_root.exists() or not project_root.is_dir():
        print(f"[ERROR] 无效的工程目录: {project_root}", file=sys.stderr)
        return 1

    skill_root = locate_skill_root()
    work_root = resolve_explicit_or_default(
        args.work_root,
        default_work_root_for(project_root),
    )
    normalize_stage_dir = stage_dir(work_root, STAGE_NORMALIZE)
    json_out = normalize_stage_dir / "normalization-report.json"
    md_out = normalize_stage_dir / "normalization-report.md"
    precheck_stage_dir = stage_dir(work_root, STAGE_PRECHECK)
    normalized_source_root = (normalize_stage_dir / "source_snapshot").resolve()

    precheck_json_path = resolve_explicit_or_stage_input(
        args.precheck_json,
        work_root,
        STAGE_PRECHECK,
        "precheck-report.json",
        legacy_filename="precheck-report.json",
    )
    if not precheck_json_path.exists():
        precheck_json_path = (project_root / "precheck-report.json").resolve()
    precheck_report = load_json_if_exists(precheck_json_path)
    used_precheck_report = precheck_report is not None

    action_log: list[ActionRecord] = []
    action_log.extend(check_rule_files(skill_root))

    # 若 precheck 已明确 FAIL，则本阶段不应继续假装正常。
    if precheck_report and precheck_report.get("status") == STATUS_FAIL:
        action_log.append(
            ActionRecord(
                action_type="abort_due_to_precheck_fail",
                severity=SEVERITY_ERROR,
                file=safe_relative(precheck_json_path, project_root),
                line=None,
                message="precheck-report.json 显示预检查失败，规范化阶段已中止。",
                details={"precheck_status": precheck_report.get("status")},
            )
        )
        report = NormalizationReport(
            status=STATUS_FAIL,
            can_continue=False,
            used_precheck_report=used_precheck_report,
            project_root=str(project_root),
            work_root=str(work_root),
            source_main_tex=None,
            normalized_main_tex=None,
            tex_files_processed=[],
            tex_files_processed_count=0,
            tex_files_modified_count=0,
            actions=[asdict(a) for a in action_log],
            file_summaries=[],
            metrics=build_action_metrics(action_log),
            summary={
                "status": STATUS_FAIL,
                "reason": "precheck_failed",
            },
            recommendations=[
                "先修复 precheck-report.json 中的 ERROR 级问题，再重新运行 normalize_tex.py。"
            ],
        )

        persist_normalization_report(
            work_root=work_root,
            report=report,
            json_out=json_out,
            md_out=md_out,
        )

        print(f"[{report.status}] normalization aborted because precheck failed.")
        print(f"JSON report: {json_out}")
        print(f"Markdown report: {md_out}")
        return 1

    # 解析主文件
    main_tex, main_actions = resolve_main_tex(project_root, args.main_tex, precheck_report)
    action_log.extend(main_actions)
    if main_tex is None:
        report = NormalizationReport(
            status=STATUS_FAIL,
            can_continue=False,
            used_precheck_report=used_precheck_report,
            project_root=str(project_root),
            work_root=str(work_root),
            source_main_tex=None,
            normalized_main_tex=None,
            tex_files_processed=[],
            tex_files_processed_count=0,
            tex_files_modified_count=0,
            actions=[asdict(a) for a in action_log],
            file_summaries=[],
            metrics=build_action_metrics(action_log),
            summary={"status": STATUS_FAIL, "reason": "main_tex_not_resolved"},
            recommendations=[
                "显式传入 --main-tex，或先运行 precheck.py 生成 precheck-report.json。"
            ],
        )

        persist_normalization_report(
            work_root=work_root,
            report=report,
            json_out=json_out,
            md_out=md_out,
        )

        print(f"[{report.status}] normalization failed: main TeX file could not be resolved.")
        print(f"JSON report: {json_out}")
        print(f"Markdown report: {md_out}")
        return 1

    # 解析工作目录
    safety_error = ensure_work_root_is_safe(project_root, work_root)
    if safety_error:
        action_log.append(
            ActionRecord(
                action_type="unsafe_work_root",
                severity=SEVERITY_ERROR,
                file=safe_relative(work_root, project_root),
                line=None,
                message=safety_error,
            )
        )
        report = NormalizationReport(
            status=STATUS_FAIL,
            can_continue=False,
            used_precheck_report=used_precheck_report,
            project_root=str(project_root),
            work_root=str(work_root),
            source_main_tex=safe_relative(main_tex, project_root),
            normalized_main_tex=None,
            tex_files_processed=[],
            tex_files_processed_count=0,
            tex_files_modified_count=0,
            actions=[asdict(a) for a in action_log],
            file_summaries=[],
            metrics=build_action_metrics(action_log),
            summary={"status": STATUS_FAIL, "reason": "unsafe_work_root"},
            recommendations=[
                "将工作目录设置到原始工程目录之外，例如工程同级目录。"
            ],
        )

        persist_normalization_report(
            work_root=work_root,
            report=report,
            json_out=json_out,
            md_out=md_out,
        )

        print(f"[{report.status}] normalization failed: unsafe work root.")
        print(f"JSON report: {json_out}")
        print(f"Markdown report: {md_out}")
        return 1

    # 处理已存在工作目录
    work_root.mkdir(parents=True, exist_ok=True)
    if normalize_stage_dir.exists():
        if args.force:
            shutil.rmtree(normalize_stage_dir)
            action_log.append(
                ActionRecord(
                    action_type="remove_existing_normalize_stage",
                    severity=SEVERITY_WARN,
                    file=safe_relative(normalize_stage_dir, work_root),
                    line=None,
                    message="检测到已存在的 normalize 阶段目录，已按 --force 删除后重建。",
                )
            )
        else:
            action_log.append(
                ActionRecord(
                    action_type="normalize_stage_already_exists",
                    severity=SEVERITY_ERROR,
                    file=safe_relative(normalize_stage_dir, work_root),
                    line=None,
                    message="normalize 阶段目录已存在；若确认可覆盖，请使用 --force。",
                )
            )
            report = NormalizationReport(
                status=STATUS_FAIL,
                can_continue=False,
                used_precheck_report=used_precheck_report,
                project_root=str(project_root),
                work_root=str(work_root),
                source_main_tex=safe_relative(main_tex, project_root),
                normalized_main_tex=None,
                tex_files_processed=[],
                tex_files_processed_count=0,
                tex_files_modified_count=0,
                actions=[asdict(a) for a in action_log],
                file_summaries=[],
                metrics=build_action_metrics(action_log),
                summary={"status": STATUS_FAIL, "reason": "normalize_stage_exists"},
                recommendations=[
                    "使用 --force 允许删除旧的 stage_normalize，或改用新的 --work-root。"
                ],
            )

            persist_normalization_report(
                work_root=work_root,
                report=report,
                json_out=json_out,
                md_out=md_out,
            )

            print(f"[{report.status}] normalization failed: normalize stage already exists.")
            print(f"JSON report: {json_out}")
            print(f"Markdown report: {md_out}")
            return 1

    normalize_stage_dir.mkdir(parents=True, exist_ok=True)

    # 复制工程到 stage_normalize/source_snapshot
    copy_project_tree(project_root, normalized_source_root)
    action_log.append(
        ActionRecord(
            action_type="copy_project_tree",
            severity=SEVERITY_INFO,
            file=safe_relative(normalized_source_root, work_root),
            line=None,
            message="已将原始工程复制到规范化工作目录。",
        )
    )

    # 若 precheck 报告来自 project_root 或 legacy 路径，则在 stage_precheck 做一份镜像，便于后续阶段统一读取。
    if precheck_report is not None:
        try:
            precheck_stage_dir.mkdir(parents=True, exist_ok=True)
            staged_precheck_json = (precheck_stage_dir / "precheck-report.json").resolve()
            staged_precheck_md = (precheck_stage_dir / "precheck-report.md").resolve()
            write_json(staged_precheck_json, precheck_report)

            source_precheck_md = precheck_json_path.with_suffix(".md")
            if source_precheck_md.exists():
                write_text_file(staged_precheck_md, read_text_file(source_precheck_md))
            else:
                write_text_file(staged_precheck_md, "# Precheck Report\n\n(原始 Markdown 报告缺失，仅保留 JSON 镜像。)\n")
            precheck_json_path = staged_precheck_json
        except Exception as exc:
            action_log.append(
                ActionRecord(
                    action_type="mirror_precheck_to_stage_failed",
                    severity=SEVERITY_WARN,
                    file=safe_relative(precheck_stage_dir, work_root),
                    line=None,
                    message="无法将 precheck 报告镜像到 stage_precheck。",
                    details={"error": str(exc)},
                )
            )

    # 解析规范化后主文件路径
    normalized_main_tex = (normalized_source_root / safe_relative(main_tex, project_root)).resolve()
    if not normalized_main_tex.exists():
        action_log.append(
            ActionRecord(
                action_type="normalized_main_tex_missing_after_copy",
                severity=SEVERITY_ERROR,
                file=safe_relative(normalized_main_tex, normalized_source_root),
                line=None,
                message="复制完成后未在工作目录中找到主 TeX 文件。",
            )
        )

        report = NormalizationReport(
            status=STATUS_FAIL,
            can_continue=False,
            used_precheck_report=used_precheck_report,
            project_root=str(project_root),
            work_root=str(work_root),
            source_main_tex=safe_relative(main_tex, project_root),
            normalized_main_tex=safe_relative(normalized_main_tex, normalized_source_root),
            tex_files_processed=[],
            tex_files_processed_count=0,
            tex_files_modified_count=0,
            actions=[asdict(a) for a in action_log],
            file_summaries=[],
            metrics=build_action_metrics(action_log),
            summary={"status": STATUS_FAIL, "reason": "normalized_main_tex_missing"},
            recommendations=[
                "检查复制后的工作目录结构，确认主文件相对路径是否正确。"
            ],
        )

        persist_normalization_report(
            work_root=work_root,
            report=report,
            json_out=json_out,
            md_out=md_out,
            normalized_source_root=normalized_source_root,
            normalized_main_tex=normalized_main_tex,
        )

        print(f"[{report.status}] normalization failed: normalized main TeX missing.")
        print(f"JSON report: {json_out}")
        print(f"Markdown report: {md_out}")
        return 1

    # 收集目标 TeX 文件
    target_tex_files = collect_target_tex_files(project_root, normalized_source_root, precheck_report)
    if not target_tex_files:
        action_log.append(
            ActionRecord(
                action_type="no_target_tex_files",
                severity=SEVERITY_ERROR,
                file=safe_relative(normalized_source_root, work_root),
                line=None,
                message="工作目录中没有可供规范化的 TeX 文件。",
            )
        )

        report = NormalizationReport(
            status=STATUS_FAIL,
            can_continue=False,
            used_precheck_report=used_precheck_report,
            project_root=str(project_root),
            work_root=str(work_root),
            source_main_tex=safe_relative(main_tex, project_root),
            normalized_main_tex=safe_relative(normalized_main_tex, normalized_source_root),
            tex_files_processed=[],
            tex_files_processed_count=0,
            tex_files_modified_count=0,
            actions=[asdict(a) for a in action_log],
            file_summaries=[],
            metrics=build_action_metrics(action_log),
            summary={"status": STATUS_FAIL, "reason": "no_target_tex_files"},
            recommendations=[
                "检查 precheck-report.json 中的 scanned_tex_files，或确认工程中确实存在 .tex 文件。"
            ],
        )

        persist_normalization_report(
            work_root=work_root,
            report=report,
            json_out=json_out,
            md_out=md_out,
            normalized_source_root=normalized_source_root,
            normalized_main_tex=normalized_main_tex,
        )

        print(f"[{report.status}] normalization failed: no target TeX files found.")
        print(f"JSON report: {json_out}")
        print(f"Markdown report: {md_out}")
        return 1

    # 收集零参数安全宏
    macros, macro_actions = collect_zero_arg_macros(target_tex_files, normalized_source_root)
    action_log.extend(macro_actions)

    # 逐文件规范化
    file_summaries: list[FileSummary] = []
    modified_count = 0

    for tex_file in target_tex_files:
        try:
            modified, file_actions = process_tex_file(tex_file, normalized_source_root, macros)
            action_log.extend(file_actions)

            actions_by_type = Counter(action.action_type for action in file_actions)
            file_summaries.append(
                FileSummary(
                    file=safe_relative(tex_file, normalized_source_root),
                    modified=modified,
                    action_count=len(file_actions),
                    actions_by_type=dict(sorted(actions_by_type.items())),
                )
            )
            if modified:
                modified_count += 1

        except Exception as exc:
            action_log.append(
                ActionRecord(
                    action_type="process_tex_file_failed",
                    severity=SEVERITY_ERROR,
                    file=safe_relative(tex_file, normalized_source_root),
                    line=None,
                    message="规范化该 TeX 文件时发生异常。",
                    details={"error": str(exc)},
                )
            )
            file_summaries.append(
                FileSummary(
                    file=safe_relative(tex_file, normalized_source_root),
                    modified=False,
                    action_count=0,
                    actions_by_type={},
                )
            )

    # 最终状态判定
    error_count = sum(1 for action in action_log if action.severity == SEVERITY_ERROR)
    warn_count = sum(1 for action in action_log if action.severity == SEVERITY_WARN)
    info_count = sum(1 for action in action_log if action.severity == SEVERITY_INFO)

    if error_count > 0:
        status = STATUS_FAIL
        can_continue = False
    elif warn_count > 0:
        status = STATUS_PASS_WITH_WARNINGS
        can_continue = True
    else:
        status = STATUS_PASS
        can_continue = True

    action_counts = Counter(action.action_type for action in action_log)

    recommendations: list[str] = []
    if status == STATUS_FAIL:
        recommendations.append("先修复规范化阶段的 ERROR 级问题，再进入 Pandoc 主转换。")
    else:
        recommendations.append("可将规范化后的工作副本作为 Pandoc 主转换输入。")

    if any(action.action_type in {"normalize_autoref", "normalize_cref"} for action in action_log):
        recommendations.append("交叉引用已做保守降级；后续应在 Word 中重点核对图表公式引用。")

    if any(action.action_type == "reorder_label_after_caption" for action in action_log):
        recommendations.append("float 中的 caption/label 顺序已部分规范化，有利于后续引用恢复。")

    if any(action.action_type == "expand_zero_arg_macro" for action in action_log):
        recommendations.append("部分零参数宏已展开；若正文存在复杂自定义命令，仍需在后续阶段重点复核。")

    if any(action.action_type == "downgrade_algorithm_environment" for action in action_log):
        recommendations.append("algorithm 环境已降级为块级结构；请在 Word 中复核伪代码标题、步骤换行与引用文本。")

    if any(action.action_type.startswith("skip_") for action in action_log):
        recommendations.append("存在被跳过的复杂宏定义；这些对象不应被视为已自动处理。")

    report = NormalizationReport(
        status=status,
        can_continue=can_continue,
        used_precheck_report=used_precheck_report,
        project_root=str(project_root),
        work_root=str(work_root),
        source_main_tex=safe_relative(main_tex, project_root),
        normalized_main_tex=safe_relative(normalized_main_tex, normalized_source_root),
        tex_files_processed=[safe_relative(path, normalized_source_root) for path in target_tex_files],
        tex_files_processed_count=len(target_tex_files),
        tex_files_modified_count=modified_count,
        actions=[asdict(action) for action in action_log],
        file_summaries=[asdict(summary) for summary in file_summaries],
        metrics={
            "action_count": len(action_log),
            "distinct_action_type_count": len(action_counts),
            "error_count": error_count,
            "warn_count": warn_count,
            "info_count": info_count,
            "safe_zero_arg_macro_count": len(macros),
            "tex_files_processed_count": len(target_tex_files),
            "tex_files_modified_count": modified_count,
        },
        summary={
            "status": status,
            "can_continue": can_continue,
            "used_precheck_report": used_precheck_report,
            "source_main_tex": safe_relative(main_tex, project_root),
            "normalized_main_tex": safe_relative(normalized_main_tex, normalized_source_root),
            "normalized_source_root": str(normalized_source_root),
        },
        recommendations=recommendations,
    )

    persist_normalization_report(
        work_root=work_root,
        report=report,
        json_out=json_out,
        md_out=md_out,
        normalized_source_root=normalized_source_root,
        normalized_main_tex=normalized_main_tex,
    )

    print(f"[{report.status}] normalization completed.")
    print(f"Source main TeX: {report.source_main_tex}")
    print(f"Normalized main TeX: {report.normalized_main_tex}")
    print(f"Processed TeX files: {report.tex_files_processed_count}")
    print(f"Modified TeX files: {report.tex_files_modified_count}")
    print(f"Errors: {report.metrics['error_count']}")
    print(f"Warnings: {report.metrics['warn_count']}")
    print(f"Infos: {report.metrics['info_count']}")
    print(f"Work root: {work_root}")
    print(f"Normalized source root: {normalized_source_root}")
    print(f"JSON report: {json_out}")
    print(f"Markdown report: {md_out}")

    return 1 if report.status == STATUS_FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
