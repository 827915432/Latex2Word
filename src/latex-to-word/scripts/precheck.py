#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
precheck.py

功能概述
--------
本脚本用于在 LaTeX -> Word 主转换之前，对 LaTeX 工程执行“只读式”预检查。

它解决的问题包括：
1. 主文件是否存在或可自动识别；
2. 多文件工程的 \\input / \\include 链是否闭合；
3. 图片资源是否缺失；
4. 参考文献资源是否缺失；
5. \\label / \\ref / \\eqref / \\cite 等引用关系是否存在明显问题；
6. 工程中是否存在高风险对象，例如：
   - 复杂表格
   - TikZ / PGFPlots
   - 子图
   - 算法环境
   - 代码环境
   - 条件编译
   - 自定义命令 / 自定义环境
7. 是否具备进入“规范化 -> Pandoc 主转换”阶段的基本条件。

脚本设计原则
------------
- 只读：绝不修改用户源文件；
- 可追踪：所有问题都转化为结构化 finding；
- 可自动化：输出 JSON + Markdown，便于后续脚本消费；
- 少假设：尽量基于文件内容和工程结构给出可解释判断；
- 不夸大成功：有问题就明确写出来，不用模糊表述。

依赖
----
- Python 3.9+
- 仅使用标准库，不依赖第三方包

典型用法
--------
python scripts/precheck.py --project-root D:/work/my-paper --main-tex main.tex

或让脚本自动识别主文件：

python scripts/precheck.py --project-root D:/work/my-paper

输出
----
默认会在 work-root 的 `stage_precheck/` 下生成：
- precheck-report.json
- precheck-report.md

若未显式传入 `--work-root`，默认 work-root 为：
`<project-root-parent>/<project-name>__latex_to_word_work/`

退出码
------
- 0: PASS 或 PASS_WITH_WARNINGS
- 1: FAIL
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable, Optional

from pipeline_layout import (
    STAGE_PRECHECK,
    default_work_root_for_project,
    stage_dir,
    update_manifest_artifacts,
    update_stage_manifest,
)


# -----------------------------------------------------------------------------
# 常量定义
# -----------------------------------------------------------------------------

# 支持自动补全的图片扩展名。
# 这里不做过度设计，只覆盖 LaTeX 工程里最常见的静态资源格式。
IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".pdf", ".svg", ".eps", ".bmp", ".tif", ".tiff"]

# 支持自动补全的 TeX 文件扩展名。
TEX_EXTENSIONS = [".tex"]

# 支持自动补全的 bibliography 扩展名。
BIB_EXTENSIONS = [".bib"]

# 预检查阶段使用的 finding 级别。
SEVERITY_INFO = "INFO"
SEVERITY_WARN = "WARN"
SEVERITY_ERROR = "ERROR"

# 最终总体状态。
STATUS_PASS = "PASS"
STATUS_PASS_WITH_WARNINGS = "PASS_WITH_WARNINGS"
STATUS_FAIL = "FAIL"

# 在 supported_envs.md 中可视为常见且预期出现的环境。
# 该集合的目的不是覆盖全部 LaTeX 环境，而是减少明显噪声。
KNOWN_ENVIRONMENTS = {
    # 文档结构
    "document",
    # 图片 / 表格
    "figure",
    "figure*",
    "table",
    "table*",
    "tabular",
    "tabular*",
    "tabularx",
    "array",
    "longtable",
    # 数学
    "equation",
    "equation*",
    "align",
    "align*",
    "gather",
    "gather*",
    "multline",
    "multline*",
    "split",
    "cases",
    "matrix",
    "pmatrix",
    "bmatrix",
    "vmatrix",
    "Vmatrix",
    # 定理类
    "theorem",
    "lemma",
    "proposition",
    "corollary",
    "definition",
    "remark",
    "example",
    "proof",
    # 列表
    "itemize",
    "enumerate",
    "description",
    # 引文 / 代码
    "quote",
    "quotation",
    "verbatim",
    "lstlisting",
    "minted",
    # 算法
    "algorithm",
    "algorithmic",
    "algpseudocode",
    # 子图与绘图
    "subfigure",
    "subtable",
    "tikzpicture",
    "axis",
    # bibliography
    "thebibliography",
}

# 预检查中重点关注的高风险环境。
HIGH_RISK_ENVIRONMENTS = {
    "longtable",
    "tabularx",
    "algorithm",
    "algorithmic",
    "algpseudocode",
    "algorithm2e",
    "verbatim",
    "lstlisting",
    "minted",
    "subfigure",
    "subtable",
    "tikzpicture",
    "axis",
}

# 表格相关高风险命令。
HIGH_RISK_TABLE_COMMANDS = {
    "multirow",
    "multicolumn",
    "cline",
    "cmidrule",
    "toprule",
    "midrule",
    "bottomrule",
}

# 常见 cite 类命令。这里只抽取“需要检查引用键”的命令集合。
CITE_COMMANDS = (
    "cite",
    "citep",
    "citet",
    "parencite",
    "textcite",
    "autocite",
    "footcite",
    "supercite",
)

# 常见 ref 类命令。
REF_COMMANDS = (
    "ref",
    "eqref",
    "pageref",
    "autoref",
    "cref",
    "Cref",
    "nameref",
)

# 条件编译命令前缀；若存在，说明工程可能依赖构建上下文。
CONDITIONAL_PREFIXES = ("if", "ifx", "ifdefined", "ifnum", "ifdim", "ifodd", "ifmmode")

# 规则文件。预检查脚本并不解析规则语义，但会检查这些文件是否存在，
# 以便尽早暴露 skill 目录结构损坏的问题。
REQUIRED_RULE_FILES = [
    "rules/supported_envs.md",
    "rules/downgrade_policy.md",
    "rules/acceptance_criteria.md",
]


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
        问题严重级别，取值为 INFO / WARN / ERROR。
    code:
        简短、稳定的机器可读代码，便于后续脚本过滤和汇总。
    message:
        面向人的问题描述。
    file:
        问题关联的文件相对路径；若无法绑定到某个文件，可为 None。
    line:
        近似行号；若无法可靠定位，则为 None。
    details:
        附加信息，保留为字典，便于 JSON 报告中结构化呈现。
    """
    severity: str
    code: str
    message: str
    file: Optional[str] = None
    line: Optional[int] = None
    details: dict = field(default_factory=dict)


@dataclass
class FileRecord:
    """
    表示一个已成功读取并纳入扫描范围的 TeX 文件。

    raw_text:
        原始文本。
    clean_text:
        去除注释后的文本。这里保留原始行数结构，使正则匹配结果仍能映射回近似行号。
    """
    path: Path
    raw_text: str
    clean_text: str


@dataclass
class PrecheckReport:
    """
    最终报告对象。该对象既用于 JSON 输出，也用于 Markdown 渲染。

    设计目标：
    - 字段扁平清晰；
    - 后续脚本无需再猜测预检查结论；
    - 既能支持机器消费，也能支持人阅读。
    """
    status: str
    can_continue: bool
    project_root: str
    skill_root: str
    main_tex: Optional[str]
    scanned_tex_files: list[str]
    scanned_tex_file_count: int
    findings: list[dict]
    metrics: dict
    inventory: dict
    summary: dict
    recommendations: list[str]


# -----------------------------------------------------------------------------
# 工具函数
# -----------------------------------------------------------------------------

def build_argument_parser() -> argparse.ArgumentParser:
    """
    构造命令行参数解析器。

    当前脚本只暴露最必要的参数，不引入多余控制开关。
    这样做有两个好处：
    1. 降低用户误用复杂度；
    2. 保证 skill 内部流程的一致性。
    """
    parser = argparse.ArgumentParser(
        description="Precheck a LaTeX project before LaTeX -> Word conversion."
    )
    parser.add_argument(
        "--project-root",
        required=True,
        help="LaTeX 工程根目录。",
    )
    parser.add_argument(
        "--main-tex",
        default=None,
        help="主 TeX 文件路径。可为相对 project-root 的路径；若不提供，将自动识别。",
    )
    parser.add_argument(
        "--work-root",
        default=None,
        help="流水线工作目录；默认 <project-root-parent>/<project-name>__latex_to_word_work。",
    )
    parser.add_argument(
        "--json-out",
        default=None,
        help="JSON 报告输出路径；默认输出到 <work-root>/stage_precheck/precheck-report.json",
    )
    parser.add_argument(
        "--md-out",
        default=None,
        help="Markdown 报告输出路径；默认输出到 <work-root>/stage_precheck/precheck-report.md",
    )
    return parser


def safe_relative(path: Path, root: Path) -> str:
    """
    将绝对路径尽量转换为相对于 root 的路径字符串。
    若转换失败，则返回标准化绝对路径字符串。

    这样做的目的，是让报告内容更适合用户阅读，
    同时避免在 JSON 中混入太多冗长绝对路径。
    """
    try:
        return str(path.resolve().relative_to(root.resolve()))
    except Exception:
        return str(path.resolve())


def read_text_file(path: Path) -> str:
    """
    以 UTF-8 优先的策略读取文本文件。

    说明：
    - LaTeX 工程常见编码并不统一；
    - 这里使用多编码回退以提高容错性；
    - 若全部失败，再抛出异常，由上层记录为 finding。
    """
    encodings = ("utf-8", "utf-8-sig", "gbk", "cp936", "latin-1")
    last_error: Optional[Exception] = None
    for encoding in encodings:
        try:
            return path.read_text(encoding=encoding)
        except Exception as exc:  # pragma: no cover - 属于容错路径
            last_error = exc
    raise RuntimeError(f"无法读取文件: {path}") from last_error


def strip_latex_comments(text: str) -> str:
    """
    移除 LaTeX 注释，同时尽量保留原始行号结构。

    实现策略：
    - 按行处理；
    - 删除未被反斜杠转义的 '%' 之后的内容；
    - 保留换行符数量，便于后续 match 偏移映射回近似行号。

    注意：
    - 该实现是工程预检查级别的近似处理，不是完整的 TeX 词法分析器；
    - 对 verbatim / minted 一类环境中的 '%'，理论上不应视为注释；
      但在预检查阶段，这种近似通常足够且可接受。
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
        if cut_index is not None:
            stripped_lines.append(line[:cut_index])
        else:
            stripped_lines.append(line)
    return "\n".join(stripped_lines)


def line_number_from_offset(text: str, offset: int) -> int:
    """
    根据字符偏移估算所在行号，行号从 1 开始。

    这里无需实现列号，因为当前 skill 的后续脚本主要只需要行级定位。
    """
    return text.count("\n", 0, offset) + 1


def split_csv_payload(payload: str) -> list[str]:
    """
    解析形如 '{a,b,c}' 中的内容载荷，返回去空白后的条目列表。

    用于处理：
    - \\bibliography{a,b}
    - \\cite{key1,key2}
    - \\cref{fig:a,tab:b}
    """
    items = []
    for item in payload.split(","):
        token = item.strip()
        if token:
            items.append(token)
    return items


def resolve_path_with_extensions(base_dir: Path, target: str, extensions: Iterable[str]) -> Optional[Path]:
    """
    在给定目录下尝试解析资源路径。

    解析逻辑：
    1. 若 target 自带后缀，则直接检查该路径；
    2. 若 target 无后缀，则依次尝试附加给定扩展名；
    3. 若 target 已是可存在路径，则返回其 resolve 后结果；
    4. 若全部失败，返回 None。

    说明：
    - 这里不做额外目录搜索，不跨目录猜测；
    - 解析应尽量与 LaTeX 的“相对当前文件目录”习惯一致。
    """
    target_path = Path(target)

    # 情况 1：用户已显式写出扩展名。
    if target_path.suffix:
        candidate = (base_dir / target_path).resolve()
        return candidate if candidate.exists() else None

    # 情况 2：无扩展名，按约定扩展补全。
    for ext in extensions:
        candidate = (base_dir / f"{target}{ext}").resolve()
        if candidate.exists():
            return candidate

    # 情况 3：某些文件名本身可能是目录 / 特殊写法；这里仅做最后一次直接检查。
    candidate = (base_dir / target_path).resolve()
    return candidate if candidate.exists() else None


def detect_main_tex(project_root: Path) -> tuple[Optional[Path], list[Finding]]:
    """
    自动识别主 TeX 文件。

    识别策略：
    - 在 project_root 下递归寻找 .tex 文件；
    - 按以下规则打分：
      * 含 \\documentclass: +5
      * 含 \\begin{document}: +5
      * 文件名为 main.tex: +3
      * 文件名与工程目录同名: +2
    - 选择得分最高者；
    - 若没有候选，返回 None；
    - 若候选过于模糊，仍返回最优文件，但记录 WARN。

    说明：
    自动识别不可能覆盖所有工程，但该启发式对常见论文/报告工程足够稳定。
    """
    findings: list[Finding] = []
    tex_files = sorted(project_root.rglob("*.tex"))
    if not tex_files:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="NO_TEX_FILES",
                message="在工程目录下未发现任何 .tex 文件。",
                file=None,
                line=None,
            )
        )
        return None, findings

    scored: list[tuple[int, Path]] = []
    for tex_file in tex_files:
        try:
            text = read_text_file(tex_file)
        except Exception as exc:
            findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="UNREADABLE_TEX_CANDIDATE",
                    message=f"候选 TeX 文件无法读取，自动识别时已忽略：{tex_file.name}",
                    file=safe_relative(tex_file, project_root),
                    line=None,
                    details={"error": str(exc)},
                )
            )
            continue

        score = 0
        if re.search(r"\\documentclass(?:\[[^\]]*\])?\{[^}]+\}", text):
            score += 5
        if r"\begin{document}" in text:
            score += 5
        if tex_file.name.lower() == "main.tex":
            score += 3
        if tex_file.stem.lower() == project_root.name.lower():
            score += 2

        scored.append((score, tex_file))

    if not scored:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="NO_READABLE_TEX_CANDIDATES",
                message="发现了 .tex 文件，但没有可读取的候选主文件。",
            )
        )
        return None, findings

    scored.sort(key=lambda item: (item[0], str(item[1]).lower()), reverse=True)
    best_score, best_file = scored[0]

    if best_score <= 0:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="MAIN_TEX_NOT_DETECTABLE",
                message="无法自动识别主 TeX 文件，请显式传入 --main-tex。",
            )
        )
        return None, findings

    # 若最高分并列，给出告警，但仍选择排序后的第一个。
    tied = [path for score, path in scored if score == best_score]
    if len(tied) > 1:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="MAIN_TEX_AMBIGUOUS",
                message="自动识别到多个可能的主文件，已按启发式选择其中一个。",
                file=safe_relative(best_file, project_root),
                details={
                    "selected": safe_relative(best_file, project_root),
                    "candidates": [safe_relative(p, project_root) for p in tied],
                },
            )
        )

    return best_file.resolve(), findings


def parse_include_targets(clean_text: str) -> list[tuple[str, int]]:
    """
    提取 \\input{...} / \\include{...} 目标。

    返回列表元素为 (target, line_number)。

    这里只处理最常见的写法：
    - \\input{path}
    - \\include{path}

    若工程使用复杂宏包包装 include 逻辑，属于高风险情形，会在其他检查中暴露。
    """
    pattern = re.compile(r"\\(?:input|include)\{([^}]+)\}")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        target = match.group(1).strip()
        line_no = line_number_from_offset(clean_text, match.start())
        if target:
            results.append((target, line_no))
    return results


def parse_includegraphics_targets(clean_text: str) -> list[tuple[str, int]]:
    """
    提取 \\includegraphics 命令中的图片路径。

    支持带可选参数与不带可选参数两种常见写法：
    - \\includegraphics{fig/a}
    - \\includegraphics[width=0.8\\textwidth]{fig/a}

    返回 (target, line_number) 列表。
    """
    pattern = re.compile(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        target = match.group(1).strip()
        line_no = line_number_from_offset(clean_text, match.start())
        if target:
            results.append((target, line_no))
    return results


def parse_bibliography_targets(clean_text: str) -> list[tuple[str, int, str]]:
    """
    提取 bibliography 资源声明。

    支持两类常见写法：
    1. \\bibliography{refs,more_refs}
    2. \\addbibresource{refs.bib}

    返回 (target, line_number, source_command) 列表。
    """
    results: list[tuple[str, int, str]] = []

    bibliography_pattern = re.compile(r"\\bibliography\{([^}]+)\}")
    for match in bibliography_pattern.finditer(clean_text):
        line_no = line_number_from_offset(clean_text, match.start())
        for token in split_csv_payload(match.group(1)):
            results.append((token, line_no, "bibliography"))

    addbib_pattern = re.compile(r"\\addbibresource(?:\[[^\]]*\])?\{([^}]+)\}")
    for match in addbib_pattern.finditer(clean_text):
        line_no = line_number_from_offset(clean_text, match.start())
        token = match.group(1).strip()
        if token:
            results.append((token, line_no, "addbibresource"))

    return results


def parse_labels(clean_text: str) -> list[tuple[str, int]]:
    """
    提取 \\label{...}。

    预检查只做“显式标签存在性”判断，不试图理解标签类型。
    """
    pattern = re.compile(r"\\label\{([^}]+)\}")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        label = match.group(1).strip()
        line_no = line_number_from_offset(clean_text, match.start())
        if label:
            results.append((label, line_no))
    return results


def parse_refs(clean_text: str) -> list[tuple[str, int, str]]:
    """
    提取各种 ref 类命令。

    返回 (label, line_number, command_name) 列表。
    """
    pattern = re.compile(r"\\(" + "|".join(REF_COMMANDS) + r")\{([^}]+)\}")
    results: list[tuple[str, int, str]] = []
    for match in pattern.finditer(clean_text):
        command = match.group(1)
        payload = match.group(2)
        line_no = line_number_from_offset(clean_text, match.start())
        for label in split_csv_payload(payload):
            results.append((label, line_no, command))
    return results


def parse_cites(clean_text: str) -> list[tuple[str, int, str]]:
    """
    提取 cite 类命令中的 citation keys。

    该正则允许最多两个可选参数，覆盖常见的：
    - \\cite{a}
    - \\cite[p.1]{a}
    - \\cite[see][p.1]{a}
    """
    pattern = re.compile(
        r"\\(" + "|".join(CITE_COMMANDS) + r")\*?(?:\[[^\]]*\]){0,2}\{([^}]+)\}"
    )
    results: list[tuple[str, int, str]] = []
    for match in pattern.finditer(clean_text):
        command = match.group(1)
        payload = match.group(2)
        line_no = line_number_from_offset(clean_text, match.start())
        for key in split_csv_payload(payload):
            results.append((key, line_no, command))
    return results


def parse_begin_environments(clean_text: str) -> list[tuple[str, int]]:
    """
    提取 \\begin{...} 环境名。

    用途：
    - 统计高风险环境；
    - 识别 theorem / algorithm / code / tikz 等对象；
    - 帮助发现未知环境使用规模。
    """
    pattern = re.compile(r"\\begin\{([A-Za-z*@]+)\}")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        env_name = match.group(1).strip()
        line_no = line_number_from_offset(clean_text, match.start())
        if env_name:
            results.append((env_name, line_no))
    return results


def parse_custom_environment_definitions(clean_text: str) -> list[tuple[str, int]]:
    """
    提取 \\newenvironment / \\renewenvironment 定义的环境名。

    这类环境通常是后续转换的高风险来源。
    """
    pattern = re.compile(r"\\(?:newenvironment|renewenvironment)\{([A-Za-z*@]+)\}")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        env_name = match.group(1).strip()
        line_no = line_number_from_offset(clean_text, match.start())
        if env_name:
            results.append((env_name, line_no))
    return results


def parse_custom_command_definitions(clean_text: str) -> list[tuple[str, int]]:
    """
    提取常见自定义命令定义。

    这里只统计和标记，不尝试理解命令语义。
    """
    pattern = re.compile(
        r"\\(?:newcommand|renewcommand|providecommand|DeclareMathOperator|NewDocumentCommand|RenewDocumentCommand)\*?\s*\{?\\([A-Za-z@]+)\}?"
    )
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        cmd_name = match.group(1).strip()
        line_no = line_number_from_offset(clean_text, match.start())
        if cmd_name:
            results.append((cmd_name, line_no))
    return results


def parse_conditionals(clean_text: str) -> list[tuple[str, int]]:
    """
    提取条件编译命令。

    说明：
    - 若工程大量依赖条件编译，预检查应提醒用户该工程依赖构建上下文；
    - 这里不求完整 TeX 语义，只做风险识别。
    """
    pattern = re.compile(r"\\(if[A-Za-z@]*|else|fi)\b")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        token = match.group(1)
        line_no = line_number_from_offset(clean_text, match.start())
        if token:
            results.append((token, line_no))
    return results


def parse_table_risk_commands(clean_text: str) -> list[tuple[str, int]]:
    """
    提取高风险表格相关命令，例如 \\multirow 和 \\multicolumn。

    这类命令一旦出现，就说明表格应被视为高风险。
    """
    pattern = re.compile(r"\\(" + "|".join(HIGH_RISK_TABLE_COMMANDS) + r")\b")
    results: list[tuple[str, int]] = []
    for match in pattern.finditer(clean_text):
        cmd_name = match.group(1)
        line_no = line_number_from_offset(clean_text, match.start())
        if cmd_name:
            results.append((cmd_name, line_no))
    return results


def extract_bib_keys_from_text(text: str) -> set[str]:
    """
    从 .bib 文件中提取 citation keys。

    简化策略：
    - 匹配形如 @article{key, ...} 的开头；
    - 对大小写不敏感；
    - 忽略更复杂的 BibTeX / BibLaTeX 细节。

    该实现不尝试成为完整的 bibliography 解析器，
    但足以支持预检查阶段的“明显缺键”判断。
    """
    pattern = re.compile(r"@\w+\s*\{\s*([^,\s]+)\s*,", re.IGNORECASE)
    return {match.group(1).strip() for match in pattern.finditer(text) if match.group(1).strip()}


def write_json(path: Path, payload: dict) -> None:
    """
    写出 JSON 报告。

    使用 UTF-8，确保中文报告可读。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def render_markdown_report(report: PrecheckReport) -> str:
    """
    将预检查结果渲染为 Markdown 报告。

    该报告面向用户与开发者共同阅读，因此强调：
    - 概况清晰；
    - 风险清单可定位；
    - 后续动作明确。
    """
    lines: list[str] = []
    lines.append("# Precheck Report")
    lines.append("")
    lines.append(f"- Status: **{report.status}**")
    lines.append(f"- Can continue: **{report.can_continue}**")
    lines.append(f"- Project root: `{report.project_root}`")
    lines.append(f"- Skill root: `{report.skill_root}`")
    lines.append(f"- Main TeX: `{report.main_tex or 'N/A'}`")
    lines.append(f"- Scanned TeX files: **{report.scanned_tex_file_count}**")
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

    lines.append("## Inventory")
    lines.append("")
    for key, value in report.inventory.items():
        if isinstance(value, list):
            lines.append(f"- {key}:")
            if value:
                for item in value:
                    lines.append(f"  - {item}")
            else:
                lines.append("  - (none)")
        else:
            lines.append(f"- {key}: **{value}**")
    lines.append("")

    lines.append("## Findings")
    lines.append("")
    if not report.findings:
        lines.append("- No findings.")
        lines.append("")
    else:
        # 按严重级别分组，便于用户优先处理 ERROR / WARN。
        grouped: dict[str, list[dict]] = defaultdict(list)
        for finding in report.findings:
            grouped[finding["severity"]].append(finding)

        for severity in [SEVERITY_ERROR, SEVERITY_WARN, SEVERITY_INFO]:
            if severity not in grouped:
                continue
            lines.append(f"### {severity}")
            lines.append("")
            for item in grouped[severity]:
                location = ""
                if item.get("file"):
                    location = f"`{item['file']}`"
                    if item.get("line"):
                        location += f":{item['line']}"
                code = item.get("code", "UNKNOWN")
                message = item.get("message", "")
                if location:
                    lines.append(f"- **[{code}]** {message} ({location})")
                else:
                    lines.append(f"- **[{code}]** {message}")
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


def write_markdown(path: Path, content: str) -> None:
    """
    写出 Markdown 报告。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


# -----------------------------------------------------------------------------
# 核心扫描逻辑
# -----------------------------------------------------------------------------

def locate_skill_root() -> Path:
    """
    定位 skill 根目录。

    目录约定：
    - 本脚本位于 <skill_root>/scripts/precheck.py
    - 因此上两级目录的父目录即为 skill 根目录

    若未来目录布局变化，应修改此函数，而不是在多个地方散写路径规则。
    """
    return Path(__file__).resolve().parents[1]


def check_rule_files(skill_root: Path) -> list[Finding]:
    """
    检查 rules/ 下的必需规则文件是否存在。

    这一步不是为了在预检查中解析规则语义，
    而是为了尽早发现 skill 目录不完整的问题。
    """
    findings: list[Finding] = []
    for relative in REQUIRED_RULE_FILES:
        path = (skill_root / relative).resolve()
        if not path.exists():
            findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="RULE_FILE_MISSING",
                    message="缺少规则文件；后续脚本行为可能不完整。",
                    file=safe_relative(path, skill_root),
                    line=None,
                    details={"required_file": relative},
                )
            )
    return findings


def resolve_main_tex(project_root: Path, main_tex_arg: Optional[str]) -> tuple[Optional[Path], list[Finding]]:
    """
    根据用户参数或自动识别结果确定主 TeX 文件。
    """
    findings: list[Finding] = []

    if main_tex_arg:
        candidate = Path(main_tex_arg)
        if not candidate.is_absolute():
            candidate = (project_root / candidate).resolve()
        if not candidate.exists():
            findings.append(
                Finding(
                    severity=SEVERITY_ERROR,
                    code="MAIN_TEX_NOT_FOUND",
                    message="指定的主 TeX 文件不存在。",
                    file=safe_relative(candidate, project_root),
                )
            )
            return None, findings
        if candidate.suffix.lower() != ".tex":
            findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="MAIN_TEX_NON_TEX_SUFFIX",
                    message="指定的主文件不是 .tex 后缀，请确认是否正确。",
                    file=safe_relative(candidate, project_root),
                )
            )
        return candidate.resolve(), findings

    auto_main, auto_findings = detect_main_tex(project_root)
    findings.extend(auto_findings)
    return auto_main, findings


def collect_tex_closure(project_root: Path, main_tex: Path) -> tuple[dict[Path, FileRecord], list[Finding]]:
    """
    从主文件出发，递归收集所有可达 TeX 文件。

    收集逻辑：
    - 读取当前文件；
    - 提取 \\input / \\include；
    - 按当前文件所在目录解析相对路径；
    - 若目标存在，则递归；
    - 若目标缺失，则记录 ERROR；
    - 用 visited 避免重复处理与循环引用。

    这是预检查的核心步骤之一，因为后续所有统计都建立在“主文件可达闭包”上。
    """
    findings: list[Finding] = []
    records: dict[Path, FileRecord] = {}
    visited: set[Path] = set()
    stack: list[Path] = [main_tex.resolve()]

    while stack:
        tex_file = stack.pop()
        tex_file = tex_file.resolve()
        if tex_file in visited:
            continue
        visited.add(tex_file)

        try:
            raw_text = read_text_file(tex_file)
        except Exception as exc:
            findings.append(
                Finding(
                    severity=SEVERITY_ERROR,
                    code="TEX_FILE_READ_ERROR",
                    message="TeX 文件无法读取。",
                    file=safe_relative(tex_file, project_root),
                    details={"error": str(exc)},
                )
            )
            continue

        clean_text = strip_latex_comments(raw_text)
        records[tex_file] = FileRecord(path=tex_file, raw_text=raw_text, clean_text=clean_text)

        include_targets = parse_include_targets(clean_text)
        for target, line_no in include_targets:
            resolved = resolve_path_with_extensions(tex_file.parent, target, TEX_EXTENSIONS)
            if resolved is None:
                # LaTeX 里 \include 常常不写 .tex，这里已经做了常见补全。
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="MISSING_INCLUDED_TEX",
                        message="引用的 TeX 子文件不存在，输入链不闭合。",
                        file=safe_relative(tex_file, project_root),
                        line=line_no,
                        details={"target": target},
                    )
                )
                continue
            stack.append(resolved)

    return records, findings


def analyze_project(
    project_root: Path,
    skill_root: Path,
    main_tex: Path,
    records: dict[Path, FileRecord],
    initial_findings: list[Finding],
) -> PrecheckReport:
    """
    基于已收集的 TeX 闭包执行工程级分析，并生成最终报告对象。

    该函数负责：
    - 聚合 labels / refs / cites / bibliography；
    - 检查图片缺失、bib 缺失；
    - 检查明显未定义的引用键；
    - 识别高风险对象；
    - 生成最终状态、建议和结构化报告。
    """
    findings: list[Finding] = list(initial_findings)

    # ---------------------------
    # 工程级统计容器
    # ---------------------------
    all_labels: dict[str, list[tuple[Path, int]]] = defaultdict(list)
    all_refs: list[tuple[str, Path, int, str]] = []
    all_cites: list[tuple[str, Path, int, str]] = []
    all_bib_paths: set[Path] = set()
    image_hits: list[Path] = []
    missing_images: list[tuple[Path, int, str]] = []
    environment_counter: Counter[str] = Counter()
    custom_environment_defs: dict[str, list[tuple[Path, int]]] = defaultdict(list)
    custom_command_defs: dict[str, list[tuple[Path, int]]] = defaultdict(list)
    conditional_tokens: list[tuple[str, Path, int]] = []
    table_risk_commands: list[tuple[str, Path, int]] = []
    thebibliography_used = False

    # 用于复杂表格粗略判定：只要某个文件中同时出现长表、多行/多列等结构，
    # 即将其视为复杂表格文件。
    complex_table_files: set[Path] = set()

    # ---------------------------
    # 单文件分析
    # ---------------------------
    for tex_path, record in records.items():
        clean_text = record.clean_text

        # 文档结构完整性提示：主文件最好含 \documentclass 和 \begin{document}
        if tex_path == main_tex:
            if not re.search(r"\\documentclass(?:\[[^\]]*\])?\{[^}]+\}", clean_text):
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="MAIN_TEX_NO_DOCUMENTCLASS",
                        message="主文件中未发现 \\documentclass，主文件可能不正确或工程结构不完整。",
                        file=safe_relative(tex_path, project_root),
                    )
                )
            if r"\begin{document}" not in clean_text:
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="MAIN_TEX_NO_BEGIN_DOCUMENT",
                        message="主文件中未发现 \\begin{document}，无法确认其为可编译入口。",
                        file=safe_relative(tex_path, project_root),
                    )
                )

        # 标签
        for label, line_no in parse_labels(clean_text):
            all_labels[label].append((tex_path, line_no))

        # ref
        for label, line_no, command in parse_refs(clean_text):
            all_refs.append((label, tex_path, line_no, command))

        # cite
        for key, line_no, command in parse_cites(clean_text):
            all_cites.append((key, tex_path, line_no, command))

        # bibliography 目标
        for target, line_no, source_command in parse_bibliography_targets(clean_text):
            resolved = resolve_path_with_extensions(tex_path.parent, target, BIB_EXTENSIONS)
            if resolved is None:
                findings.append(
                    Finding(
                        severity=SEVERITY_ERROR,
                        code="MISSING_BIB_RESOURCE",
                        message="参考文献资源缺失。",
                        file=safe_relative(tex_path, project_root),
                        line=line_no,
                        details={"target": target, "source_command": source_command},
                    )
                )
            else:
                all_bib_paths.add(resolved)

        # thebibliography 环境
        for env_name, _line_no in parse_begin_environments(clean_text):
            environment_counter[env_name] += 1
            if env_name == "thebibliography":
                thebibliography_used = True

        # 自定义环境定义
        for env_name, line_no in parse_custom_environment_definitions(clean_text):
            custom_environment_defs[env_name].append((tex_path, line_no))

        # 自定义命令定义
        for cmd_name, line_no in parse_custom_command_definitions(clean_text):
            custom_command_defs[cmd_name].append((tex_path, line_no))

        # 条件编译
        for token, line_no in parse_conditionals(clean_text):
            conditional_tokens.append((token, tex_path, line_no))

        # 表格高风险命令
        file_table_risk_commands = parse_table_risk_commands(clean_text)
        for cmd_name, line_no in file_table_risk_commands:
            table_risk_commands.append((cmd_name, tex_path, line_no))

        # 粗略复杂表格文件判定
        # 这里只要“高风险表格环境 + 高风险表格命令”同时出现，即认为该文件含复杂表格。
        file_envs = {env for env, _ in parse_begin_environments(clean_text)}
        if (
            ("longtable" in file_envs or "tabularx" in file_envs)
            and len(file_table_risk_commands) > 0
        ) or (
            "longtable" in file_envs
            and "tabular" in file_envs
        ):
            complex_table_files.add(tex_path)

        # 图片
        for image_target, line_no in parse_includegraphics_targets(clean_text):
            resolved_image = resolve_path_with_extensions(tex_path.parent, image_target, IMAGE_EXTENSIONS)
            if resolved_image is None:
                missing_images.append((tex_path, line_no, image_target))
            else:
                image_hits.append(resolved_image)

    # ---------------------------
    # 重复标签检查
    # ---------------------------
    duplicate_labels = {label: locations for label, locations in all_labels.items() if len(locations) > 1}
    for label, locations in duplicate_labels.items():
        first_file, first_line = locations[0]
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="DUPLICATE_LABEL",
                message="发现重复标签；后续交叉引用可能不稳定。",
                file=safe_relative(first_file, project_root),
                line=first_line,
                details={
                    "label": label,
                    "locations": [
                        {"file": safe_relative(path, project_root), "line": line_no}
                        for path, line_no in locations
                    ],
                },
            )
        )

    # ---------------------------
    # 未定义 label 引用检查
    # ---------------------------
    defined_label_names = set(all_labels.keys())
    for label, tex_path, line_no, command in all_refs:
        if label not in defined_label_names:
            findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="UNDEFINED_LABEL_REFERENCE",
                    message="发现引用了未定义的标签。",
                    file=safe_relative(tex_path, project_root),
                    line=line_no,
                    details={"label": label, "command": command},
                )
            )

    # ---------------------------
    # bibliography / cite 检查
    # ---------------------------
    bib_keys: set[str] = set()
    unreadable_bib_files: list[str] = []

    for bib_path in sorted(all_bib_paths):
        try:
            bib_text = read_text_file(bib_path)
            bib_keys.update(extract_bib_keys_from_text(bib_text))
        except Exception as exc:
            unreadable_bib_files.append(safe_relative(bib_path, project_root))
            findings.append(
                Finding(
                    severity=SEVERITY_WARN,
                    code="UNREADABLE_BIB_FILE",
                    message="参考文献文件存在，但无法读取或解析。",
                    file=safe_relative(bib_path, project_root),
                    details={"error": str(exc)},
                )
            )

    cite_count = len(all_cites)
    if cite_count > 0 and not all_bib_paths and not thebibliography_used:
        findings.append(
            Finding(
                severity=SEVERITY_ERROR,
                code="CITE_WITHOUT_BIBLIOGRAPHY",
                message="文档中存在 \\cite，但未发现 .bib 资源，也未发现 thebibliography 环境。",
                file=safe_relative(main_tex, project_root),
                details={"cite_count": cite_count},
            )
        )

    if all_bib_paths and bib_keys:
        for key, tex_path, line_no, command in all_cites:
            if key not in bib_keys:
                findings.append(
                    Finding(
                        severity=SEVERITY_WARN,
                        code="UNDEFINED_CITATION_KEY",
                        message="发现未在 .bib 中定义的引用键。",
                        file=safe_relative(tex_path, project_root),
                        line=line_no,
                        details={"key": key, "command": command},
                    )
                )

    # 若使用了手工 thebibliography，则 cite 键不一定能从 .bib 验证。
    if cite_count > 0 and thebibliography_used and not all_bib_paths:
        findings.append(
            Finding(
                severity=SEVERITY_INFO,
                code="MANUAL_BIBLIOGRAPHY_DETECTED",
                message="检测到 thebibliography 环境；citation key 的自动完整性验证能力有限。",
                file=safe_relative(main_tex, project_root),
            )
        )

    # ---------------------------
    # 缺图检查
    # ---------------------------
    for tex_path, line_no, image_target in missing_images:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="MISSING_IMAGE",
                message="图片资源缺失。",
                file=safe_relative(tex_path, project_root),
                line=line_no,
                details={"target": image_target},
            )
        )

    # 当缺图比例过高时，升级为 ERROR。
    total_image_commands = len(image_hits) + len(missing_images)
    if total_image_commands > 0:
        missing_ratio = len(missing_images) / total_image_commands
        if len(missing_images) >= 3 and missing_ratio >= 0.5:
            findings.append(
                Finding(
                    severity=SEVERITY_ERROR,
                    code="TOO_MANY_MISSING_IMAGES",
                    message="缺失图片过多，已影响继续转换的可靠性。",
                    file=safe_relative(main_tex, project_root),
                    details={
                        "missing_images": len(missing_images),
                        "total_includegraphics": total_image_commands,
                        "missing_ratio": round(missing_ratio, 4),
                    },
                )
            )

    # ---------------------------
    # 高风险对象检查
    # ---------------------------
    high_risk_env_hits = {env: count for env, count in environment_counter.items() if env in HIGH_RISK_ENVIRONMENTS}
    if high_risk_env_hits:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="HIGH_RISK_ENVIRONMENTS",
                message="检测到高风险 LaTeX 环境，后续转换需重点复核。",
                file=safe_relative(main_tex, project_root),
                details={"environments": high_risk_env_hits},
            )
        )

    if complex_table_files:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="COMPLEX_TABLES_DETECTED",
                message="检测到复杂表格文件，表格在 Word 中可能需要人工修复。",
                file=safe_relative(main_tex, project_root),
                details={
                    "files": [safe_relative(path, project_root) for path in sorted(complex_table_files)]
                },
            )
        )

    if conditional_tokens:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="CONDITIONAL_COMPILATION_DETECTED",
                message="检测到条件编译命令，工程可能依赖构建上下文；转换结果需重点核对。",
                file=safe_relative(main_tex, project_root),
                details={"conditional_token_count": len(conditional_tokens)},
            )
        )

    if custom_environment_defs:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="CUSTOM_ENVIRONMENTS_DETECTED",
                message="检测到自定义环境定义，可能需要降级或人工修复。",
                file=safe_relative(main_tex, project_root),
                details={
                    "custom_environments": {
                        name: [
                            {"file": safe_relative(path, project_root), "line": line_no}
                            for path, line_no in locations
                        ]
                        for name, locations in sorted(custom_environment_defs.items())
                    }
                },
            )
        )

    if custom_command_defs:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="CUSTOM_COMMANDS_DETECTED",
                message="检测到自定义命令定义；若大量参与正文结构，后续转换风险会升高。",
                file=safe_relative(main_tex, project_root),
                details={"custom_command_count": len(custom_command_defs)},
            )
        )

    # 未知环境提示。为了避免噪声，仅在“环境已被使用，但既不是 KNOWN_ENVIRONMENTS，
    # 也不是自定义环境定义”的情况下记录。
    custom_env_names = set(custom_environment_defs.keys())
    unknown_used_envs = {
        env: count
        for env, count in environment_counter.items()
        if env not in KNOWN_ENVIRONMENTS and env not in custom_env_names
    }
    if unknown_used_envs:
        findings.append(
            Finding(
                severity=SEVERITY_WARN,
                code="UNKNOWN_ENVIRONMENTS_DETECTED",
                message="检测到未在当前支持集合中登记的环境，需人工确认其转换风险。",
                file=safe_relative(main_tex, project_root),
                details={"unknown_environments": unknown_used_envs},
            )
        )

    # ---------------------------
    # 状态判定
    # ---------------------------
    error_count = sum(1 for f in findings if f.severity == SEVERITY_ERROR)
    warn_count = sum(1 for f in findings if f.severity == SEVERITY_WARN)
    info_count = sum(1 for f in findings if f.severity == SEVERITY_INFO)

    if error_count > 0:
        status = STATUS_FAIL
        can_continue = False
    elif warn_count > 0:
        status = STATUS_PASS_WITH_WARNINGS
        can_continue = True
    else:
        status = STATUS_PASS
        can_continue = True

    # ---------------------------
    # 建议生成
    # ---------------------------
    recommendations: list[str] = []

    if status == STATUS_FAIL:
        recommendations.append("先修复所有 ERROR 级问题，再进入规范化和 Pandoc 主转换阶段。")
    else:
        recommendations.append("可进入 normalize_tex.py 阶段，但需保留本预检查报告供后续脚本引用。")

    if missing_images:
        recommendations.append("补齐缺失图片资源，或确认这些图片是否应按降级策略处理。")

    if cite_count > 0 and (not all_bib_paths and not thebibliography_used):
        recommendations.append("补充 .bib 文件或确认是否应改用 thebibliography。")

    if duplicate_labels:
        recommendations.append("修复重复标签，避免 Word 中交叉引用恢复不稳定。")

    if custom_environment_defs or custom_command_defs:
        recommendations.append("在规范化阶段重点处理自定义命令/环境，避免 Pandoc 主转换时静默降级。")

    if high_risk_env_hits or complex_table_files:
        recommendations.append("将复杂表格、子图、TikZ、算法和代码环境列为后续人工重点核查对象。")

    if conditional_tokens:
        recommendations.append("确认当前工程所需的条件编译分支；必要时先生成明确的展开版本再转换。")

    # ---------------------------
    # 指标与清单
    # ---------------------------
    scanned_tex_files = sorted(safe_relative(path, project_root) for path in records.keys())

    metrics = {
        "tex_file_count": len(records),
        "label_count": len(all_labels),
        "ref_count": len(all_refs),
        "cite_count": len(all_cites),
        "bibliography_file_count": len(all_bib_paths),
        "bib_key_count": len(bib_keys),
        "image_reference_count": total_image_commands,
        "missing_image_count": len(missing_images),
        "environment_count": sum(environment_counter.values()),
        "custom_environment_count": len(custom_environment_defs),
        "custom_command_count": len(custom_command_defs),
        "conditional_token_count": len(conditional_tokens),
        "complex_table_file_count": len(complex_table_files),
        "finding_info_count": info_count,
        "finding_warn_count": warn_count,
        "finding_error_count": error_count,
    }

    inventory = {
        "bibliography_files": [safe_relative(path, project_root) for path in sorted(all_bib_paths)],
        "unreadable_bibliography_files": unreadable_bib_files,
        "high_risk_environments": dict(sorted(high_risk_env_hits.items())),
        "complex_table_files": [safe_relative(path, project_root) for path in sorted(complex_table_files)],
        "custom_environments": sorted(custom_environment_defs.keys()),
        "custom_commands": sorted(custom_command_defs.keys()),
        "unknown_used_environments": sorted(unknown_used_envs.keys()),
    }

    summary = {
        "status": status,
        "can_continue": can_continue,
        "main_tex_detected": safe_relative(main_tex, project_root),
        "tex_files_scanned": len(records),
        "errors": error_count,
        "warnings": warn_count,
        "infos": info_count,
    }

    # 为保证报告稳定可比较，按严重级别 + code + file + line 排序。
    severity_order = {SEVERITY_ERROR: 0, SEVERITY_WARN: 1, SEVERITY_INFO: 2}
    findings.sort(
        key=lambda f: (
            severity_order.get(f.severity, 99),
            f.code,
            f.file or "",
            f.line or -1,
            f.message,
        )
    )

    report = PrecheckReport(
        status=status,
        can_continue=can_continue,
        project_root=str(project_root.resolve()),
        skill_root=str(skill_root.resolve()),
        main_tex=safe_relative(main_tex, project_root),
        scanned_tex_files=scanned_tex_files,
        scanned_tex_file_count=len(scanned_tex_files),
        findings=[asdict(f) for f in findings],
        metrics=metrics,
        inventory=inventory,
        summary=summary,
        recommendations=recommendations,
    )
    return report


# -----------------------------------------------------------------------------
# 主入口
# -----------------------------------------------------------------------------

def main() -> int:
    """
    主入口函数。

    流程顺序固定如下：
    1. 解析命令行参数；
    2. 定位工程根目录与 skill 根目录；
    3. 检查规则文件；
    4. 确定主 TeX 文件；
    5. 从主文件收集 TeX 闭包；
    6. 对工程执行分析并生成报告；
    7. 写出 JSON 和 Markdown 报告；
    8. 向 stdout 输出摘要；
    9. 根据总体状态返回退出码。
    """
    parser = build_argument_parser()
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    if not project_root.exists() or not project_root.is_dir():
        print(f"[ERROR] 无效的工程目录: {project_root}", file=sys.stderr)
        return 1

    skill_root = locate_skill_root()
    work_root = Path(args.work_root).resolve() if args.work_root else default_work_root_for_project(project_root)
    precheck_stage_dir = stage_dir(work_root, STAGE_PRECHECK)

    json_out = (
        Path(args.json_out).resolve()
        if args.json_out
        else (precheck_stage_dir / "precheck-report.json").resolve()
    )
    md_out = (
        Path(args.md_out).resolve()
        if args.md_out
        else (precheck_stage_dir / "precheck-report.md").resolve()
    )

    initial_findings: list[Finding] = []

    # 规则文件存在性检查
    initial_findings.extend(check_rule_files(skill_root))

    # 主文件识别
    main_tex, main_findings = resolve_main_tex(project_root, args.main_tex)
    initial_findings.extend(main_findings)

    if main_tex is None:
        # 连主文件都没有，无法继续进行闭包收集。
        status = STATUS_FAIL
        can_continue = False
        report = PrecheckReport(
            status=status,
            can_continue=can_continue,
            project_root=str(project_root),
            skill_root=str(skill_root),
            main_tex=None,
            scanned_tex_files=[],
            scanned_tex_file_count=0,
            findings=[asdict(f) for f in initial_findings],
            metrics={},
            inventory={},
            summary={
                "status": status,
                "can_continue": can_continue,
                "main_tex_detected": None,
                "tex_files_scanned": 0,
                "errors": sum(1 for f in initial_findings if f.severity == SEVERITY_ERROR),
                "warnings": sum(1 for f in initial_findings if f.severity == SEVERITY_WARN),
                "infos": sum(1 for f in initial_findings if f.severity == SEVERITY_INFO),
            },
            recommendations=[
                "显式指定 --main-tex，或修复工程目录中的 .tex 入口文件组织。"
            ],
        )
        write_json(json_out, asdict(report))
        write_markdown(md_out, render_markdown_report(report))
        try:
            update_stage_manifest(
                work_root,
                STAGE_PRECHECK,
                status=report.status,
                can_continue=report.can_continue,
                artifacts={
                    "precheck_report_json": json_out,
                    "precheck_report_md": md_out,
                    "project_root": project_root,
                },
                summary=report.summary,
                metrics=report.metrics,
            )
            update_manifest_artifacts(
                work_root,
                "reports",
                {
                    "precheck_report_json": json_out,
                    "precheck_report_md": md_out,
                },
            )
        except Exception:
            pass
        print(f"[{status}] precheck failed: main TeX file could not be resolved.")
        print(f"JSON report: {json_out}")
        print(f"Markdown report: {md_out}")
        return 1

    # TeX 闭包收集
    records, closure_findings = collect_tex_closure(project_root, main_tex)
    initial_findings.extend(closure_findings)

    # 工程分析
    report = analyze_project(
        project_root=project_root,
        skill_root=skill_root,
        main_tex=main_tex,
        records=records,
        initial_findings=initial_findings,
    )

    # 写出报告
    report_dict = asdict(report)
    write_json(json_out, report_dict)
    write_markdown(md_out, render_markdown_report(report))
    try:
        update_stage_manifest(
            work_root,
            STAGE_PRECHECK,
            status=report.status,
            can_continue=report.can_continue,
            artifacts={
                "precheck_report_json": json_out,
                "precheck_report_md": md_out,
                "project_root": project_root,
                "main_tex": report.main_tex,
            },
            summary=report.summary,
            metrics=report.metrics,
        )
        update_manifest_artifacts(
            work_root,
            "reports",
            {
                "precheck_report_json": json_out,
                "precheck_report_md": md_out,
            },
        )
    except Exception:
        pass

    # 向控制台打印摘要，便于 Codex / CLI / 用户快速判断。
    print(f"[{report.status}] Precheck completed.")
    print(f"Main TeX: {report.main_tex}")
    print(f"Work root: {work_root}")
    print(f"Scanned TeX files: {report.scanned_tex_file_count}")
    print(f"Errors: {report.metrics.get('finding_error_count', 0)}")
    print(f"Warnings: {report.metrics.get('finding_warn_count', 0)}")
    print(f"Infos: {report.metrics.get('finding_info_count', 0)}")
    print(f"JSON report: {json_out}")
    print(f"Markdown report: {md_out}")

    # FAIL 返回 1，其余返回 0，便于上层流程判断是否继续。
    return 1 if report.status == STATUS_FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
