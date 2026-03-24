#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_with_pandoc.py

功能概述
--------
本脚本用于在“规范化工作副本”基础上，调用 Pandoc 完成 LaTeX -> Word (.docx)
的主体转换。它是 Windows 11 下的原生入口，避免依赖 Bash。

它解决的问题包括：
1. 检查 Pandoc / reference.docx 等主转换依赖；
2. 读取 normalization-report.json，确认规范化阶段没有失败；
3. 自动确定主 TeX 文件、输出 docx 路径、参考样式模板路径；
4. 自动收集 bibliography 文件；
5. 自动构造 Pandoc 所需的 --resource-path；
6. 执行 Pandoc 主转换并保存完整日志；
7. 生成 JSON + Markdown 两份转换报告。

设计边界
--------
本脚本只负责“主体转换”，不做以下事情：
- 不修改原始工程；
- 不回写规范化前的源工程；
- 不做 docx 质量验收；
- 不生成人工修复清单；
- 不替代 postcheck_docx.py 的职责。

依赖
----
- Python 3.9+
- Pandoc

典型用法
--------
python scripts/convert_with_pandoc.py --work-root D:/work/my-paper__latex_to_word_work

或显式指定输出：
python scripts/convert_with_pandoc.py ^
  --work-root D:/work/my-paper__latex_to_word_work ^
  --output-docx D:/work/my-paper__latex_to_word_work/stage_convert/output.docx

输出
----
默认在 work-root 的 `stage_convert/` 下生成：
- output.docx
- pandoc-conversion.log
- pandoc-conversion-report.json
- pandoc-conversion-report.md

退出码
------
- 0: PASS 或 PASS_WITH_WARNINGS
- 1: FAIL
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shlex
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from docx_postprocess import run_docx_postprocess
from pipeline_common import locate_skill_root, read_text_file, split_csv_payload
from pipeline_layout import (
    STAGE_CONVERT,
    STAGE_NORMALIZE,
    best_effort_update_manifest,
    resolve_explicit_or_default,
    resolve_explicit_or_stage_input,
    resolve_explicit_or_stage_output,
    stage_dir,
)


STATUS_PASS = "PASS"
STATUS_PASS_WITH_WARNINGS = "PASS_WITH_WARNINGS"
STATUS_FAIL = "FAIL"

IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".pdf", ".svg", ".eps", ".bmp", ".tif", ".tiff"]
RESOURCE_DIR_LIMIT_DEFAULT = 200
COMMAND_LENGTH_LIMIT_DEFAULT = 7800


@dataclass
class ConversionContext:
    work_root: Path
    normalization_json: Path
    main_tex: Path
    tex_files: list[Path]
    output_docx: Path
    reference_doc: Path
    log_file: Path
    report_json: Path
    report_md: Path
    postprocess_report_json: Path
    metadata_json: Path
    resource_dirs_txt: Path
    bibs_txt: Path
    normalization_status: str
    normalization_can_continue: bool
    tex_file_count: int
    bibliographies: list[Path]
    thebibliography_used: bool
    cite_count: int
    resource_dirs: list[Path]
    resource_dirs_priority: list[Path]
    resource_dirs_original_count: int
    resource_dir_limit: int
    resource_dir_strategy: str
    command_length_limit: int
    command_length_estimate: int
    heading_numbering_mode: str
    guard_warnings: list[str]
    normalization_report: dict
    normalized_source_root: Path
    postprocess_report: Optional[dict]


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run Pandoc conversion from normalized LaTeX working copy to DOCX."
    )
    parser.add_argument("--work-root", required=True, help="规范化工作目录。应包含 normalization-report.json。")
    parser.add_argument("--main-tex", default=None, help="显式指定主 TeX 文件。默认从 normalization-report.json 读取。")
    parser.add_argument(
        "--normalization-json",
        default=None,
        help="规范化报告 JSON 路径。默认优先 <work-root>/stage_normalize/normalization-report.json，再回退 <work-root>/normalization-report.json",
    )
    parser.add_argument("--output-docx", default=None, help="输出 docx 路径。默认 <work-root>/stage_convert/output.docx")
    parser.add_argument(
        "--reference-doc",
        default=None,
        help="Word 样式模板路径。默认 <skill-root>/templates/reference.docx",
    )
    parser.add_argument("--log-file", default=None, help="Pandoc 日志路径。默认 <work-root>/stage_convert/pandoc-conversion.log")
    parser.add_argument(
        "--report-json",
        default=None,
        help="转换报告 JSON 路径。默认 <work-root>/stage_convert/pandoc-conversion-report.json",
    )
    parser.add_argument(
        "--report-md",
        default=None,
        help="转换报告 Markdown 路径。默认 <work-root>/stage_convert/pandoc-conversion-report.md",
    )
    parser.add_argument(
        "--max-resource-dirs",
        type=int,
        default=RESOURCE_DIR_LIMIT_DEFAULT,
        help=f"resource-path 目录数量保护阈值（默认 {RESOURCE_DIR_LIMIT_DEFAULT}）。",
    )
    parser.add_argument(
        "--max-command-length",
        type=int,
        default=COMMAND_LENGTH_LIMIT_DEFAULT,
        help=f"Pandoc 命令长度保护阈值（默认 {COMMAND_LENGTH_LIMIT_DEFAULT}）。",
    )
    parser.add_argument(
        "--heading-numbering-mode",
        choices=["template", "pandoc"],
        default="template",
        help=(
            "章节编号策略：template=依赖 reference.docx 的样式编号（不注入文本编号）；"
            "pandoc=使用 --number-sections 注入文本编号。默认 template。"
        ),
    )
    return parser


def resolve_bib(base_dir: Path, target: str) -> Optional[Path]:
    candidate = Path(target)
    if candidate.suffix:
        resolved = (base_dir / candidate).resolve()
        return resolved if resolved.exists() else None
    resolved = (base_dir / f"{target}.bib").resolve()
    return resolved if resolved.exists() else None


def collect_bibliographies_and_cites(tex_files: list[Path]) -> tuple[list[Path], bool, int]:
    addbib_pattern = re.compile(r"\\addbibresource(?:\[[^\]]*\])?\{([^}]+)\}")
    bibliography_pattern = re.compile(r"\\bibliography\{([^}]+)\}")
    thebibliography_pattern = re.compile(r"\\begin\{thebibliography\}")
    cite_pattern = re.compile(
        r"\\(?:cite|citep|citet|parencite|textcite|autocite|footcite|supercite)\*?(?:\[[^\]]*\]){0,2}\{([^}]+)\}"
    )

    bibliographies: set[Path] = set()
    thebibliography_used = False
    cite_count = 0

    for tex_file in tex_files:
        try:
            text = read_text_file(tex_file)
        except Exception:
            continue

        if thebibliography_pattern.search(text):
            thebibliography_used = True

        cite_count += len(cite_pattern.findall(text))

        for match in addbib_pattern.finditer(text):
            target = match.group(1).strip()
            resolved = resolve_bib(tex_file.parent, target)
            if resolved is not None:
                bibliographies.add(resolved)

        for match in bibliography_pattern.finditer(text):
            for target in split_csv_payload(match.group(1)):
                resolved = resolve_bib(tex_file.parent, target)
                if resolved is not None:
                    bibliographies.add(resolved)

    return sorted(bibliographies), thebibliography_used, cite_count


def unique_paths_in_order(paths: list[Path]) -> list[Path]:
    """
    路径去重并保持原有顺序（统一为 resolve 后路径）。
    """
    result: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        resolved = path.resolve()
        key = str(resolved).lower() if os.name == "nt" else str(resolved)
        if key in seen:
            continue
        seen.add(key)
        result.append(resolved)
    return result


def build_priority_resource_dirs(
    *,
    main_tex: Path,
    normalized_source_root: Path,
    tex_files: list[Path],
    bibliographies: list[Path],
) -> list[Path]:
    """
    构建 resource-path 优先目录集合。

    优先级（从高到低）：
    1. normalized_source_root；
    2. main.tex 所在目录；
    3. 所有 TeX 文件所在目录；
    4. bibliography 文件所在目录。
    """
    candidates: list[Path] = [normalized_source_root.resolve(), main_tex.parent.resolve()]
    candidates.extend(path.parent.resolve() for path in tex_files)
    candidates.extend(path.parent.resolve() for path in bibliographies)
    return unique_paths_in_order(candidates)


def apply_resource_dir_count_guard(
    *,
    all_resource_dirs: list[Path],
    priority_dirs: list[Path],
    max_resource_dirs: int,
) -> tuple[list[Path], str, list[str]]:
    """
    对 resource-path 目录数量做阈值保护，超限时按优先级裁剪。
    """
    warnings: list[str] = []
    limit = max(1, int(max_resource_dirs))
    all_unique = unique_paths_in_order(all_resource_dirs)

    if len(all_unique) <= limit:
        return all_unique, "full", warnings

    selected = unique_paths_in_order(priority_dirs)
    selected_keys = {str(path).lower() if os.name == "nt" else str(path) for path in selected}
    for path in all_unique:
        if len(selected) >= limit:
            break
        key = str(path).lower() if os.name == "nt" else str(path)
        if key in selected_keys:
            continue
        selected.append(path)
        selected_keys.add(key)

    selected = selected[:limit]
    warnings.append(
        "resource-path 目录数超限，已按优先级裁剪："
        f" original={len(all_unique)}, limit={limit}, selected={len(selected)}。"
    )
    return selected, "capped_by_priority", warnings


def collect_conversion_context(args: argparse.Namespace) -> ConversionContext:
    skill_root = locate_skill_root()
    work_root = Path(args.work_root).resolve()
    if not work_root.exists() or not work_root.is_dir():
        raise RuntimeError(f"无效的工作目录: {work_root}")

    convert_stage_dir = stage_dir(work_root, STAGE_CONVERT)
    convert_stage_dir.mkdir(parents=True, exist_ok=True)

    normalization_json = resolve_explicit_or_stage_input(
        args.normalization_json,
        work_root,
        STAGE_NORMALIZE,
        "normalization-report.json",
        legacy_filename="normalization-report.json",
    )
    if not normalization_json.exists():
        raise RuntimeError(f"未找到 normalization-report.json: {normalization_json}")

    output_docx = resolve_explicit_or_stage_output(
        args.output_docx,
        work_root,
        STAGE_CONVERT,
        "output.docx",
    )
    reference_doc = resolve_explicit_or_default(
        args.reference_doc,
        (skill_root / "templates/reference.docx").resolve(),
    )
    log_file = resolve_explicit_or_stage_output(
        args.log_file,
        work_root,
        STAGE_CONVERT,
        "pandoc-conversion.log",
    )
    report_json = resolve_explicit_or_stage_output(
        args.report_json,
        work_root,
        STAGE_CONVERT,
        "pandoc-conversion-report.json",
    )
    report_md = resolve_explicit_or_stage_output(
        args.report_md,
        work_root,
        STAGE_CONVERT,
        "pandoc-conversion-report.md",
    )
    postprocess_report_json = resolve_explicit_or_stage_output(
        None,
        work_root,
        STAGE_CONVERT,
        "docx-postprocess-report.json",
    )

    if not reference_doc.exists():
        raise RuntimeError(f"未找到 reference.docx: {reference_doc}")

    normalization_report = json.loads(read_text_file(normalization_json))
    normalization_status = str(normalization_report.get("status", ""))
    normalization_can_continue = bool(normalization_report.get("can_continue", False))

    if normalization_status == STATUS_FAIL or not normalization_can_continue:
        raise RuntimeError(
            f"规范化阶段状态不可继续：status={normalization_status}, can_continue={normalization_can_continue}"
        )

    normalized_source_root = None
    if isinstance(normalization_report.get("summary"), dict):
        normalized_source_root_value = normalization_report["summary"].get("normalized_source_root")
        if isinstance(normalized_source_root_value, str) and normalized_source_root_value.strip():
            normalized_source_root = Path(normalized_source_root_value).resolve()

    if normalized_source_root is None:
        candidate_source_root = (stage_dir(work_root, STAGE_NORMALIZE) / "source_snapshot").resolve()
        if candidate_source_root.exists():
            normalized_source_root = candidate_source_root
        else:
            normalized_source_root = work_root

    main_tex_arg = args.main_tex.strip() if isinstance(args.main_tex, str) else ""
    normalized_main_tex = normalization_report.get("normalized_main_tex")

    if main_tex_arg:
        main_tex = Path(main_tex_arg)
        if not main_tex.is_absolute():
            main_tex = (normalized_source_root / main_tex).resolve()
    else:
        if not normalized_main_tex:
            raise RuntimeError("normalization-report.json 中缺少 normalized_main_tex。")
        main_tex_candidate = Path(str(normalized_main_tex))
        if main_tex_candidate.is_absolute():
            main_tex = main_tex_candidate.resolve()
        else:
            main_tex = (normalized_source_root / main_tex_candidate).resolve()
            if not main_tex.exists():
                # 兼容历史版本 normalization-report（相对 work_root）。
                legacy_main_tex = (work_root / main_tex_candidate).resolve()
                if legacy_main_tex.exists():
                    main_tex = legacy_main_tex

    if not main_tex.exists():
        raise RuntimeError(f"规范化后的主 TeX 文件不存在：{main_tex}")

    tex_files: list[Path] = []
    processed_tex = normalization_report.get("tex_files_processed") or []
    if isinstance(processed_tex, list):
        for rel in processed_tex:
            rel_path = Path(str(rel))
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

    bibliographies, thebibliography_used, cite_count = collect_bibliographies_and_cites(tex_files)
    resource_dirs_all = sorted(
        {path.resolve() for path in normalized_source_root.rglob("*") if path.is_dir()}
        | {normalized_source_root.resolve()}
    )
    resource_dirs_priority = build_priority_resource_dirs(
        main_tex=main_tex,
        normalized_source_root=normalized_source_root,
        tex_files=tex_files,
        bibliographies=bibliographies,
    )
    resource_dirs, resource_dir_strategy, guard_warnings = apply_resource_dir_count_guard(
        all_resource_dirs=resource_dirs_all,
        priority_dirs=resource_dirs_priority,
        max_resource_dirs=args.max_resource_dirs,
    )

    if cite_count > 0 and not bibliographies and not thebibliography_used:
        raise RuntimeError("检测到文内引用，但未发现 bibliography 资源，也未发现 thebibliography 环境。")

    metadata_json = (convert_stage_dir / ".pandoc_metadata.json").resolve()
    resource_dirs_txt = (convert_stage_dir / ".pandoc_resource_dirs.txt").resolve()
    bibs_txt = (convert_stage_dir / ".pandoc_bibliographies.txt").resolve()

    return ConversionContext(
        work_root=work_root,
        normalization_json=normalization_json,
        main_tex=main_tex,
        tex_files=tex_files,
        output_docx=output_docx,
        reference_doc=reference_doc,
        log_file=log_file,
        report_json=report_json,
        report_md=report_md,
        postprocess_report_json=postprocess_report_json,
        metadata_json=metadata_json,
        resource_dirs_txt=resource_dirs_txt,
        bibs_txt=bibs_txt,
        normalization_status=normalization_status,
        normalization_can_continue=normalization_can_continue,
        tex_file_count=len(tex_files),
        bibliographies=bibliographies,
        thebibliography_used=thebibliography_used,
        cite_count=cite_count,
        resource_dirs=resource_dirs,
        resource_dirs_priority=resource_dirs_priority,
        resource_dirs_original_count=len(resource_dirs_all),
        resource_dir_limit=max(1, int(args.max_resource_dirs)),
        resource_dir_strategy=resource_dir_strategy,
        command_length_limit=max(256, int(args.max_command_length)),
        command_length_estimate=0,
        heading_numbering_mode=str(args.heading_numbering_mode),
        guard_warnings=list(guard_warnings),
        normalization_report=normalization_report,
        normalized_source_root=normalized_source_root,
        postprocess_report=None,
    )


def format_command_for_log(command: list[str]) -> str:
    if os.name == "nt":
        return subprocess.list2cmdline(command)
    return " ".join(shlex.quote(part) for part in command)


def estimate_command_length(command: list[str]) -> int:
    """
    估算命令行长度（用于阈值保护）。
    """
    return len(format_command_for_log(command))


def build_pandoc_command(context: ConversionContext, resource_dirs: list[Path]) -> list[str]:
    resource_path_value = os.pathsep.join(str(path) for path in resource_dirs) if resource_dirs else str(context.work_root)

    command: list[str] = [
        "pandoc",
        str(context.main_tex),
        "--from=latex",
        "--to=docx",
        "--standalone",
        f"--output={context.output_docx}",
        f"--reference-doc={context.reference_doc}",
        f"--resource-path={resource_path_value}",
        "--toc",
    ]
    if context.heading_numbering_mode == "pandoc":
        command.append("--number-sections")

    if context.bibliographies:
        command.extend(["--citeproc", "-M", "link-citations=true"])
        for bib in context.bibliographies:
            command.append(f"--bibliography={bib}")

    return command


def build_compact_resource_dirs(context: ConversionContext) -> list[Path]:
    """
    构建紧凑型 resource-path 目录集合，用于命令长度超限时降级。
    """
    compact_candidates: list[Path] = [
        context.normalized_source_root,
        context.main_tex.parent,
    ]
    compact_candidates.extend(path.parent for path in context.bibliographies)
    return unique_paths_in_order(compact_candidates)


def select_safe_pandoc_command(context: ConversionContext) -> list[str]:
    """
    选择满足阈值保护的 Pandoc 命令：
    1. 优先使用当前 resource_dirs；
    2. 若命令长度超限，按候选策略降级 resource-path。
    """
    max_len = max(256, int(context.command_length_limit))

    candidates: list[tuple[str, list[Path]]] = []
    seen_signatures: set[tuple[str, ...]] = set()

    def add_candidate(label: str, dirs: list[Path]) -> None:
        normalized = unique_paths_in_order(dirs)
        if not normalized:
            normalized = [context.normalized_source_root]
        signature = tuple(
            (str(path).lower() if os.name == "nt" else str(path))
            for path in normalized
        )
        if signature in seen_signatures:
            return
        seen_signatures.add(signature)
        candidates.append((label, normalized))

    add_candidate("selected", context.resource_dirs)
    add_candidate("priority", context.resource_dirs_priority)
    add_candidate("compact", build_compact_resource_dirs(context))
    add_candidate("root_only", [context.normalized_source_root])

    for index, (label, dirs) in enumerate(candidates):
        command = build_pandoc_command(context, dirs)
        cmd_len = estimate_command_length(command)
        if cmd_len <= max_len:
            context.resource_dirs = dirs
            context.command_length_estimate = cmd_len
            if index > 0:
                context.resource_dir_strategy = f"{context.resource_dir_strategy}+cmdlen:{label}"
                context.guard_warnings.append(
                    "Pandoc 命令长度超限触发降级："
                    f" selected_strategy={label}, command_length={cmd_len}, limit={max_len}, "
                    f"resource_dirs={len(dirs)}。"
                )
            return command

    # 理论上 root_only 一般已足够收敛；若仍超限则保留最后候选并给出告警。
    fallback_label, fallback_dirs = candidates[-1]
    fallback_command = build_pandoc_command(context, fallback_dirs)
    fallback_len = estimate_command_length(fallback_command)
    context.resource_dirs = fallback_dirs
    context.command_length_estimate = fallback_len
    context.resource_dir_strategy = f"{context.resource_dir_strategy}+cmdlen:{fallback_label}_overflow"
    context.guard_warnings.append(
        "Pandoc 命令长度仍高于阈值："
        f" command_length={fallback_len}, limit={max_len}；已使用最小化 resource-path 继续执行。"
    )
    return fallback_command


def write_context_debug_artifacts(context: ConversionContext, command: list[str]) -> None:
    """
    写出调试辅助文件（metadata/resource_dirs/bibliographies）。
    """
    metadata_payload = {
        "normalization_status": context.normalization_status,
        "normalization_can_continue": context.normalization_can_continue,
        "main_tex": str(context.main_tex),
        "tex_file_count": context.tex_file_count,
        "bibliography_file_count": len(context.bibliographies),
        "bibliographies": [str(path) for path in context.bibliographies],
        "thebibliography_used": context.thebibliography_used,
        "cite_count": context.cite_count,
        "resource_dir_count": len(context.resource_dirs),
        "resource_dirs_original_count": context.resource_dirs_original_count,
        "resource_dir_limit": context.resource_dir_limit,
        "resource_dir_strategy": context.resource_dir_strategy,
        "resource_dirs": [str(path) for path in context.resource_dirs],
        "command_length_limit": context.command_length_limit,
        "command_length_estimate": context.command_length_estimate,
        "heading_numbering_mode": context.heading_numbering_mode,
        "guard_warnings": context.guard_warnings,
        "normalized_source_root": str(context.normalized_source_root),
        "command_preview": format_command_for_log(command),
        "docx_postprocess_report_json": str(context.postprocess_report_json),
    }
    context.metadata_json.write_text(json.dumps(metadata_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    context.resource_dirs_txt.write_text("\n".join(str(path) for path in context.resource_dirs), encoding="utf-8")
    context.bibs_txt.write_text("\n".join(str(path) for path in context.bibliographies), encoding="utf-8")


def write_log_header(context: ConversionContext, command: list[str]) -> None:
    context.log_file.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "=== Pandoc Conversion Log ===",
        f"Skill root: {locate_skill_root()}",
        f"Work root: {context.work_root}",
        f"Main TeX: {context.main_tex}",
        f"Output DOCX: {context.output_docx}",
        f"Reference DOCX: {context.reference_doc}",
        f"Normalization JSON: {context.normalization_json}",
        f"Bibliography file count: {len(context.bibliographies)}",
        f"Resource dir count: {len(context.resource_dirs)}",
        f"Resource dir original count: {context.resource_dirs_original_count}",
        f"Resource dir strategy: {context.resource_dir_strategy}",
        f"Resource dir limit: {context.resource_dir_limit}",
        f"Command length estimate: {context.command_length_estimate}",
        f"Command length limit: {context.command_length_limit}",
        f"Heading numbering mode: {context.heading_numbering_mode}",
        f"Guard warning count: {len(context.guard_warnings)}",
        "",
        "Command:",
        f"  {format_command_for_log(command)}",
        "",
        "=== Pandoc stdout/stderr ===",
    ]
    context.log_file.write_text("\n".join(lines) + "\n", encoding="utf-8")


def run_pandoc(context: ConversionContext, command: list[str]) -> int:
    print("[INFO] 开始执行 Pandoc 主转换...")
    with context.log_file.open("a", encoding="utf-8") as log_fp:
        process = subprocess.Popen(
            command,
            cwd=str(context.normalized_source_root),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        assert process.stdout is not None
        for line in process.stdout:
            print(line, end="")
            log_fp.write(line)
        return process.wait()


def write_postprocess_report_json(context: ConversionContext, payload: dict) -> None:
    context.postprocess_report_json.parent.mkdir(parents=True, exist_ok=True)
    context.postprocess_report_json.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def run_docx_postprocess_step(context: ConversionContext) -> None:
    """
    在 Pandoc 成功后执行 docx 结构后处理：
    - 图表题注补 SEQ；
    - 公式编号补 SEQ；
    - 尝试修复缺失书签以改善内部跳转。
    """
    if not context.output_docx.exists() or context.output_docx.stat().st_size == 0:
        payload = {
            "status": "SKIPPED",
            "reason": "output_docx_missing_or_empty",
            "output_docx": str(context.output_docx),
        }
        context.postprocess_report = payload
        write_postprocess_report_json(context, payload)
        return

    try:
        result = run_docx_postprocess(
            docx_path=context.output_docx,
            tex_files=context.tex_files,
        )
        payload = {
            "status": "PASS",
            **result.to_dict(),
        }
        context.postprocess_report = payload
        write_postprocess_report_json(context, payload)

        if payload.get("warnings"):
            for warning in payload["warnings"]:
                context.guard_warnings.append(f"DOCX 后处理：{warning}")

        modified = bool(payload.get("modified", False))
        metrics = payload.get("metrics", {}) if isinstance(payload.get("metrics"), dict) else {}
        print(f"[INFO] DOCX 后处理完成：modified={modified}")
        if metrics:
            print(
                "[INFO] DOCX 后处理统计："
                f" figure_seq={metrics.get('figure_caption_seq_added', 0)},"
                f" table_seq={metrics.get('table_caption_seq_added', 0)},"
                f" equation_seq={metrics.get('equation_seq_added', 0)},"
                f" bookmark_added={metrics.get('bookmark_added_total', 0)},"
                f" missing_anchor_after={metrics.get('missing_anchor_count_after', 0)}"
            )
    except Exception as exc:
        message = f"DOCX 后处理失败（已降级为告警）：{exc}"
        context.guard_warnings.append(message)
        payload = {
            "status": "PASS_WITH_WARNINGS",
            "modified": False,
            "warnings": [message],
            "metrics": {},
            "details": {},
        }
        context.postprocess_report = payload
        write_postprocess_report_json(context, payload)


def determine_status(context: ConversionContext, pandoc_exit_code: int) -> tuple[str, bool, str, list[str]]:
    status = STATUS_PASS
    can_continue = True
    failure_reason = ""
    warnings: list[str] = list(context.guard_warnings)

    if pandoc_exit_code != 0:
        status = STATUS_FAIL
        can_continue = False
        failure_reason = "pandoc_exit_nonzero"

    if not context.output_docx.exists():
        status = STATUS_FAIL
        can_continue = False
        failure_reason = "output_docx_missing"
    elif context.output_docx.stat().st_size == 0:
        status = STATUS_FAIL
        can_continue = False
        failure_reason = "output_docx_empty"

    if (
        status != STATUS_FAIL
        and context.cite_count > 0
        and len(context.bibliographies) == 0
        and context.thebibliography_used
    ):
        status = STATUS_PASS_WITH_WARNINGS
        warnings.append("检测到 thebibliography 环境；Pandoc 未启用 .bib 驱动的 citeproc，参考文献与链接效果需在后检查阶段重点核对。")

    if status != STATUS_FAIL and len(context.bibliographies) == 0 and context.cite_count == 0:
        warnings.append("未检测到 .bib bibliography 文件；若文档确实不使用引用系统，这不是问题。")
        if status == STATUS_PASS:
            status = STATUS_PASS_WITH_WARNINGS

    if status != STATUS_FAIL and warnings and status == STATUS_PASS:
        status = STATUS_PASS_WITH_WARNINGS

    return status, can_continue, failure_reason, warnings


def write_conversion_reports(
    context: ConversionContext,
    pandoc_exit_code: int,
    status: str,
    can_continue: bool,
    failure_reason: str,
    warnings: list[str],
) -> None:
    report = {
        "status": status,
        "can_continue": can_continue,
        "work_root": str(context.work_root),
        "main_tex": str(context.main_tex),
        "output_docx": str(context.output_docx),
        "reference_doc": str(context.reference_doc),
        "normalization_json": str(context.normalization_json),
        "pandoc_log": str(context.log_file),
        "docx_postprocess_report_json": str(context.postprocess_report_json),
        "pandoc_exit_code": pandoc_exit_code,
        "failure_reason": failure_reason if failure_reason else None,
        "metrics": {
            "bibliography_file_count": len(context.bibliographies),
            "resource_dir_count": len(context.resource_dirs),
            "resource_dirs_original_count": context.resource_dirs_original_count,
            "resource_dir_limit": context.resource_dir_limit,
            "tex_file_count": context.tex_file_count,
            "cite_count": context.cite_count,
            "postprocess_warning_count": len(context.postprocess_report.get("warnings", []))
            if isinstance(context.postprocess_report, dict)
            else 0,
        },
        "inventory": {
            "bibliographies": [str(path) for path in context.bibliographies],
            "resource_dirs": [str(path) for path in context.resource_dirs],
            "resource_dir_strategy": context.resource_dir_strategy,
            "thebibliography_used": context.thebibliography_used,
        },
        "summary": {
            "normalization_status": context.normalization_report.get("status"),
            "normalization_can_continue": context.normalization_report.get("can_continue"),
            "pandoc_exit_code": pandoc_exit_code,
            "output_exists": context.output_docx.exists(),
            "output_size_bytes": context.output_docx.stat().st_size if context.output_docx.exists() else 0,
            "command_length_estimate": context.command_length_estimate,
            "command_length_limit": context.command_length_limit,
            "heading_numbering_mode": context.heading_numbering_mode,
            "docx_postprocess_modified": (
                bool(context.postprocess_report.get("modified", False))
                if isinstance(context.postprocess_report, dict)
                else False
            ),
        },
        "warnings": warnings,
        "recommendations": [],
    }

    if status == STATUS_FAIL:
        report["recommendations"].append("先查看 pandoc-conversion.log 中的错误信息，再修复阻塞问题后重试。")
    else:
        report["recommendations"].append("可进入 postcheck_docx.py 阶段，对生成的 docx 做结构级后检查。")

    if len(context.bibliographies) > 0:
        report["recommendations"].append("参考文献已作为 Pandoc 主转换输入；仍应在后检查阶段核对引用跳转与样式。")

    if warnings:
        report["recommendations"].append("存在转换告警；请在后检查和人工修复阶段重点核对这些对象。")
    if context.heading_numbering_mode == "template":
        report["recommendations"].append(
            "当前采用 reference.docx 管理章节编号；请确认模板中的 Heading 样式已配置多级编号。"
        )

    context.report_json.parent.mkdir(parents=True, exist_ok=True)
    context.report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = [
        "# Pandoc Conversion Report",
        "",
        f"- Status: **{report['status']}**",
        f"- Can continue: **{report['can_continue']}**",
        f"- Work root: `{report['work_root']}`",
        f"- Main TeX: `{report['main_tex']}`",
        f"- Output DOCX: `{report['output_docx']}`",
        f"- Reference DOCX: `{report['reference_doc']}`",
        f"- Pandoc log: `{report['pandoc_log']}`",
        f"- DOCX postprocess report: `{report['docx_postprocess_report_json']}`",
        "",
        "## Summary",
        "",
    ]
    for key, value in report["summary"].items():
        lines.append(f"- {key}: **{value}**")

    lines.extend(["", "## Metrics", ""])
    for key, value in report["metrics"].items():
        lines.append(f"- {key}: **{value}**")

    lines.extend(["", "## Warnings", ""])
    if report["warnings"]:
        for warning in report["warnings"]:
            lines.append(f"- {warning}")
    else:
        lines.append("- (none)")

    lines.extend(["", "## Recommendations", ""])
    if report["recommendations"]:
        for recommendation in report["recommendations"]:
            lines.append(f"- {recommendation}")
    else:
        lines.append("- (none)")
    lines.append("")

    context.report_md.parent.mkdir(parents=True, exist_ok=True)
    context.report_md.write_text("\n".join(lines), encoding="utf-8")


def run() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()

    if not shutil_which("pandoc"):
        print("[ERROR] 未找到 pandoc，请先安装并确保其已加入 PATH。", file=sys.stderr)
        return 1

    try:
        context = collect_conversion_context(args)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1

    command = select_safe_pandoc_command(context)
    write_context_debug_artifacts(context, command)
    write_log_header(context, command)
    pandoc_exit_code = run_pandoc(context, command)
    if pandoc_exit_code == 0:
        run_docx_postprocess_step(context)
    status, can_continue, failure_reason, warnings = determine_status(context, pandoc_exit_code)
    write_conversion_reports(
        context=context,
        pandoc_exit_code=pandoc_exit_code,
        status=status,
        can_continue=can_continue,
        failure_reason=failure_reason,
        warnings=warnings,
    )
    best_effort_update_manifest(
        context.work_root,
        stage=STAGE_CONVERT,
        status=status,
        can_continue=can_continue,
        artifacts={
            "output_docx": context.output_docx,
            "pandoc_log": context.log_file,
            "conversion_report_json": context.report_json,
            "conversion_report_md": context.report_md,
            "pandoc_metadata_json": context.metadata_json,
            "pandoc_resource_dirs_txt": context.resource_dirs_txt,
            "pandoc_bibliographies_txt": context.bibs_txt,
            "docx_postprocess_report_json": context.postprocess_report_json,
            "normalized_source_root": context.normalized_source_root,
        },
        summary={
            "pandoc_exit_code": pandoc_exit_code,
            "warning_count": len(warnings),
            "resource_dir_strategy": context.resource_dir_strategy,
            "guard_warning_count": len(context.guard_warnings),
            "heading_numbering_mode": context.heading_numbering_mode,
        },
        metrics={
            "tex_file_count": context.tex_file_count,
            "bibliography_file_count": len(context.bibliographies),
            "resource_dir_count": len(context.resource_dirs),
            "resource_dirs_original_count": context.resource_dirs_original_count,
            "resource_dir_limit": context.resource_dir_limit,
            "command_length_estimate": context.command_length_estimate,
            "command_length_limit": context.command_length_limit,
            "cite_count": context.cite_count,
            "docx_postprocess_missing_anchor_after": (
                int(context.postprocess_report.get("metrics", {}).get("missing_anchor_count_after", 0))
                if isinstance(context.postprocess_report, dict)
                else 0
            ),
        },
        notes=warnings,
        top_level_artifacts={
            "deliverables": {
                "output_docx": context.output_docx,
            },
            "reports": {
                "pandoc_conversion_report_json": context.report_json,
                "pandoc_conversion_report_md": context.report_md,
                "docx_postprocess_report_json": context.postprocess_report_json,
            },
            "logs": {
                "pandoc_conversion_log": context.log_file,
            },
            "debug": {
                "pandoc_metadata_json": context.metadata_json,
                "pandoc_resource_dirs_txt": context.resource_dirs_txt,
                "pandoc_bibliographies_txt": context.bibs_txt,
            },
        },
    )

    print("[INFO] Pandoc 主转换结束。")
    print(f"[INFO] Status: {status}")
    print(f"[INFO] Main TeX: {context.main_tex}")
    print(f"[INFO] Output DOCX: {context.output_docx}")
    print(f"[INFO] Log file: {context.log_file}")
    print(f"[INFO] JSON report: {context.report_json}")
    print(f"[INFO] Markdown report: {context.report_md}")

    if status == STATUS_FAIL:
        print(f"[ERROR] Pandoc 主转换失败。请先检查日志：{context.log_file}", file=sys.stderr)
        return 1
    return 0


def shutil_which(binary_name: str) -> Optional[str]:
    for path in os.environ.get("PATH", "").split(os.pathsep):
        if not path:
            continue
        candidate = Path(path) / binary_name
        if candidate.exists():
            return str(candidate)
        if os.name == "nt":
            exe_candidate = candidate.with_suffix(".exe")
            if exe_candidate.exists():
                return str(exe_candidate)
    return None


if __name__ == "__main__":
    sys.exit(run())
