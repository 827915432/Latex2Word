#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pipeline_layout.py

流水线阶段目录与 run-manifest 管理工具。

设计目标：
1. 统一 stage_* 目录命名与路径解析；
2. 统一 run-manifest.json 的创建、更新与持久化；
3. 在不破坏旧路径兼容的前提下，支持脚本逐步迁移到阶段化输出。
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


STAGE_PRECHECK = "precheck"
STAGE_NORMALIZE = "normalize"
STAGE_CONVERT = "convert"
STAGE_POSTCHECK = "postcheck"
STAGE_CHECKLIST = "checklist"

STAGE_TO_DIRNAME = {
    STAGE_PRECHECK: "stage_precheck",
    STAGE_NORMALIZE: "stage_normalize",
    STAGE_CONVERT: "stage_convert",
    STAGE_POSTCHECK: "stage_postcheck",
    STAGE_CHECKLIST: "stage_checklist",
}

MANIFEST_FILENAME = "run-manifest.json"
MANIFEST_SCHEMA_VERSION = "1.0"


def now_iso_utc() -> str:
    """
    生成 UTC 时间戳（ISO 8601）。
    """
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def default_work_root_for_project(project_root: Path) -> Path:
    """
    根据工程目录生成默认工作目录。
    """
    return (project_root.parent / f"{project_root.name}__latex_to_word_work").resolve()


def stage_dir(work_root: Path, stage: str) -> Path:
    """
    获取指定阶段目录路径。
    """
    if stage not in STAGE_TO_DIRNAME:
        raise ValueError(f"unknown stage: {stage}")
    return (work_root / STAGE_TO_DIRNAME[stage]).resolve()


def ensure_layout_dirs(work_root: Path) -> dict[str, Path]:
    """
    确保阶段目录与用户视图目录存在。
    """
    work_root = work_root.resolve()
    work_root.mkdir(parents=True, exist_ok=True)

    dirs = {
        "work_root": work_root,
        STAGE_PRECHECK: stage_dir(work_root, STAGE_PRECHECK),
        STAGE_NORMALIZE: stage_dir(work_root, STAGE_NORMALIZE),
        STAGE_CONVERT: stage_dir(work_root, STAGE_CONVERT),
        STAGE_POSTCHECK: stage_dir(work_root, STAGE_POSTCHECK),
        STAGE_CHECKLIST: stage_dir(work_root, STAGE_CHECKLIST),
        "deliverables": (work_root / "deliverables").resolve(),
        "reports": (work_root / "reports").resolve(),
        "logs": (work_root / "logs").resolve(),
        "debug": (work_root / "debug").resolve(),
    }

    for path in dirs.values():
        path.mkdir(parents=True, exist_ok=True)

    return dirs


def manifest_path(work_root: Path) -> Path:
    """
    run-manifest.json 路径。
    """
    return (work_root.resolve() / MANIFEST_FILENAME).resolve()


def _init_manifest_payload(work_root: Path) -> dict:
    """
    初始化 manifest 结构。
    """
    work_root = work_root.resolve()
    ts = now_iso_utc()
    layout = {
        "stage_dirs": {
            stage: str(stage_dir(work_root, stage))
            for stage in [
                STAGE_PRECHECK,
                STAGE_NORMALIZE,
                STAGE_CONVERT,
                STAGE_POSTCHECK,
                STAGE_CHECKLIST,
            ]
        },
        "deliverables_dir": str((work_root / "deliverables").resolve()),
        "reports_dir": str((work_root / "reports").resolve()),
        "logs_dir": str((work_root / "logs").resolve()),
        "debug_dir": str((work_root / "debug").resolve()),
    }
    return {
        "schema_version": MANIFEST_SCHEMA_VERSION,
        "work_root": str(work_root),
        "run_id": ts.replace(":", "").replace("-", ""),
        "created_at": ts,
        "updated_at": ts,
        "layout": layout,
        "stages": {},
        "artifacts": {
            "deliverables": {},
            "reports": {},
            "logs": {},
            "debug": {},
        },
    }


def load_manifest(work_root: Path) -> dict:
    """
    读取 manifest；若不存在则返回初始化结构（不落盘）。
    """
    path = manifest_path(work_root)
    if not path.exists():
        return _init_manifest_payload(work_root)

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        payload = _init_manifest_payload(work_root)

    if not isinstance(payload, dict):
        payload = _init_manifest_payload(work_root)

    # 基础字段兜底，避免历史版本缺字段时出错。
    payload.setdefault("schema_version", MANIFEST_SCHEMA_VERSION)
    payload.setdefault("work_root", str(work_root.resolve()))
    payload.setdefault("run_id", now_iso_utc().replace(":", "").replace("-", ""))
    payload.setdefault("created_at", now_iso_utc())
    payload.setdefault("updated_at", now_iso_utc())
    payload.setdefault("layout", _init_manifest_payload(work_root)["layout"])
    payload.setdefault("stages", {})
    payload.setdefault(
        "artifacts",
        {
            "deliverables": {},
            "reports": {},
            "logs": {},
            "debug": {},
        },
    )
    return payload


def save_manifest(work_root: Path, payload: dict) -> Path:
    """
    写出 manifest。
    """
    path = manifest_path(work_root)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload["updated_at"] = now_iso_utc()
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _normalize_artifact_values(artifacts: dict) -> dict:
    """
    将 artifact 值规范化为字符串路径或基础类型。
    """
    normalized: dict = {}
    for key, value in artifacts.items():
        if isinstance(value, Path):
            normalized[key] = str(value.resolve())
        elif isinstance(value, list):
            tmp: list = []
            for item in value:
                if isinstance(item, Path):
                    tmp.append(str(item.resolve()))
                else:
                    tmp.append(item)
            normalized[key] = tmp
        else:
            normalized[key] = value
    return normalized


def update_stage_manifest(
    work_root: Path,
    stage: str,
    *,
    status: Optional[str] = None,
    can_continue: Optional[bool] = None,
    artifacts: Optional[dict] = None,
    metrics: Optional[dict] = None,
    summary: Optional[dict] = None,
    notes: Optional[list[str]] = None,
) -> Path:
    """
    更新某个阶段的 manifest 条目。
    """
    payload = load_manifest(work_root)
    stage_entry = payload.setdefault("stages", {}).get(stage, {})
    if not isinstance(stage_entry, dict):
        stage_entry = {}

    stage_entry["stage_dir"] = str(stage_dir(work_root.resolve(), stage))
    if status is not None:
        stage_entry["status"] = status
    if can_continue is not None:
        stage_entry["can_continue"] = bool(can_continue)
    if artifacts is not None:
        stage_entry["artifacts"] = _normalize_artifact_values(artifacts)
    if metrics is not None:
        stage_entry["metrics"] = metrics
    if summary is not None:
        stage_entry["summary"] = summary
    if notes is not None:
        stage_entry["notes"] = notes

    stage_entry["updated_at"] = now_iso_utc()
    payload["stages"][stage] = stage_entry
    return save_manifest(work_root, payload)


def update_manifest_artifacts(work_root: Path, section: str, artifacts: dict) -> Path:
    """
    更新 manifest 顶层 artifacts。
    """
    payload = load_manifest(work_root)
    payload.setdefault("artifacts", {})
    if section not in payload["artifacts"] or not isinstance(payload["artifacts"][section], dict):
        payload["artifacts"][section] = {}
    payload["artifacts"][section].update(_normalize_artifact_values(artifacts))
    return save_manifest(work_root, payload)


def stage_default_or_legacy(
    work_root: Path,
    stage: str,
    filename: str,
    *,
    legacy_filename: Optional[str] = None,
) -> Path:
    """
    读取型路径解析：
    - 优先使用阶段目录中的文件；
    - 若不存在且 legacy 存在，则回退 legacy；
    - 若都不存在，返回阶段目录目标路径（供后续写入或报错展示）。
    """
    staged = (stage_dir(work_root, stage) / filename).resolve()
    if staged.exists():
        return staged

    if legacy_filename:
        legacy = (work_root.resolve() / legacy_filename).resolve()
        if legacy.exists():
            return legacy

    return staged


def resolve_explicit_or_default(explicit_path: Optional[str], default_path: Path) -> Path:
    """
    解析“显式参数优先，否则默认路径”。
    """
    if explicit_path:
        return Path(explicit_path).resolve()
    return default_path.resolve()


def resolve_explicit_or_stage_output(
    explicit_path: Optional[str],
    work_root: Path,
    stage: str,
    filename: str,
) -> Path:
    """
    输出型路径解析：
    - 显式参数优先；
    - 否则返回 <work-root>/<stage_dir>/<filename>。
    """
    default_path = (stage_dir(work_root, stage) / filename).resolve()
    return resolve_explicit_or_default(explicit_path, default_path)


def resolve_explicit_or_stage_input(
    explicit_path: Optional[str],
    work_root: Path,
    stage: str,
    filename: str,
    *,
    legacy_filename: Optional[str] = None,
) -> Path:
    """
    读取型路径解析：
    - 显式参数优先；
    - 否则按 stage_default_or_legacy 规则回退。
    """
    if explicit_path:
        return Path(explicit_path).resolve()
    return stage_default_or_legacy(
        work_root,
        stage,
        filename,
        legacy_filename=legacy_filename,
    )


def best_effort_update_manifest(
    work_root: Path,
    *,
    stage: str,
    status: Optional[str] = None,
    can_continue: Optional[bool] = None,
    artifacts: Optional[dict] = None,
    metrics: Optional[dict] = None,
    summary: Optional[dict] = None,
    notes: Optional[list[str]] = None,
    top_level_artifacts: Optional[dict[str, dict]] = None,
) -> None:
    """
    以“尽力而为”方式更新 manifest：
    - 更新阶段条目；
    - 再按 section 更新顶层 artifacts；
    - 任一步失败都不抛出异常，避免阻塞主流程。
    """
    try:
        update_stage_manifest(
            work_root,
            stage,
            status=status,
            can_continue=can_continue,
            artifacts=artifacts,
            metrics=metrics,
            summary=summary,
            notes=notes,
        )
        if top_level_artifacts:
            for section, section_artifacts in top_level_artifacts.items():
                if section_artifacts is None:
                    continue
                update_manifest_artifacts(work_root, section, section_artifacts)
    except Exception:
        return
