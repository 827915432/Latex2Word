#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stage_reporting.py

统一“阶段报告落盘 + manifest 更新”的模板逻辑。
"""

from __future__ import annotations

from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any, Optional

from pipeline_common import write_json, write_markdown
from pipeline_layout import best_effort_update_manifest


def _to_payload(report_obj: Any) -> dict:
    if isinstance(report_obj, dict):
        return report_obj
    if is_dataclass(report_obj):
        return asdict(report_obj)
    raise TypeError("report_obj 必须是 dict 或 dataclass 实例。")


def persist_stage_report(
    *,
    work_root: Path,
    stage: str,
    report_obj: Any,
    markdown_text: str,
    report_json_path: Path,
    report_md_path: Path,
    status: str,
    can_continue: bool,
    artifacts: Optional[dict] = None,
    summary: Optional[dict] = None,
    metrics: Optional[dict] = None,
    notes: Optional[list[str]] = None,
    top_level_artifacts: Optional[dict[str, dict]] = None,
) -> dict:
    """
    写出阶段报告并以 best-effort 方式更新 manifest。
    """
    payload = _to_payload(report_obj)
    write_json(report_json_path, payload)
    write_markdown(report_md_path, markdown_text)

    best_effort_update_manifest(
        work_root,
        stage=stage,
        status=status,
        can_continue=can_continue,
        artifacts=artifacts or {},
        summary=summary or {},
        metrics=metrics or {},
        notes=notes,
        top_level_artifacts=top_level_artifacts,
    )
    return payload
