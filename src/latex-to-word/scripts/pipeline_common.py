#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pipeline_common.py

跨脚本复用的通用工具：
- 文本/JSON 读写
- 常见路径处理
- 轻量字符串拆分
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Optional


def locate_skill_root() -> Path:
    """
    定位 skill 根目录。

    当前目录约定：
    <skill_root>/scripts/pipeline_common.py
    """
    return Path(__file__).resolve().parents[1]


def safe_relative(path: Path, root: Path) -> str:
    """
    将路径尽量转为相对路径；失败时返回绝对路径字符串。
    """
    try:
        return str(path.resolve().relative_to(root.resolve()))
    except Exception:
        return str(path.resolve())


def read_text_file(path: Path) -> str:
    """
    以 UTF-8 优先、多编码回退方式读取文本文件。
    """
    encodings = ("utf-8", "utf-8-sig", "gbk", "cp936", "latin-1")
    last_error: Optional[Exception] = None
    for encoding in encodings:
        try:
            return path.read_text(encoding=encoding)
        except Exception as exc:  # pragma: no cover - 容错路径
            last_error = exc
    raise RuntimeError(f"无法读取文本文件: {path}") from last_error


def write_text_file(path: Path, text: str, *, newline: str = "\n") -> None:
    """
    写出 UTF-8 文本文件，可指定换行符风格。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline=newline) as f:
        f.write(text)


def write_json(path: Path, payload: dict) -> None:
    """
    写出 UTF-8 JSON（缩进 2，保留非 ASCII 字符）。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_markdown(path: Path, content: str) -> None:
    """
    写出 UTF-8 Markdown。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def load_json_if_exists(path: Path) -> Optional[dict]:
    """
    若 JSON 文件存在则读取，不存在返回 None。
    """
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def split_csv_payload(payload: str) -> list[str]:
    """
    解析逗号分隔字符串，返回去空白后的非空条目。
    """
    return [item.strip() for item in payload.split(",") if item.strip()]
