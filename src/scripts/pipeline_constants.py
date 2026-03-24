#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pipeline_constants.py

跨阶段共享的常量定义，避免在各脚本中重复维护。
"""

from __future__ import annotations


STATUS_PASS = "PASS"
STATUS_PASS_WITH_WARNINGS = "PASS_WITH_WARNINGS"
STATUS_FAIL = "FAIL"

SEVERITY_INFO = "INFO"
SEVERITY_WARN = "WARN"
SEVERITY_ERROR = "ERROR"

SEVERITY_ORDER = [SEVERITY_ERROR, SEVERITY_WARN, SEVERITY_INFO]

REQUIRED_RULE_FILES = [
    "rules/supported_envs.md",
    "rules/downgrade_policy.md",
    "rules/acceptance_criteria.md",
]
