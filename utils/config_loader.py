"""加载配置。优先级：config.local.yaml → config.yaml → config.example.yaml。

真实敏感配置放 config.local.yaml（已 .gitignore）；仓库只保留脱敏的
config.example.yaml。config.yaml 兼容旧用法，也被 .gitignore 排除。
"""
from __future__ import annotations

import os
from typing import Any, Dict

import yaml

_CONFIG_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CANDIDATES = ["config.local.yaml", "config.yaml", "config.example.yaml"]


def load_config(path: str = None) -> Dict[str, Any]:
    if path is not None:
        with open(path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    for name in _CANDIDATES:
        p = os.path.join(_CONFIG_DIR, name)
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return yaml.safe_load(f) or {}
    raise FileNotFoundError("未找到配置文件（config.local.yaml / config.yaml / config.example.yaml）")


def default_config() -> Dict[str, Any]:
    return load_config()


# 占位符/示例标记：命中则视为配置未完成
_PLACEHOLDER_MARKERS = ["YOUR_", "PLACEHOLDER", "example.com", "example.org", "example.net"]


def validate_business_config(config: Dict[str, Any]) -> list:
    """校验业务配置是否已填写真实付款信息。

    返回问题列表（空 = 就绪）。若存在占位符/示例/空值则不可生成正式版。
    """
    problems = []
    payee = config.get("payment", {}).get("payee", {}) or {}
    payer_name = config.get("payment", {}).get("payer", {}).get("name", "") if isinstance(config.get("payment", {}).get("payer"), dict) else ""
    email_to = config.get("email", {}).get("to", []) or []

    checks = [
        ("payer", payer_name),
        ("payee.name", payee.get("name", "")),
        ("payee.bank", payee.get("bank", "")),
        ("payee.account", payee.get("account", "")),
        ("payee.bank_address", payee.get("bank_address", "")),
        ("payee.swift", payee.get("swift", "")),
        ("payee.bank_code", payee.get("bank_code", "")),
    ]
    for field, value in checks:
        s = str(value).strip()
        if not s:
            problems.append(f"{field} 为空")
        elif any(m in s for m in _PLACEHOLDER_MARKERS):
            problems.append(f"{field} 含占位符/示例值")

    if not email_to:
        problems.append("email.to 为空")
    else:
        for addr in email_to:
            a = str(addr).strip()
            if not a or any(m in a for m in _PLACEHOLDER_MARKERS):
                problems.append(f"email.to 含占位邮箱：{addr!r}")

    return problems


def config_is_ready(config: Dict[str, Any]) -> bool:
    return not validate_business_config(config)
