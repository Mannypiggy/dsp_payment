"""可选 LLM 客户端（用于 KPI 原因 AI 总结模式）。

API Key 从 .env 读取，不写 config.yaml、不提交仓库。
无 API Key 时返回 None，调用方回退到规则总结模式。
"""
from __future__ import annotations

import os
from typing import Optional

import requests

_ENV_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")


def load_env(path: str = None) -> dict:
    env = {}
    p = path or _ENV_PATH
    if not os.path.exists(p):
        return env
    with open(p, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


class LLMClient:
    """极简 Anthropic Messages API 客户端。"""

    def __init__(self, api_key: str, model: str = "claude-sonnet-5",
                 base_url: str = "https://api.anthropic.com"):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url.rstrip("/")

    def generate(self, system: str, user: str) -> str:
        resp = requests.post(
            f"{self.base_url}/v1/messages",
            headers={
                "x-api-key": self.api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": self.model,
                "max_tokens": 1024,
                "system": system,
                "messages": [{"role": "user", "content": user}],
            },
            timeout=60,
        )
        resp.raise_for_status()
        data = resp.json()
        parts = data.get("content", [])
        return "".join(p.get("text", "") for p in parts if p.get("type") == "text").strip()


def get_llm_client() -> Optional[LLMClient]:
    """从 .env 读取 API Key，返回 LLMClient 或 None（回退规则模式）。"""
    env = load_env()
    key = env.get("ANTHROPIC_API_KEY") or env.get("DSP_LLM_API_KEY")
    if not key:
        return None
    model = env.get("DSP_LLM_MODEL", "claude-sonnet-5")
    base_url = env.get("DSP_LLM_BASE_URL", "https://api.anthropic.com")
    return LLMClient(key, model, base_url)
