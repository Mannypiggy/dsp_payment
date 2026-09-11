"""测试：业务配置校验（示例/占位值阻断正式生成）。"""
import copy
import os

import pytest

from pipeline import ConfigNotReadyError, run_pipeline
from utils.config_loader import config_is_ready, load_config, validate_business_config


def _example_config():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return load_config(os.path.join(root, "config.example.yaml"))


def test_example_config_has_problems():
    problems = validate_business_config(_example_config())
    assert problems
    assert not config_is_ready(_example_config())


def test_example_config_blocks_pipeline(tmp_path):
    config = _example_config()
    inputs = {
        "month_str": "2026-07", "output_dir": str(tmp_path), "rebate_rate": 0.08,
        "month_confirmed": False, "payment_due_date": "2026-08-20",
        "kpi_targets": config["kpi"],
        "order_summary_path": None,
    }
    with pytest.raises(ConfigNotReadyError):
        run_pipeline(inputs, config)


def test_local_config_is_ready():
    config = load_config()  # config.local.yaml（本机真实配置）
    assert config_is_ready(config)


def test_example_com_detected():
    config = _example_config()
    assert any("example" in p for p in validate_business_config(config))


def test_your_placeholder_detected():
    config = _example_config()
    assert any("YOUR_" in p or "占位" in p for p in validate_business_config(config))


def test_empty_bank_account_detected():
    config = copy.deepcopy(_example_config())
    config["payment"]["payee"]["account"] = ""
    problems = validate_business_config(config)
    assert any("account" in p for p in problems)


def test_real_local_can_generate():
    """真实 local 配置无问题，可进入流程（不抛 ConfigNotReadyError）。"""
    config = load_config()
    assert config_is_ready(config)
    # 不实际跑完整流程（需要真实文件），仅验证 config_is_ready
