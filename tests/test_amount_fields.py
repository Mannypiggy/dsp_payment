"""测试：金额口径字段级绑定（规格 12 节 114-118）。"""
import pytest

from conftest import build_july_ps
from rules import amount_fields


def test_dsp_cost_and_payment_distinct():
    ps = build_july_ps()
    assert float(ps.dsp_cost) == pytest.approx(49025.52, abs=0.01)
    assert float(ps.payment_amount) == pytest.approx(45103.48, abs=0.01)
    assert abs(float(ps.dsp_cost) - float(ps.payment_amount)) > 1000  # 明确不同


def test_email_dsp_cost_uses_dsp_cost():
    ps = build_july_ps()
    v = amount_fields.resolve_amount(ps, "email.dsp_cost")
    assert float(v) == pytest.approx(float(ps.dsp_cost), rel=1e-9)


def test_email_payment_uses_payment_amount():
    ps = build_july_ps()
    v = amount_fields.resolve_amount(ps, "email.payment_amount")
    assert float(v) == pytest.approx(float(ps.payment_amount), rel=1e-9)


def test_spend_detail_uses_payment_amount():
    """《DSP花费明细》申请支付金额 = 对账单结算金额（payment_amount），不是 dsp_cost。"""
    ps = build_july_ps()
    v = amount_fields.resolve_amount(ps, "spend_detail.apply_amount")
    assert float(v) == pytest.approx(float(ps.payment_amount), rel=1e-9)
    assert float(v) == pytest.approx(45103.48, abs=0.01)


def test_spend_ratio_uses_payment_amount():
    ps = build_july_ps()
    v = amount_fields.resolve_amount(ps, "spend_ratio.apply_amount")
    assert float(v) == pytest.approx(float(ps.payment_amount), rel=1e-9)


def test_mapping_keys_exist():
    for k in ("email.dsp_cost", "email.rebate_amount", "email.payment_amount",
              "spend_detail.apply_amount", "spend_ratio.apply_amount"):
        assert k in amount_fields.AMOUNT_FIELD_MAPPING


def test_no_global_apply_basis():
    ps = build_july_ps()
    assert not hasattr(ps, "apply_amount")
    assert not hasattr(ps, "apply_amount_basis")
