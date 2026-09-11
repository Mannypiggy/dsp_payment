"""测试：返点 / 付款金额口径（规格 78 节 47-49 + Decimal + billing_base_cost）。"""
from decimal import Decimal

import pytest

from conftest import build_july_ps
from processors import email_processor
from utils import money


def test_billing_base_cost_is_q2_dsp_cost():
    """billing_base_cost = q2(dsp_cost)（本轮新增）。"""
    ps = build_july_ps()
    assert ps.billing_base_cost == money.q2(ps.dsp_cost)
    assert ps.billing_base_cost == Decimal("49025.52")


def test_rebate_raw_full_precision():
    ps = build_july_ps()
    assert ps.rebate_raw == Decimal("3922.0416")


def test_rebate_amount_is_q2_rebate_raw():
    ps = build_july_ps()
    assert ps.rebate_amount == money.q2(ps.rebate_raw)
    assert ps.rebate_amount == Decimal("3922.04")


def test_payment_raw_full_precision():
    ps = build_july_ps()
    assert ps.payment_raw == Decimal("45103.4784")


def test_payment_amount_is_q2_payment_raw():
    ps = build_july_ps()
    assert ps.payment_amount == money.q2(ps.payment_raw)
    assert ps.payment_amount == Decimal("45103.48")


def test_rebate_calculation():
    ps = build_july_ps()
    assert float(ps.rebate_amount) == pytest.approx(3922.04, abs=0.01)


def test_payment_calculation():
    ps = build_july_ps()
    assert float(ps.payment_amount) == pytest.approx(45103.48, abs=0.01)


def test_email_payment_matches_summary():
    ps = build_july_ps()
    email = email_processor.generate_email(ps)
    assert "$45,103.48 USD" in email["text"]


def test_money_fields_are_decimal():
    """核心财务金额是 Decimal（158）。"""
    ps = build_july_ps()
    for v in (ps.dsp_cost, ps.billing_base_cost, ps.rebate_raw, ps.rebate_amount,
              ps.payment_raw, ps.payment_amount, ps.onsite_cost, ps.onsite_sales,
              ps.combined_cost, ps.combined_sales):
        assert isinstance(v, Decimal)


def test_rebate_round_half_up_boundary():
    """第三位小数为 5 时 ROUND_HALF_UP（161）。"""
    assert money.q2(Decimal("1.005")) == Decimal("1.01")
    assert money.q2(Decimal("1.004")) == Decimal("1.00")
