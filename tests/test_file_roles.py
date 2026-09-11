"""测试：文件角色（数据源来自原始报表，不来自对账单）（规格 172-181）。"""
import os

import openpyxl
import pytest

from conftest import JULY_BASE, build_july_ps, july_reconciliation, july_statement
from processors import (order_summary_processor, promoted_ads_processor,
                        reconciliation_processor, statement_processor)


# ---- 对账单不要求含 DSP数据源 / 已推广数据源 Sheet ----
def test_statement_has_no_dsp_src_sheet(july_statement):
    """付款对账单不含 DSP数据源 Sheet（172 前提）。"""
    wb = openpyxl.load_workbook(july_statement, read_only=True)
    assert not any("数据源" in sn for sn in wb.sheetnames)


def test_statement_read_without_dsp_src(july_statement):
    """对账单无 DSP数据源 Sheet 仍能读取付款金额（172/173）。"""
    stmt = statement_processor.read_statement(july_statement)
    assert stmt is not None
    assert stmt.dsp_cost == pytest.approx(49025.52, abs=0.01)
    assert stmt.rebate == pytest.approx(3922.0416, abs=0.01)
    assert stmt.payment == pytest.approx(45103.4784, abs=0.01)


def test_no_missing_sheet_error_for_statement(july_statement, tmp_path):
    """对账单不含 DSP数据源 Sheet 不再触发 MissingSheetError（179）。"""
    ps = build_july_ps()
    # 把对账单当作 DSP+SA 模板传入 → 不含数据源 Sheet → 不报错，返回 None
    out, logs, _mod = reconciliation_processor.process_dsp_sa_template(
        july_statement, ps, str(tmp_path), use_com=False)
    assert out is None
    assert any("不包含" in l for l in logs)


# ---- Order Summary / 已推广-汇总 独立生成数据 ----
def test_order_summary_independent_dsp_data():
    """Order Summary 可独立生成 DSP 数据，不依赖对账单（174/175）。"""
    os_path = os.path.join(JULY_BASE, "HUION_US_(Advertiser)_Order_Summary_20260810015543.xlsx")
    orders = order_summary_processor.parse_order_summary(os_path, "Amazon美国站哥贝尔店")
    assert len(orders) == 11
    assert sum(float(o.total_cost) for o in orders) == pytest.approx(49025.52, abs=0.01)


def test_promoted_ads_independent_sa_data():
    """已推广-汇总可独立生成 PromotedAd 数据（176/177）。"""
    ps = build_july_ps()
    assert len(ps.promoted_ads) == 355
    assert float(ps.onsite_cost) == pytest.approx(145222.40, abs=0.01)
    assert float(ps.onsite_sales) == pytest.approx(620136.07, abs=0.01)


# ---- 对账单仅参与付款金额核对 ----
def test_statement_only_for_payment_verification(july_statement):
    """对账单只用于付款金额核对（178）。"""
    ps = build_july_ps()
    stmt = statement_processor.read_statement(july_statement)
    checks = statement_processor.verify_statement(ps, stmt)
    statuses = {c["name"]: c["status"] for c in checks}
    assert statuses.get("对账单 DSP费用") == "pass"
    assert statuses.get("对账单 返点金额") == "pass"
    assert statuses.get("对账单 结算金额") == "pass"


def test_statement_mismatch_detected(july_statement):
    """对账单金额与程序不符时 → fail。"""
    ps = build_july_ps()
    from processors.statement_processor import StatementData
    bad = StatementData(dsp_cost=50000.0, rebate=4000.0, payment=46000.0)
    checks = statement_processor.verify_statement(ps, bad)
    assert any(c["status"] == "fail" for c in checks)


# ---- DSP+SA 分析模板（可选）----
def test_analysis_template_writes_and_refreshes(july_reconciliation, tmp_path):
    """分析模板存在时正常写数据源并刷新 Pivot（180）。"""
    ps = build_july_ps()
    out, logs, modified = reconciliation_processor.process_dsp_sa_template(
        july_reconciliation, ps, str(tmp_path), use_com=False)
    assert out is not None
    assert any("DSP数据源" in m for m in modified)


def test_no_analysis_template_dsp_still_works(tmp_path):
    """分析模板未上传时 DSP 基础计算仍可继续（181）。"""
    ps = build_july_ps()
    assert float(ps.dsp_cost) == pytest.approx(49025.52, abs=0.01)
    assert float(ps.payment_amount) == pytest.approx(45103.48, abs=0.01)
