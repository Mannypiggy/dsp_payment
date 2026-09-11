"""本轮返工新增回归测试：Pivot 保留 / 表头保留 / DSP花费明细用 dsp_cost / 附件统一 / KPI 模板可选。

聚焦本次返工修复的 7 个根因，用真实 2026-07 数据验证。
"""
import os
import zipfile

import openpyxl
import pytest

from conftest import JULY_BASE, build_july_ps, july_reconciliation, july_spend_detail
from pipeline import run_pipeline
from processors import reconciliation_processor, spend_detail_processor
from utils import excel_com, excel_utils
from utils.config_loader import load_config

_OS = os.path.join(JULY_BASE, "HUION_US_(Advertiser)_Order_Summary_20260810015543.xlsx")
_INVOICE = os.path.join(JULY_BASE, "Invoice_M001098_HUION GLOBAL (HK) LIMITED.pdf")


# ---- 一、PivotTable/PivotCache 数量 before == after（规格 17 节）----
def test_pivot_counts_preserved_openpyxl(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before = excel_utils.pivot_signature(july_reconciliation)
    out, _logs, _mod = reconciliation_processor.process_dsp_sa_template(
        july_reconciliation, ps, str(tmp_path), use_com=False)
    after = excel_utils.pivot_signature(out)
    assert after["pivot_tables"] == before["pivot_tables"] == 2
    assert after["pivot_caches"] == before["pivot_caches"] == 2
    assert after["pivot_cache_records"] == before["pivot_cache_records"] == 2


def test_pivot_counts_preserved_com(july_reconciliation, tmp_path):
    if not excel_com.com_available():
        pytest.skip("本机无 Excel COM")
    ps = build_july_ps()
    before = excel_utils.pivot_signature(july_reconciliation)
    out, _logs, _mod = reconciliation_processor.process_dsp_sa_template(
        july_reconciliation, ps, str(tmp_path), use_com=True)
    after = excel_utils.pivot_signature(out)
    assert after["pivot_tables"] == before["pivot_tables"]
    assert after["pivot_caches"] == before["pivot_caches"]


def test_pivot_parts_byte_identical_after_openpyxl_restore(july_reconciliation, tmp_path):
    """openpyxl 回退后 PivotTable XML 与模板逐字节一致（只允许 refreshOnLoad 注入到 cacheDefinition）。"""
    ps = build_july_ps()
    out, _logs, _mod = reconciliation_processor.process_dsp_sa_template(
        july_reconciliation, ps, str(tmp_path), use_com=False)
    tz = zipfile.ZipFile(july_reconciliation)
    oz = zipfile.ZipFile(out)
    for n in ("xl/pivotTables/pivotTable1.xml", "xl/pivotTables/pivotTable2.xml",
              "xl/pivotTables/_rels/pivotTable1.xml.rels", "xl/pivotTables/_rels/pivotTable2.xml.rels"):
        assert tz.read(n) == oz.read(n), f"{n} 被 openpyxl 改动"


# ---- 二、数据源表头第一行不被清空（规格 2 节）----
def test_dsp_source_header_row1_preserved(july_reconciliation, tmp_path):
    ps = build_july_ps()
    out, _logs, _mod = reconciliation_processor.process_dsp_sa_template(
        july_reconciliation, ps, str(tmp_path), use_com=False)
    ws = openpyxl.load_workbook(out)["DSP数据源"]
    assert ws.cell(1, 1).value == "Order"
    assert ws.cell(1, 3).value == "Total cost"


def test_promo_source_header_row1_preserved(july_reconciliation, tmp_path):
    ps = build_july_ps()
    out, _logs, _mod = reconciliation_processor.process_dsp_sa_template(
        july_reconciliation, ps, str(tmp_path), use_com=False)
    ws = openpyxl.load_workbook(out)["已推广广告数据 数据源"]
    assert ws.cell(1, 1).value == "店铺名称"
    assert ws.cell(1, 13).value == "花费-本币"


# ---- 三、DSP花费明细「申请支付金额（原币）」= 对账单结算金额（45103.48）----
def test_spend_detail_apply_amount_is_payment(july_spend_detail, tmp_path):
    ps = build_july_ps()
    out, _logs = spend_detail_processor.process_spend_detail(july_spend_detail, ps, str(tmp_path))
    ws = openpyxl.load_workbook(out)["示例"]
    assert float(ws["E3"].value) == pytest.approx(45103.48, abs=0.01)
    assert float(ws["E3"].value) != pytest.approx(49025.52, abs=0.01)


# ---- 四、发票状态：邮件/ZIP 同源，不再误判缺失（规格 12 节）----
def test_email_invoice_consistent_with_zip(tmp_path):
    config = load_config()
    inputs = {
        "month_str": "2026-07",
        "output_dir": str(tmp_path),
        "rebate_rate": 0.08,
        "use_com": False,
        "enable_screenshots": False,
        "month_confirmed": True,
        "payment_due_date": "2026-08-20",
        "kpi_targets": config["kpi"],
        "order_summary_path": _OS,
        "promoted_ads_path": None,
        "dsp_sa_template_path": None,
        "spend_detail_template_path": None,
        "spend_ratio_template_path": None,
        "meeting_record_path": None,
        "invoice_path": _INVOICE,
    }
    r = run_pipeline(inputs, config)
    email = r["outputs"]["email"]
    assert "Invoice发票" in email["text"]
    assert "缺少发票" not in email["text"]
    # ZIP 里确实含发票 PDF
    z = zipfile.ZipFile(r["outputs"]["zip"])
    assert any("Invoice" in n and n.endswith(".pdf") for n in z.namelist())


# ---- 五、未上传 KPI 模板不生成 KPI Excel（规格 5/15 节）----
def test_no_kpi_template_no_kpi_excel(tmp_path):
    config = load_config()
    inputs = {
        "month_str": "2026-07",
        "output_dir": str(tmp_path),
        "rebate_rate": 0.08,
        "use_com": False,
        "enable_screenshots": False,
        "month_confirmed": True,
        "payment_due_date": "2026-08-20",
        "kpi_targets": config["kpi"],
        "order_summary_path": _OS,
        "promoted_ads_path": None,
        "dsp_sa_template_path": None,
        "spend_detail_template_path": None,
        "spend_ratio_template_path": None,
        "meeting_record_path": None,
        "invoice_path": None,
        "kpi_template_path": None,
    }
    r = run_pipeline(inputs, config)
    assert "kpi" not in r["outputs"]
    assert not any("DSP_KPI-" in n for n in os.listdir(str(tmp_path)))


# ---- 六、KPI 邮件用 final（最终汇报值），不是 calculated（参考计算值）----
def test_email_uses_kpi_final_not_calculated():
    from processors import email_processor, kpi_processor
    ps = build_july_ps()
    kpi_processor.compute_and_reconcile(
        ps, targets=load_config()["kpi"],
        template_values={"Total ROAS": 7.03, "Consideration CPDPV": 0.37, "Conversion ROAS": 10.89},
        diff_tolerance=0.02)
    ps.payment_due_date = "2026-08-20"
    ps.payment_due_date_confirmed = True
    ps.kpi_reason_final = {"Total ROAS": "x", "Consideration CPDPV": "y", "Conversion ROAS": "z"}
    e = email_processor.generate_email(ps)
    # 精确校验 KPI 表的「实际」列 = final（最终汇报值），而不是 calculated（参考计算值）
    _headers, rows = email_processor._kpi_table(ps)
    by_name = {r[0]: r for r in rows}
    assert by_name["Total ROAS"][2] == "7.03"
    assert by_name["Consideration CPDPV"][2] == "0.37"
    assert by_name["Conversion ROAS"][2] == "10.89"
    assert by_name["Conversion ROAS"][2] != "10.90"


# ---- 七、AI 不可用时正式邮件不得出现技术性 fallback ----
def test_email_no_ai_garbage():
    from processors import email_processor
    ps = build_july_ps()
    ps.kpi_reason_final = {}  # 未确认原因
    e = email_processor.generate_email(ps)
    for bad in ("未启用AI", "API Key", "AI总结失败", "出单报告"):
        assert bad not in e["text"]
    assert "[DSP KPI截图：02_DSP_KPI.png]" in e["text"]


# ---- 八、截止付款日期未确认 → 待确认，不能出现具体日期 ----
def test_email_due_date_unconfirmed_no_date():
    from processors import email_processor
    ps = build_july_ps()
    ps.payment_due_date = ""
    ps.payment_due_date_confirmed = False
    e = email_processor.generate_email(ps)
    assert "截止付款时间：待确认" in e["text"]
    assert "2026-10" not in e["text"]
    assert "2026-09" not in e["text"]
