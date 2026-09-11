"""测试：文件识别 + 缺文件处理 + ZIP（规格 78 节 58-61）。"""
import os

import pytest

from conftest import JULY_BASE
from pipeline import run_pipeline
from utils import file_utils
from utils.config_loader import load_config

_OS = os.path.join(JULY_BASE, "HUION_US_(Advertiser)_Order_Summary_20260810015543.xlsx")
_REC = os.path.join(JULY_BASE, "HUION-US-ENTITY-对账单-自用.xlsx")
_SD = os.path.join(JULY_BASE, "DSP花费明细-202607.xlsx")
_SR = os.path.join(JULY_BASE, "亚马逊美国DSP店铺花费占比-2026-07.xls")
_MEETING = r"C:\Users\41250\Desktop\2026年7月Huion-US对账单\2026年7月Huion-US对账单\测试\7月dsp会议记录.docx"
_INVOICE = os.path.join(JULY_BASE, "Invoice_M001098_HUION GLOBAL (HK) LIMITED.pdf")


def _base_inputs(tmp_path, config):
    return {
        "month_str": "2026-07",
        "output_dir": str(tmp_path),
        "rebate_rate": 0.08,
        "use_com": False,
        "enable_screenshots": False,
        "payment_due_date": "2026-08-20",
        "kpi_targets": {"total_roas": 6.5, "consideration_cpdpv": 0.68, "conversion_roas": 8.8},
        "order_summary_path": _OS,
        "promoted_ads_path": None,
        "dsp_sa_template_path": _REC,
        "spend_detail_template_path": _SD,
        "spend_ratio_template_path": _SR,
        "meeting_record_path": _MEETING,
        "invoice_path": _INVOICE,
    }


def test_missing_order_summary(tmp_path):
    config = load_config()
    inputs = _base_inputs(tmp_path, config)
    inputs["order_summary_path"] = None
    with pytest.raises(FileNotFoundError):
        run_pipeline(inputs, config)


def test_missing_template(tmp_path):
    config = load_config()
    inputs = _base_inputs(tmp_path, config)
    inputs["dsp_sa_template_path"] = None
    inputs["spend_detail_template_path"] = None
    inputs["spend_ratio_template_path"] = None
    r = run_pipeline(inputs, config)
    assert any("DSP+SA 分析模板" in l for l in r["logs"])


def test_missing_invoice_allows_draft(tmp_path):
    config = load_config()
    inputs = _base_inputs(tmp_path, config)
    inputs["invoice_path"] = None
    r = run_pipeline(inputs, config)
    checks = r["outputs"]["validation"]
    inv = next(c for c in checks if c.name == "发票")
    assert inv.status == "warn"
    assert os.path.exists(r["outputs"]["zip"])


def test_zip_generated(tmp_path):
    config = load_config()
    inputs = _base_inputs(tmp_path, config)
    r = run_pipeline(inputs, config)
    assert os.path.exists(r["outputs"]["zip"])
    assert r["outputs"]["zip"].endswith("DSP付款申请_202607.zip")


class TestFileDetection:
    def test_order_summary(self):
        assert file_utils.detect_file_kind(_OS) == "order_summary"

    def test_reconciliation(self):
        assert file_utils.detect_file_kind(_REC) == "reconciliation"

    def test_spend_detail(self):
        assert file_utils.detect_file_kind(_SD) == "spend_detail"

    def test_spend_ratio(self):
        assert file_utils.detect_file_kind(_SR) == "spend_ratio"

    def test_invoice(self):
        assert file_utils.detect_file_kind(_INVOICE) == "invoice"

    def test_meeting(self):
        assert file_utils.detect_file_kind(_MEETING) == "meeting_record"
