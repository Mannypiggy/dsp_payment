"""测试：申请月份一致性强校验 + 完整月检查（规格 12 节 123-125 + 148-152）。"""
import datetime
import os

import openpyxl
import pytest

from conftest import JULY_BASE
from pipeline import MonthMismatchError, run_pipeline
from processors import month_check
from rules import month_rules
from utils.config_loader import load_config

_OS = os.path.join(JULY_BASE, "HUION_US_(Advertiser)_Order_Summary_20260810015543.xlsx")


def _make_inputs(tmp_path, config, month="2026-07", month_confirmed=False):
    return {
        "month_str": month,
        "output_dir": str(tmp_path),
        "rebate_rate": 0.08,
        "use_com": False,
        "enable_screenshots": False,
        "month_confirmed": month_confirmed,
        "payment_due_date": "2026-08-20",
        "kpi_targets": config["kpi"],
        "order_summary_path": _OS,
        "promoted_ads_path": None,
        "dsp_sa_template_path": None,
        "spend_detail_template_path": None,
        "spend_ratio_template_path": None,
        "meeting_record_path": None,
        "invoice_path": None,
    }


def _partial_os(tmp_path, start="2026-07-01", end="2026-07-20"):
    """构造一个指定区间的 Order Summary 文件。"""
    import datetime as dt
    p = os.path.join(str(tmp_path), f"_os_{start}_{end}.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Interval start", "Interval end", "Campaign name", "Total cost", "Impressions", "ROAS"])
    ws.append([dt.datetime.fromisoformat(start), dt.datetime.fromisoformat(end), "HUION_Conversion_X", 100, 1000, 2.0])
    wb.save(p)
    return p


def test_detect_interval(july_order_summary):
    from processors import order_summary_processor
    interval = order_summary_processor.detect_interval(july_order_summary)
    assert interval is not None
    assert interval[0].date() == datetime.date(2026, 7, 1)
    assert interval[1].date() == datetime.date(2026, 7, 31)


def test_matching_month_passes(tmp_path):
    config = load_config()
    r = run_pipeline(_make_inputs(tmp_path, config, month="2026-07"), config)
    assert r["ps"].month == "2026-07"


def test_mismatched_order_summary_month_blocked(tmp_path):
    config = load_config()
    with pytest.raises(MonthMismatchError) as e:
        run_pipeline(_make_inputs(tmp_path, config, month="2026-08"), config)
    assert "2026-08" in str(e.value)


def test_full_month_pass(tmp_path):
    """完整月通过（148）。"""
    ps = month_rules.parse_month("2026-07")
    c = month_check.check_order_summary(ps, (datetime.datetime(2026, 7, 1), datetime.datetime(2026, 7, 31)))
    assert c.status == "PASS"


def test_partial_month_requires_confirm(tmp_path):
    """同月但非完整周期 → WARNING 需确认（149）。"""
    ps = month_rules.parse_month("2026-07")
    c = month_check.check_order_summary(ps, (datetime.datetime(2026, 7, 1), datetime.datetime(2026, 7, 20)))
    assert c.status == "WARNING"
    assert c.needs_confirm


def test_no_date_requires_confirm(tmp_path):
    """无法识别月份 → WARNING 需确认（150）。"""
    ps = month_rules.parse_month("2026-07")
    c = month_check.check_order_summary(ps, None)
    assert c.status == "WARNING"
    assert c.needs_confirm


def test_partial_month_blocked_without_confirm(tmp_path):
    """非完整月未确认 → 阻止（151）。"""
    config = load_config()
    inputs = _make_inputs(tmp_path, config, month="2026-07", month_confirmed=False)
    inputs["order_summary_path"] = _partial_os(tmp_path)
    with pytest.raises(MonthMismatchError):
        run_pipeline(inputs, config)


def test_partial_month_allowed_with_confirm(tmp_path):
    """非完整月人工确认后允许继续（152）。"""
    config = load_config()
    inputs = _make_inputs(tmp_path, config, month="2026-07", month_confirmed=True)
    inputs["order_summary_path"] = _partial_os(tmp_path)
    r = run_pipeline(inputs, config)
    assert r["ps"].month == "2026-07"


def test_mismatched_promoted_month_blocked(tmp_path):
    """已推广广告数据月份不一致 → 阻止（125）。"""
    config = load_config()
    src = os.path.join(str(tmp_path), "已推广_202606.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["类型", "曝光量", "点击", "花费-本币", "广告销售额-本币", "直接销售额-本币", "日期"])
    ws.append(["SP", 100, 10, 50, 200, 180, "2026-06-15"])
    wb.save(src)
    inputs = _make_inputs(tmp_path, config, month="2026-07")
    inputs["promoted_ads_path"] = src
    with pytest.raises(MonthMismatchError):
        run_pipeline(inputs, config)
