"""真实 2026-07 历史数据完整回归测试（规格 79 节）。

历史数字只用于测试，不写死进生产代码。
"""
import os

import pytest

from conftest import (JULY_BASE, build_july_ps, july_order_summary,
                      july_reconciliation, july_spend_detail, july_spend_ratio,
                      july_invoice, july_meeting, july_promoted_extract)
from pipeline import run_pipeline
from utils.config_loader import load_config


class TestDSPHistorical:
    def test_dsp_impressions(self):
        ps = build_july_ps()
        assert ps.dsp_impressions == 13669440

    def test_dsp_cost(self):
        ps = build_july_ps()
        assert float(ps.dsp_cost) == pytest.approx(49025.52, abs=0.01)

    def test_dsp_sales(self):
        ps = build_july_ps()
        assert float(ps.dsp_total_sales) == pytest.approx(345000.39, abs=0.01)

    def test_dsp_roas(self):
        ps = build_july_ps()
        assert ps.dsp_roas == pytest.approx(7.037, abs=0.005)


class TestStrategyHistorical:
    def test_consideration(self):
        c = build_july_ps().consideration
        assert c.impressions == 5294118
        assert c.dpv == 52672
        assert c.atc == 1306
        assert float(c.promoted_sales) == pytest.approx(15938.50, abs=0.01)
        assert float(c.total_cost) == pytest.approx(19607.96, abs=0.01)
        assert c.total_units_sold == 236   # 总销量 = Total units sold
        assert float(c.total_sales) == pytest.approx(24291.87, abs=0.01)

    def test_conversion(self):
        c = build_july_ps().conversion
        assert c.impressions == 8375322
        assert c.dpv == 91463
        assert c.atc == 7953
        assert float(c.promoted_sales) == pytest.approx(281382.35, abs=0.01)
        assert float(c.total_cost) == pytest.approx(29417.57, abs=0.01)
        assert c.total_units_sold == 3223
        assert float(c.total_sales) == pytest.approx(320708.52, abs=0.01)


class TestOnsiteHistorical:
    def test_onsite(self):
        ps = build_july_ps()
        assert ps.onsite_impressions == 16405160
        assert float(ps.onsite_cost) == pytest.approx(145222.40, abs=0.01)
        assert float(ps.onsite_sales) == pytest.approx(620136.07, abs=0.01)
        assert ps.onsite_roas == pytest.approx(4.270, abs=0.005)

    def test_combined(self):
        ps = build_july_ps()
        assert float(ps.combined_cost) == pytest.approx(194247.92, abs=0.01)
        assert float(ps.combined_sales) == pytest.approx(965136.46, abs=0.01)
        assert ps.combined_roas == pytest.approx(4.969, abs=0.005)


class TestRebateHistorical:
    def test_rebate(self):
        ps = build_july_ps()
        assert float(ps.rebate_rate) == 0.08
        assert float(ps.rebate_amount) == pytest.approx(3922.04, abs=0.01)

    def test_payment(self):
        ps = build_july_ps()
        assert float(ps.payment_amount) == pytest.approx(45103.48, abs=0.01)


class TestKPIHistorical:
    def test_kpi_computed(self):
        ps = build_july_ps()
        r = {k.name: k for k in ps.kpi_results}
        assert r["Total ROAS"].calculated == pytest.approx(7.037, abs=0.005)
        assert r["Consideration CPDPV"].calculated == pytest.approx(0.372, abs=0.005)
        assert r["Conversion ROAS"].calculated == pytest.approx(10.902, abs=0.005)

    def test_kpi_reconcile_keeps_template_value(self):
        """历史模板 7.03/10.89 vs 计算 7.04/10.90，差异 <= 0.02 → 保留模板值。"""
        from rules.kpi_rules import compute_kpis, reconcile_template_kpi
        ps = build_july_ps()
        results = compute_kpis(ps, targets=load_config()["kpi"])
        reconcile_template_kpi(results, {
            "Total ROAS": 7.03, "Consideration CPDPV": 0.37, "Conversion ROAS": 10.89,
        }, 0.02)
        r = {k.name: k for k in results}
        assert r["Total ROAS"].diff_status == "consistent"
        assert r["Total ROAS"].final == 7.03
        assert r["Conversion ROAS"].final == 10.89
        assert r["Consideration CPDPV"].final == 0.37


class TestEndToEnd:
    def test_full_pipeline_all_pass(self, tmp_path, july_promoted_extract, july_order_summary,
                                    july_reconciliation, july_spend_detail, july_spend_ratio,
                                    july_meeting, july_invoice):
        config = load_config()
        inputs = {
            "month_str": "2026-07",
            "output_dir": str(tmp_path),
            "rebate_rate": 0.08,
            "month_confirmed": True,  # 提取的已推广数据源无日期列，人工确认
            "payment_due_date": "2026-08-20",
            "kpi_targets": config["kpi"],
            "order_summary_path": july_order_summary,
            "promoted_ads_path": july_promoted_extract,
            "dsp_sa_template_path": july_reconciliation,
            "spend_detail_template_path": july_spend_detail,
            "spend_ratio_template_path": july_spend_ratio,
            "meeting_record_path": july_meeting,
            "invoice_path": july_invoice,
        }
        r = run_pipeline(inputs, config)
        # 所有校验通过（无 fail）
        checks = r["outputs"]["validation"]
        assert not any(c.status == "fail" for c in checks), [c.message for c in checks if c.status == "fail"]
        # 输出文件齐全（HTML 邮件为主版本）
        for key in ("reconciliation", "spend_detail", "spend_ratio", "email_html", "email_txt", "zip"):
            assert key in r["outputs"] and os.path.exists(r["outputs"][key])
        # 未上传 KPI 模板 → 不生成 KPI Excel
        assert "kpi" not in r["outputs"]
        # 格式保护通过
        assert r["outputs"]["reconciliation_format"]["ok"]
        assert r["outputs"]["spend_detail_format"]["ok"]
