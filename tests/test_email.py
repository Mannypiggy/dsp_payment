"""测试：邮件 DSP Strategy 汇总 / 站内汇总 / DSP+SA 综合（规格 78 节 27-46）。

用真实 2026-07 历史数据核对。
"""
import pytest

from conftest import build_july_ps
from processors import email_processor


@pytest.fixture(scope="module")
def ps():
    return build_july_ps()


class TestDSPStrategy:
    def test_consideration(self, ps):
        c = ps.consideration
        assert c.impressions == 5294118
        assert c.dpv == 52672
        assert c.atc == 1306
        assert float(c.promoted_sales) == pytest.approx(15938.50, abs=0.01)
        assert float(c.total_cost) == pytest.approx(19607.96, abs=0.01)
        assert c.total_units_sold == 236
        assert float(c.total_sales) == pytest.approx(24291.87, abs=0.01)

    def test_conversion(self, ps):
        c = ps.conversion
        assert c.impressions == 8375322
        assert c.dpv == 91463
        assert c.atc == 7953
        assert float(c.promoted_sales) == pytest.approx(281382.35, abs=0.01)
        assert float(c.total_cost) == pytest.approx(29417.57, abs=0.01)
        assert c.total_units_sold == 3223
        assert float(c.total_sales) == pytest.approx(320708.52, abs=0.01)

    def test_strategy_total(self, ps):
        c, v = ps.consideration, ps.conversion
        assert c.impressions + v.impressions == ps.dsp_impressions == 13669440
        assert abs(c.total_cost + v.total_cost - ps.dsp_cost) <= 0.01
        assert abs(c.total_sales + v.total_sales - ps.dsp_total_sales) <= 0.01

    def test_dpv_rate(self, ps):
        c = ps.consideration
        assert c.dpv_rate == pytest.approx(52672 / 5294118, rel=1e-6)

    def test_atc_rate(self, ps):
        c = ps.consideration
        assert c.atc_rate == pytest.approx(1306 / 52672, rel=1e-6)

    def test_promoted_roas(self, ps):
        c = ps.consideration
        assert c.promoted_roas == pytest.approx(15938.50 / 19607.9553, rel=1e-4)

    def test_total_roas(self, ps):
        assert ps.dsp_roas == pytest.approx(345000.39 / 49025.52, rel=1e-4)
        assert ps.dsp_roas == pytest.approx(7.037, abs=0.005)

    def test_cost_share(self, ps):
        c = ps.consideration
        assert float(c.total_cost / ps.dsp_cost) == pytest.approx(0.39995, abs=0.001)

    def test_sales_share(self, ps):
        c = ps.consideration
        assert float(c.total_sales / ps.dsp_total_sales) == pytest.approx(0.07041, abs=0.001)


class TestOnsite:
    def _channel(self, ps, t):
        return next(ch for ch in ps.onsite_channels if ch.ad_type == t)

    def test_sb(self, ps):
        ch = self._channel(ps, "SB")
        assert ch.impressions == 545084
        assert ch.cost == pytest.approx(10600.31, abs=0.01)
        assert ch.sales == pytest.approx(69960.74, abs=0.01)

    def test_sbv(self, ps):
        ch = self._channel(ps, "SBV")
        assert ch.impressions == 166756
        assert ch.cost == pytest.approx(4679.64, abs=0.01)
        assert ch.sales == pytest.approx(16587.52, abs=0.01)

    def test_sd(self, ps):
        ch = self._channel(ps, "SD")
        assert ch.impressions == 1245725
        assert ch.cost == pytest.approx(6017.98, abs=0.01)
        assert ch.sales == pytest.approx(14177.54, abs=0.01)

    def test_sp(self, ps):
        ch = self._channel(ps, "SP")
        assert ch.impressions == 14447595
        assert ch.cost == pytest.approx(123924.47, abs=0.01)
        assert ch.sales == pytest.approx(519410.27, abs=0.01)

    def test_no_sb2_in_email(self, ps):
        email = email_processor.generate_email(ps)
        assert "SB2" not in email["text"]

    def test_sd_uses_corrected_sales(self, ps):
        ch = self._channel(ps, "SD")
        # SD 销售额 = 直接成交销售额（修正后）
        assert ch.sales == pytest.approx(14177.54, abs=0.01)

    def test_ctr_total_clicks_over_impressions(self, ps):
        ctr = ps.onsite_clicks / ps.onsite_impressions
        assert ctr == pytest.approx(95637 / 16405160, rel=1e-6)

    def test_roas_total_sales_over_cost(self, ps):
        roas = ps.onsite_sales / ps.onsite_cost
        assert float(roas) == pytest.approx(620136.07 / 145222.40, rel=1e-4)


class TestCombined:
    def test_combined_cost(self, ps):
        assert float(ps.combined_cost) == pytest.approx(194247.92, abs=0.01)

    def test_combined_sales(self, ps):
        assert float(ps.combined_sales) == pytest.approx(965136.46, abs=0.01)

    def test_combined_roas(self, ps):
        assert ps.combined_roas == pytest.approx(965136.46 / 194247.92, rel=1e-4)
        assert ps.combined_roas == pytest.approx(4.969, abs=0.005)
