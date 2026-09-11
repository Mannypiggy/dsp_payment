"""测试：KPI 计算 + 达标判断 + 口径核对（规格 78 节 16-21）。"""
from models.payment_summary import KPIResult, PaymentSummary, StrategySummary
from rules import kpi_rules


def _ps(cost=100.0, sales=650.0, cons_cost=40.0, cons_dpv=60.0, conv_sales=600.0, conv_cost=60.0):
    ps = PaymentSummary()
    ps.dsp_cost = cost
    ps.dsp_total_sales = sales
    ps.consideration = StrategySummary(strategy="Consideration", total_cost=cons_cost, dpv=cons_dpv)
    ps.conversion = StrategySummary(strategy="Conversion", total_cost=conv_cost, total_sales=conv_sales)
    return ps


class TestKPICalc:
    def test_total_roas(self):
        results = kpi_rules.compute_kpis(_ps(cost=100, sales=650))
        r = {x.name: x for x in results}
        assert abs(r["Total ROAS"].calculated - 6.5) < 1e-9

    def test_cpdpv(self):
        results = kpi_rules.compute_kpis(_ps(cons_cost=40, cons_dpv=60))
        r = {x.name: x for x in results}
        assert abs(r["Consideration CPDPV"].calculated - 40 / 60) < 1e-9

    def test_conversion_roas(self):
        results = kpi_rules.compute_kpis(_ps(conv_cost=60, conv_sales=600))
        r = {x.name: x for x in results}
        assert abs(r["Conversion ROAS"].calculated - 10.0) < 1e-9

    def test_roas_higher_better(self):
        ps = _ps(cost=100, sales=650)
        results = kpi_rules.compute_kpis(ps, targets={"total_roas": 6.0})
        r = {x.name: x for x in results}
        assert r["Total ROAS"].better_when_higher is True
        assert r["Total ROAS"].is_met is True  # 6.5 >= 6.0

    def test_cpdpv_lower_better(self):
        ps = _ps(cons_cost=40, cons_dpv=60)
        results = kpi_rules.compute_kpis(ps, targets={"consideration_cpdpv": 0.7})
        r = {x.name: x for x in results}
        assert r["Consideration CPDPV"].better_when_higher is False
        assert r["Consideration CPDPV"].is_met is True  # 0.667 <= 0.7


class TestReconcile:
    def _result(self, name, computed):
        return KPIResult(name=name, calculated=computed, final=computed, better_when_higher=True)

    def test_diff_within_tolerance_keep_template(self):
        rs = [self._result("Total ROAS", 7.04)]
        kpi_rules.reconcile_template_kpi(rs, {"Total ROAS": 7.03}, 0.02)
        assert rs[0].diff_status == "consistent"
        assert rs[0].final == 7.03  # 优先保留模板值

    def test_diff_exceeds_tolerance(self):
        rs = [self._result("Total ROAS", 7.50)]
        kpi_rules.reconcile_template_kpi(rs, {"Total ROAS": 7.03}, 0.02)
        assert rs[0].diff_status == "inconsistent"

    def test_no_template_use_computed(self):
        rs = [self._result("Total ROAS", 7.04)]
        kpi_rules.reconcile_template_kpi(rs, {}, 0.02)
        assert rs[0].diff_status == "no_template"
        assert rs[0].final == 7.04
