"""KPI 计算规则：Total ROAS / Consideration CPDPV / Conversion ROAS。"""
from __future__ import annotations

from typing import Optional

from models.payment_summary import KPIResult, PaymentSummary


def compute_kpis(ps: PaymentSummary, targets: Optional[dict] = None,
                 diff_tolerance: float = 0.02) -> list:
    """计算三个 KPI，返回 KPIResult 列表并写回 ps。

    - Total ROAS           = DSP Total Sales / DSP Total Cost（越高越好）
    - Consideration CPDPV  = Consideration Total Cost / Consideration Total DPV（越低越好）
    - Conversion ROAS      = Conversion Total Sales / Conversion Total Cost（越高越好）
    """
    if targets is None:
        targets = {}

    total_roas = float(ps.dsp_total_sales / ps.dsp_cost) if ps.dsp_cost else 0.0

    cons = ps.consideration
    if cons is not None and cons.dpv:
        cpdpv = float(cons.total_cost) / cons.dpv
    else:
        cpdpv = 0.0

    conv = ps.conversion
    if conv is not None and conv.total_cost:
        conversion_roas = float(conv.total_sales / conv.total_cost)
    else:
        conversion_roas = 0.0

    specs = [
        ("Total ROAS", "total_roas", total_roas, True),
        ("Consideration CPDPV", "consideration_cpdpv", cpdpv, False),
        ("Conversion ROAS", "conversion_roas", conversion_roas, True),
    ]

    results = []
    for name, key, computed, higher_better in specs:
        target = float(targets.get(key, 0.0))
        is_met = (computed >= target) if higher_better else (computed <= target)

        r = KPIResult(
            name=name,
            target=target,
            calculated=computed,
            final=computed,
            is_met=is_met,
            better_when_higher=higher_better,
        )
        results.append(r)

    ps.kpi_results = results
    ps.kpi_targets = {r.name: r.target for r in results}
    ps.kpi_calculated = {r.name: r.calculated for r in results}
    ps.kpi_final = {r.name: r.final for r in results}
    ps.kpi_actual = {r.name: r.final for r in results}
    ps.kpi_status = {r.name: ("达标" if r.is_met else "未达标") for r in results}
    return results


def reconcile_template_kpi(results: list, template_values: Optional[dict],
                           diff_tolerance: float = 0.02) -> list:
    """KPI 口径保护：把模板已有 KPI 实际值与程序计算值对比。

    - 差异 <= tolerance：口径基本一致，优先保留模板值。
    - 差异 >  tolerance：⚠ 差异，保留模板值但标记 diff_status=inconsistent。
    - 模板无值：使用程序计算值。
    """
    template_values = template_values or {}
    for r in results:
        tv = template_values.get(r.name)
        if tv is None:
            r.template_value = None
            r.diff_status = "no_template"
            r.final = r.calculated
            continue
        r.template_value = float(tv)
        r.diff = round(r.calculated - r.template_value, 4)
        if abs(r.diff) <= diff_tolerance:
            r.diff_status = "consistent"
            r.final = r.template_value  # 优先保留模板值
        else:
            r.diff_status = "inconsistent"
            r.final = r.template_value  # 不静默覆盖，等待人工确认
    return results
