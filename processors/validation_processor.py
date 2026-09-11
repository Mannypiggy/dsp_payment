"""数据总校验：在允许下载「最终版」之前必须全部通过。"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List

from models.payment_summary import PaymentSummary
from rules.promoted_ads_mapping import ALLOWED_TYPES
from utils import money

EPS = 0.01


@dataclass
class CheckResult:
    name: str
    status: str   # pass / warn / fail
    message: str = ""


def validate(ps: PaymentSummary) -> List[CheckResult]:
    results: List[CheckResult] = []

    # 1. Strategy 识别
    unknown = [o.order for o in ps.dsp_orders if o.strategy not in ("Consideration", "Conversion")]
    if unknown:
        results.append(CheckResult("Strategy识别", "fail",
                                   f"存在未识别 Strategy 的 Order：{unknown}"))
    else:
        results.append(CheckResult("Strategy识别", "pass",
                                   f"{len(ps.dsp_orders)}/{len(ps.dsp_orders)} 全部识别"))

    # 2. DSP Cost / Sales 内部一致（单一来源，从 Order Summary 直接聚合）
    cost_ok = abs(float(ps.dsp_cost) - sum(o.total_cost for o in ps.dsp_orders)) <= EPS
    results.append(CheckResult("DSP Cost 核对", "pass" if cost_ok else "fail",
                               f"DSP Cost = {ps.dsp_cost:.2f}"))

    # 3. Strategy 汇总：Consideration + Conversion = Total
    if ps.consideration and ps.conversion:
        diff_cost = abs(ps.consideration.total_cost + ps.conversion.total_cost - ps.dsp_cost)
        diff_sales = abs(ps.consideration.total_sales + ps.conversion.total_sales - ps.dsp_total_sales)
        diff_imp = abs(ps.consideration.impressions + ps.conversion.impressions - ps.dsp_impressions)
        ok = diff_cost <= EPS and diff_sales <= EPS and diff_imp <= 0.5
        results.append(CheckResult("DSP Strategy 汇总", "pass" if ok else "fail",
                                   f"Cost差={diff_cost:.4f}, Sales差={diff_sales:.4f}"))
    else:
        results.append(CheckResult("DSP Strategy 汇总", "warn", "缺少 Consideration/Conversion 数据"))

    # 4. 广告类型校验（不得有 SB2）
    bad_types = [a.ad_type for a in ps.promoted_ads if a.ad_type not in ALLOWED_TYPES]
    if bad_types:
        results.append(CheckResult("广告类型校验", "fail", f"存在非法类型：{set(bad_types)}"))
    else:
        results.append(CheckResult("广告类型校验", "pass", "无 SB2 残留"))

    # 5. SD 校验：SD 销售额 = 直接成交销售额
    sd_bad = [a for a in ps.promoted_ads if a.ad_type == "SD" and abs(a.sales - a.direct_sales) > EPS]
    if sd_bad:
        results.append(CheckResult("SD 销售额修正", "fail", f"{len(sd_bad)} 条 SD 未修正"))
    else:
        results.append(CheckResult("SD 销售额修正", "pass", "全部 SD 已修正"))

    # 6. DSP+SA 综合
    combined_cost_ok = abs(ps.dsp_cost + money.D(ps.onsite_cost) - money.D(ps.combined_cost)) <= EPS
    combined_sales_ok = abs(ps.dsp_total_sales + money.D(ps.onsite_sales) - money.D(ps.combined_sales)) <= EPS
    ok = combined_cost_ok and combined_sales_ok
    results.append(CheckResult("DSP+SA 综合", "pass" if ok else "fail",
                               f"综合花费={ps.combined_cost:.2f}，综合销售额={ps.combined_sales:.2f}"))

    # 7. 返点 / 付款（口径：billing_base_cost=q2(dsp_cost)，rebate/payment 均 q2 舍入）
    billing_base_cost = money.q2(ps.dsp_cost)
    rebate_ok = (ps.billing_base_cost == billing_base_cost
                 and abs(ps.billing_base_cost * ps.rebate_rate - ps.rebate_raw) <= EPS
                 and abs(money.q2(ps.rebate_raw) - ps.rebate_amount) <= EPS)
    payment_ok = (abs(ps.billing_base_cost - ps.rebate_raw - ps.payment_raw) <= EPS
                  and abs(money.q2(ps.payment_raw) - ps.payment_amount) <= EPS)
    results.append(CheckResult("返点计算", "pass" if rebate_ok else "fail",
                               f"返点={ps.rebate_amount:.2f}"))
    results.append(CheckResult("付款金额", "pass" if payment_ok else "fail",
                               f"付款={ps.payment_amount:.2f}"))

    # 8. 发票 / 截止日期 / KPI 原因（非阻断，但影响「最终版」状态）
    if not ps.has_invoice:
        results.append(CheckResult("发票", "warn", "缺少发票（允许生成草稿）"))
    else:
        results.append(CheckResult("发票", "pass", "已上传"))
    if not ps.payment_due_date or not ps.payment_due_date_confirmed:
        results.append(CheckResult("截止付款日期", "warn", "待确认"))
    else:
        results.append(CheckResult("截止付款日期", "pass", ps.payment_due_date))
    if not ps.kpi_reasons_confirmed():
        results.append(CheckResult("KPI原因", "warn", "待确认"))
    else:
        results.append(CheckResult("KPI原因", "pass", "三项原因已确认"))

    return results


def has_fail(results: List[CheckResult]) -> bool:
    return any(r.status == "fail" for r in results)
