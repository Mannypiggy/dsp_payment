"""金额口径映射：每个目标字段明确绑定自己的业务金额语义。

禁止用全局开关同时影响多个金额位置。金额语义只有三类原始量，
其余都是派生量，避免 49,025.52 与 45,103.48 写错位置。
"""
from __future__ import annotations

from typing import List

# 金额语义中文说明（供页面展示）
AMOUNT_SEMANTICS = {
    "dsp_cost": "DSP原始广告费用 = SUM(Total cost)",
    "rebate_amount": "返点金额 = dsp_cost × rebate_rate",
    "payment_amount": "最终付款金额 = dsp_cost − rebate_amount",
    "onsite_cost": "站内花费 = SUM(花费-本币)",
    "onsite_sales": "站内销售额 = SUM(修正后销售额-本币)",
    "combined_cost": "DSP+站内花费 = dsp_cost + onsite_cost",
    "combined_sales": "DSP+站内销售额 = dsp_total_sales + onsite_sales",
}

# 每个目标字段 -> 金额语义（固定，不随配置切换）
AMOUNT_FIELD_MAPPING = {
    # 邮件
    "email.dsp_cost": "dsp_cost",
    "email.rebate_amount": "rebate_amount",
    "email.payment_amount": "payment_amount",
    "email.onsite_cost": "onsite_cost",
    "email.onsite_sales": "onsite_sales",
    "email.combined_cost": "combined_cost",
    "email.combined_sales": "combined_sales",
    # DSP花费明细「申请支付金额（原币）」= 对账单的结算金额（扣除返点后的最终付款金额）
    "spend_detail.apply_amount": "payment_amount",
    # 店铺花费占比「申请费用总额」= 扣除返点后的最终付款金额
    "spend_ratio.apply_amount": "payment_amount",
}

# 供页面展示的字段 -> 语义说明
FIELD_LABELS = {
    "email.dsp_cost": "邮件-DSP累计费用",
    "email.rebate_amount": "邮件-返点金额",
    "email.payment_amount": "邮件-付款金额",
    "email.onsite_cost": "邮件-站内累计费用",
    "email.onsite_sales": "邮件-站内销售额",
    "email.combined_cost": "邮件-DSP+站内累计费用",
    "email.combined_sales": "邮件-DSP+站内销售额",
    "spend_detail.apply_amount": "DSP花费明细-申请支付金额",
    "spend_ratio.apply_amount": "店铺花费占比-申请费用",
}


def resolve_amount(ps, field_key: str) -> float:
    """按映射返回目标字段应写入的金额。"""
    semantic = AMOUNT_FIELD_MAPPING.get(field_key)
    if semantic is None:
        raise ValueError(f"未定义的金额字段：{field_key}")
    return getattr(ps, semantic)


def binding_table(ps) -> List[dict]:
    """返回 [(字段, 语义, 金额)] 供页面核对。"""
    out = []
    for field_key, semantic in AMOUNT_FIELD_MAPPING.items():
        out.append({
            "field": FIELD_LABELS.get(field_key, field_key),
            "semantic": semantic,
            "description": AMOUNT_SEMANTICS.get(semantic, semantic),
            "value": round(getattr(ps, semantic), 2),
        })
    return out
