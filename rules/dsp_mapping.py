"""DSP Order Summary 字段映射与 Strategy 识别规则。"""
from __future__ import annotations

from typing import Dict, List, Optional


# 标准化字段 -> 源表头候选（按优先级）
DSP_FIELD_ALIASES: Dict[str, List[str]] = {
    "order": ["Order name", "Campaign name", "Order", "订单名称", "订单"],
    "order_id": ["Campaign ID", "Order ID", "CampaignId", "订单ID"],
    "total_cost": ["Total cost", "总花费金额（原币）"],
    "impressions": ["Impressions", "展示量"],
    "clicks": ["Click-throughs", "Clicks", "点击量"],
    "ctr": ["CTR", "点击率"],
    "total_dpv": ["Total DPV", "详情页浏览量"],
    "total_dpvr": ["Total DPVR", "详情页浏览率"],
    "total_atc": ["Total ATC", "加购量"],
    "total_atcr": ["Total ATCR", "加购率"],
    "promoted_sales": ["Sales USD", "广告销售额"],
    "promoted_roas": ["ROAS", "广告投资回报率"],
    "total_purchases": ["Total purchases", "总购买量"],
    "total_purchase_rate": ["Total purchase rate", "总购买率"],
    "total_units_sold": ["Total units sold", "总销量"],
    "total_sales": ["Total product sales", "总销售额"],
    "total_roas": ["Total ROAS", "总投资回报率"],
}


def detect_strategy(order_name: str) -> str:
    """根据 Order/Campaign 标题识别 Strategy。

    标题包含 Consideration -> Consideration
    标题包含 Conversion   -> Conversion
    否则返回 ""（未识别，调用方须停止最终文件生成）。
    """
    name = (order_name or "").strip()
    if not name:
        return ""
    upper = name.upper()
    # 先判断 Consideration，再判断 Conversion；两个词互斥优先级：包含 Consideration 优先。
    has_consideration = "CONSIDERATION" in upper
    has_conversion = "CONVERSION" in upper
    if has_consideration and not has_conversion:
        return "Consideration"
    if has_conversion and not has_consideration:
        return "Conversion"
    if has_consideration and has_conversion:
        # 理论上不应出现同时包含两者；保守按 Consideration 出现位置更靠前者判断，
        # 但更安全是标记未识别。
        return ""
    return ""


def build_field_map(header_row: List[str]) -> Dict[str, Optional[int]]:
    """把表头行映射为 标准化字段 -> 列索引(0-based)。

    找不到的字段值为 None，调用方据此输出 WARNING。
    """
    header_norm = [(str(h) if h is not None else "").strip() for h in header_row]
    result: Dict[str, Optional[int]] = {}
    for field, aliases in DSP_FIELD_ALIASES.items():
        result[field] = None
        for alias in aliases:
            if alias in header_norm:
                result[field] = header_norm.index(alias)
                break
    return result


def find_header_row(rows: List[List], max_scan: int = 5) -> Optional[int]:
    """在首若干行中定位表头行：包含 Total cost 且包含 Impressions 的行。"""
    for i, row in enumerate(rows[:max_scan]):
        norm = [str(c) if c is not None else "" for c in row]
        joined = "|".join(norm)
        if "Total cost" in joined and "Impressions" in joined:
            return i
    return None
