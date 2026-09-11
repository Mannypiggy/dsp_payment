"""已推广广告字段映射 + SB2 转换 + SD 销售额修正规则。"""
from __future__ import annotations

from typing import Dict, List, Optional

from models.payment_summary import PromotedAd


# 标准化字段 -> 源表头候选（按优先级）
PROMOTED_FIELD_ALIASES: Dict[str, List[str]] = {
    "type": ["类型", "广告类型", "广告产品类型"],
    "portfolio": ["广告组合名称", "广告组合", "广告组合名", "组合名称"],
    "campaign": ["广告活动名称", "广告活动", "广告活动名", "活动名称"],
    "ad_group": ["广告组", "广告组名称", "广告组名"],
    "target_type": ["广告组投放类型", "投放类型", "投放方式"],
    "asin": ["ASIN", "asin", "商品ASIN"],
    "msku": ["MSKU", "msku", "SKU"],
    "status": ["广告有效状态", "状态", "有效状态"],
    "impressions": ["曝光量", "曝光", "Impressions", "展示量"],
    "clicks": ["点击", "点击量", "Clicks"],
    "cost": ["花费-本币", "花费(本币)", "花费", "Spend", "花费（本币）"],
    "sales": ["广告销售额-本币", "销售额本币", "广告销售额(本币)", "广告销售额（本币）", "销售额-本币"],
    "orders": ["广告订单", "订单数量", "订单数", "Orders", "广告订单数"],
    "direct_sales": ["直接销售额-本币", "直接成交销售额-本币", "直接销售额(本币)", "直接成交销售额"],
}


def convert_type(raw_type: str, campaign_name: str = "", portfolio_name: str = "") -> str:
    """广告类型转换：SP->SP, SD->SD, SB2->(SBV|SB)。

    若类型=SB2：名称（广告活动名称优先，其次广告组合名称）含 SBV -> SBV，否则 -> SB。
    其余类型原样返回。
    """
    t = (raw_type or "").strip().upper()
    if t == "SB2":
        name = f"{campaign_name or ''} {portfolio_name or ''}".upper()
        if "SBV" in name:
            return "SBV"
        return "SB"
    return t


def apply_sd_rule(ad: PromotedAd) -> PromotedAd:
    """SD 特殊规则：SD 的 销售额-本币 = 直接成交销售额-本币。

    只影响 SD，不修改其他类型。
    """
    if ad.ad_type == "SD":
        ad.sales = ad.direct_sales
    return ad


def build_field_map(header_row: List[str]) -> Dict[str, Optional[int]]:
    header_norm = [(str(h) if h is not None else "").strip() for h in header_row]
    result: Dict[str, Optional[int]] = {}
    for field, aliases in PROMOTED_FIELD_ALIASES.items():
        result[field] = None
        for alias in aliases:
            if alias in header_norm:
                result[field] = header_norm.index(alias)
                break
    return result


def find_header_row(rows: List[List], max_scan: int = 10) -> Optional[int]:
    """定位表头行：包含「类型」且包含「花费」或「销售额」的行。"""
    for i, row in enumerate(rows[:max_scan]):
        norm = [str(c) if c is not None else "" for c in row]
        joined = "|".join(norm)
        if "类型" in joined and ("花费" in joined or "销售额" in joined):
            return i
    return None


ALLOWED_TYPES = {"SP", "SD", "SB", "SBV"}
