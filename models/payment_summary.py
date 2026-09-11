"""统一数据对象 PaymentSummary。

所有 Excel / KPI / 页面 / 邮件 / ZIP 都从这一份结果对象读取，
禁止各模块各自重新计算一套数据。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class DSPOrder:
    """DSP 单个 Order（来自 Amazon DSP Order Summary 的一行）。"""
    order: str                 # Order / Campaign name
    order_id: str = ""         # Campaign ID
    strategy: str = ""         # Consideration / Conversion
    total_cost: float = 0.0
    impressions: float = 0.0
    clicks: float = 0.0
    ctr: float = 0.0
    total_dpv: float = 0.0
    total_dpvr: float = 0.0
    total_atc: float = 0.0
    total_atcr: float = 0.0
    promoted_sales: float = 0.0   # Sales USD
    promoted_roas: float = 0.0
    total_purchases: float = 0.0
    total_purchase_rate: float = 0.0
    total_units_sold: float = 0.0
    total_sales: float = 0.0      # Total product sales
    total_roas: float = 0.0
    store: str = ""               # 核算维度（店铺）


@dataclass
class PromotedAd:
    """已推广广告的一行（已转换 SB2、已应用 SD 修正）。"""
    store: str = ""
    country: str = ""
    ad_type: str = ""             # SP / SD / SB / SBV（禁止残留 SB2）
    portfolio: str = ""
    campaign: str = ""
    ad_group: str = ""
    target_type: str = ""
    asin: str = ""
    msku: str = ""
    status: str = ""
    impressions: float = 0.0
    clicks: float = 0.0
    cost: float = 0.0             # 花费-本币
    sales: float = 0.0            # 销售额-本币（SD 已修正为直接成交销售额）
    orders: float = 0.0
    direct_sales: float = 0.0     # 直接成交销售额-本币


@dataclass
class StrategySummary:
    """邮件 / DSP+SA 中单个 Strategy（Consideration / Conversion）的汇总。"""
    strategy: str = ""
    impressions: float = 0.0
    dpv: float = 0.0
    atc: float = 0.0
    promoted_sales: float = 0.0
    total_cost: float = 0.0
    total_purchases: float = 0.0     # Total purchases（原始口径）
    total_units_sold: float = 0.0    # Total units sold（邮件「总销量」口径）
    total_sales: float = 0.0

    @property
    def dpv_rate(self) -> float:
        return self.dpv / self.impressions if self.impressions else 0.0

    @property
    def atc_rate(self) -> float:
        return self.atc / self.dpv if self.dpv else 0.0

    @property
    def promoted_roas(self) -> float:
        return float(self.promoted_sales / self.total_cost) if self.total_cost else 0.0

    @property
    def total_roas(self) -> float:
        return float(self.total_sales / self.total_cost) if self.total_cost else 0.0


@dataclass
class SAChannelSummary:
    """站内广告单个类型（SB/SBV/SD/SP）的汇总。"""
    ad_type: str = ""
    impressions: float = 0.0
    clicks: float = 0.0
    cost: float = 0.0
    sales: float = 0.0            # 修正后销售额
    orders: float = 0.0

    @property
    def ctr(self) -> float:
        return self.clicks / self.impressions if self.impressions else 0.0

    @property
    def cvr(self) -> float:
        return self.orders / self.clicks if self.clicks else 0.0

    @property
    def roas(self) -> float:
        return self.sales / self.cost if self.cost else 0.0


@dataclass
class KPIResult:
    """单个 KPI 的计算结果。

    明确区分两个值：
      - calculated：程序按 Order Summary 重新计算值，只用于差异校验 / 后台「参考计算值」提示；
      - final：最终汇报值（用户/模板确认后），正式业务输出（页面/KPI表/邮件/原因标题）一律读它。
    """
    name: str = ""
    target: float = 0.0
    calculated: float = 0.0        # 程序计算值（参考，不进正式输出）
    template_value: Optional[float] = None  # 模板已有 KPI 实际值
    final: float = 0.0             # 最终汇报值（用户/模板确认后；未确认时=calculated）
    is_met: bool = False
    better_when_higher: bool = True  # 越高越好；CPDPV 为 False
    # 口径对比：diff = calculated - template_value
    diff: Optional[float] = None
    diff_status: str = ""          # consistent / inconsistent / no_template
    reason: str = ""               # 达标/未达标原因（AI 提出，可编辑，非最终确认）


@dataclass
class PaymentSummary:
    """DSP 月度付款申请的统一结果对象。"""

    # 月份
    month: str = ""                # "2026-07"
    year: int = 0
    month_num: int = 0
    month_code: str = ""           # "202607"
    display_month: str = ""        # "2026年7月"
    month_type: str = ""           # 大促月份 / 其他月份

    # DSP 汇总
    dsp_orders: List[DSPOrder] = field(default_factory=list)
    dsp_impressions: float = 0.0
    dsp_dpv: float = 0.0
    dsp_atc: float = 0.0
    dsp_purchases: float = 0.0
    dsp_cost: float = 0.0
    dsp_promoted_sales: float = 0.0
    dsp_total_sales: float = 0.0
    dsp_roas: float = 0.0

    consideration: Optional[StrategySummary] = None
    conversion: Optional[StrategySummary] = None

    # 站内（已推广）汇总
    promoted_ads: List[PromotedAd] = field(default_factory=list)
    onsite_impressions: float = 0.0
    onsite_clicks: float = 0.0
    onsite_orders: float = 0.0
    onsite_cost: float = 0.0
    onsite_sales: float = 0.0
    onsite_roas: float = 0.0
    onsite_channels: List[SAChannelSummary] = field(default_factory=list)  # SB/SBV/SD/SP

    # DSP + 站内
    combined_cost: float = 0.0
    combined_sales: float = 0.0
    combined_roas: float = 0.0

    # 返点 / 付款（金额均为 Decimal）
    # 口径：billing_base_cost = q2(dsp_cost)
    #       rebate_raw = billing_base_cost × rebate_rate
    #       payment_raw = billing_base_cost − rebate_raw
    #       rebate_amount = q2(rebate_raw)
    #       payment_amount = q2(payment_raw)
    rebate_rate: float = 0.08
    billing_base_cost: float = 0.0    # = q2(dsp_cost)，计费基数
    rebate_raw: float = 0.0           # 全精度返点
    payment_raw: float = 0.0          # 全精度付款
    rebate_amount: float = 0.0        # 最终返点（q2 舍入）
    payment_amount: float = 0.0       # 最终付款（q2 舍入）
    payment_cycle_workdays: int = 20

    # KPI
    kpi_targets: dict = field(default_factory=dict)
    kpi_actual: dict = field(default_factory=dict)
    kpi_status: dict = field(default_factory=dict)
    kpi_reason: dict = field(default_factory=dict)
    kpi_results: List[KPIResult] = field(default_factory=list)
    # 明确区分：kpi_calculated（计算值，仅参考） vs kpi_final（最终汇报值，正式输出读它）
    kpi_calculated: dict = field(default_factory=dict)   # name -> calculated
    kpi_final: dict = field(default_factory=dict)        # name -> final
    # KPI 原因最终确认版（用户手动填写或 AI 生成后确认），正式邮件只读它
    kpi_reason_final: dict = field(default_factory=dict)  # name -> 最终原因文本

    # 付款方 / 收款方 / 截止日期
    payer: str = ""
    payee: dict = field(default_factory=dict)
    payment_due_date: str = ""     # 用户确认后填写，未确认为空
    payment_due_date_confirmed: bool = False  # 用户手动确认后为 True，才允许进入正式邮件

    # 附件 / 文件
    attachments: List[str] = field(default_factory=list)
    has_invoice: bool = False
    has_meeting_record: bool = False

    # 提示
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def add_warning(self, msg: str):
        self.warnings.append(msg)

    def add_error(self, msg: str):
        self.errors.append(msg)

    def kpi_reasons_confirmed(self) -> bool:
        """三个 KPI 原因是否都已确认（kpi_reason_final 三项齐全）。"""
        names = {r.name for r in self.kpi_results}
        return bool(names) and all(self.kpi_reason_final.get(n) for n in names)

    def is_final_ready(self) -> bool:
        """无阻断性错误 + 截止日期已确认 + KPI 原因已确认，才允许标记最终版。"""
        return (not self.errors
                and bool(self.payment_due_date)
                and self.payment_due_date_confirmed
                and self.kpi_reasons_confirmed())
