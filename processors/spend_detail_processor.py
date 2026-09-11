"""DSP 花费明细处理：复制模板，只改 B3/E3 与 H:X 数据，绝不改任何 merge range。

模板自身的核算维度（F 列，如 F3:F8=哥贝尔店 / F9:F13=DeePaint店）、店铺品牌（G 列）、
国家（C）、币种（D）全部沿用模板。程序不推导店铺、不 unmerge / remerge、不根据 Order
数量重新调整合并区域。若当月 Order 数量与模板可填写行数不一致，直接报错要求人工确认。
"""
from __future__ import annotations

import os
from typing import List, Tuple

import openpyxl

from models.payment_summary import PaymentSummary
from rules import amount_fields
from utils import money

# DSP花费明细 列 -> DSPOrder 字段（H 起为数据列）
DETAIL_COLUMNS = {
    "H": "order",
    "I": "order_id",
    "J": "total_cost",
    "K": "impressions",
    "L": "clicks",
    "M": "ctr",
    "N": "total_dpv",
    "O": "total_dpvr",
    "P": "total_atc",
    "Q": "total_atcr",
    "R": "promoted_sales",
    "S": "promoted_roas",
    "T": "total_purchases",
    "U": "total_purchase_rate",
    "V": "total_units_sold",
    "W": "total_sales",
    "X": "total_roas",
}

DATA_START_ROW = 3


def _merged_signature(ws) -> List[str]:
    """合并区域的完整签名（含具体 Range，不只看数量）。"""
    return sorted(str(rng) for rng in ws.merged_cells.ranges)


def _fillable_rows(ws) -> int:
    """从模板合并区域推断可填写的明细行数（数据区从 DATA_START_ROW 开始）。

    以「起始行 = DATA_START_ROW」的合并区域（如 B3:B13、E3:E13、G3:G13）的最大结束行
    确定数据区总行数；F 列的分段合并（F3:F8 / F9:F13）不影响总行数。
    """
    max_end = DATA_START_ROW - 1
    for rng in ws.merged_cells.ranges:
        if rng.min_row == DATA_START_ROW:
            max_end = max(max_end, rng.max_row)
    return max_end - DATA_START_ROW + 1


def process_spend_detail(template_path: str, ps: PaymentSummary, output_dir: str,
                         output_name: str = None, store_default: str = "Amazon美国站哥贝尔店",
                         brand: str = "HUION", stores: list = None) -> Tuple[str, List[str]]:
    logs: List[str] = []
    wb = openpyxl.load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]

    orders = ps.dsp_orders
    n = len(orders)

    before_merged = _merged_signature(ws)
    fillable = _fillable_rows(ws)

    # Order 数量与模板可填写行数不一致 → 停止，绝不为了塞数据调整 merge。
    if n != fillable:
        raise ValueError(
            f"DSP花费明细模板行数与当月Order数量不一致（模板可填写 {fillable} 行，当月 {n} 条 Order），"
            "请人工确认模板结构。"
        )

    # 写元数据：B3 实际年月、C3 国家、D3 币种、E3 申请支付金额（=对账单结算金额）、
    # F 核算维度、G 店铺品牌；再写 H:X Order 明细。merge range 保持不变。
    ws["B3"] = ps.month_code
    ws["C3"] = "美国"
    ws["D3"] = "USD"
    ws["E3"] = money.q2(amount_fields.resolve_amount(ps, "spend_detail.apply_amount"))
    ws["G3"] = brand
    _fill_store_column(ws, stores if stores else [store_default])

    for i, o in enumerate(orders):
        r = DATA_START_ROW + i
        for col, attr in DETAIL_COLUMNS.items():
            ws.cell(row=r, column=openpyxl.utils.column_index_from_string(col),
                    value=getattr(o, attr))

    # 校验 merge range 处理前后完全一致（不是只比较数量）。
    after_merged = _merged_signature(ws)
    if before_merged != after_merged:
        raise ValueError(f"❌ DSP花费明细合并区域发生变化：\n{before_merged}\n→\n{after_merged}")

    if output_name is None:
        output_name = f"DSP花费明细-{ps.month_code}_已处理.xlsx"
    out_path = os.path.join(output_dir, output_name)
    wb.save(out_path)
    logs.append(f"✅ DSP花费明细：{n} 条（merge range 保持不变，核算维度/店铺品牌已填写）")
    return out_path, logs


def _fill_store_column(ws, stores: list) -> None:
    """把 F 列（核算维度）各合并区域的左上角单元格按顺序填上店铺名。不改变 merge 结构。"""
    tops = sorted({rng.min_row for rng in ws.merged_cells.ranges
                   if rng.min_col == 6 and rng.min_row >= DATA_START_ROW})
    for i, row in enumerate(tops):
        if i < len(stores):
            ws.cell(row=row, column=6, value=stores[i])
