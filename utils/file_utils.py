"""文件类型识别：通过 Sheet 名 / 表头 / 结构判断文件用途，不依赖文件名。"""
from __future__ import annotations

import os
from typing import Optional

import openpyxl

FILE_KINDS = [
    "order_summary",   # Amazon DSP Order Summary
    "promoted_ads",    # 已推广-汇总（源数据，含 SB2）
    "reconciliation",  # 对账单模板（含 DSP数据源 / 已推广广告数据 数据源 / DSP+SA）
    "spend_detail",    # DSP花费明细模板
    "spend_ratio",     # 店铺花费占比模板
    "invoice",         # 发票 PDF
    "meeting_record",  # 会议记录 docx
]


def detect_file_kind(path: str) -> Optional[str]:
    """根据文件内容识别用途。返回 FILE_KINDS 之一或 None。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return "invoice"
    if ext in (".docx", ".doc"):
        return "meeting_record"
    if ext == ".xls":
        return _detect_xls_kind(path)
    if ext == ".xlsx":
        return _detect_xlsx_kind(path)
    return None


def _detect_xlsx_kind(path: str) -> Optional[str]:
    try:
        wb = openpyxl.load_workbook(path, data_only=True)  # 普通模式：read_only 下 max_column 不可靠
    except Exception:
        return None
    sheets = [s.strip() for s in wb.sheetnames]

    # 对账单：含 DSP数据源 / 已推广广告数据 数据源 / DSP+SA
    has_dsp_src = any("dsp数据源" in s.lower() for s in sheets)
    has_promo_src = any("已推广广告数据" in s and "数据源" in s for s in sheets)
    has_dsp_sa = any(s.replace(" ", "").lower() == "dsp+sa" for s in sheets)
    if has_dsp_src and has_promo_src:
        return "reconciliation"

    # 花费明细：含「示例」或表头含「申请支付金额」
    if any("示例" in s for s in sheets):
        return "spend_detail"

    # Order Summary：单 Sheet 表头含 Total cost + Impressions + ROAS
    if len(sheets) == 1:
        headers = _read_header_cells(wb[wb.sheetnames[0]])
        joined = "|".join(headers)
        if "Total cost" in joined and "Impressions" in joined and ("ROAS" in joined):
            return "order_summary"

    # 已推广广告：表头含「类型」且含「花费」或「销售额」
    for sn in sheets:
        headers = _read_header_cells(wb[sn], max_rows=10)
        joined = "|".join(headers)
        if "类型" in joined and ("花费" in joined or "销售额" in joined):
            return "promoted_ads"

    return None


def _read_header_cells(ws, max_rows: int = 3) -> list:
    out = []
    for i, row in enumerate(ws.iter_rows(max_row=max_rows, values_only=True)):
        for v in row:
            if v is not None:
                out.append(str(v).strip())
    return out


def _detect_xls_kind(path: str) -> Optional[str]:
    try:
        import xlrd
        wb = xlrd.open_workbook(path)
    except Exception:
        return None
    for name in wb.sheet_names():
        if "分配DSP费用" in name or "店铺花费占比" in name:
            return "spend_ratio"
    return None
