"""月份规则：月份编码、大促月/其他月份判断。"""
from __future__ import annotations

from typing import Tuple

from models.payment_summary import PaymentSummary


def parse_month(month_str: str, peak_months=None) -> PaymentSummary:
    """解析 "YYYY-MM" 得到月份信息。

    返回一个已填充月份字段的 PaymentSummary。
    """
    if peak_months is None:
        peak_months = [7, 10, 11, 12]

    s = month_str.strip()
    parts = s.split("-")
    if len(parts) != 2:
        raise ValueError(f"月份格式应为 YYYY-MM，实际为 {month_str!r}")
    year = int(parts[0])
    month_num = int(parts[1])
    if month_num < 1 or month_num > 12:
        raise ValueError(f"月份超出范围：{month_num}")

    ps = PaymentSummary()
    ps.month = f"{year:04d}-{month_num:02d}"
    ps.year = year
    ps.month_num = month_num
    ps.month_code = f"{year:04d}{month_num:02d}"
    ps.display_month = f"{year}年{month_num}月"

    if month_num in peak_months:
        ps.month_type = "大促月份"
    else:
        ps.month_type = "其他月份"
    return ps


def month_label(month_num: int, peak_months=None) -> str:
    if peak_months is None:
        peak_months = [7, 10, 11, 12]
    return "大促月份" if month_num in peak_months else "其他月份"


def excel_date_serial(year: int, month: int) -> float:
    """把年月转成 Excel 日期序列号（当月 1 号）。

    Excel 序列号：1900-01-01 = 1（含 1900 闰年 bug，与 xlrd 口径一致）。
    """
    import datetime
    base = datetime.date(1899, 12, 30)
    d = datetime.date(year, month, 1)
    return float((d - base).days)
