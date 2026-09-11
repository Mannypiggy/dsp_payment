"""月份防呆：三态检查（PASS / WARNING_CONFIRM_REQUIRED / FAIL）+ 完整月检查。

无法验证 ≠ 验证通过。无法识别日期 → 要求人工确认。
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass

from models.payment_summary import PaymentSummary


@dataclass
class MonthCheck:
    source: str
    status: str      # PASS / WARNING / FAIL
    message: str
    needs_confirm: bool = False


def last_day_of_month(year: int, month: int) -> datetime.date:
    if month == 12:
        return datetime.date(year, 12, 31)
    return datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)


def check_order_summary(ps: PaymentSummary, interval) -> MonthCheck:
    source = "DSP Order Summary"
    if interval is None:
        return MonthCheck(source, "WARNING", "无法自动验证该文件的数据月份", needs_confirm=True)
    start, end = interval
    if start.year != ps.year or start.month != ps.month_num:
        return MonthCheck(source, "FAIL",
                          f"数据月份 {start.year}-{start.month:02d} 与申请月份 {ps.month} 不一致")
    first = datetime.date(ps.year, ps.month_num, 1)
    last = last_day_of_month(ps.year, ps.month_num)
    if start.date() == first and end.date() == last:
        return MonthCheck(source, "PASS", f"完整月：{first} ~ {last}")
    return MonthCheck(source, "WARNING",
                      f"不是完整月份数据：检测区间 {start.date()} ~ {end.date()}（申请 {ps.month}）",
                      needs_confirm=True)


def check_promoted(ps: PaymentSummary, data_month) -> MonthCheck:
    source = "已推广广告数据"
    if data_month is None:
        return MonthCheck(source, "WARNING", "无法自动验证该文件的数据月份", needs_confirm=True)
    y, m = data_month
    if (y, m) != (ps.year, ps.month_num):
        return MonthCheck(source, "FAIL", f"数据月份 {y}-{m:02d} 与申请月份 {ps.month} 不一致")
    return MonthCheck(source, "PASS", f"数据月份 {y}-{m:02d}")
