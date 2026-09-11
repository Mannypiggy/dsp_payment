"""财务金额 Decimal 工具。

核心财务金额统一使用 Decimal，避免二进制浮点误差。
舍入统一 ROUND_HALF_UP（四舍五入），最终保留 2 位。
"""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

TWOPLACES = Decimal("0.01")
ROUND_HALF_UP = ROUND_HALF_UP


def D(value) -> Decimal:
    """把数值转为 Decimal（float 走字符串，避免二进制误差）。"""
    if isinstance(value, Decimal):
        return value
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value))


def q2(value) -> Decimal:
    """四舍五入保留 2 位（ROUND_HALF_UP）。"""
    return D(value).quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def q2_display(value) -> str:
    """2 位小数显示字符串（千分位）。"""
    return f"{q2(value):,.2f}"
