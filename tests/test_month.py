"""测试：月份类型判断 + 月份编码（规格 78 节 22-26）。"""
import pytest

from rules import month_rules


class TestMonthType:
    def test_july_peak(self):
        assert month_rules.month_label(7) == "大促月份"

    def test_oct_peak(self):
        assert month_rules.month_label(10) == "大促月份"

    def test_nov_peak(self):
        assert month_rules.month_label(11) == "大促月份"

    def test_dec_peak(self):
        assert month_rules.month_label(12) == "大促月份"

    def test_other_normal(self):
        for m in [1, 2, 3, 4, 5, 6, 8, 9]:
            assert month_rules.month_label(m) == "其他月份"


class TestParseMonth:
    def test_parse(self):
        ps = month_rules.parse_month("2026-07")
        assert ps.year == 2026
        assert ps.month_num == 7
        assert ps.month_code == "202607"
        assert ps.display_month == "2026年7月"
        assert ps.month_type == "大促月份"

    def test_no_leading_zero_display(self):
        ps = month_rules.parse_month("2026-07")
        assert ps.display_month == "2026年7月"  # 不是 2026年07月

    def test_invalid(self):
        with pytest.raises(ValueError):
            month_rules.parse_month("2026/07")
