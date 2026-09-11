"""测试：DSP Strategy 识别 + 字段映射（规格 78 节 1-6）。"""
from rules import dsp_mapping


class TestStrategy:
    def test_consideration(self):
        assert dsp_mapping.detect_strategy("HUION_混合_Consideration_CPDPV_2025") == "Consideration"

    def test_conversion(self):
        assert dsp_mapping.detect_strategy("HUION_混合_Conversion_ROAS_2025") == "Conversion"

    def test_unrecognized(self):
        assert dsp_mapping.detect_strategy("HUION_混合_Unknown_2025") == ""

    def test_empty(self):
        assert dsp_mapping.detect_strategy("") == ""

    def test_case_insensitive(self):
        assert dsp_mapping.detect_strategy("huion_混合_consideration_cpdpv") == "Consideration"


class TestFieldMapping:
    HEADER = ["Interval start", "Interval end", "Campaign name", "Campaign ID", "Total cost",
              "Impressions", "eCPM", "Avg. impression frequency", "Click-throughs", "CTR",
              "Purchases", "Units sold", "Sales USD", "ROAS", "Total DPV", "Total DPVR",
              "Total ATC", "Total ATCR", "Total purchases", "Total purchase rate",
              "Total units sold", "Total product sales", "Total ROAS"]

    def test_mapping(self):
        m = dsp_mapping.build_field_map(self.HEADER)
        assert m["order"] == 2            # Campaign name
        assert m["order_id"] == 3         # Campaign ID
        assert m["total_cost"] == 4
        assert m["impressions"] == 5
        assert m["clicks"] == 8           # Click-throughs
        assert m["ctr"] == 9
        assert m["promoted_sales"] == 12  # Sales USD
        assert m["promoted_roas"] == 13   # ROAS
        assert m["total_dpv"] == 14
        assert m["total_atc"] == 16
        assert m["total_purchases"] == 18
        assert m["total_units_sold"] == 20
        assert m["total_sales"] == 21     # Total product sales
        assert m["total_roas"] == 22

    def test_missing_field_is_none(self):
        m = dsp_mapping.build_field_map(["Total cost", "Impressions"])
        assert m["order"] is None

    def test_find_header_row(self):
        rows = [["junk"], ["Interval start", "Total cost", "Impressions", "ROAS"], ["x"]]
        assert dsp_mapping.find_header_row(rows) == 1
