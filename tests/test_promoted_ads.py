"""测试：广告类型转换 + SD 销售额修正（规格 78 节 7-15）。"""
from models.payment_summary import PromotedAd
from rules import promoted_ads_mapping as pm


class TestTypeConversion:
    def test_sp(self):
        assert pm.convert_type("SP") == "SP"

    def test_sd(self):
        assert pm.convert_type("SD") == "SD"

    def test_sb2_to_sbv(self):
        assert pm.convert_type("SB2", "SBV-RES-品牌词-精准") == "SBV"

    def test_sb2_to_sb(self):
        assert pm.convert_type("SB2", "引流-Brand-SB-产品词") == "SB"

    def test_sb2_by_portfolio(self):
        assert pm.convert_type("SB2", "", "MK-SBV") == "SBV"

    def test_no_sb2_left(self):
        # 任何 SB2 都被转换为 SB 或 SBV
        for name in ["SBV-x", "普通SB", "brand SB 引流"]:
            assert pm.convert_type("SB2", name) in ("SB", "SBV")


class TestSDRule:
    def _ad(self, t, sales, direct):
        return PromotedAd(ad_type=t, sales=sales, direct_sales=direct)

    def test_sd_overwritten(self):
        ad = self._ad("SD", 100.0, 80.0)
        pm.apply_sd_rule(ad)
        assert ad.sales == 80.0

    def test_sp_unaffected(self):
        ad = self._ad("SP", 100.0, 80.0)
        pm.apply_sd_rule(ad)
        assert ad.sales == 100.0

    def test_sb_unaffected(self):
        ad = self._ad("SB", 100.0, 80.0)
        pm.apply_sd_rule(ad)
        assert ad.sales == 100.0

    def test_sbv_unaffected(self):
        ad = self._ad("SBV", 100.0, 80.0)
        pm.apply_sd_rule(ad)
        assert ad.sales == 100.0


class TestFieldMap:
    def test_source_format(self):
        h = ["类型", "广告组合名称", "广告活动名称", "广告组", "曝光量", "点击",
             "花费-本币", "广告销售额-本币", "直接销售额-本币", "广告订单"]
        m = pm.build_field_map(h)
        assert m["type"] == 0
        assert m["sales"] == 7           # 广告销售额-本币
        assert m["direct_sales"] == 8    # 直接销售额-本币
        assert m["orders"] == 9

    def test_target_format(self):
        h = ["店铺名称", "国家", "类型", "广告组合", "广告活动", "广告组", "广告组投放类型",
             "ASIN", "MSKU", "广告有效状态", "曝光量", "点击", "花费-本币",
             "销售额-本币", "广告订单", "直接成交销售额-本币"]
        m = pm.build_field_map(h)
        assert m["type"] == 2
        assert m["sales"] == 13          # 销售额-本币（兼容目标格式）
        assert m["direct_sales"] == 15
