"""测试：Excel 截图（最终阶段）——COM 逻辑可 Mock，截图文件存在，邮件/ZIP 集成。"""
import os
import zipfile

import pytest

from utils import excel_screenshot


# ---------------------------------------------------------------------------
# COM 逻辑 Mock 测试
# ---------------------------------------------------------------------------
def test_wait_calculation_done():
    class App:
        CalculationState = 0  # xlDone
    assert excel_screenshot.wait_calculation(App(), timeout=1) is True


def test_wait_calculation_timeout():
    class App:
        CalculationState = 1  # 永远在计算
    assert excel_screenshot.wait_calculation(App(), timeout=0.2) is False


def test_refresh_workbook_calls_refresh_and_calc():
    calls = {"refresh": 0, "calc": 0}

    class App:
        CalculationState = 0
        def CalculateFullRebuild(self):
            calls["calc"] += 1

    class Wb:
        def RefreshAll(self):
            calls["refresh"] += 1

    excel_screenshot.refresh_workbook(App(), Wb())
    assert calls["refresh"] == 1
    assert calls["calc"] == 1


def test_refresh_workbook_timeout_raises():
    class App:
        CalculationState = 1
        def CalculateFullRebuild(self):
            pass
    class Wb:
        def RefreshAll(self):
            pass
    with pytest.raises(TimeoutError):
        excel_screenshot.refresh_workbook(App(), Wb(), timeout=0.2)


def test_generate_all_quits_even_on_error(monkeypatch, tmp_path):
    """截图过程抛异常时，finally 仍必须 Quit 释放 EXCEL.EXE。"""
    class App:
        Ready = True
        def __init__(self):
            self.quitted = False
        def Quit(self):
            self.quitted = True

    app = App()
    monkeypatch.setattr(excel_screenshot, "com_available", lambda: True)
    monkeypatch.setattr(excel_screenshot, "_dispatch", lambda: app)

    def boom(*a, **k):
        raise RuntimeError("boom")
    monkeypatch.setattr(excel_screenshot, "_screenshot_reconciliation", boom)

    rec = tmp_path / "rec.xlsx"
    rec.write_bytes(b"x")
    with pytest.raises(RuntimeError):
        excel_screenshot.generate_all_screenshots(str(rec), None, None, False, str(tmp_path), {})
    assert app.quitted is True


def test_find_range_dynamic():
    """动态定位：Strategy 表头 → 总计结束行，结束列取表头最后一个非空列。"""
    class _Rng:
        Row = 1
        Column = 1
        Value = (("Strategy", "曝光", "总投资回报率"),
                 ("Consideration-引流", 1, 2),
                 ("Conversion-转化", 1, 2),
                 ("总计", 1, 2))
    class _Ws:
        def __init__(self):
            self.UsedRange = _Rng()
    assert excel_screenshot.find_range(_Ws(), "Strategy", ("总计", "Grand Total")) == "A1:C4"


def test_find_range_offsets_when_usedrange_not_from_a1():
    """UsedRange 顶部有空行时（Row/Column 偏移），必须换算回真实坐标，不能 off-by-one。"""
    class _Rng:
        Row = 2
        Column = 1
        Value = (("DSP",), ("Strategy", "曝光", "总投资回报率"),
                 ("Consideration-引流", 1, 2),
                 ("Conversion-转化", 1, 2),
                 ("总计", 1, 2))
    class _Ws:
        def __init__(self):
            self.UsedRange = _Rng()
    # 数组里 Strategy 是第 2 行，但真实 Sheet 是第 3 行；总计数组第 5 行 → 真实第 6 行
    assert excel_screenshot.find_range(_Ws(), "Strategy", ("总计", "Grand Total")) == "A3:C6"


def test_missing_labels():
    class _Rng:
        Value = (("类型",), ("SB",), ("SBV",), ("SD",), ("SP",), ("总计",))
    class _Ws:
        def __init__(self):
            self.UsedRange = _Rng()
    assert excel_screenshot.missing_labels(_Ws(), ("SB", "SBV", "SD", "SP", "总计"), col=1) == []


def test_missing_labels_detects_missing():
    class _Rng:
        Value = (("类型",), ("SB",), ("总计",))  # 缺 SBV/SD/SP
    class _Ws:
        def __init__(self):
            self.UsedRange = _Rng()
    miss = excel_screenshot.missing_labels(_Ws(), ("SB", "SBV", "SD", "SP", "总计"), col=1)
    assert "SBV" in miss and "SD" in miss and "SP" in miss


# ---------------------------------------------------------------------------
# 真实 COM 截图（本机无 Excel 时跳过）
# ---------------------------------------------------------------------------
@pytest.mark.skipif(not excel_screenshot.com_available(), reason="本机无 Excel COM")
def test_real_screenshot_dsp_and_onsite(july_reconciliation, tmp_path):
    from conftest import build_july_ps
    ps = build_july_ps()
    res = excel_screenshot.generate_all_screenshots(
        july_reconciliation, None, None, False, str(tmp_path), {},
        expected_totals={"dsp": float(ps.dsp_cost), "onsite": float(ps.onsite_cost)})
    dsp_path, dsp_status, dsp_msg = res["dsp"]
    onsite_path, onsite_status, onsite_msg = res["onsite"]
    assert dsp_status == "success", dsp_msg
    assert onsite_status == "success", onsite_msg
    assert os.path.exists(dsp_path) and os.path.getsize(dsp_path) > 0
    assert os.path.exists(onsite_path) and os.path.getsize(onsite_path) > 0


# ---------------------------------------------------------------------------
# PNG 内容校验（空白图必须判失败）
# ---------------------------------------------------------------------------
def test_validate_png_content_detects_blank(tmp_path):
    from PIL import Image
    # 模拟空白图：白底 + 细边框（与实际空白截图一致）
    img = Image.new("RGB", (100, 50), "white")
    for x in range(100):
        img.putpixel((x, 0), (0, 0, 0))
        img.putpixel((x, 49), (0, 0, 0))
    for y in range(50):
        img.putpixel((0, y), (0, 0, 0))
        img.putpixel((99, y), (0, 0, 0))
    p = tmp_path / "blank.png"
    img.save(str(p))
    ok, info = excel_screenshot.validate_png_content(str(p))
    assert not ok, f"空白图必须判失败，却返回 ok：{info}"


def test_validate_png_content_accepts_real(tmp_path):
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (200, 100), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([10, 10, 190, 90], outline="black")
    d.text((20, 40), "Strategy Consideration Conversion", fill="black")
    for i in range(50):
        d.rectangle([i * 4, 60, i * 4 + 3, 80], fill=(i * 5 % 256, i * 7 % 256, i * 11 % 256))
    p = tmp_path / "real.png"
    img.save(str(p))
    ok, info = excel_screenshot.validate_png_content(str(p))
    assert ok, f"真实表格图必须通过，却判失败：{info}"


# ---------------------------------------------------------------------------
# ZIP 集成
# ---------------------------------------------------------------------------
def test_zip_includes_email_images(tmp_path):
    from processors import package_processor
    from models.payment_summary import PaymentSummary
    ps = PaymentSummary(month_code="202607")
    png = tmp_path / "01_DSP广告效果.png"
    png.write_bytes(b"\x89PNG\r\n\x1a\n" + b"x" * 20)
    zip_path = package_processor.build_zip(str(tmp_path), [], ps, image_files=[str(png)])
    with zipfile.ZipFile(zip_path) as z:
        assert "email_images/01_DSP广告效果.png" in z.namelist()
