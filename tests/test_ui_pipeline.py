"""UI 路径集成测试：与 Streamlit「生成」按钮调用完全相同的 entry point（run_pipeline），
最终验证真实 ZIP 的物理内容（不检查内存变量）。

覆盖：COM 线程初始化后截图成功、email_images 存在、HTML 3 图 base64 内嵌、无错误字符串、
最终包使用正式文件名。
"""
import os
import zipfile

import pytest

from conftest import JULY_BASE, build_july_ps
from pipeline import run_pipeline
from utils import excel_screenshot
from utils.config_loader import load_config

_BASE = JULY_BASE


@pytest.mark.skipif(not excel_screenshot.com_available(), reason="本机无 Excel COM")
def test_ui_pipeline_final_zip(tmp_path):
    os_path = os.path.join(_BASE, "HUION_US_(Advertiser)_Order_Summary_20260810015543.xlsx")
    rec_path = os.path.join(_BASE, "HUION-US-ENTITY-对账单-自用.xlsx")
    sd_path = os.path.join(_BASE, "DSP花费明细-202607.xlsx")
    sr_path = os.path.join(_BASE, "亚马逊美国DSP店铺花费占比-2026-07.xls")
    invoice_path = os.path.join(_BASE, "Invoice_M001098_HUION GLOBAL (HK) LIMITED.pdf")
    meeting_path = r"C:\Users\41250\Desktop\2026年7月Huion-US对账单\2026年7月Huion-US对账单\测试\7月dsp会议记录.docx"
    if not all(os.path.exists(p) for p in (os_path, rec_path, sd_path, sr_path)):
        pytest.skip("缺少真实 2026-07 文件")

    # 提取已推广数据源（与 conftest 一致，模拟入口端）
    import openpyxl
    pa_path = os.path.join(str(tmp_path), "已推广数据源_提取.xlsx")
    wb = openpyxl.load_workbook(rec_path, data_only=True)
    ws = wb["已推广广告数据 数据源"]
    w2 = openpyxl.Workbook()
    ws2 = w2.active
    for row in ws.iter_rows(values_only=True):
        ws2.append(list(row))
    w2.save(pa_path)

    config = load_config()
    ps = build_july_ps()
    inputs = {
        "month_str": "2026-07",
        "output_dir": str(tmp_path),
        "rebate_rate": 0.08,
        "use_com": True,
        "month_confirmed": True,
        "payment_due_date": None,
        "payment_due_date_confirmed": False,
        "ai_mode": False,
        "kpi_targets": config["kpi"],
        "kpi_template_values": {"Total ROAS": 7.03, "Consideration CPDPV": 0.37, "Conversion ROAS": 10.89},
        "kpi_reasons_final": {},
        "order_summary_path": os_path,
        "promoted_ads_path": pa_path,
        "statement_path": None,
        "dsp_sa_template_path": rec_path,
        "spend_detail_template_path": sd_path,
        "spend_ratio_template_path": sr_path,
        "meeting_record_path": meeting_path,
        "invoice_path": invoice_path,
        "kpi_template_path": None,
    }
    r = run_pipeline(inputs, config)
    outputs = r["outputs"]

    # 无 KPI 模板 → 3 张必有截图必须成功
    screenshots = outputs["screenshots"]
    for key in ("dsp", "onsite", "spend_ratio"):
        assert screenshots[key][1] == "success", f"{key} 截图失败：{screenshots[key][2]}"

    # 最终 ZIP 物理内容
    zip_path = outputs["zip"]
    assert os.path.exists(zip_path)
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
    for png in ("email_images/01_DSP广告效果.png", "email_images/03_站内广告效果.png",
                "email_images/04_店铺花费占比.png"):
        assert png in names, f"ZIP 缺少 {png}"
    assert any(n.endswith(".html") for n in names)
    assert any(n.endswith(".txt") for n in names)

    # HTML：3 张图 base64 内嵌 + 无错误字符串
    html = open(outputs["email_html"], encoding="utf-8").read()
    assert html.count("<img") == 3
    assert html.count("data:image/png;base64,") == 3
    for bad in ("Pivot刷新失败", "截图失败", "被呼叫方拒绝接收呼叫", "Open.RefreshAll",
                "Traceback", "COM Error"):
        assert bad not in html, f"HTML 含错误字符串：{bad}"

    # 截图成功 → 最终包用正式文件名
    assert outputs.get("final_ok") is True
    assert os.path.basename(zip_path) == "DSP付款申请_202607.zip"
