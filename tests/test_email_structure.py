"""测试：邮件结构与格式（规格 38-65 节）。"""
from conftest import build_july_ps
from processors import email_processor


def _email():
    ps = build_july_ps()
    ps.payment_due_date = "2026-08-20"
    ps.payment_due_date_confirmed = True
    return ps, email_processor.generate_email(ps)


def test_subject():
    ps, e = _email()
    assert e["subject"] == "SparkX-HUION_US-AMAZON DSP_2026_7月 付款申请"


def test_no_leading_zero_month():
    _, e = _email()
    assert "2026_07月" not in e["subject"]
    assert "2026_7月" in e["subject"]


def test_opening():
    _, e = _email()
    assert "Dear all," in e["text"]
    assert "附件为美国亚马逊DSP广告2026年7月的发票以及对账单" in e["text"]


def test_dsp_effect_numbers():
    _, e = _email()
    assert "13,669,440" in e["text"]
    assert "$49,025.52 USD" in e["text"]
    assert "$345,000.39 USD" in e["text"]


def test_strategy_table_replaced_by_screenshot():
    """DSP Strategy 表格不再重绘，改为截图占位。"""
    _, e = _email()
    assert "[DSP广告效果截图：01_DSP广告效果.png]" in e["text"]
    assert "Consideration-引流" not in e["text"]


def test_no_sb2():
    _, e = _email()
    assert "SB2" not in e["text"]


def test_onsite_replaced_by_screenshot():
    """站内广告表格不再重绘，改为截图占位。"""
    _, e = _email()
    assert "[站内广告效果截图：03_站内广告效果.png]" in e["text"]
    assert "SBV" not in e["text"]


def test_combined_numbers():
    _, e = _email()
    assert "$194,247.92 USD" in e["text"]
    assert "$965,136.46 USD" in e["text"]


def test_rebate_section():
    _, e = _email()
    assert "返点情况" in e["text"]
    assert "0.08 × $49,025.52 USD = $3,922.04 USD" in e["text"]


def test_payment_amount_equals_summary():
    ps, e = _email()
    assert f"${ps.payment_amount:,.2f} USD" in e["text"]


def test_application_reason():
    _, e = _email()
    assert "申请原因：HUION美国亚马逊DSP 2026-7月广告花费" in e["text"]


def test_payee_info():
    _, e = _email()
    assert "SPARKX MARKETING CO., LIMITED" in e["text"]
    assert "124-787771-838" in e["text"]
    assert "HSBCHKHHHKH" in e["text"]


def test_payer():
    _, e = _email()
    assert "付款方：HUION GLOBAL (HK) LIMITED" in e["text"]


def test_no_payment_cycle_auto_date():
    """不再根据「20个工作日」自动决定付款日期（规格 13 节）。"""
    _, e = _email()
    assert "付款周期" not in e["text"]


def test_due_date_confirmed():
    _, e = _email()  # _email() 里已设 payment_due_date = 2026-08-20
    assert "截止付款时间：2026-08-20" in e["text"]


def test_due_date_unconfirmed():
    ps = build_july_ps()
    ps.payment_due_date = ""
    e = email_processor.generate_email(ps)
    assert "截止付款时间：待确认" in e["text"]


def test_html_no_business_table():
    """HTML 不再出现业务 table（DSP/KPI/站内都改截图）。"""
    _, e = _email()
    assert "<table" not in e["html"]


def test_txt_screenshot_placeholders():
    """TXT 里 4 张截图均为占位行。"""
    _, e = _email()
    for s in ("01_DSP广告效果.png", "02_DSP_KPI.png", "03_站内广告效果.png", "04_店铺花费占比.png"):
        assert s in e["text"]


def test_html_embeds_base64_image(tmp_path):
    """提供截图后，HTML 用 base64 内嵌 <img>。"""
    ps = build_july_ps()
    png = tmp_path / "01_DSP广告效果.png"
    png.write_bytes(b"\x89PNG\r\n\x1a\n" + b"x" * 20)
    e = email_processor.generate_email(ps, screenshots={"dsp": (str(png), "success", "")})
    assert "<img" in e["html"]
    assert "data:image/png;base64," in e["html"]


def test_missing_invoice_warning():
    ps = build_july_ps()
    ps.has_invoice = False
    e = email_processor.generate_email(ps)
    assert "⚠ 缺少发票" in e["text"]
