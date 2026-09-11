"""付款申请邮件生成。只生成，不发送。

最终阶段：正文中的 DSP 广告效果 / KPI / 站内广告效果 / 店铺花费占比 不再重绘 HTML/Markdown
表格，而是直接内嵌最终 Excel 对应区域的截图（PNG）。

输出：
- html：主版本，图片用 base64 内嵌（<img src="data:image/png;base64,...">），单独打开也可完整预览；
- text：纯文本，数据区域用 [xxx截图：文件名] 占位，不再输出大量数据表。
"""
from __future__ import annotations

import base64
import os
from typing import Dict, List

from models.payment_summary import PaymentSummary


def _num(x, nd=2) -> str:
    return f"{x:,.{nd}f}"


def _int(x) -> str:
    return f"{int(round(x)):,}"


def _money(x) -> str:
    return f"${_num(x)} USD"


def _pct(x, nd=2) -> str:
    """比率 -> 百分数字符串：0.0099 -> 0.99%。"""
    return f"{x * 100:.{nd}f}%"


def subject(ps: PaymentSummary) -> str:
    return f"SparkX-HUION_US-AMAZON DSP_{ps.year}_{ps.month_num}月 付款申请"


# 截图槽位：key -> (输出文件名, 中文描述)
IMAGE_META = {
    "dsp": ("01_DSP广告效果.png", "DSP广告效果截图"),
    "kpi": ("02_DSP_KPI.png", "DSP KPI截图"),
    "onsite": ("03_站内广告效果.png", "站内广告效果截图"),
    "spend_ratio": ("04_店铺花费占比.png", "店铺花费占比截图"),
}


def _dsp_effect(ps: PaymentSummary) -> str:
    return (
        "累计曝光次数 " + _int(ps.dsp_impressions) + " 次，"
        "累计费用 " + _money(ps.dsp_cost) + "，\n"
        "带来总销售额是 " + _money(ps.dsp_total_sales) + "，\n"
        "总投资回报率为 " + _num(ps.dsp_roas) + "，详情如下："
    )


def _kpi_heading(ps: PaymentSummary) -> str:
    return f"{ps.month_num}月DSP KPI 完成情况如下："


def _kpi_table(ps: PaymentSummary) -> tuple:
    """KPI 表格数据（保留供测试/调试；正式邮件正文已改用截图）。"""
    headers = ["KPI", "目标", "实际", "状态"]
    rows = []
    for r in ps.kpi_results:
        rows.append([r.name, f"{r.target:g}", _num(r.final), "达标" if r.is_met else "未达标"])
    return headers, rows


def _onsite_effect(ps: PaymentSummary) -> str:
    return (
        "累计曝光次数 " + _int(ps.onsite_impressions) + " 次，\n"
        "累计费用 " + _money(ps.onsite_cost) + "，\n"
        "带来的销售额是 " + _money(ps.onsite_sales) + "，\n"
        "投资回报率为 " + _num(ps.onsite_roas) + "，详情如下："
    )


def _combined_effect(ps: PaymentSummary) -> str:
    return (
        f"{ps.month_num}月累计费用 {_money(ps.combined_cost)}，\n"
        f"带来的销售额是 {_money(ps.combined_sales)}，\n"
        f"投资回报率为 {_num(ps.combined_roas)}。"
    )


def _rebate_section(ps: PaymentSummary) -> str:
    rate = ps.rebate_rate
    return (
        f"{ps.month_num}月费用：{_money(ps.dsp_cost)}\n\n"
        "返点金额：\n"
        f"{rate:g} × {_money(ps.dsp_cost)} = {_money(ps.rebate_amount)}\n\n"
        "账单金额：\n"
        f"{_money(ps.dsp_cost)} - {_money(ps.rebate_amount)} = {_money(ps.payment_amount)}"
    )


def _attachments(ps: PaymentSummary) -> str:
    lines = ["1. 对账单", "2. DSP花费明细"]
    if ps.has_invoice:
        lines.append("3. Invoice发票")
    else:
        lines.append("3. Invoice发票（⚠ 缺少发票）")
    lines.append("4. 亚马逊美国DSP店铺花费占比")
    return "\n".join(lines)


def _payee_info(ps: PaymentSummary) -> str:
    p = ps.payee or {}
    return (
        f"开户名称：{p.get('name', '')}\n\n"
        f"开户银行：\n{p.get('bank', '')}\n\n"
        f"开户账号：\n{p.get('account', '')}\n\n"
        f"银行地址：\n{p.get('bank_address', '')}\n\n"
        f"SWIFT代码：\n{p.get('swift', '')}\n\n"
        f"HSBC银行代码：\n{p.get('bank_code', '')}"
    )


# ---------------------------------------------------------------------------
# 内容块：text / heading / image
# image 块 = ("image", desc, filename, path_or_None, message)
# ---------------------------------------------------------------------------

def _blocks(ps: PaymentSummary, screenshots: Dict) -> List:
    """按历史邮件模板顺序组装内容块（数据区域用截图）。"""
    screenshots = screenshots or {}
    due = ps.payment_due_date if ps.payment_due_date_confirmed else "待确认"

    def img(key):
        filename, desc = IMAGE_META[key]
        v = screenshots.get(key)
        path, status, msg = (v if v else (None, "", ""))
        return ("image", desc, filename, path, msg)

    blocks: List = []
    blocks.append(("text", "Dear all,"))
    blocks.append(("text", f"附件为美国亚马逊DSP广告{ps.year}年{ps.month_num}月的发票以及对账单，请查收。数据核对无误，麻烦安排支付。"))
    blocks.append(("text", "附件：\n" + _attachments(ps)))
    blocks.append(("heading", "DSP广告效果"))
    blocks.append(("text", _dsp_effect(ps)))
    blocks.append(img("dsp"))
    blocks.append(("heading", "DSP KPI完成情况"))
    blocks.append(("text", _kpi_heading(ps)))
    blocks.append(img("kpi"))
    blocks.append(("heading", "站内广告效果"))
    blocks.append(("text", _onsite_effect(ps)))
    blocks.append(img("onsite"))
    blocks.append(("heading", "DSP+站内广告效果"))
    blocks.append(("text", _combined_effect(ps)))
    blocks.append(("heading", "返点情况"))
    blocks.append(("text", _rebate_section(ps)))
    blocks.append(("text", f"申请原因：HUION美国亚马逊DSP {ps.year}-{ps.month_num}月广告花费"))
    blocks.append(("text", f"付款方：{ps.payer}"))
    blocks.append(("text", f"付款金额：{_money(ps.payment_amount)}"))
    blocks.append(("text", f"截止付款时间：{due}"))
    blocks.append(("heading", "收款方信息"))
    blocks.append(("text", _payee_info(ps)))
    blocks.append(("heading", "店铺花费占比"))
    blocks.append(("text", "店铺花费占比情况如下："))
    blocks.append(img("spend_ratio"))
    return blocks


def _image_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def _render_text(blocks: List) -> str:
    out = []
    for b in blocks:
        kind = b[0]
        if kind == "text":
            out.append(b[1])
        elif kind == "heading":
            out.append(b[1] + "：")
        elif kind == "image":
            _kind, desc, filename, path, msg = b
            # TXT 固定用文件名占位（规格 26 节）；是否生成由 HTML/页面提示
            out.append(f"[{desc}：{filename}]")
    return "\n\n".join(out)


def _render_html(blocks: List) -> str:
    import html as _html
    parts = []
    for b in blocks:
        kind = b[0]
        if kind == "text":
            parts.append("<p>" + _html.escape(b[1]).replace("\n", "<br>") + "</p>")
        elif kind == "heading":
            parts.append(f"<h3>{_html.escape(b[1])}</h3>")
        elif kind == "image":
            _kind, desc, filename, path, msg = b
            if path and os.path.exists(path):
                b64 = _image_base64(path)
                parts.append(
                    f'<img src="data:image/png;base64,{b64}" '
                    f'alt="{_html.escape(desc)}" style="max-width:100%;height:auto;border:1px solid #ccc;">'
                )
            else:
                note = msg or "未生成"
                parts.append(f'<p style="color:#b00020;">⚠ {_html.escape(desc)}：{_html.escape(note)}</p>')
    return (
        "<html><head><meta charset='utf-8'></head><body style='font-family:Arial,sans-serif;font-size:14px;'>"
        + "\n".join(parts) + "</body></html>"
    )


def _email_status(ps: PaymentSummary, screenshots: Dict) -> tuple:
    """邮件状态：final / draft，并给草稿原因。"""
    screenshots = screenshots or {}
    notes: List[str] = []

    for key, (filename, desc) in IMAGE_META.items():
        v = screenshots.get(key)
        if not v or not v[0] or not os.path.exists(v[0]):
            notes.append(f"缺少{desc}")
    if not ps.payment_due_date_confirmed:
        notes.append("待确认截止付款日期")
    if not ps.kpi_reasons_confirmed():
        notes.append("待补KPI原因")
    if not ps.has_invoice:
        notes.append("缺少发票")
    if ps.errors:
        notes.append("存在阻断性错误")

    if not notes:
        return "final", ""
    return "draft", "、".join(notes)


def generate_email(ps: PaymentSummary, config: Dict = None, screenshots: Dict = None) -> Dict[str, str]:
    """生成完整邮件。

    screenshots = {key: (path, status, message)}，key ∈ {dsp, kpi, onsite, spend_ratio}。
    返回 {subject, to, cc, text, html, status, status_note}。
    """
    config = config or {}
    email_cfg = config.get("email", {})
    to = email_cfg.get("to", [])
    cc = email_cfg.get("cc", [])

    blocks = _blocks(ps, screenshots)
    status, status_note = _email_status(ps, screenshots)
    return {
        "subject": subject(ps),
        "to": to,
        "cc": cc,
        "text": _render_text(blocks),
        "html": _render_html(blocks),
        "status": status,
        "status_note": status_note,
    }
