"""DSP 月度付款申请助手 — Streamlit 主页面。"""
from __future__ import annotations

import datetime
import os
import uuid

import streamlit as st

from utils.config_loader import load_config, validate_business_config
from pipeline import run_pipeline
from processors import email_processor, package_processor, validation_processor

st.set_page_config(page_title="DSP 月度付款申请助手", layout="wide")

WORKDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(WORKDIR, exist_ok=True)


def _new_run_dir() -> str:
    """每次运行创建独立目录，避免旧/新结果混用。"""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    rid = uuid.uuid4().hex[:6]
    d = os.path.join(WORKDIR, "runs", f"{ts}_{rid}")
    os.makedirs(d, exist_ok=True)
    return d


def save_uploaded(uploaded, run_dir: str, subdir: str) -> str:
    d = os.path.join(run_dir, subdir)
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, uploaded.name)
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    return path


def _parse_kpi_final_values(final_roas: str, final_cpdpv: str, final_conv_roas: str) -> dict:
    """把可选的 KPI 最终汇报值解析为 dict（空/非法则跳过该项）。"""
    out = {}
    mapping = {
        "Total ROAS": final_roas,
        "Consideration CPDPV": final_cpdpv,
        "Conversion ROAS": final_conv_roas,
    }
    for name, raw in mapping.items():
        s = (raw or "").strip()
        if not s:
            continue
        try:
            out[name] = float(s)
        except ValueError:
            continue
    return out


@st.cache_data(show_spinner=False)
def _config():
    return load_config()


config = _config()

st.title("DSP 月度付款申请助手")

# 业务配置校验：示例/占位值只能测试，不能生成正式付款申请
_config_problems = validate_business_config(config)
if _config_problems:
    st.warning(
        "⚠ 尚未配置真实付款信息。当前正在使用示例配置，请先创建 config.local.yaml "
        "并填写真实付款/收款信息。（配置未完成 / 测试模式）"
    )
    with st.expander("配置问题明细"):
        for p in _config_problems:
            st.text(f"• {p}")

# ================= Step 1：申请月份 =================
st.header("Step 1 · 申请月份")
c1, c2 = st.columns(2)
with c1:
    month_str = st.selectbox(
        "申请月份 (YYYY-MM)",
        options=[f"{y}-{m:02d}" for y in (2025, 2026, 2027) for m in range(1, 13)],
        index=[f"{y}-{m:02d}" for y in (2025, 2026, 2027) for m in range(1, 13)].index("2026-08"),
    )
with c2:
    st.caption("大促月：7 / 10 / 11 / 12；其余为其他月份。")

# ================= Step 2：上传文件 =================
st.header("Step 2 · 上传文件")

col_a, col_b = st.columns(2)
with col_a:
    order_summary = st.file_uploader("A. Amazon DSP Order Summary（*.xlsx）", type=["xlsx"], key="os")
    promoted_ads = st.file_uploader("B. 已推广广告数据（已推广-汇总*.xlsx）", type=["xlsx"], key="pa")
    statement = st.file_uploader("C. 付款对账单（*.xlsx，用于核对DSP费用/返点/付款金额）", type=["xlsx"], key="stmt")
    dsp_sa_template = st.file_uploader("D. DSP+SA 分析模板（可选，*.xlsx）", type=["xlsx"], key="rec")
with col_b:
    spend_detail = st.file_uploader("E. DSP花费明细模板（*.xlsx）", type=["xlsx"], key="sd")
    spend_ratio = st.file_uploader("F. 店铺花费占比模板（*.xls / *.xlsx）", type=["xls", "xlsx"], key="sr")
    meeting_record = st.file_uploader("G. DSP会议记录（*.docx）", type=["docx"], key="mr")
    invoice = st.file_uploader("H. 发票 Invoice（*.pdf）", type=["pdf"], key="inv")
    kpi_template = st.file_uploader("I. KPI 模板（*.xlsx，可选；未上传不生成 KPI Excel）", type=["xlsx"], key="kpit")

st.caption("说明：DSP数据源由 Order Summary 自动生成，站内数据源由已推广-汇总生成；"
           "付款对账单只用于核对最终付款金额，不要求包含 DSP数据源 Sheet。"
           "DSP+SA 分析模板为可选项，用于写入数据源并刷新 Pivot 报表。")

# ================= 配置 =================
with st.expander("⚙ 配置（KPI 目标 / 返点 / 邮件 / 截止日期）", expanded=True):
    cc1, cc2, cc3, cc4 = st.columns(4)
    with cc1:
        total_roas_t = st.number_input("Total ROAS 目标", value=float(config["kpi"]["total_roas"]), step=0.1)
        consideration_cpdpv_t = st.number_input("Consideration CPDPV 目标", value=float(config["kpi"]["consideration_cpdpv"]), step=0.01)
    with cc2:
        conversion_roas_t = st.number_input("Conversion ROAS 目标", value=float(config["kpi"]["conversion_roas"]), step=0.1)
        rebate_rate = st.number_input("返点比例", value=float(config["payment"]["rebate_rate"]), step=0.01, min_value=0.0, max_value=1.0)
    with cc3:
        # 截止付款日期无默认值，初始为 None/空；未确认时邮件显示「待确认」
        if "payment_due_date" not in st.session_state:
            st.session_state["payment_due_date"] = ""
        payment_due_date = st.text_input("截止付款日期（用户确认）", key="payment_due_date", placeholder="如 2026-09-15")
        ai_mode = st.checkbox("KPI 原因 AI 总结模式（需在 .env 配置 API Key）", value=False)
        month_confirmed = st.checkbox("我确认所有数据文件为该月完整月数据（月份无法自动验证/非完整月时勾选）", value=False)
    with cc4:
        st.write("收件人：", ", ".join(config["email"]["to"]))
        st.write("抄送：", ", ".join(config["email"]["cc"]))

with st.expander("🎯 KPI 最终汇报值（可选，模板/人工确认；留空则用程序计算值）", expanded=False):
    st.caption("正式 KPI 表/邮件/页面使用这里的「最终汇报值」。程序会另行计算并提示差异，不擅自覆盖。")
    kc1, kc2, kc3 = st.columns(3)
    with kc1:
        final_roas = st.text_input("Total ROAS 最终汇报值", value="", placeholder="如 7.03（留空=计算值）")
    with kc2:
        final_cpdpv = st.text_input("Consideration CPDPV 最终汇报值", value="", placeholder="如 0.37（留空=计算值）")
    with kc3:
        final_conv_roas = st.text_input("Conversion ROAS 最终汇报值", value="", placeholder="如 10.89（留空=计算值）")

# ================= 生成 =================
st.divider()
if st.button("生成本月 DSP 付款申请资料", type="primary", use_container_width=True):
    with st.spinner("处理中……"):
        run_dir = _new_run_dir()  # 每次运行独立目录，禁止复用旧 output
        inputs = {
            "month_str": month_str,
            "output_dir": run_dir,
            "rebate_rate": rebate_rate,
            "payment_due_date": payment_due_date.strip() or None,
            "payment_due_date_confirmed": bool(payment_due_date.strip()),
            "ai_mode": ai_mode,
            "month_confirmed": month_confirmed,
            "kpi_targets": {
                "total_roas": total_roas_t,
                "consideration_cpdpv": consideration_cpdpv_t,
                "conversion_roas": conversion_roas_t,
            },
            "kpi_template_values": _parse_kpi_final_values(final_roas, final_cpdpv, final_conv_roas),
            "order_summary_path": save_uploaded(order_summary, run_dir, "uploads") if order_summary else None,
            "promoted_ads_path": save_uploaded(promoted_ads, run_dir, "uploads") if promoted_ads else None,
            "statement_path": save_uploaded(statement, run_dir, "uploads") if statement else None,
            "dsp_sa_template_path": save_uploaded(dsp_sa_template, run_dir, "uploads") if dsp_sa_template else None,
            "spend_detail_template_path": save_uploaded(spend_detail, run_dir, "uploads") if spend_detail else None,
            "spend_ratio_template_path": save_uploaded(spend_ratio, run_dir, "uploads") if spend_ratio else None,
            "meeting_record_path": save_uploaded(meeting_record, run_dir, "uploads") if meeting_record else None,
            "invoice_path": save_uploaded(invoice, run_dir, "uploads") if invoice else None,
            "kpi_template_path": save_uploaded(kpi_template, run_dir, "uploads") if kpi_template else None,
        }
        try:
            result = run_pipeline(inputs, config)
            st.session_state["result"] = result
            st.session_state["run_dir"] = run_dir
        except Exception as e:
            st.error(f"处理失败：{e}")
            st.session_state["result"] = None

result = st.session_state.get("result")
if result is None:
    st.info("上传文件后点击「生成本月 DSP 付款申请资料」。")
    st.stop()

ps = result["ps"]
outputs = result["outputs"]

# ================= 运行日志 =================
st.header("运行日志")
for line in result["logs"]:
    st.text(line)

# ================= 付款申请摘要 =================
st.header("付款申请摘要")
sc1, sc2, sc3 = st.columns(3)
with sc1:
    st.metric("申请月份", ps.month)
    st.metric("月份类型", ps.month_type)
    st.metric("DSP 曝光", f"{ps.dsp_impressions:,.0f}")
    st.metric("DSP 费用", f"${ps.dsp_cost:,.2f}")
    st.metric("DSP 销售额", f"${ps.dsp_total_sales:,.2f}")
    st.metric("DSP ROAS", f"{ps.dsp_roas:.2f}")
with sc2:
    st.metric("站内花费", f"${ps.onsite_cost:,.2f}")
    st.metric("站内销售额", f"${ps.onsite_sales:,.2f}")
    st.metric("站内 ROAS", f"{ps.onsite_roas:.2f}")
    st.metric("DSP+站内花费", f"${ps.combined_cost:,.2f}")
    st.metric("DSP+站内销售额", f"${ps.combined_sales:,.2f}")
    st.metric("DSP+站内 ROAS", f"{ps.combined_roas:.2f}")
with sc3:
    st.metric("返点比例", f"{ps.rebate_rate:.0%}")
    st.metric("返点金额", f"${ps.rebate_amount:,.2f}")
    st.metric("最终付款金额", f"${ps.payment_amount:,.2f}")
    st.metric("截止付款日期", ps.payment_due_date or "⚠ 未确认")

# ================= 金额口径核对 =================
st.header("金额口径")
st.warning("DSP原始费用 ≠ 最终付款金额")
amt_cols = st.columns(3)
amt_cols[0].metric("DSP原始费用", f"${ps.dsp_cost:,.2f}")
amt_cols[1].metric("返点比例", f"{ps.rebate_rate:.0%}")
amt_cols[2].metric("返点金额", f"${ps.rebate_amount:,.2f}")
st.metric("最终付款金额", f"${ps.payment_amount:,.2f}")

st.subheader("每个字段最终写入金额")
bindings = outputs.get("amount_bindings", [])
rows = []
for b in bindings:
    rows.append({
        "字段": b["field"],
        "金额语义": b["description"],
        "写入值": f"${b['value']:,.2f}",
    })
st.dataframe(rows, use_container_width=True, hide_index=True)

# ================= KPI =================
st.header("KPI 完成情况")
kpi_rows = []
for r in ps.kpi_results:
    kpi_rows.append({
        "KPI": r.name,
        "目标": r.target,
        "最终汇报值": r.final,
        "参考计算值": round(r.calculated, 4),
        "状态": "达标" if r.is_met else "未达标",
        "口径": {"no_template": "无模板值，用计算值", "consistent": "口径基本一致",
                "inconsistent": "⚠ 口径存在差异，请人工确认"}.get(r.diff_status, ""),
    })
st.dataframe(kpi_rows, use_container_width=True)

st.header("KPI 原因")
mode = outputs.get("kpi_reason_mode", "none")
if ps.has_meeting_record and mode == "ai":
    st.success("AI 已生成提议原因，可编辑后确认。")
elif ps.has_meeting_record:
    st.warning("⚠ KPI原因尚未生成（AI 不可用），请手动填写后确认。")
else:
    st.info("未上传 DSP 会议记录，需手动填写 KPI 原因。")

by_name = {r.name: r for r in ps.kpi_results}
reason_inputs = {}
rc1, rc2, rc3 = st.columns(3)
with rc1:
    r = by_name.get("Total ROAS")
    reason_inputs["Total ROAS"] = st.text_area(
        "Total ROAS 原因", value=ps.kpi_reason_final.get("Total ROAS") or (r.reason if r else ""),
        height=170, key="reason_roas")
with rc2:
    r = by_name.get("Consideration CPDPV")
    reason_inputs["Consideration CPDPV"] = st.text_area(
        "Consideration CPDPV 原因", value=ps.kpi_reason_final.get("Consideration CPDPV") or (r.reason if r else ""),
        height=170, key="reason_cpdpv")
with rc3:
    r = by_name.get("Conversion ROAS")
    reason_inputs["Conversion ROAS"] = st.text_area(
        "Conversion ROAS 原因", value=ps.kpi_reason_final.get("Conversion ROAS") or (r.reason if r else ""),
        height=170, key="reason_conv")

if st.button("确认KPI原因（写入正式邮件）", key="confirm_reasons"):
    ps.kpi_reason_final = {k: v.strip() for k, v in reason_inputs.items()}
    _email = email_processor.generate_email(ps, config)
    outputs["email"] = _email
    email_html = os.path.join(WORKDIR, f"DSP付款申请邮件-{ps.month_code}.html")
    with open(email_html, "w", encoding="utf-8") as f:
        f.write(_email["html"])
    outputs["email_html"] = email_html
    email_txt = os.path.join(WORKDIR, f"DSP付款申请邮件-{ps.month_code}.txt")
    with open(email_txt, "w", encoding="utf-8") as f:
        f.write(_email["text"])
    outputs["email_txt"] = email_txt
    files = [v for k, v in outputs.items() if k in
             ("reconciliation", "spend_detail", "spend_ratio", "kpi", "email_html", "email_txt", "invoice")]
    outputs["zip"] = package_processor.build_zip(WORKDIR, files, ps)
    st.session_state["result"] = result
    st.rerun()

# ================= 邮件预览 =================
st.header("付款申请邮件预览")
email = outputs["email"]
st.subheader(f"主题：{email['subject']}")
if email.get("status") == "final":
    st.success("✅ 最终付款申请邮件")
else:
    st.warning(f"⚠ 草稿：{email.get('status_note', '')}")

# 按真实顺序：正文 / 图片 / 正文 / 图片 ……（截图用 st.image 展示）
screenshots = outputs.get("screenshots", {}) or {}
blocks = email_processor._blocks(ps, screenshots)
for b in blocks:
    kind = b[0]
    if kind == "text":
        st.markdown(b[1].replace("\n", "  \n"))
    elif kind == "heading":
        st.subheader(b[1])
    elif kind == "image":
        _kind, desc, filename, path, msg = b
        if path and os.path.exists(path):
            st.image(path, caption=filename)
            st.success(f"✅ {desc}：截图正常")
        else:
            st.warning(f"⚠ {desc}：{msg or '生成失败'}")

# 截图状态总览
st.caption("—— 截图状态 ——")
for key, (filename, desc) in email_processor.IMAGE_META.items():
    v = screenshots.get(key)
    if v and v[0] and os.path.exists(v[0]):
        st.text(f"✅ {desc}：{os.path.basename(v[0])}")
    else:
        st.text(f"⚠ {desc}：{v[2] if v else '未生成'}")
st.download_button("下载 HTML（图片已内嵌）",
                   open(outputs["email_html"], "rb"),
                   file_name=os.path.basename(outputs["email_html"]),
                   key="dl_email_html_preview")

# ================= 月份防呆 =================
st.header("月份防呆")
for mc in outputs.get("month_checks", []):
    icon = {"PASS": "✅", "WARNING": "⚠", "FAIL": "❌"}.get(mc["status"], "❓")
    st.text(f"{icon} {mc['source']}：{mc['message']}")

# ================= 校验 =================
st.header("自动校验")
checks = outputs.get("validation", [])
for ch in checks:
    icon = {"pass": "✅", "warn": "⚠", "fail": "❌"}[ch.status]
    st.text(f"{icon} {ch.name}：{ch.message}")

# 格式保护
for key, label in (("reconciliation_format", "对账单格式保护"), ("spend_detail_format", "花费明细格式保护")):
    fc = outputs.get(key)
    if fc is not None:
        if fc["ok"]:
            st.success(f"{label}：通过")
        else:
            st.error(f"{label}：失败 → {'; '.join(fc['differences'])}")

# DSP+SA 实际修改范围（仅数据源，不覆盖 Pivot/公式）
if "reconciliation_modified" in outputs:
    st.subheader("对账单实际修改范围")
    for rng in outputs["reconciliation_modified"]:
        st.text(f"✏ {rng}")
    st.caption("DSP+SA 汇总表（Pivot/GETPIVOTDATA 公式）未被程序覆盖。")

# ================= Checklist =================
st.header("DSP 付款申请材料")
checklist = package_processor.build_checklist(ps)
ready_map = {
    "对账单": "reconciliation",
    "DSP花费明细": "spend_detail",
    "Invoice发票": "invoice",
    "亚马逊美国DSP店铺花费占比": "spend_ratio",
    "KPI": "kpi",
    "付款申请邮件": "email_html",
    "付款申请单": None,
}
for item in checklist:
    k = ready_map.get(item)
    done = bool(k and outputs.get(k) and os.path.exists(outputs.get(k)))
    st.text(f"{'✅' if done else '⬜'} {item}")

# ================= 下载 =================
st.header("下载")
final_ok = outputs.get("final_ok", False)
if final_ok:
    st.success("✅ 全部校验通过（含截图 + 金额 + 附件），可下载最终付款申请包。")
else:
    st.error("❌ 邮件截图生成失败或存在未确认项，当前生成的是【调试包/草稿】。")
    for p in outputs.get("final_problems", []):
        st.text(f"• {p}")

for label, key in (
    ("📦 付款申请包 ZIP", "zip"),
    ("📄 对账单", "reconciliation"),
    ("📄 DSP花费明细", "spend_detail"),
    ("📄 店铺花费占比", "spend_ratio"),
    ("📄 KPI", "kpi"),
    ("📧 付款申请邮件 HTML", "email_html"),
    ("📧 付款申请邮件 TXT", "email_txt"),
):
    p = outputs.get(key)
    if p and os.path.exists(p):
        with open(p, "rb") as f:
            st.download_button(label, f, file_name=os.path.basename(p), key=f"dl_{key}")

st.caption("打印资料：3类（1. 对账单  2. 申请邮件  3. 发票）。资料确认后提交财务。")
