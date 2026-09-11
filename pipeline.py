"""一键处理流水线：上传文件 → 一键处理 → 自动核对 → 输出付款申请包。

这是 Streamlit app 与底层 processor 之间的编排层。
"""
from __future__ import annotations

import os
from typing import Dict, List

from models.payment_summary import PaymentSummary
from rules import amount_fields, month_rules
from processors import (
    email_processor,
    kpi_processor,
    meeting_processor,
    month_check,
    order_summary_processor,
    package_processor,
    promoted_ads_processor,
    reconciliation_processor,
    spend_detail_processor,
    spend_ratio_processor,
    statement_processor,
    validation_processor,
)
from utils import config_loader, excel_screenshot, excel_utils, validation


class MonthMismatchError(ValueError):
    """申请月份与数据月份不一致，禁止正式生成。"""


class ConfigNotReadyError(ValueError):
    """业务配置未完成（使用示例/占位值），禁止生成正式付款申请。"""


def run_pipeline(inputs: Dict, config: Dict) -> Dict:
    """执行完整付款申请流程。inputs 见 app.py 的调用处。"""
    logs: List[str] = []
    outputs: Dict[str, str] = {}

    # ---- 0. 业务配置校验（示例/占位值禁止生成正式版）----
    problems = config_loader.validate_business_config(config)
    if problems:
        raise ConfigNotReadyError(
            "❌ 尚未配置真实付款信息\n\n"
            "当前正在使用示例配置，请先创建 config.local.yaml 并填写真实付款/收款信息。\n\n"
            + "\n".join(f"  - {p}" for p in problems)
        )

    os.makedirs(inputs["output_dir"], exist_ok=True)

    # ---- 1. 月份 ----
    month_str = inputs["month_str"]
    ps = month_rules.parse_month(month_str, peak_months=config.get("month", {}).get("peak_months"))
    logs.append(f"✅ 申请月份：{ps.month}（{ps.month_type}）")

    # ---- 2. 配置参数（金额口径不再全局切换，见 rules/amount_fields.py）----
    rebate_rate = float(inputs.get("rebate_rate") or config.get("payment", {}).get("rebate_rate", 0.08))
    ps.payment_cycle_workdays = int(config.get("payment", {}).get("payment_cycle_workdays", 20))
    ps.payer = config.get("payment", {}).get("payer", {}).get("name", "")
    ps.payee = config.get("payment", {}).get("payee", {})
    ps.payment_due_date = inputs.get("payment_due_date") or ""
    ps.payment_due_date_confirmed = bool(inputs.get("payment_due_date_confirmed")) and bool(ps.payment_due_date)
    store_default = config.get("store", {}).get("default_store", "Amazon美国站哥贝尔店")

    # ---- 2.5 附件统一注册（发票必须在邮件生成前注册，避免邮件/ZIP/页面三套状态不一致）----
    inv_path = inputs.get("invoice_path")
    if inv_path and os.path.exists(inv_path):
        ps.has_invoice = True
        outputs["invoice"] = inv_path
        logs.append("✅ 发票已上传")
    else:
        logs.append("⚠ 缺少发票（允许生成草稿）")

    # ---- 3. 月份防呆：三态检查 + 完整月检查 ----
    os_path = inputs.get("order_summary_path")
    if not os_path or not os.path.exists(os_path):
        raise FileNotFoundError("缺少 Order Summary 文件")

    month_checks = []
    month_checks.append(month_check.check_order_summary(ps, order_summary_processor.detect_interval(os_path)))

    pa_path = inputs.get("promoted_ads_path")
    if pa_path and os.path.exists(pa_path):
        month_checks.append(month_check.check_promoted(ps, promoted_ads_processor.detect_data_month(pa_path)))

    _apply_month_checks(month_checks, inputs, outputs, logs)

    # ---- 4. 解析 Order Summary ----
    orders = order_summary_processor.parse_order_summary(os_path, store_default)
    order_summary_processor.aggregate(ps, orders, rebate_rate=rebate_rate)
    logs.append(f"✅ Order Summary：{len(orders)} 条")

    # Strategy 校验（未识别必须停止）
    unknown = [o.order for o in orders if o.strategy not in ("Consideration", "Conversion")]
    if unknown:
        raise ValueError(f"Strategy 未识别：{unknown}")

    # ---- 5. 解析已推广广告 ----
    if pa_path and os.path.exists(pa_path):
        ads = promoted_ads_processor.parse_promoted_ads(pa_path)
        promoted_ads_processor.aggregate_onsite(ps, ads)
        logs.append(f"✅ 已推广广告：{len(ads)} 条")
    else:
        logs.append("⚠ 未提供已推广广告数据，站内指标为空")

    # ---- 5. 计算 KPI + 口径核对 ----
    targets = inputs.get("kpi_targets") or config.get("kpi", {})
    template_values = inputs.get("kpi_template_values")
    kpi_results = kpi_processor.compute_and_reconcile(ps, targets=targets,
                                                       template_values=template_values,
                                                       diff_tolerance=float(config.get("kpi", {}).get("diff_tolerance", 0.02)))
    logs.append("✅ KPI 计算完成")

    # ---- 6. 会议记录 → KPI 原因（AI 只生成「提议原因」，最终确认走 kpi_reason_final）----
    mr_path = inputs.get("meeting_record_path")
    ai_mode = inputs.get("ai_mode", False)
    if mr_path and os.path.exists(mr_path):
        ps.has_meeting_record = True
        text = meeting_processor.parse_meeting_text(mr_path)
        _, kpi_mode = meeting_processor.generate_reasons(text, kpi_results, use_ai=ai_mode)
        outputs["kpi_reason_mode"] = kpi_mode
        if kpi_mode == "ai":
            logs.append("✅ 会议记录 AI 语义总结已生成（仍需确认后进入正式邮件）")
        else:
            logs.append("⚠ AI 不可用，KPI 原因需人工填写并确认")
    else:
        outputs["kpi_reason_mode"] = "none"
        logs.append("⚠ 未上传 DSP 会议记录")

    # ---- 6.5 KPI 原因最终确认版（来自用户手动填写/确认；未确认为空）----
    ps.kpi_reason_final = dict(inputs.get("kpi_reasons_final") or {})

    # ---- 8. DSP+SA 分析模板（可选；数据源来自 Order Summary / 已推广-汇总，不依赖此模板）----
    rec_path = inputs.get("dsp_sa_template_path")
    if rec_path and os.path.exists(rec_path):
        use_com = inputs.get("use_com", True)
        rec_out, rec_logs, modified_ranges = reconciliation_processor.process_dsp_sa_template(
            rec_path, ps, inputs["output_dir"], use_com=use_com)
        logs.extend(rec_logs)
        if rec_out is not None:
            outputs["reconciliation"] = rec_out
            outputs["reconciliation_modified"] = modified_ranges
            outputs["reconciliation_format"] = _format_check(rec_path, rec_out, allow_merge_change=False)
    else:
        logs.append("ℹ 未提供 DSP+SA 分析模板（可选），DSP/站内数据已由原始报表直接生成")

    # ---- 8.5 付款对账单核对（独立文件，只做金额核对）----
    statement_path = inputs.get("statement_path")
    if statement_path and os.path.exists(statement_path):
        stmt = statement_processor.read_statement(statement_path)
        statement_checks = statement_processor.verify_statement(ps, stmt)
        outputs["statement_checks"] = statement_checks
        for sc in statement_checks:
            icon = {"pass": "✅", "warn": "⚠", "fail": "❌"}.get(sc["status"], "?")
            logs.append(f"{icon} {sc['name']}：{sc['message']}")
    else:
        outputs["statement_checks"] = []
        logs.append("⚠ 未提供付款对账单，无法自动核对付款金额")

    # ---- 9. 花费明细 ----
    sd_path = inputs.get("spend_detail_template_path")
    if sd_path and os.path.exists(sd_path):
        sd_out, sd_logs = spend_detail_processor.process_spend_detail(
            sd_path, ps, inputs["output_dir"],
            store_default=store_default,
            brand=config.get("store", {}).get("brand", "HUION"),
            stores=config.get("store", {}).get("stores"))
        logs.extend(sd_logs)
        outputs["spend_detail"] = sd_out
        outputs["spend_detail_format"] = _format_check(sd_path, sd_out, allow_merge_change=True)
    else:
        logs.append("⚠ 未提供 DSP 花费明细模板")

    # ---- 10. 店铺占比（店铺/比例来自模板，不根据 Order 推导）----
    sr_path = inputs.get("spend_ratio_template_path")
    if sr_path and os.path.exists(sr_path):
        sr_out, sr_logs = spend_ratio_processor.process_spend_ratio(
            sr_path, ps, inputs["output_dir"],
            applicant=config.get("spend_ratio", {}).get("applicant", "余曼妮"))
        logs.extend(sr_logs)
        outputs["spend_ratio"] = sr_out
    else:
        logs.append("⚠ 未提供店铺花费占比模板")

    # ---- 11. KPI 模板（可选；未上传不生成 KPI Excel，禁止程序自行设计 KPI 表）----
    kpi_template_path = inputs.get("kpi_template_path")
    if kpi_template_path and os.path.exists(kpi_template_path):
        kpi_out, kpi_logs = kpi_processor.fill_kpi_template(kpi_template_path, ps, inputs["output_dir"])
        logs.extend(kpi_logs)
        outputs["kpi"] = kpi_out
    else:
        logs.append("ℹ 未上传 KPI 模板，仅计算/展示 KPI，不生成 KPI Excel")

    # ---- 11.5 邮件截图（用 Excel COM 对最终 Excel 指定区域截图）----
    enable_screenshots = inputs.get("enable_screenshots", True)
    if enable_screenshots:
        screenshots = excel_screenshot.generate_all_screenshots(
            outputs.get("reconciliation"),
            outputs.get("kpi"),
            outputs.get("spend_ratio"),
            reasons_confirmed=ps.kpi_reasons_confirmed(),
            output_dir=inputs["output_dir"],
            config=config,
            expected_totals={"dsp": float(ps.dsp_cost), "onsite": float(ps.onsite_cost)},
        )
        outputs["screenshots"] = screenshots
        for key, (path, status, message) in screenshots.items():
            if status == "success":
                logs.append(f"✅ 截图成功：{os.path.basename(path)}")
            elif status == "skip":
                logs.append(f"⚠ 跳过截图：{message}")
            else:
                logs.append(f"❌ 截图失败：{message}")
    else:
        screenshots = {key: (None, "skip", "截图已禁用") for key, _ in excel_screenshot.PNG_SLOTS}
        outputs["screenshots"] = screenshots
        logs.append("ℹ 截图功能已禁用（enable_screenshots=False）")
    outputs["email_images_dir"] = os.path.join(inputs["output_dir"], "email_images")

    # ---- 12. 生成邮件（截图 base64 内嵌）+ 保存 HTML/TXT ----
    email = email_processor.generate_email(ps, config, screenshots=screenshots)
    outputs["email"] = email
    email_html = os.path.join(inputs["output_dir"], f"DSP付款申请邮件-{ps.month_code}.html")
    with open(email_html, "w", encoding="utf-8") as f:
        f.write(email["html"])
    outputs["email_html"] = email_html
    email_txt = os.path.join(inputs["output_dir"], f"DSP付款申请邮件-{ps.month_code}.txt")
    with open(email_txt, "w", encoding="utf-8") as f:
        f.write(email["text"])
    outputs["email_txt"] = email_txt
    logs.append(f"✅ 邮件已生成（HTML 主版本 + TXT 纯文本）：{os.path.basename(email_html)}")

    # ---- 13. 校验 ----
    checks = validation_processor.validate(ps)
    # 合并付款对账单核对结果
    stmt_results = [validation_processor.CheckResult(
        name=sc["name"], status=sc["status"], message=sc["message"])
        for sc in outputs.get("statement_checks", [])]
    checks = checks + stmt_results
    outputs["validation"] = checks
    outputs["amount_bindings"] = amount_fields.binding_table(ps)

    # ---- 15. 最终输出校验 + ZIP（截图失败/错误字符串命中时禁止生成最终包）----
    final_ok, problems = _validate_final_outputs(outputs, screenshots, enable_screenshots=enable_screenshots)
    outputs["final_ok"] = final_ok
    outputs["final_problems"] = problems

    files = [v for k, v in outputs.items() if k in
             ("reconciliation", "spend_detail", "spend_ratio", "kpi", "email_html", "email_txt", "invoice")]
    image_files = [v[0] for v in screenshots.values() if v and v[0] and os.path.exists(v[0])]

    zip_name = f"DSP付款申请_{ps.month_code}.zip" if final_ok else f"DSP付款申请_{ps.month_code}_调试包.zip"
    zip_path = package_processor.build_zip(inputs["output_dir"], files, ps,
                                           zip_name=zip_name, image_files=image_files)
    outputs["zip"] = zip_path
    outputs["zip_is_final"] = final_ok

    # ZIP 生成后重新解压验收（不能只报告创建成功）
    zip_ok, zip_missing = _verify_zip_contents(zip_path, screenshots, bool(outputs.get("kpi")))
    outputs["zip_verified"] = zip_ok
    outputs["zip_missing"] = zip_missing

    if final_ok and zip_ok:
        logs.append(f"✅ 最终 ZIP 已生成并通过解压验收：{os.path.basename(zip_path)}")
    else:
        for p in problems:
            logs.append(f"❌ {p}")
        if not zip_ok:
            logs.append(f"❌ ZIP 解压验收失败，缺少：{zip_missing}")
        logs.append(f"⚠ 生成的是调试包（非最终）：{os.path.basename(zip_path)}")

    ps.attachments = package_processor.list_attachments(ps)

    return {"ps": ps, "outputs": outputs, "logs": logs}


def _apply_month_checks(checks, inputs: Dict, outputs: Dict, logs: List[str]) -> None:
    """应用月份三态检查：FAIL 阻断；WARNING 需人工确认，否则阻断。"""
    outputs["month_checks"] = [
        {"source": c.source, "status": c.status, "message": c.message} for c in checks
    ]
    fails = [c for c in checks if c.status == "FAIL"]
    if fails:
        msgs = "\n".join(f"  - {c.source}: {c.message}" for c in fails)
        raise MonthMismatchError(f"❌ 数据月份与申请月份不一致\n\n{msgs}\n\n请重新上传正确月份的数据。")

    warnings = [c for c in checks if c.status == "WARNING"]
    if warnings and not inputs.get("month_confirmed"):
        msgs = "\n".join(f"  - {c.source}: {c.message}" for c in warnings)
        raise MonthMismatchError(
            f"⚠ 需要人工确认月份数据\n\n{msgs}\n\n"
            "请确认后勾选「我确认该文件为完整月数据」再重试。"
        )
    for c in warnings:
        logs.append(f"⚠ {c.source}: {c.message}（已人工确认）")
    for c in checks:
        if c.status == "PASS":
            logs.append(f"✅ {c.source}: {c.message}")


def _load(path: str):
    import openpyxl
    return openpyxl.load_workbook(path)


def _format_check(template_path: str, output_path: str, allow_merge_change: bool):
    before = excel_utils.workbook_signature(_load(template_path))
    after = excel_utils.workbook_signature(_load(output_path))
    result = validation.compare_signatures(before, after, allow_merge_change=allow_merge_change)
    return {"ok": result.ok, "differences": result.differences}


def _required_screenshot_keys(has_kpi: bool) -> List[str]:
    """本应存在的截图：DSP/站内/店铺占比 必有；KPI 仅当上传了 KPI 模板时必存在。"""
    keys = ["dsp", "onsite", "spend_ratio"]
    if has_kpi:
        keys.append("kpi")
    return keys


def _validate_final_outputs(outputs: Dict, screenshots: Dict, enable_screenshots: bool = True) -> tuple:
    """最终输出校验：截图成功 + PNG 内容非空白 + HTML img/base64 + 错误字符串扫描。"""
    if not enable_screenshots:
        return (True, [])
    problems: List[str] = []
    required = _required_screenshot_keys(bool(outputs.get("kpi")))

    for key in required:
        v = screenshots.get(key)
        if not v or v[0] is None or v[1] != "success":
            problems.append(f"截图失败（{key}）：{v[2] if v else '无结果'}")
        elif not os.path.exists(v[0]) or os.path.getsize(v[0]) <= 0:
            problems.append(f"截图文件缺失或为空：{v[0]}")
        else:
            ok, info = excel_screenshot.validate_png_content(v[0])
            if not ok:
                problems.append(f"截图内容异常（{key}）：{info}")

    html_path = outputs.get("email_html")
    if html_path and os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        img_count = html.count("<img")
        b64_count = html.count("data:image/png;base64,")
        expected = len(required)
        if img_count < expected:
            problems.append(f"HTML <img> 数量不足：{img_count} < {expected}")
        if b64_count < expected:
            problems.append(f"HTML base64 数量不足：{b64_count} < {expected}")
        # 解码 HTML 中 base64 PNG，再次验证内容非空白
        import base64
        import re
        b64_list = re.findall(r"data:image/png;base64,([A-Za-z0-9+/=]+)", html)
        for i, b64 in enumerate(b64_list):
            try:
                data = base64.b64decode(b64)
            except Exception as e:
                problems.append(f"HTML base64 PNG 解码失败（第{i + 1}张）：{e}")
                continue
            ok, info = excel_screenshot.validate_png_bytes(data)
            if not ok:
                problems.append(f"HTML 内嵌 PNG 内容异常（第{i + 1}张）：{info}")
        for bad in ("Pivot刷新失败", "截图失败", "被呼叫方拒绝接收呼叫",
                    "Open.RefreshAll", "Traceback", "COM Error"):
            if bad in html:
                problems.append(f"HTML 含错误字符串：{bad}")
    else:
        problems.append("HTML 文件缺失")

    return (len(problems) == 0, problems)


def _verify_zip_contents(zip_path: str, screenshots: Dict, has_kpi: bool) -> tuple:
    """ZIP 生成后重新解压，检查 email_images 等关键文件是否存在。"""
    import zipfile
    expected_pngs = []
    for key in _required_screenshot_keys(has_kpi):
        v = screenshots.get(key)
        if v and v[0]:
            expected_pngs.append(os.path.join("email_images", os.path.basename(v[0])))
    try:
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()
    except Exception as e:
        return (False, [f"ZIP 无法打开：{e}"])
    missing = [n for n in expected_pngs if n not in names]
    return (len(missing) == 0, missing)
