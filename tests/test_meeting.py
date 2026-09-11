"""测试：会议记录读取 + Evidence 提取 + KPI 原因生成。

本轮返工：无 AI 时不再生成技术性 fallback（如「未启用 AI / 没有 API Key / AI 总结失败」），
原因保持为空，由页面提示并让用户手动填写。
"""
from conftest import build_july_ps
from processors import meeting_processor


def test_parse_meeting(july_meeting):
    text = meeting_processor.parse_meeting_text(july_meeting)
    assert len(text) > 100
    assert "返场" in text


def test_extract_evidence(july_meeting):
    text = meeting_processor.parse_meeting_text(july_meeting)
    ev = meeting_processor.extract_evidence(text)
    assert set(ev.keys()) == {"overall", "consideration", "conversion"}
    assert len(ev["consideration"]) > 0
    assert len(ev["conversion"]) > 0


def test_evidence_only_contains_meeting_facts(july_meeting):
    """Evidence 只包含会议记录真实存在的句子（135）。"""
    text = meeting_processor.parse_meeting_text(july_meeting)
    ev = meeting_processor.extract_evidence(text)
    for cat, facts in ev.items():
        for f in facts:
            assert f in text, f"Evidence 句子不在会议记录中：{f}"


def test_no_ai_reasons_empty(july_meeting):
    """无 AI 时不生成任何原因（留空待人工填写），mode 为 none。"""
    ps = build_july_ps()
    text = meeting_processor.parse_meeting_text(july_meeting)
    reasons, mode = meeting_processor.generate_reasons(text, ps.kpi_results, use_ai=False)
    assert mode == "none"
    for k in ("Total ROAS", "Consideration CPDPV", "Conversion ROAS"):
        assert reasons[k] == ""
        assert ps.kpi_results[[r.name for r in ps.kpi_results].index(k)].reason == ""


def test_no_ai_no_garbage(july_meeting):
    """无 AI 时不得出现「出单报告 / 未启用AI / API Key / AI总结失败」等技术性提示。"""
    ps = build_july_ps()
    text = meeting_processor.parse_meeting_text(july_meeting)
    reasons, _ = meeting_processor.generate_reasons(text, ps.kpi_results, use_ai=False)
    for reason in reasons.values():
        assert "出单报告" not in reason
        assert "未启用" not in reason
        assert "API Key" not in reason
        assert "AI" not in reason
        assert "总结失败" not in reason


def test_no_api_key_mode_none(july_meeting):
    """无 API Key（use_ai=True 但本机无 .env）→ mode=none，不报错、不留垃圾。"""
    ps = build_july_ps()
    text = meeting_processor.parse_meeting_text(july_meeting)
    reasons, mode = meeting_processor.generate_reasons(text, ps.kpi_results, use_ai=True)
    assert mode == "none"
    assert all(reason == "" for reason in reasons.values())


def test_format_reasons_structure(july_meeting):
    ps = build_july_ps()
    out = meeting_processor.format_reasons(ps.kpi_results)
    assert out.startswith("一、Total ROAS")
    assert "二、Consideration CPDPV" in out
    assert "三、Conversion ROAS" in out


def test_format_reasons_uses_final_value(july_meeting):
    """原因标题里的「实际」用 final（最终汇报值），不是 calculated。"""
    ps = build_july_ps()
    from rules.kpi_rules import reconcile_template_kpi
    reconcile_template_kpi(ps.kpi_results, {"Total ROAS": 7.03}, 0.02)
    ps.kpi_reason_final = {"Total ROAS": "xx", "Consideration CPDPV": "xx", "Conversion ROAS": "xx"}
    out = meeting_processor.format_reasons(ps.kpi_results, ps.kpi_reason_final)
    assert "实际7.03" in out


def test_missing_meeting_no_reason():
    ps = build_july_ps()
    for r in ps.kpi_results:
        r.reason = ""
    out = meeting_processor.format_reasons(ps.kpi_results)
    assert "（未生成原因）" in out
