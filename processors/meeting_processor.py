"""DSP 会议记录分析：把会议记录语义总结为三个 KPI 的达标/未达标原因。

- AI 模式（需在 .env 配置 API Key）：真正调用 LLM，按每个 KPI 各自的聚焦点做语义总结，
  输出「完整、自然的运营汇报语言」，禁止原句拼接、禁止跨 KPI 串内容。生成的是「提议原因」，
  仍需用户确认后写入 PaymentSummary.kpi_reason_final 才进入正式邮件。
- 无 AI：**不生成任何技术性 fallback**（如「未启用 AI / 没有 API Key / AI 总结失败」）。
  原因保持为空，由页面提示「KPI 原因尚未生成」并让用户手动填写。
"""
from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from models.payment_summary import KPIResult


# 每个 KPI 各自的聚焦点（规格：三段内容不能互相乱串）。
KPI_FOCUS = {
    "Total ROAS": ["整体转化", "返场/促销", "预算结构", "高效资源", "低效资源优化"],
    "Consideration CPDPV": ["P+Search", "竞品搜索", "RFM", "引流成本", "DPV效率"],
    "Conversion ROAS": ["P+", "Brand View", "GIA", "RTG", "Contextual Category", "高效/低效转化资源"],
}


def parse_meeting_text(docx_path: str) -> str:
    """读取 docx 会议记录，返回全文（段落 + 表格）。"""
    from docx import Document
    doc = Document(docx_path)
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _split_sentences(text: str) -> List[str]:
    raw = re.split(r"[。；;\n]+", text)
    cleaned = []
    for s in raw:
        s = s.strip()
        if len(s) < 4:
            continue
        if s.rstrip("：:").strip() == "":
            continue
        cleaned.append(s)
    return cleaned


def extract_evidence(text: str) -> Dict[str, List[str]]:
    """从会议记录提取结构化事实（仅供 AI 提示参考，不再用于拼句）。

    返回 {"overall": [...], "consideration": [...], "conversion": [...]}。
    """
    sentences = _split_sentences(text)
    evidence: Dict[str, List[str]] = {"overall": [], "consideration": [], "conversion": []}

    consideration_kw = ["p+search", "p+ search", "竞品搜索", "竞品", "rfm",
                        "引流成本", "拉新", "cpdpv", "dpv", "新客率"]
    conversion_kw = ["brand view", "gia", "rtg", "contextual", "高单价", "热销",
                     "低效", "预算控制", "下调", "转化"]
    overall_kw = ["返场", "活动", "整体转化", "转化率回升", "转化回升", "出单主力", "出单"]

    def classify(s: str) -> List[str]:
        low = s.lower()
        cats = []
        if any(k in low for k in consideration_kw):
            cats.append("consideration")
        if any(k in low for k in conversion_kw):
            cats.append("conversion")
        if any(k in low for k in overall_kw):
            cats.append("overall")
        return cats

    for s in sentences:
        for cat in classify(s):
            if s not in evidence[cat]:
                evidence[cat].append(s)
    return evidence


def _generate_ai_reasons(text: str, kpi_results: List[KPIResult], llm) -> Dict[str, str]:
    """AI 总结模式：每个 KPI 单独一次 LLM 调用，只聚焦该 KPI 的聚焦点。"""
    reasons: Dict[str, str] = {}
    system = (
        "你是亚马逊 DSP 广告运营，负责撰写月度付款申请邮件里的 KPI 达标/未达标原因。"
        "只能依据下方提供的会议记录生成，严禁虚构会议记录中不存在的产品、促销、预算、"
        "人群、销量、ROAS、CPDPV 等信息。"
        "输出必须是完整、自然的运营汇报语言（100～250 字的一段话），"
        "不要机械复制会议原句，不要输出列表或编号。"
    )
    for r in kpi_results:
        focus = KPI_FOCUS.get(r.name, [])
        focus_text = "、".join(focus)
        user = (
            f"KPI：{r.name}\n"
            f"目标：{r.target:g}\n"
            f"实际：{r.final:.2f}\n"
            f"状态：{'达标' if r.is_met else '未达标'}\n\n"
            f"本段只聚焦以下维度：{focus_text}。\n"
            f"禁止混入其他 KPI（尤其 Consideration CPDPV 的引流/CPDPV 内容）的表述。\n\n"
            f"会议记录：\n{text}\n\n"
            f"请生成 {r.name} 的达标/未达标原因。"
        )
        try:
            reasons[r.name] = llm.generate(system, user)
        except Exception:
            reasons[r.name] = ""
    return reasons


def generate_reasons(text: str, kpi_results: List[KPIResult],
                     use_ai: bool = False) -> Tuple[Dict[str, str], str]:
    """生成三个 KPI 的「提议原因」。返回 (reasons_dict, mode)。

    mode ∈ {"ai", "none"}：
      - "ai"：LLM 已生成提议原因（仍需用户确认后写入 kpi_reason_final）；
      - "none"：AI 不可用，不生成任何内容，原因留空由用户手动填写。
    """
    if use_ai:
        from utils.llm_client import get_llm_client
        llm = get_llm_client()
        if llm is not None:
            reasons = _generate_ai_reasons(text, kpi_results, llm)
            mode = "ai"
        else:
            reasons = {r.name: "" for r in kpi_results}
            mode = "none"
    else:
        reasons = {r.name: "" for r in kpi_results}
        mode = "none"

    for r in kpi_results:
        r.reason = reasons.get(r.name, "")
    return reasons, mode


def format_reasons(kpi_results: List[KPIResult], kpi_reason_final: Optional[dict] = None) -> str:
    """把 KPI 原因格式化为三段文案。优先读 kpi_reason_final（已确认），否则读 r.reason。"""
    blocks = []
    for i, r in enumerate(kpi_results, start=1):
        status = "达标" if r.is_met else "未达标"
        cn_num = "一二三四五六"[i - 1] if i <= 6 else str(i)
        body = (kpi_reason_final or {}).get(r.name) or r.reason or "（未生成原因）"
        head = f"{cn_num}、{r.name} {status}原因（目标{r.target:g}，实际{r.final:.2f}）："
        blocks.append(f"{head}\n{body}")
    return "\n\n".join(blocks)
