"""测试：Excel 格式保护（规格 78 节 50-57）。

格式保护用 openpyxl 回退路径（use_com=False）快速验证；
COM 路径的 Pivot 完整性在 test_com.py 单独验证。
"""
import hashlib
import os

import openpyxl

from conftest import build_july_ps, july_reconciliation, july_spend_detail
from processors import reconciliation_processor, spend_detail_processor
from utils import excel_utils, validation


def _md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def _run_rec(template, ps, tmp_path):
    out, _logs, _mod = reconciliation_processor.process_dsp_sa_template(
        template, ps, str(tmp_path), use_com=False)
    return out


def test_sheet_order_and_names_preserved(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before = openpyxl.load_workbook(july_reconciliation).sheetnames
    out = _run_rec(july_reconciliation, ps, tmp_path)
    after = openpyxl.load_workbook(out).sheetnames
    assert before == after
    assert before == ["DSP+SA ", "DSP数据源", "已推广广告数据 数据源"]


def test_header_unchanged(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before = openpyxl.load_workbook(july_reconciliation)
    out = _run_rec(july_reconciliation, ps, tmp_path)
    after = openpyxl.load_workbook(out)
    for sn in before.sheetnames:
        bws, aws = before[sn], after[sn]
        for c in range(1, bws.max_column + 1):
            assert bws.cell(1, c).value == aws.cell(1, c).value, f"{sn} 表头列 {c} 变化"


def test_formulas_preserved(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before_sig = excel_utils.workbook_signature(openpyxl.load_workbook(july_reconciliation))
    out = _run_rec(july_reconciliation, ps, tmp_path)
    after_sig = excel_utils.workbook_signature(openpyxl.load_workbook(out))
    res = validation.compare_signatures(before_sig, after_sig)
    formula_diffs = [d for d in res.differences if "公式" in d]
    assert formula_diffs == []


def test_merged_cells_preserved(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before_sig = excel_utils.workbook_signature(openpyxl.load_workbook(july_reconciliation))
    out = _run_rec(july_reconciliation, ps, tmp_path)
    after_sig = excel_utils.workbook_signature(openpyxl.load_workbook(out))
    res = validation.compare_signatures(before_sig, after_sig)
    merge_diffs = [d for d in res.differences if "合并" in d]
    assert merge_diffs == []


def test_column_structure_preserved(july_reconciliation, tmp_path):
    """列数/行数（结构）不变。列宽具体值因 Excel 自动列宽重拟合不逐值比较。"""
    ps = build_july_ps()
    before = openpyxl.load_workbook(july_reconciliation)
    out = _run_rec(july_reconciliation, ps, tmp_path)
    after = openpyxl.load_workbook(out)
    for sn in before.sheetnames:
        assert before[sn].max_column == after[sn].max_column, f"{sn} 列数变化"


def test_freeze_panes_preserved(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before_sig = excel_utils.workbook_signature(openpyxl.load_workbook(july_reconciliation))
    out = _run_rec(july_reconciliation, ps, tmp_path)
    after_sig = excel_utils.workbook_signature(openpyxl.load_workbook(out))
    res = validation.compare_signatures(before_sig, after_sig)
    freeze_diffs = [d for d in res.differences if "冻结" in d]
    assert freeze_diffs == []


def test_original_template_not_overwritten(july_reconciliation, tmp_path):
    ps = build_july_ps()
    before_md5 = _md5(july_reconciliation)
    reconciliation_processor.process_dsp_sa_template(july_reconciliation, ps, str(tmp_path), use_com=False)
    assert _md5(july_reconciliation) == before_md5


def test_spend_detail_template_not_overwritten(july_spend_detail, tmp_path):
    ps = build_july_ps()
    before_md5 = _md5(july_spend_detail)
    spend_detail_processor.process_spend_detail(july_spend_detail, ps, str(tmp_path))
    assert _md5(july_spend_detail) == before_md5


def test_spend_detail_merged_cells_preserved_for_same_count(july_spend_detail, tmp_path):
    ps = build_july_ps()
    out, _ = spend_detail_processor.process_spend_detail(july_spend_detail, ps, str(tmp_path))
    after = openpyxl.load_workbook(out)["示例"]
    after_merges = sorted(str(m) for m in after.merged_cells.ranges)
    for m in ("B3:B13", "C3:C13", "D3:D13", "E3:E13", "G3:G13"):
        assert m in after_merges, f"{m} 合并丢失"
    # F 列保持两个独立合并区域，绝不允许被合并成 F3:F13
    assert "F3:F8" in after_merges
    assert "F9:F13" in after_merges
    assert "F3:F13" not in after_merges
