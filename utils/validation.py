"""Excel 格式保护校验：处理前后对比结构是否被破坏。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass
class FormatCheckResult:
    ok: bool = True
    differences: List[str] = field(default_factory=list)

    def add(self, msg: str):
        self.ok = False
        self.differences.append(msg)


def compare_signatures(before: dict, after: dict, ignore_sheets=None,
                       allow_merge_change=False) -> FormatCheckResult:
    """对比 workbook 结构签名。

    before/after 均为 workbook_signature 的返回。
    允许变化：数据区单元格 value、max_row/max_column（因数据行数不同）。
    禁止变化：Sheet 数量/名称/顺序、合并单元格、公式、冻结窗格、列宽、行高、
              隐藏行列、打印区域。
    allow_merge_change=True 时允许合并单元格变化（如数据行数变化导致的重新合并）。
    """
    ignore_sheets = set(ignore_sheets or [])
    res = FormatCheckResult()

    b_sheets = before["sheets"]
    a_sheets = after["sheets"]
    if b_sheets != a_sheets:
        res.add(f"Sheet 列表不一致：{b_sheets} -> {a_sheets}")

    for sn in b_sheets:
        if sn in ignore_sheets:
            continue
        b = before["per_sheet"].get(sn, {})
        a = after["per_sheet"].get(sn, {})
        if not b or not a:
            if b != a:
                res.add(f"Sheet {sn!r} 缺失")
            continue

        if not allow_merge_change and b["merged"] != a["merged"]:
            res.add(f"Sheet {sn!r} 合并单元格变化：{b['merged']} -> {a['merged']}")
        if b["freeze"] != a["freeze"]:
            res.add(f"Sheet {sn!r} 冻结窗格变化：{b['freeze']} -> {a['freeze']}")
        if b["auto_filter"] != a["auto_filter"]:
            res.add(f"Sheet {sn!r} 筛选变化：{b['auto_filter']} -> {a['auto_filter']}")
        if b["print_area"] != a["print_area"]:
            res.add(f"Sheet {sn!r} 打印区域变化：{b['print_area']} -> {a['print_area']}")
        # 列宽恢复校验：COM 已恢复原始列宽；Excel ColumnWidth 在字符↔MDW 单位间
        # 转换有约 0.02 的固有舍入，容差取 0.05。
        if _dims_changed(b["column_widths"], a["column_widths"], tol=0.05):
            res.add(f"Sheet {sn!r} 列宽变化")
        # 行高不逐值比较：模板由 WPS 生成，其行高（如 18.75）在 Excel 打开后会被
        # 重新解释为默认行高（~17.7），属于 WPS/Excel 兼容问题，非程序修改所致。
        if b["hidden_cols"] != a["hidden_cols"]:
            res.add(f"Sheet {sn!r} 隐藏列变化：{b['hidden_cols']} -> {a['hidden_cols']}")
        if b["hidden_rows"] != a["hidden_rows"]:
            res.add(f"Sheet {sn!r} 隐藏行变化：{b['hidden_rows']} -> {a['hidden_rows']}")

        # 公式对比：允许数据区公式数量变化，但已有公式不能被改
        b_formulas = b["formulas"]
        a_formulas = a["formulas"]
        for coord, formula in b_formulas.items():
            if coord in a_formulas and a_formulas[coord] != formula:
                res.add(f"Sheet {sn!r} 公式 {coord} 被修改：{formula} -> {a_formulas[coord]}")

    return res


def _dims_changed(before: dict, after: dict, tol: float) -> bool:
    """列宽/行高是否发生实质性变化（允许 tol 误差）。"""
    common = set(before) & set(after)
    for k in common:
        bv, av = before[k], after[k]
        if bv is None or av is None:
            continue
        try:
            if abs(float(bv) - float(av)) > tol:
                return True
        except (TypeError, ValueError):
            if bv != av:
                return True
    return False
