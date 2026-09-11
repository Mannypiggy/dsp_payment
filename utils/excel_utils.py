"""openpyxl 工具：Sheet 识别、清空明细、计算模式设置、格式保护对比。"""
from __future__ import annotations

import copy
from typing import List, Optional

import openpyxl


def normalize_sheet_name(name: str) -> str:
    return (name or "").strip()


def find_sheet(wb: openpyxl.Workbook, candidates: List[str]) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
    """按候选名（支持去空格）查找 Sheet。"""
    normalized = {normalize_sheet_name(sn): sn for sn in wb.sheetnames}
    for c in candidates:
        key = normalize_sheet_name(c)
        if key in normalized:
            return wb[normalized[key]]
    return None


def set_calc_mode_auto(wb: openpyxl.Workbook) -> None:
    """设置 Excel 打开时自动重算（版本兼容）。"""
    props = wb.calculation
    if hasattr(props, "calcMode"):
        props.calcMode = "auto"
    if hasattr(props, "fullCalcOnLoad"):
        props.fullCalcOnLoad = True
    if hasattr(props, "forceFullCalc"):
        props.forceFullCalc = True


def set_pivot_refresh_on_load(wb: openpyxl.Workbook) -> None:
    """设置所有 pivot cache 在打开文件时自动刷新。

    openpyxl 不解析 Pivot 表（wb._pivots 为空），此处仅为兼容旧接口的 no-op；
    真正的 refreshOnLoad 在保存后由 set_pivot_refresh_on_load_post_save 写入 XML。
    """
    try:
        for cache in getattr(wb, "_pivots", []):
            definition = getattr(cache, "cacheDefinition", None)
            if definition is not None and hasattr(definition, "refreshOnLoad"):
                definition.refreshOnLoad = True
    except Exception:
        pass


def set_pivot_refresh_on_load_post_save(path: str) -> None:
    """openpyxl 保存后，向 pivotCacheDefinition XML 添加 refreshOnLoad="1"。"""
    import re
    import shutil
    import zipfile

    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", item.filename):
                    s = data.decode("utf-8")
                    if "refreshOnLoad" not in s:
                        s = re.sub(r"<pivotCacheDefinition\b",
                                   '<pivotCacheDefinition refreshOnLoad="1"', s, count=1)
                    data = s.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp, path)


def set_refresh_on_load_post_save(path: str) -> None:
    """保存后统一注入「Excel 打开时自动刷新」所需的两项设置：

    1. 所有 pivotCacheDefinition 增加 refreshOnLoad="1"（Pivot 打开时从数据源刷新）；
    2. xl/workbook.xml 的 calcPr 设置 fullCalcOnLoad="1" + calcMode="auto"
       （GETPIVOTDATA 等公式打开时自动重算）。

    直接改 XML，不经过 openpyxl 往返，因此不会破坏 PivotTable/PivotCache 定义。
    """
    import re
    import shutil
    import zipfile

    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", item.filename):
                    s = data.decode("utf-8")
                    if "refreshOnLoad" not in s:
                        s = re.sub(r"<pivotCacheDefinition\b",
                                   '<pivotCacheDefinition refreshOnLoad="1"', s, count=1)
                    data = s.encode("utf-8")
                elif item.filename == "xl/workbook.xml":
                    s = data.decode("utf-8")
                    if "<calcPr" in s:
                        s = re.sub(
                            r"<calcPr[^>]*/>",
                            '<calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>',
                            s, count=1,
                        )
                    else:
                        s = s.replace(
                            "</workbook>",
                            '<calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
                            1,
                        )
                    data = s.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp, path)


def pivot_signature(path: str) -> dict:
    """统计工作簿中的 PivotTable / PivotCache 数量（用于处理前后一致性校验）。"""
    import re
    import zipfile

    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()
    return {
        "pivot_tables": len([n for n in names if re.match(r"xl/pivotTables/pivotTable\d+\.xml$", n)]),
        "pivot_caches": len([n for n in names if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", n)]),
        "pivot_cache_records": len([n for n in names if re.match(r"xl/pivotCache/pivotCacheRecords\d+\.xml$", n)]),
    }


def restore_pivot_parts(template_path: str, output_path: str) -> None:
    """把模板中所有 PivotTable / PivotCache 相关 XML（含 .rels）原样复制回输出文件。

    openpyxl 往返会把 pivotTable→pivotCache、pivotCacheDefinition→pivotCacheRecords 的
    关系 ID 打乱/串位，从而破坏 Pivot。本函数在 openpyxl 保存后，用 zip 级手术把
    模板的 xl/pivotTables/ 与 xl/pivotCache/ 两个子树（含各自的 _rels）整体覆盖回输出，
    保证 Pivot 定义与模板逐字节一致（仅数据源 Sheet 由程序写入）。
    """
    import shutil
    import zipfile

    PIVOT_PREFIXES = ("xl/pivotTables/", "xl/pivotCache/")
    with zipfile.ZipFile(template_path, "r") as zt:
        tpl = {n: zt.read(n) for n in zt.namelist() if n.startswith(PIVOT_PREFIXES)}

    tmp = output_path + ".tmp"
    with zipfile.ZipFile(output_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in tpl:
                    zout.writestr(item, tpl[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, output_path)


def clear_data_cells(ws, first_data_row: int, last_data_row: int,
                     min_col: int = 1, max_col: Optional[int] = None) -> int:
    """只清空数据单元格的 value，不删除行列、不改格式。

    返回清空的单元格数量。
    """
    max_col = max_col or ws.max_column
    count = 0
    for r in range(first_data_row, last_data_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                cell.value = None
                count += 1
    return count


def sheet_signature(ws) -> dict:
    """采集 Sheet 的格式签名，用于处理前后对比（格式保护校验）。"""
    return {
        "name": ws.title,
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "merged": sorted(str(m) for m in ws.merged_cells.ranges),
        "freeze": ws.freeze_panes,
        "auto_filter": ws.auto_filter.ref if ws.auto_filter else None,
        "print_area": ws.print_area,
        "column_widths": {k: v.width for k, v in ws.column_dimensions.items() if v.width is not None},
        "row_heights": {k: v.height for k, v in ws.row_dimensions.items() if v.height is not None},
        "hidden_cols": [k for k, v in ws.column_dimensions.items() if v.hidden],
        "hidden_rows": [k for k, v in ws.row_dimensions.items() if v.hidden],
        # 公式单元格：记录 (坐标 -> 公式) 用于对比
        "formulas": _collect_formulas(ws),
    }


def _collect_formulas(ws) -> dict:
    out = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                out[cell.coordinate] = cell.value
    return out


def workbook_signature(wb: openpyxl.Workbook) -> dict:
    return {
        "sheets": wb.sheetnames,
        "per_sheet": {sn: sheet_signature(wb[sn]) for sn in wb.sheetnames},
    }
