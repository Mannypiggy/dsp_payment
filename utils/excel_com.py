"""Windows Excel COM 桥接（可选）。

优先用 pywin32 + 本机 Excel 原格式处理 .xls / 刷新 Pivot。
COM 不可用时调用方回退到 openpyxl 兼容方案，并明确提示。
"""
from __future__ import annotations

import os
from typing import List, Tuple


def com_available() -> bool:
    """检查是否可用 Excel COM。"""
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    import pythoncom
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Quit()
        return True
    except Exception:
        return False
    finally:
        pythoncom.CoUninitialize()


def _dispatch():
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    return app


def edit_xls_via_com(template_path: str, output_path: str,
                     updates: List[Tuple[int, int, object]]) -> bool:
    """用 Excel 打开 .xls 模板的临时副本，另存为新文件，只改指定单元格，保留 .xls 格式。

    先复制模板到临时文件再打开，绝不直接打开/改写用户原文件。
    updates: [(row_1based, col_1based, value), ...]
    """
    import pythoncom
    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        import win32com.client  # noqa: F401
        app = _dispatch()
        wb = _open_template_copy(app, template_path)
        # 另存为 .xls（保留格式）—— 用 56 = xlExcel8(.xls)
        wb.SaveAs(os.path.abspath(output_path), FileFormat=56)
        ws = wb.Worksheets(1)
        for row, col, value in updates:
            ws.Cells(row, col).Value = value
        wb.Save()
        return True
    except Exception:
        return False
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _open_template_copy(app, template_path: str):
    """把模板复制到临时文件再打开，避免 Excel 直接改写用户原文件。"""
    import shutil
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), "_com_" + os.path.basename(template_path))
    shutil.copy(template_path, tmp)
    return app.Workbooks.Open(os.path.abspath(tmp))


def _find_sheet_by_name(wb, candidates):
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if (ws.Name or "").strip() in [c.strip() for c in candidates]:
            return ws
    return None


def _write_sheet_data(ws, ncols: int, rows) -> None:
    """清空第 2 行起的明细数据，再批量写入 rows（每行 ncols 个值）。"""
    nrows = len(rows)
    # 清空（覆盖一个足够大的范围，避免旧数据残留）
    clear_end = max(200, nrows + 1)
    ws.Range(ws.Cells(2, 1), ws.Cells(clear_end, ncols)).ClearContents()
    if nrows == 0:
        return
    data = [tuple(row) for row in rows]
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + nrows, ncols)).Value = data


def _record_dimensions(wb) -> dict:
    """记录所有 Sheet 的列宽 / 自定义行高 / 隐藏行列（刷新前）。"""
    dims = {}
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        try:
            ur = ws.UsedRange
            ncols = ur.Columns.Count
            nrows = ur.Rows.Count
        except Exception:
            ncols, nrows = 50, 500
        ncols = max(1, min(int(ncols), 200))
        nrows = max(1, min(int(nrows), 5000))
        col_widths, row_heights = {}, {}
        hidden_cols, hidden_rows = set(), set()
        for c in range(1, ncols + 1):
            col_widths[c] = ws.Columns(c).ColumnWidth
            if ws.Columns(c).Hidden:
                hidden_cols.add(c)
        for r in range(1, nrows + 1):
            # 只记录自定义行高，避免把默认行高写成显式
            try:
                if not ws.Rows(r).UseStandardHeight:
                    row_heights[r] = ws.Rows(r).RowHeight
            except Exception:
                pass
            if ws.Rows(r).Hidden:
                hidden_rows.add(r)
        dims[ws.Name] = {"col_widths": col_widths, "row_heights": row_heights,
                         "hidden_cols": hidden_cols, "hidden_rows": hidden_rows,
                         "ncols": ncols, "nrows": nrows}
    return dims


def _restore_dimensions(wb, dims: dict) -> None:
    """恢复列宽 / 自定义行高 / 隐藏行列（刷新后）。"""
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        d = dims.get(ws.Name)
        if d is None:
            continue
        for c, w in d["col_widths"].items():
            ws.Columns(c).ColumnWidth = w
        for r, h in d["row_heights"].items():
            ws.Rows(r).RowHeight = h
        for c in range(1, d["ncols"] + 1):
            ws.Columns(c).Hidden = (c in d["hidden_cols"])
        for r in range(1, d["nrows"] + 1):
            ws.Rows(r).Hidden = (r in d["hidden_rows"])


def _wait_for_refresh(app, timeout: int = 120) -> bool:
    """等待刷新完成；优先 CalculateUntilAsyncQueriesDone，否则轮询 CalculationState。"""
    import time
    if hasattr(app, "CalculateUntilAsyncQueriesDone"):
        try:
            app.CalculateUntilAsyncQueriesDone()
            return True
        except Exception:
            pass
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if app.CalculationState == 0:  # xlDone
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def write_reconciliation_via_com(template_path: str, output_path: str,
                                 dsp_rows, promo_rows, timeout: int = 120,
                                 write_dsp: bool = True, write_promo: bool = True) -> bool:
    """用 Excel COM 打开 DSP+SA 分析模板的临时副本，写入数据源（可选），另存 .xlsx。

    write_dsp / write_promo 控制是否写入对应数据源 Sheet（缺 Sheet 时不写、不报错）。

    只更新两个数据源 Sheet 的第 2 行起明细（第 1 行表头保留），
    **不 RefreshAll、不重建/删除 PivotTable/PivotCache、不把 Pivot 转成普通表**。
    真正的 Pivot 刷新交给 Excel 打开文件时执行（调用方在保存后向 pivotCacheDefinition
    注入 refreshOnLoad + 设置 fullCalcOnLoad）。
    """
    import pythoncom
    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        import win32com.client  # noqa: F401
        app = _dispatch()
        wb = _open_template_copy(app, template_path)
        ws_dsp = _find_sheet_by_name(wb, ["DSP数据源", "SHEET-DSP数据源"]) if write_dsp else None
        ws_promo = _find_sheet_by_name(wb, ["已推广广告数据 数据源", "SHEET-已推广广告数据 数据源"]) if write_promo else None

        if ws_dsp is not None:
            _write_sheet_data(ws_dsp, 15, dsp_rows)
        if ws_promo is not None:
            _write_sheet_data(ws_promo, 16, promo_rows)

        wb.SaveAs(os.path.abspath(output_path), FileFormat=51)  # 51 = xlOpenXMLWorkbook
        return True
    except Exception:
        return False
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
