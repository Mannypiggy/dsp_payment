"""Excel 区域截图（最终阶段）：用 Excel COM 打开最终 Excel，刷新后 CopyPicture 导出 PNG。

关键：Streamlit 的按钮处理不在主线程，所有 COM 调用必须显式 CoInitialize / CoUninitialize，
否则会报 RPC_E_CALL_REJECTED（-2147418111，被呼叫方拒绝接收呼叫）或对象错乱（如 Open.RefreshAll）。

只读最终 Excel 的视觉呈现（字体/边框/填充/合并/列宽/行高/数字格式/Pivot 布局），
禁止 pandas/matplotlib/HTML 重画。截图后删除临时 ChartObject、不保存，业务文件与截图前一致。
"""
from __future__ import annotations

import os
import time
from typing import Dict, List, Optional, Tuple

RPC_E_CALL_REJECTED = -2147418111  # 0x80010001

PNG_SLOTS = [
    ("dsp", "01_DSP广告效果.png"),
    ("kpi", "02_DSP_KPI.png"),
    ("onsite", "03_站内广告效果.png"),
    ("spend_ratio", "04_店铺花费占比.png"),
]


def _is_rpc_rejected(e: Exception) -> bool:
    try:
        import pywintypes
        if isinstance(e, pywintypes.com_error):
            hr = getattr(e, "hresult", 0)
            return hr == RPC_E_CALL_REJECTED or hr == 0x80010001
    except Exception:
        pass
    return "被呼叫方拒绝接收呼叫" in str(e) or "RPC_E_CALL_REJECTED" in str(e).upper()


def com_retry(func, retries: int = 10, delay: float = 0.5):
    """COM 忙（RPC_E_CALL_REJECTED）时重试，最多 retries 次。"""
    last = None
    for _ in range(retries):
        try:
            return func()
        except Exception as e:
            if _is_rpc_rejected(e):
                last = e
                time.sleep(delay)
                continue
            raise
    raise last


def com_available() -> bool:
    import pythoncom
    pythoncom.CoInitialize()
    try:
        import win32com.client
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
    # CopyPicture 依赖 Excel UI 真实渲染：Visible=False 或 ScreenUpdating=False 都会导出空白图，
    # 尤其是旧版 .xls（店铺占比）。截图阶段必须 Visible=True + ScreenUpdating=True。
    app.Visible = True
    app.ScreenUpdating = True
    app.DisplayAlerts = False
    return app


def _wait_ready(app, timeout: float = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if app.Ready:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def wait_calculation(app, timeout: float = 60) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if app.CalculationState == 0:  # xlDone
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return False


def _wait_after_refresh(app, timeout: float = 60) -> None:
    if hasattr(app, "CalculateUntilAsyncQueriesDone"):
        try:
            com_retry(lambda: app.CalculateUntilAsyncQueriesDone())
            return
        except Exception:
            pass
    if not wait_calculation(app, timeout):
        raise TimeoutError(f"Excel 刷新/计算超时（>{timeout}s）")


def refresh_workbook(app, wb, timeout: float = 60) -> None:
    """RefreshAll + CalculateFullRebuild，等待完成。"""
    com_retry(lambda: wb.RefreshAll())
    com_retry(lambda: app.CalculateFullRebuild())
    _wait_after_refresh(app, timeout)


def export_excel_range_as_png(workbook, sheet_name: str, range_address: str,
                              output_path: str) -> str:
    """CopyPicture → 临时 ChartObject → Chart.Export PNG → 删除临时对象。

    关键：Excel 必须 Visible=True（CopyPicture 依赖 UI 渲染），并显式激活/选中目标 Range，
    CopyPicture(xlBitmap) 后等剪贴板、Paste 后再等，确保导出的是真实 Excel 内容而非空白图。
    """
    xlScreen = 1
    xlBitmap = 2
    ws = workbook.Worksheets(sheet_name)
    rng = ws.Range(range_address)

    workbook.Activate()
    ws.Activate()
    com_retry(lambda: rng.Select())

    # 临时 AutoFit 第一列（标签列），确保 Strategy/Consideration/总计 等文字完整可读。
    # 只用于截图渲染，列宽在 finally 里恢复，禁止保存回业务 Workbook。
    first_col = ws.Columns(rng.Column)
    orig_width = None
    try:
        orig_width = first_col.ColumnWidth
    except Exception:
        pass
    try:
        com_retry(lambda: first_col.AutoFit())
    except Exception:
        pass
    time.sleep(0.2)

    try:
        com_retry(lambda: rng.CopyPicture(Appearance=xlScreen, Format=xlBitmap))
        time.sleep(0.5)  # 等剪贴板就绪

        co = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
        try:
            chart = co.Chart
            com_retry(lambda: chart.Paste())
            time.sleep(0.5)  # Paste 可能异步
            com_retry(lambda: chart.Export(os.path.abspath(output_path), "PNG"))
        finally:
            try:
                co.Delete()
            except Exception:
                pass
    finally:
        # 恢复第一列原列宽
        if orig_width is not None:
            try:
                first_col.ColumnWidth = orig_width
            except Exception:
                pass
    return output_path


def validate_png_content(path: str) -> tuple:
    """用 Pillow 验证 PNG 文件是否为空白图。返回 (ok, info_str)。"""
    from PIL import Image
    try:
        img = Image.open(path).convert("RGB")
    except Exception as e:
        return (False, f"无法读取 PNG：{e}")
    return _validate_png_img(img, os.path.getsize(path))


def validate_png_bytes(data: bytes) -> tuple:
    """用 Pillow 验证 PNG 字节流（如 HTML 中 base64 解码后）是否为空白图。"""
    import io
    from PIL import Image
    try:
        img = Image.open(io.BytesIO(data)).convert("RGB")
    except Exception as e:
        return (False, f"无法读取 PNG 字节：{e}")
    return _validate_png_img(img, len(data))


def _validate_png_img(img, byte_size: int) -> tuple:
    """核心校验：空白图特征 = unique colors 极少 + 非白像素占比≈0 + 灰度方差≈0。"""
    w, h = img.size
    px = list(img.getdata())
    total = len(px)
    if total == 0:
        return (False, "图片为空")

    unique = len(set(px))
    non_white = sum(1 for r, g, b in px if r < 245 or g < 245 or b < 245)
    non_white_ratio = non_white / total
    gray = [(r + g + b) // 3 for r, g, b in px]
    mean = sum(gray) / total
    variance = sum((g - mean) ** 2 for g in gray) / total

    info = (f"尺寸={w}x{h} 字节={byte_size} unique_colors={unique} "
            f"non_white_ratio={non_white_ratio:.4f} variance={variance:.1f}")

    if unique < 20:
        return (False, f"unique colors 过少（{unique}<20），疑似空白图 — {info}")
    if non_white_ratio < 0.01:
        return (False, f"非白像素占比过低（{non_white_ratio:.4f}<0.01），疑似空白图 — {info}")
    if variance < 10:
        return (False, f"灰度方差过低（{variance:.1f}<10），疑似空白图 — {info}")
    return (True, info)


# ---------------------------------------------------------------------------
# 工具：UsedRange / 动态范围 / Pivot 标签与总值验证
# ---------------------------------------------------------------------------

def _used_range_values(ws) -> List[list]:
    ur = ws.UsedRange
    if ur is None:
        return []
    v = ur.Value
    if v is None:
        return []
    if not isinstance(v, tuple):
        return [[v]]
    if len(v) > 0 and not isinstance(v[0], tuple):
        return [list(v)]
    return [list(row) for row in v]


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def find_sheet(wb, candidates) -> object:
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if (ws.Name or "").strip() in [c.strip() for c in candidates]:
            return ws
    return None


def find_range(ws, header_marker: str, end_markers: Tuple[str, ...]) -> Optional[str]:
    # UsedRange 可能不从 A1 开始（顶部有空行时起点下移），必须用 Row/Column 把数组索引换算回真实 Sheet 坐标，
    # 否则会 off-by-one、漏掉「总计」行。
    ur = ws.UsedRange
    start_row = int(ur.Row) if ur is not None else 1
    start_col = int(ur.Column) if ur is not None else 1

    vals = _used_range_values(ws)
    if not vals:
        return None
    nrows = len(vals)
    ncols = max((len(r) for r in vals), default=1)

    def cell(r, c):
        if 1 <= r <= nrows and 1 <= c <= len(vals[r - 1]):
            return vals[r - 1][c - 1]
        return None

    header_row = None
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            if cell(r, c) is not None and str(cell(r, c)).strip() == header_marker:
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return None

    end_row = None
    for r in range(header_row + 1, nrows + 1):
        for c in range(1, ncols + 1):
            if cell(r, c) is not None and any(m in str(cell(r, c)).strip() for m in end_markers):
                end_row = r
                break
        if end_row is not None:
            break
    if end_row is None:
        end_row = header_row

    last_col = 1
    for c in range(1, ncols + 1):
        if cell(header_row, c) is not None and str(cell(header_row, c)).strip() != "":
            last_col = c

    # 换算回真实 Sheet 坐标
    real_hdr_row = start_row + header_row - 1
    real_end_row = start_row + end_row - 1
    real_last_col = start_col + last_col - 1
    return f"{_col_letter(start_col)}{real_hdr_row}:{_col_letter(real_last_col)}{real_end_row}"


def missing_labels(ws, labels: Tuple[str, ...], col: int = 1) -> List[str]:
    vals = _used_range_values(ws)
    found = set()
    for row in vals:
        if len(row) >= col and row[col - 1] is not None:
            found.add(str(row[col - 1]).strip())
    return [label for label in labels if not any(label in f for f in found)]


def verify_total_value(ws, cost_label: str, total_label: str,
                       expected: float) -> Tuple[bool, str]:
    """验证总计行的 cost_label 列值 ≈ expected（判定 Pivot 是否真正刷新成功）。"""
    vals = _used_range_values(ws)
    if not vals:
        return False, "UsedRange 为空"

    header_row, cost_col = None, None
    for r, row in enumerate(vals, start=1):
        for c, v in enumerate(row, start=1):
            if v is not None and cost_label in str(v).strip():
                header_row, cost_col = r, c
                break
        if header_row is not None:
            break
    if header_row is None:
        return False, f"找不到「{cost_label}」列"

    total_row = None
    for r in range(header_row + 1, len(vals) + 1):
        row = vals[r - 1]
        if row and row[0] is not None and total_label in str(row[0]).strip():
            total_row = r
            break
    if total_row is None:
        return False, f"找不到「{total_label}」行"

    actual = vals[total_row - 1][cost_col - 1]
    try:
        diff = abs(float(actual) - float(expected))
    except (TypeError, ValueError):
        return False, f"总计值非数字：{actual!r}"
    if diff > 0.01:
        return False, f"总计值不符：实际 {actual} vs 期望 {expected}"
    return True, f"{actual}"


# ---------------------------------------------------------------------------
# 高层编排
# ---------------------------------------------------------------------------

def _open(app, path: str):
    return com_retry(lambda: app.Workbooks.Open(os.path.abspath(path)))


def _close(wb) -> None:
    try:
        com_retry(lambda: wb.Close(False), retries=3, delay=0.3)
    except Exception:
        try:
            wb.Close(False)
        except Exception:
            pass


def _screenshot_reconciliation(app, rec_path: str, images_dir: str, scr_cfg: dict,
                               expected_totals: dict) -> dict:
    dsp_cfg = scr_cfg.get("dsp", {}) or {}
    onsite_cfg = scr_cfg.get("onsite", {}) or {}
    results = {"dsp": (None, "fail", ""), "onsite": (None, "fail", "")}
    wb = None
    try:
        wb = _open(app, rec_path)
    except Exception as e:
        msg = f"Workbook Open 失败：{e}"
        results["dsp"] = (None, "fail", msg)
        results["onsite"] = (None, "fail", msg)
        return results

    try:
        _wait_ready(app)
        refresh_workbook(app, wb)
    except Exception as e:
        msg = f"RefreshAll 失败：{e}"
        results["dsp"] = (None, "fail", msg)
        results["onsite"] = (None, "fail", msg)
        return results

    sheet = find_sheet(wb, ["DSP+SA", "SHEET-DSP+SA"])
    if sheet is None:
        results["dsp"] = (None, "fail", "找不到 DSP+SA Sheet")
        results["onsite"] = (None, "fail", "找不到 DSP+SA Sheet")
        return results

    # DSP 区域
    miss = missing_labels(sheet, ("Consideration", "Conversion", "总计"), col=1)
    if miss:
        results["dsp"] = (None, "fail", f"DSP Pivot验证失败，缺少 {miss}")
    else:
        ok, info = verify_total_value(sheet, "总花费", "总计", expected_totals.get("dsp", 0))
        if not ok:
            results["dsp"] = (None, "fail", f"DSP Pivot总值验证失败：{info}")
        else:
            rng = dsp_cfg.get("range") or find_range(sheet, "Strategy", ("总计", "Grand Total"))
            if not rng:
                results["dsp"] = (None, "fail", "无法定位 DSP 截图范围")
            else:
                try:
                    out = os.path.join(images_dir, "01_DSP广告效果.png")
                    export_excel_range_as_png(wb, sheet.Name, rng, out)
                    results["dsp"] = (out, "success", "")
                except Exception as e:
                    results["dsp"] = (None, "fail", f"DSP CopyPicture/Export 失败：{e}")

    # 站内区域
    miss = missing_labels(sheet, ("SB", "SBV", "SD", "SP", "总计"), col=1)
    if miss:
        results["onsite"] = (None, "fail", f"站内 Pivot验证失败，缺少 {miss}")
    else:
        ok, info = verify_total_value(sheet, "花费-本币", "总计", expected_totals.get("onsite", 0))
        if not ok:
            results["onsite"] = (None, "fail", f"站内 Pivot总值验证失败：{info}")
        else:
            rng = onsite_cfg.get("range") or find_range(sheet, "类型", ("总计", "Grand Total"))
            if not rng:
                results["onsite"] = (None, "fail", "无法定位站内截图范围")
            else:
                try:
                    out = os.path.join(images_dir, "03_站内广告效果.png")
                    export_excel_range_as_png(wb, sheet.Name, rng, out)
                    results["onsite"] = (out, "success", "")
                except Exception as e:
                    results["onsite"] = (None, "fail", f"站内 CopyPicture/Export 失败：{e}")

    _close(wb)
    return results


def _screenshot_kpi(app, kpi_path, images_dir, scr_cfg, reasons_confirmed) -> tuple:
    kpi_cfg = scr_cfg.get("kpi", {}) or {}
    if not kpi_path or not os.path.exists(kpi_path):
        return (None, "skip", "未上传 KPI 模板，无法生成 KPI 邮件截图")
    if not reasons_confirmed:
        return (None, "skip", "KPI原因尚未确认，KPI截图暂未生成")
    wb = None
    try:
        wb = _open(app, kpi_path)
        ws = wb.Worksheets(kpi_cfg.get("sheet", 1)) if kpi_cfg.get("sheet") else wb.Worksheets(1)
        rng = kpi_cfg.get("range") or (ws.UsedRange.Address if ws.UsedRange is not None else None)
        if not rng:
            return (None, "fail", "KPI 截图范围为空")
        out = os.path.join(images_dir, "02_DSP_KPI.png")
        export_excel_range_as_png(wb, ws.Name, rng, out)
        return (out, "success", "")
    except Exception as e:
        return (None, "fail", f"KPI 截图失败：{e}")
    finally:
        _close(wb)


def _screenshot_spend_ratio(app, sr_path, images_dir, scr_cfg) -> tuple:
    sr_cfg = scr_cfg.get("spend_ratio", {}) or {}
    if not sr_path or not os.path.exists(sr_path):
        return (None, "skip", "未生成店铺花费占比文件")
    wb = None
    try:
        wb = _open(app, sr_path)
        ws = wb.Worksheets(sr_cfg.get("sheet", 1)) if sr_cfg.get("sheet") else wb.Worksheets(1)
        rng = sr_cfg.get("range") or (ws.UsedRange.Address if ws.UsedRange is not None else None)
        if not rng:
            return (None, "fail", "店铺花费占比截图范围为空")
        out = os.path.join(images_dir, "04_店铺花费占比.png")
        export_excel_range_as_png(wb, ws.Name, rng, out)
        return (out, "success", "")
    except Exception as e:
        return (None, "fail", f"店铺花费占比截图失败：{e}")
    finally:
        _close(wb)


def generate_all_screenshots(rec_path: Optional[str], kpi_path: Optional[str],
                             sr_path: Optional[str], reasons_confirmed: bool,
                             output_dir: str, config: dict,
                             expected_totals: Optional[dict] = None) -> dict:
    """生成 4 张截图，返回 {key: (path, status, message)}。

    每次调用独立 CoInitialize/CoUninitialize（兼容 Streamlit 非主线程）。
    """
    import pythoncom

    expected_totals = expected_totals or {"dsp": 0, "onsite": 0}
    images_dir = os.path.join(output_dir, "email_images")
    os.makedirs(images_dir, exist_ok=True)
    scr_cfg = (config or {}).get("screenshots", {}) or {}

    empty = {key: (None, "fail", "") for key, _ in PNG_SLOTS}
    if not com_available():
        for key in empty:
            empty[key] = (None, "fail", "当前电脑无法调用桌面版 Microsoft Excel")
        return empty

    pythoncom.CoInitialize()
    app = None
    try:
        app = _dispatch()
        if not _wait_ready(app, timeout=15):
            for key in empty:
                empty[key] = (None, "fail", "Excel 未就绪（Ready 超时）")
            return empty

        results: dict = {}
        if rec_path and os.path.exists(rec_path):
            results.update(_screenshot_reconciliation(app, rec_path, images_dir, scr_cfg, expected_totals))
        else:
            results["dsp"] = (None, "fail", "缺少对账单文件")
            results["onsite"] = (None, "fail", "缺少对账单文件")
        results["kpi"] = _screenshot_kpi(app, kpi_path, images_dir, scr_cfg, reasons_confirmed)
        results["spend_ratio"] = _screenshot_spend_ratio(app, sr_path, images_dir, scr_cfg)

        # 内容校验：PNG 存在但为空白图时，必须判失败
        for key, (path, status, message) in list(results.items()):
            if status == "success" and path and os.path.exists(path):
                ok, info = validate_png_content(path)
                if not ok:
                    results[key] = (path, "fail", info)
        return results
    finally:
        if app is not None:
            try:
                com_retry(lambda: app.Quit(), retries=3, delay=0.3)
            except Exception:
                pass
        pythoncom.CoUninitialize()
