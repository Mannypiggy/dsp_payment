"""最终付款申请包：ZIP 打包 + 材料清单。"""
from __future__ import annotations

import os
import zipfile
from typing import List


def build_checklist(ps) -> List[str]:
    """付款材料 checklist（页面展示用）。"""
    items = [
        "对账单",
        "DSP花费明细",
        "Invoice发票",
        "亚马逊美国DSP店铺花费占比",
        "KPI",
        "付款申请邮件",
        "付款申请单",
    ]
    return items


def build_zip(output_dir: str, files: List[str], ps, zip_name: str = None,
              image_files: List[str] = None) -> str:
    """打包所有生成文件 + 用户上传的发票 + 邮件截图（email_images/）。返回 zip 路径。"""
    if zip_name is None:
        zip_name = f"DSP付款申请_{ps.month_code}.zip"
    zip_path = os.path.join(output_dir, zip_name)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            if f and os.path.exists(f):
                zf.write(f, os.path.basename(f))
        for f in (image_files or []):
            if f and os.path.exists(f):
                zf.write(f, os.path.join("email_images", os.path.basename(f)))
    return zip_path


def list_attachments(ps) -> List[str]:
    """返回附件文件名（不含路径）。"""
    names = ["对账单", "DSP花费明细"]
    if ps.has_invoice:
        names.append("Invoice发票")
    names.append("亚马逊美国DSP店铺花费占比")
    return names
