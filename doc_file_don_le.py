"""
Doc file don hang le (in Vera N7450.xlsx) va chuyen sang format Tong hop san luong.
"""

import os
import re
import openpyxl
import pandas as pd
from datetime import datetime


def phan_loai_product(level):
    """Xac dinh Product tu cot Level."""
    if not level or str(level).strip() == "":
        return "bundle"
    lv = str(level).lower().strip()
    if "set bundle" in lv or "set" in lv and "bundle" in lv:
        return "bundle"
    # Wigs theo trong luong
    for g in [400, 350, 300, 250, 200]:
        if f"{g}" in lv:
            return f"wigs {g}"
    # Co tu wig ma khong co trong luong
    if "wig" in lv:
        return "wigs 300"
    # Con lai la lace
    return "lace"


def parse_ngay(val):
    """Chuyen chuoi ngay (30.5 hoac 30/5) thanh datetime."""
    if not val:
        return None
    s = str(val).strip()
    for fmt in ["%d.%m", "%d/%m", "%d.%m.%Y", "%d/%m/%Y"]:
        try:
            d = datetime.strptime(s, fmt)
            if d.year == 1900:
                d = d.replace(year=datetime.now().year)
            return d
        except Exception:
            pass
    return None


def doc_file_don_le(filepath, ten_sale=None):
    """
    Doc file don hang le, tra ve list cac dong du lieu
    theo format Tong hop san luong.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Lay cac hang co du lieu (bo qua hang trong)
    all_rows = list(ws.iter_rows(values_only=True))
    rows = [r for r in all_rows if any(v is not None for v in r)]
    if len(rows) < 3:
        raise ValueError("File khong du du lieu")

    # Row 0: in don, INV date, ten khach
    # Row 1: ngay in don, INV date
    # Row 2: header (number, Level, Qty, Length...)
    row0 = rows[0]
    row1 = rows[1]

    # Ten khach hang - lay gia tri cuoi cung khong phai keyword trong row 0
    ten_kh = None
    keywords = ["in ", "inv", "đơn", "don", "date", "ngay"]
    for v in row0:
        if v:
            s = str(v).strip()
            sl = s.lower()
            is_kw = any(k in sl for k in keywords)
            if s and not is_kw:
                ten_kh = s

    # Ngay in don - cot B (index 1) cua row 1
    ngay_str = row1[1] if len(row1) > 1 else None
    ngay_in_don = parse_ngay(ngay_str)

    # INV date - cot D (index 3) cua row 1
    inv_date_str = row1[3] if len(row1) > 3 else None
    inv_date = parse_ngay(inv_date_str)

    # Doc cac dong hang (tu row 3 tro di, bo dong "tong")
    ket_qua = []
    for row in rows[3:]:
        if not row or all(v is None for v in row):
            continue
        # Bo dong tong
        col_b = row[1]
        if col_b and str(col_b).lower().strip() in ["tong", "tổng", "total"]:
            continue
        # So thu tu (number)
        number = row[1]
        if number is None:
            continue
        try:
            number = int(number)
        except Exception:
            continue

        level = row[2] if len(row) > 2 else None
        qty = row[3] if len(row) > 3 else 0
        length = row[4] if len(row) > 4 else None
        full_level = row[5] if len(row) > 5 else None
        quality = row[6] if len(row) > 6 else None
        color = row[7] if len(row) > 7 else None
        pattern = row[8] if len(row) > 8 else None

        try:
            qty = float(qty) if qty else 0
        except Exception:
            qty = 0

        product = phan_loai_product(level)

        dong = {
            "Ngày đưa đơn": ngay_in_don,
            "Code 1": ten_kh,
            "Product": product,
            "Number": number,
            "Level": level,
            "Qty": qty,
            "Lenght": length,
            "Full Level": full_level,
            "Quality": quality,
            "Color": color,
            "Column1": pattern,
            "INV Date": inv_date,
            "Sale": ten_sale or "",
        }
        ket_qua.append(dong)

    return ket_qua, ten_kh


def them_vao_tong_hop(du_lieu_moi, master_file):
    """
    Them cac dong moi vao file tong hop san luong.
    Neu file chua ton tai thi tao moi.
    """
    cols = ["Ngày đưa đơn", "Code 1", "Product", "Number", "Level",
            "Qty", "Lenght", "Full Level", "Quality", "Color", "Column1", "INV Date", "Sale"]

    if os.path.exists(master_file):
        df_cu = pd.read_excel(master_file, sheet_name="Tổng hợp sản lượng")
    else:
        df_cu = pd.DataFrame(columns=cols)

    df_moi = pd.DataFrame(du_lieu_moi, columns=cols)
    df_all = pd.concat([df_cu, df_moi], ignore_index=True)

    with pd.ExcelWriter(master_file, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="Tổng hợp sản lượng", index=False)

    return len(df_all)


