import os
import re
import openpyxl
import pandas as pd
from datetime import datetime


def phan_loai_product(level):
    if not level or str(level).strip() == "":
        return "bundle"
    lv = str(level).lower().strip()
    if "bundle" in lv:
        return "bundle"
    for g in [400, 350, 300, 250, 200]:
        if str(g) in lv:
            return f"wigs {g}"
    if "wig" in lv:
        return "wigs 300"
    return "lace"


def parse_ngay(val):
    if not val:
        return None
    s = str(val).strip()
    for sep in ["-", ".", "/"]:
        if sep in s:
            parts = s.split(sep)
            if len(parts) >= 2:
                try:
                    day, month = int(parts[0]), int(parts[1])
                    year = int(parts[2]) if len(parts) > 2 else datetime.now().year
                    return datetime(year, month, day)
                except Exception:
                    pass
    return None


def tim_header_row(rows):
    header_keywords = ["number", "qty", "length", "level", "type", "quality", "color"]
    for i, row in enumerate(rows):
        row_lower = [str(v).lower().strip() if v else "" for v in row]
        matches = sum(1 for kw in header_keywords if any(kw in cell for cell in row_lower))
        if matches >= 2:
            return i
    return None


def tim_ten_khach(rows, header_idx):
    # Cach 1: Tim label "customer name" -> lay gia tri hang tiep theo
    for i, row in enumerate(rows[:header_idx]):
        for j, v in enumerate(row):
            if v and "customer" in str(v).lower():
                if i + 1 < len(rows) and j < len(rows[i + 1]):
                    val = rows[i + 1][j]
                    if val and str(val).strip():
                        return str(val).strip()
                break

    # Cach 2: Tim gia tri khong phai keyword
    skip_kw = ["inv", "start", "amount", "date", "number", "type", "level",
               "qty", "length", "quality", "color", "pattern", "mau",
               "closure", "full", "gia", "don", "ngay", "customer"]

    def normalize(s):
        viet = {"d": "d", "a": "a", "e": "e", "o": "o", "u": "u", "i": "i"}
        result = s.lower()
        for ch, rep in [("\u0111","d"),("\u0131","i"),("\u0103","a"),("\u00e2","a"),
                        ("\u1eb9","e"),("\u00ea","e"),("\u1ecd","o"),("\u00f4","o"),
                        ("\u1ef3","y"),("\u01b0","u"),("\u01a1","o"),("\u1ebd","e"),
                        ("\u01b0","u")]:
            result = result.replace(ch, rep)
        return result.replace(" ", "")

    for row in rows[:header_idx]:
        for v in row:
            if v:
                s = str(v).strip()
                sl = normalize(s)
                is_kw = any(k.replace(" ", "") in sl for k in skip_kw)
                # Kiem tra them: bat dau bang "in " (in don)
                if s.lower().startswith("in ") and len(s) <= 10:
                    is_kw = True
                is_num = bool(re.match(r'^[\d\-\./]+$', s))
                if s and not is_kw and not is_num and len(s) > 2:
                    return s
    return None


def tim_ngay(rows, header_idx):
    ngay_in = None
    inv_date = None
    for row in rows[:header_idx]:
        for v in row:
            d = parse_ngay(v)
            if d:
                if ngay_in is None:
                    ngay_in = d
                elif inv_date is None and d != ngay_in:
                    inv_date = d
                    break
    return ngay_in, inv_date


def doc_file_don_le(filepath, ten_sale=None):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    rows = [r for r in all_rows if any(v is not None for v in r)]

    if len(rows) < 3:
        raise ValueError("File khong du du lieu")

    header_idx = tim_header_row(rows)
    if header_idx is None:
        raise ValueError("Khong tim thay hang header")

    ten_kh = tim_ten_khach(rows, header_idx)
    ngay_in_don, inv_date = tim_ngay(rows, header_idx)

    # Tao Code 1: DDMM + ten khach (bo khoang trang)
    if ngay_in_don and ten_kh:
        code1 = ngay_in_don.strftime("%d%m") + ten_kh.replace(" ", "")
    elif ten_kh:
        code1 = ten_kh.replace(" ", "")
    else:
        code1 = None

    # Map cot tu header
    header_row = rows[header_idx]
    col_map = {}
    for j, v in enumerate(header_row):
        if not v:
            continue
        vl = str(v).lower().strip()
        if "number" in vl or vl == "no":
            col_map["number"] = j
        elif "full" in vl:
            col_map["full_level"] = j
        elif "closure" in vl or ("level" in vl and "full" not in vl):
            col_map["level"] = j
        elif "type" in vl:
            col_map["type"] = j
            col_map["product_col"] = j
        elif "qty" in vl:
            col_map["qty"] = j
        elif "length" in vl or "lenght" in vl:
            col_map["length"] = j
        elif "quality" in vl:
            col_map["quality"] = j
        elif "color" in vl or "mau" in vl:
            col_map["color"] = j
        elif "pattern" in vl:
            col_map["pattern"] = j

    ket_qua = []
    for row in rows[header_idx + 1:]:
        if not row or all(v is None for v in row):
            continue

        num_col = col_map.get("number", 0)
        number = row[num_col] if num_col < len(row) else None
        if number is None:
            continue
        try:
            number = int(number)
        except Exception:
            continue

        type_val = row[col_map["type"]] if "type" in col_map and col_map["type"] < len(row) else None
        level = row[col_map["level"]] if "level" in col_map and col_map["level"] < len(row) else None
        if not level and type_val and "product_col" not in col_map:
            level = type_val

        qty = row[col_map["qty"]] if "qty" in col_map and col_map["qty"] < len(row) else 0
        length = row[col_map["length"]] if "length" in col_map and col_map["length"] < len(row) else None
        full_level = row[col_map["full_level"]] if "full_level" in col_map and col_map["full_level"] < len(row) else None
        quality = row[col_map["quality"]] if "quality" in col_map and col_map["quality"] < len(row) else None
        color = row[col_map["color"]] if "color" in col_map and col_map["color"] < len(row) else None
        pattern = row[col_map["pattern"]] if "pattern" in col_map and col_map["pattern"] < len(row) else None

        try:
            qty = float(qty) if qty else 0
        except Exception:
            qty = 0

        if "product_col" in col_map and type_val:
            product = phan_loai_product(type_val)
        else:
            product = phan_loai_product(level)

        dong = {
            "Ngày đưa đơn": ngay_in_don,
            "Code 1": code1,
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
    cols = ["Ngày đưa đơn", "Code 1", "Product", "Number", "Level",
            "Qty", "Lenght", "Full Level", "Quality", "Color", "Column1", "INV Date", "Sale"]

    if os.path.exists(master_file):
        try:
            try:
                df_cu = pd.read_excel(master_file, sheet_name="Tổng hợp sản lượng")
            except Exception:
                df_cu = pd.read_excel(master_file, sheet_name=0)
        except Exception:
            df_cu = pd.DataFrame(columns=cols)
    else:
        df_cu = pd.DataFrame(columns=cols)

    df_moi = pd.DataFrame(du_lieu_moi)
    df_all = pd.concat([df_cu, df_moi], ignore_index=True)

    with pd.ExcelWriter(master_file, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="Tổng hợp sản lượng", index=False)

    return len(df_all)
