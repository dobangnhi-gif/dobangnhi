"""
Doc file PDF don le (Lenh San Xuat / Production Order) cua CATHAIR FACTORY.

File PDF co watermark "CATHAIR FACTORY" in cheo de len chu, gay nhieu khi
trich xuat text. Module nay loc watermark va tra ve DUNG cung cau truc du
lieu nhu doc_file_don_le.py / tao_bao_cao_doanhthu.py (doc tu Excel),
de app.py dung chung mot luong xu ly cho ca Excel lan PDF.

Yeu cau: pip install pdfplumber
"""
import re
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# Cac chu cai xuat hien trong watermark "CATHAIR FACTORY"
_WM_LETTERS = set("CATHIRFOY")


def _clean_cell(s):
    """Loai bo cac dong watermark (mot chu cai don le) trong 1 o."""
    if s is None:
        return ""
    keep = []
    for part in str(s).split("\n"):
        p = part.strip()
        if len(p) == 1 and p in _WM_LETTERS:
            continue
        keep.append(p)
    return re.sub(r"\s+", " ", " ".join(keep)).strip()


def _so_thu_tu(first_cell):
    """Lay so thu tu cua dong hang, GIU NGUYEN duoi chu.

    Xu ly linh hoat (chiu duoc nhieu watermark), tra ve:
      - '1A', '1B', '2C' (chuoi) khi co duoi chu   -> GIU nguyen A/B, KHONG
        gop thanh 1
      - 3 (so nguyen) khi la so tron, khong co duoi -> giu nguyen kieu cu
      - None khi o dau khong phai so thu tu (hang tieu de, o 'label: value')
    """
    f = _clean_cell(first_cell)
    if not f or ":" in f:
        return None
    # Bo chu cai watermark dung rieng le con dinh trong o
    f = re.sub(r"(?<![0-9A-Za-z])[CATHIRFOY](?![0-9A-Za-z])", "", f).strip()
    # Bo chu cai watermark dinh SAT TRUOC so: 'R3' -> '3' (prefix chac chan
    # khong phai so thu tu). Con duoi chu SAU so ('1A','1B') thi GIU nguyen -
    # do la danh so phu that su, khong gop thanh 1.
    f = re.sub(r"^[CATHIRFOY]+(?=\d)", "", f)
    m = re.match(r"^(\d{1,3})([A-Za-z]?)$", f)
    if not m:
        return None
    num = int(m.group(1))
    suffix = m.group(2).upper()
    return f"{num}{suffix}" if suffix else num


def _la_dong_hang(first_cell):
    """Nhan dien dong hang chi tiet mot cach LINH HOAT.

    Truoc day app yeu cau o dau la SO TRON tuyet doi (re.fullmatch d+),
    nen bi vo khi:
      - Don danh so phu kieu '1A', '1B', '2A' (nhu file khach Nelly)
      - pdfplumber (tuy phien ban) lam dinh 1 chu cai watermark vao o so,
        vd '3R', 'R3', '3\\nH' ...
    """
    return _so_thu_tu(first_cell) is not None


def _norm_type(v):
    """bundle / closure / gift (chiu duoc nhieu ky tu watermark)."""
    s = _clean_cell(v).lower()
    if "bundle" in s:
        return "bundle"
    if "closure" in s:
        return "closure"
    if "gift" in s:
        return "gift"
    if "wig" in s:
        return "wigs"
    return _clean_cell(v)


def _norm_quality(v):
    s = _clean_cell(v).lower()
    if "raw" in s:
        return "RAW"
    if "baby" in s:
        return "Baby"
    if "donor" in s:
        return "Donor"
    return _clean_cell(v)


def _norm_color(v):
    s = _clean_cell(v).lower()
    if "natural" in s:
        return "natural color"
    if "cajun" in s:
        return "cajun"
    return _clean_cell(v)


def _norm_pattern(v):
    s = _clean_cell(v).lower()
    if "loose" in s and "wav" in s:
        return "loose wavy"
    if "bone" in s and "straight" in s:
        return "bone straight"
    if "straight" in s:
        return "straight"
    return _clean_cell(v)


def _norm_length(v):
    """Lay so inch: vd '28' -> '28\"'."""
    s = _clean_cell(v)
    m = re.search(r"(\d{1,2})", s)
    if m:
        return m.group(1) + '"'
    return s


def _norm_lace(v):
    """Lay quy cach lace: 13x6 HD, 9x6 HD, 6x6 HD, 2x6..."""
    s = _clean_cell(v)
    m = re.search(r"(\d+\s*x\s*\d+)(\s*HD)?", s, re.IGNORECASE)
    if m:
        lace = m.group(1).replace(" ", "")
        if m.group(2):
            lace += " HD"
        return lace
    return ""


def _norm_int(v):
    s = _clean_cell(v)
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def _parse_ngay(s):
    """Chuyen '06-07' / '30-06' -> datetime (DD-MM, nam hien tai)."""
    if not s:
        return None
    s = _clean_cell(s)
    m = re.search(r"(\d{1,2})\s*[-./]\s*(\d{1,2})(?:\s*[-./]\s*(\d{2,4}))?", s)
    if not m:
        return None
    day, month = int(m.group(1)), int(m.group(2))
    year = int(m.group(3)) if m.group(3) else datetime.now().year
    if year < 100:
        year += 2000
    try:
        return datetime(year, month, day)
    except Exception:
        return None


def _split_label(cell):
    """'Khach: Peace 3103' -> ('khach', 'Peace 3103')."""
    if ":" not in cell:
        return None, None
    label, _, val = cell.partition(":")
    return label.strip().lower(), val.strip()


def _doc_header_va_rows(filepath):
    """Doc PDF -> (header_info dict, danh sach dong hang da lam sach).

    Header duoc trich tu cac O trong bang trang 1 (da lam sach watermark) -
    on dinh hon nhieu so voi doc tu toan van.
    """
    if pdfplumber is None:
        raise ImportError("Chua cai pdfplumber. Chay: pip install pdfplumber")

    header = {
        "so_don": None, "inv_date": None, "ngay_in_don": None,
        "khach": None, "sale": None, "note": None, "amount": None,
    }
    line_items = []
    header_cells = []  # cac o o trang 1 (chua phai dong hang)

    full_text_parts = []
    with pdfplumber.open(filepath) as pdf:
        for pidx, page in enumerate(pdf.pages):
            full_text_parts.append(page.extract_text() or "")
            for table in page.extract_tables():
                for raw in table:
                    cells = [_clean_cell(c) for c in raw]
                    # Dong hang chi tiet: o dau la so thu tu (nhan dien linh hoat)
                    if cells and _la_dong_hang(raw[0]):
                        line_items.append(cells)
                    elif pidx == 0:
                        header_cells.extend([c for c in cells if c])

    # --- Trich header tu cac o "label: value" ---
    for cell in header_cells:
        label, val = _split_label(cell)
        if not label:
            continue
        if "amount" in label:
            num = re.sub(r"[^\d]", "", val)
            if num:
                header["amount"] = float(num)
        elif "inv" in label and "date" in label:
            header["inv_date"] = _parse_ngay(val)
        elif "ngay" in label or "ng" in label:            # Ngay in don
            header["ngay_in_don"] = _parse_ngay(val)
        elif "kh" in label:                                # Khach
            header["khach"] = header["khach"] or val
        elif label == "sale":
            header["sale"] = header["sale"] or val
        elif "note" in label:
            header["note"] = val
        elif label.startswith("s") and val and re.search(r"[A-Za-z]{2,}\d", val):
            header["so_don"] = header["so_don"] or val     # So don

    # --- Fallback: doc header tu TOAN VAN neu bang khong lay du ---
    # (chong truong hop watermark lam vo o header trong bang)
    full_text = "\n".join(full_text_parts)
    ft = re.sub(r"(?m)^\s*[CATHIRFOY]\s*$", "", full_text)  # bo dong watermark 1 chu

    def _tim(pattern):
        m = re.search(pattern, ft, re.IGNORECASE)
        return m.group(1).strip() if m else None

    # Chu y: fallback deu YEU CAU dau ':' de tranh bat nham
    # (vd 'SALES INVOICE' khong bi hieu la 'Sale: S INVOICE')
    if not header["amount"]:
        v = _tim(r"AMOUNT\s*:\s*([\d.,]+)")
        if v:
            num = re.sub(r"[^\d]", "", v)
            if num:
                header["amount"] = float(num)
    if not header["khach"]:
        header["khach"] = _tim(r"Kh[aá]ch\s*:\s*([^\n|]+?)\s*(?:\||Sale|$)")
    if not header["sale"]:
        header["sale"] = _tim(r"\bSale\s*:\s*([^\n|]+)")
    if not header["so_don"]:
        header["so_don"] = _tim(r"(INV[\d\-]+)")
    if not header["ngay_in_don"]:
        v = _tim(r"Ng[aà]y in [đd][oơ]n\s*:\s*([\d\-./]+)")
        if v:
            header["ngay_in_don"] = _parse_ngay(v)
    if not header["inv_date"]:
        v = _tim(r"INV\s*DATE\s*:\s*([\d\-./]+)")
        if v:
            header["inv_date"] = _parse_ngay(v)

    return header, line_items


# ============================================================
#  API 1: San luong  (tuong duong doc_file_don_le)
# ============================================================
def doc_file_don_le_pdf(filepath, ten_sale=None):
    """
    Tra ve (danh sach dong san pham, ten_khach) - GIONG doc_file_don_le().
    Moi dong la dict cung key voi ban Excel de app.py dung chung.
    """
    from doc_file_don_le import phan_loai_product

    header, rows = _doc_header_va_rows(filepath)
    if not rows:
        raise ValueError("Khong doc duoc dong hang nao trong PDF")

    ten_kh = header.get("khach")
    if header.get("sale"):
        ten_sale = header["sale"]

    ngay_in_don = header.get("ngay_in_don")

    # Code 1 = DDMM + ten khach (bo khoang trang) - giong ban Excel
    if ngay_in_don and ten_kh:
        code1 = ngay_in_don.strftime("%d%m") + ten_kh.replace(" ", "")
    elif ten_kh:
        code1 = ten_kh.replace(" ", "")
    else:
        code1 = None

    # Thu tu cot PDF: # Type Lace Qty Length Full-level Quality Color Pattern San Kho-xuat
    ket_qua = []
    for c in rows:
        def col(i):
            return c[i] if i < len(c) else ""

        number = _so_thu_tu(col(0))   # giu nguyen duoi chu: '1A','1B' hoac int
        if number is None:
            continue
        type_val = _norm_type(col(1))
        lace = _norm_lace(col(2))
        qty = _norm_int(col(3)) or 0
        length = _norm_length(col(4))
        full_level = _clean_cell(col(5))
        quality = _norm_quality(col(6))
        color = _norm_color(col(7))
        pattern = _norm_pattern(col(8))

        # Phan loai Product theo CA cot "Type" LAN cot "Lace".
        # Nhieu INV de cot "Type" TRONG voi hang wig, thong tin "300g wigs"
        # lai nam o cot "Lace" (vd "5x5 300g wigs"). Neu chi doc cot Type se sai:
        #   - Type trong han   -> phan_loai_product('') tra ve 'bundle'
        #   - manh '5x5' lot vao Type -> khong khop 300g/wig -> roi vao 'lace'
        # KHONG dung cot "Full level" vi no chua "wig 10" (= so hieu wig set),
        # se lam hang bundle/lace bi nhan nham thanh wig.
        type_lace = (_clean_cell(col(1)) + " " + _clean_cell(col(2))).strip()
        product = phan_loai_product(type_lace)
        loai_mau = "No Color" if "natural" in (color or "").lower() else "Color"
        code2 = (str(code1) + str(number)) if code1 else str(number)

        ket_qua.append({
            "Ngày đưa đơn": ngay_in_don,
            "Code 1 ": code1,
            "Product": product,
            "Number": number,
            "Level": lace,          # cot 'Lace' cua PDF ~ quy cach closure/level
            "Qty": float(qty),
            "Lenght": length,
            "Full Level": full_level or None,
            "Quality": quality,
            "Color": color,
            "Column1": pattern,     # cot Pattern
            "Loại màu": loai_mau,
            "Sale ": ten_sale or "",
            "Code 2": code2,
        })

    return ket_qua, ten_kh


# ============================================================
#  API 2: Doanh thu  (tuong duong doc_file_don_le_doanhthu)
# ============================================================
def doc_file_don_le_doanhthu_pdf(filepath, ten_sale=None):
    """Tra ve 1 dong doanh thu - GIONG doc_file_don_le_doanhthu()."""
    header, _ = _doc_header_va_rows(filepath)
    if header.get("sale"):
        ten_sale = header["sale"]

    return {
        "Tên khách ": header.get("khach"),
        "In đơn": header.get("ngay_in_don"),
        "INV Date": header.get("inv_date"),
        "Ngày gửi đơn": header.get("ngay_in_don"),
        "Amount": header.get("amount") or 0,
        "Số Bit ": None,
        "Sale": ten_sale or "",
    }


def la_file_pdf(filename):
    return str(filename).lower().strip().endswith(".pdf")
