import sys
import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime


def style_header(cell, bg="1F4E79", fg="FFFFFF"):
    cell.font = Font(name="Arial", bold=True, color=fg, size=10)
    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_total(cell):
    cell.font = Font(name="Arial", bold=True, size=10)
    cell.fill = PatternFill("solid", start_color="D6E4F0")
    thin = Side(style="thin")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def doc_file_doanhthu(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet_name = None
    for s in wb.sheetnames:
        if "quan" in s.lower() or "ly" in s.lower() or s == wb.sheetnames[0]:
            sheet_name = s
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]

    df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    # Tim cac cot
    col_ten = col_amount = col_sale = col_indon = col_inv = col_ngay = col_bit = None
    for col in df.columns:
        cl = col.lower().replace(" ", "").replace("\n", "")
        if "ten" in cl and ("khach" in cl or "kh" in cl):
            col_ten = col
        elif col.lower() == "amount":
            col_amount = col
        elif "sale" in cl:
            col_sale = col
        elif "indon" in cl or "inđon" in cl or ("in" in cl and "don" in cl):
            col_indon = col
        elif "inv" in cl:
            col_inv = col
        elif "ngay" in cl and "gui" in cl:
            col_ngay = col
        elif "bit" in cl:
            col_bit = col

    if not col_amount:
        raise ValueError(f"Khong tim thay cot Amount trong {os.path.basename(filepath)}")
    if not col_ten:
        col_ten = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    if not col_sale:
        col_sale = df.columns[-1]

    df[col_amount] = pd.to_numeric(df[col_amount], errors="coerce").fillna(0)
    df = df[df[col_ten].notna()].copy()

    return df, col_ten, col_amount, col_sale, col_indon, col_inv, col_ngay, col_bit


def viet_sheet_quan_ly(wb, df_all, col_ten, col_amount, col_sale, col_indon, col_inv, col_ngay, col_bit):
    ws = wb.create_sheet("Quan ly")
    headers = ["STT", "Ten khach", "In don", "INV Date", "Ngay gui don", "Amount", "So Bit", "Sale"]
    ws.append(headers)
    for cell in ws[1]:
        style_header(cell)

    cols_order = [col_ten, col_indon, col_inv, col_ngay, col_amount, col_bit, col_sale]

    for i, (_, row) in enumerate(df_all.iterrows(), start=1):
        ws.cell(i + 1, 1, i).font = Font(name="Arial", size=10)
        for j, col in enumerate(cols_order, start=2):
            if col and col in df_all.columns:
                val = row[col]
                if pd.isna(val):
                    val = None
                c = ws.cell(i + 1, j, val)
                c.font = Font(name="Arial", size=10)
                if j == 6 and isinstance(val, (int, float)):
                    c.number_format = "#,##0"

    r = len(df_all) + 2
    ws.cell(r, 1, "TONG")
    ws.cell(r, 6, f"=SUM(F2:F{r-1})")
    ws.cell(r, 6).number_format = "#,##0"
    ws.cell(r, 7, f"=SUM(G2:G{r-1})")
    for col in range(1, 9):
        style_total(ws.cell(r, col))

    widths = [6, 22, 12, 12, 12, 12, 10, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w


def viet_sheet_cong_no(wb, df_all, col_ten, col_amount):
    ws = wb.create_sheet("Cong no")
    ws.append(["STT", "Ten khach hang", "Tong tien"])
    for cell in ws[1]:
        style_header(cell)

    khach = df_all.groupby(col_ten)[col_amount].sum().reset_index()
    khach.columns = ["Ten", "Tong"]
    khach = khach.sort_values("Tong", ascending=False).reset_index(drop=True)

    for i, row in khach.iterrows():
        ws.cell(i + 2, 1, i + 1).font = Font(name="Arial", size=10)
        ws.cell(i + 2, 2, row["Ten"]).font = Font(name="Arial", size=10)
        c = ws.cell(i + 2, 3, row["Tong"])
        c.font = Font(name="Arial", size=10)
        c.number_format = "#,##0"

    r = len(khach) + 2
    ws.cell(r, 1, "TONG")
    ws.cell(r, 3, f"=SUM(C2:C{r-1})")
    ws.cell(r, 3).number_format = "#,##0"
    for col in range(1, 4):
        style_total(ws.cell(r, col))

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 15


def viet_sheet_dthu(wb, df_all, col_sale, col_amount):
    ws = wb.create_sheet("DThu")
    ws.append(["Sale", "Tong doanh thu"])
    for cell in ws[1]:
        style_header(cell)

    sale = df_all.groupby(col_sale)[col_amount].sum().reset_index()
    sale.columns = ["Sale", "Tong"]
    sale = sale.sort_values("Tong", ascending=False).reset_index(drop=True)

    for i, row in sale.iterrows():
        ws.cell(i + 2, 1, row["Sale"]).font = Font(name="Arial", size=10)
        c = ws.cell(i + 2, 2, row["Tong"])
        c.font = Font(name="Arial", size=10)
        c.number_format = "#,##0"

    r = len(sale) + 2
    ws.cell(r, 1, "TONG")
    ws.cell(r, 2, f"=SUM(B2:B{r-1})")
    ws.cell(r, 2).number_format = "#,##0"
    for col in range(1, 3):
        style_total(ws.cell(r, col))

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18


def tong_hop_doanh_thu(input_files, output_file, thang=None):
    all_dfs = []
    col_ten = col_amount = col_sale = col_indon = col_inv = col_ngay = col_bit = None

    for f in input_files:
        try:
            df, ct, ca, cs, ci, cinv, cn, cb = doc_file_doanhthu(f)
            if not col_ten and ct:
                col_ten = ct
            if not col_amount and ca:
                col_amount = ca
            if not col_sale and cs:
                col_sale = cs
            col_indon = ci or col_indon
            col_inv = cinv or col_inv
            col_ngay = cn or col_ngay
            col_bit = cb or col_bit
            all_dfs.append(df)
        except Exception as e:
            print(f"Bo qua {os.path.basename(f)}: {e}")

    if not all_dfs:
        raise ValueError("Khong co file hop le nao de tong hop")

    df_all = pd.concat(all_dfs, ignore_index=True)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    viet_sheet_quan_ly(wb, df_all, col_ten, col_amount, col_sale, col_indon, col_inv, col_ngay, col_bit)
    if col_ten and col_amount:
        viet_sheet_cong_no(wb, df_all, col_ten, col_amount)
    if col_sale and col_amount:
        viet_sheet_dthu(wb, df_all, col_sale, col_amount)

    wb.save(output_file)
    print(f"Da tao: {output_file}")
    return output_file


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Cach dung: python tao_bao_cao_doanhthu.py file1.xlsx [file2.xlsx ...] output.xlsx")
        sys.exit(1)
    tong_hop_doanh_thu(sys.argv[1:-1], sys.argv[-1])
