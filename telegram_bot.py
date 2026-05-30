import os
import logging
from datetime import datetime, timedelta, date
import calendar
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes
from tao_bao_cao import tao_bao_cao, doc_du_lieu
from tao_bao_cao_doanhthu import tong_hop_doanh_thu
import pandas as pd

BOT_TOKEN = "8816246373:AAFNJr6KWKxXWgQwBYQu5nFqGOl574m-q-c"
CHAT_ID = None
DOWNLOAD_DIR = "downloads"

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


async def luu_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg or not msg.document:
        return
    filename = msg.document.file_name or ""
    if not (filename.endswith(".xlsx") or filename.endswith(".xls")):
        return
    if CHAT_ID and msg.chat_id != CHAT_ID:
        return
    try:
        file = await context.bot.get_file(msg.document.file_id)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_name = f"{ts}_{filename}"
        input_path = os.path.join(DOWNLOAD_DIR, save_name)
        await file.download_to_drive(input_path)
        logging.info(f"Da luu: {save_name}")
        await msg.reply_text(f"Da luu file: {filename}")
    except Exception as e:
        logging.error(f"Loi luu file: {e}")


def tinh_khoang_ngay(args):
    now = datetime.now()
    if not args:
        today = now.date()
        days_since_monday = today.weekday()
        last_monday = today - timedelta(days=days_since_monday + 7)
        last_sunday = last_monday + timedelta(days=6)
        tu_ngay = datetime.combine(last_monday, datetime.min.time())
        den_ngay = datetime.combine(last_sunday, datetime.max.time().replace(microsecond=0))
        ten_ky = f"Tuan {last_monday.strftime('%d/%m')} - {last_sunday.strftime('%d/%m/%Y')}"
        return tu_ngay, den_ngay, ten_ky
    if args[0].lower() == "thang":
        first_this_month = now.replace(day=1)
        last_month_end = first_this_month - timedelta(days=1)
        last_month_start = last_month_end.replace(day=1)
        tu_ngay = datetime.combine(last_month_start, datetime.min.time())
        den_ngay = datetime.combine(last_month_end, datetime.max.time().replace(microsecond=0))
        ten_ky = f"Thang {last_month_start.strftime('%m/%Y')}"
        return tu_ngay, den_ngay, ten_ky
    if len(args) >= 2:
        try:
            year = now.year
            d1 = datetime.strptime(f"{args[0]}/{year}", "%d/%m/%Y")
            d2 = datetime.strptime(f"{args[1]}/{year}", "%d/%m/%Y")
            tu_ngay = d1.replace(hour=0, minute=0, second=0)
            den_ngay = d2.replace(hour=23, minute=59, second=59)
            ten_ky = f"{d1.strftime('%d/%m')} - {d2.strftime('%d/%m/%Y')}"
            return tu_ngay, den_ngay, ten_ky
        except ValueError:
            return None, None, None
    return None, None, None


def tim_file_theo_thang(thang, nam=None):
    now = datetime.now()
    if nam is None:
        nam = now.year
    tu_ngay = datetime(nam, thang, 1)
    if thang == 12:
        den_ngay = datetime(nam + 1, 1, 1) - timedelta(seconds=1)
    else:
        den_ngay = datetime(nam, thang + 1, 1) - timedelta(seconds=1)
    files = []
    for f in os.listdir(DOWNLOAD_DIR):
        if not (f.endswith(".xlsx") or f.endswith(".xls")):
            continue
        if f.startswith("temp_") or f.startswith("BaoCao") or f.startswith("DoanhThu"):
            continue
        fpath = os.path.join(DOWNLOAD_DIR, f)
        mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
        if tu_ngay <= mtime <= den_ngay:
            files.append(fpath)
    return files, tu_ngay, den_ngay


async def tao_bao_cao_tuan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if CHAT_ID and msg.chat_id != CHAT_ID:
        return
    args = context.args or []
    tu_ngay, den_ngay, ten_ky = tinh_khoang_ngay(args)
    if tu_ngay is None:
        await msg.reply_text(
            "Cu phap:\n"
            "/baocao - bao cao tuan truoc\n"
            "/baocao thang - bao cao thang truoc\n"
            "/baocao 18/5 24/5 - bao cao theo ngay cu the"
        )
        return
    await msg.reply_text(f"Dang tong hop: {ten_ky}...")
    files = []
    for f in os.listdir(DOWNLOAD_DIR):
        if f.endswith(".xlsx") or f.endswith(".xls"):
            if f.startswith("temp_") or f.startswith("BaoCao") or f.startswith("DoanhThu"):
                continue
            fpath = os.path.join(DOWNLOAD_DIR, f)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
            if tu_ngay <= mtime <= den_ngay:
                files.append(fpath)
    if not files:
        await msg.reply_text(f"Khong tim thay file Excel nao trong: {ten_ky}")
        return
    dfs = []
    for f in files:
        try:
            df = doc_du_lieu(f)
            dfs.append(df)
        except Exception as e:
            logging.warning(f"Bo qua {f}: {e}")
    if not dfs:
        await msg.reply_text("Khong doc duoc du lieu tu cac file Excel.")
        return
    try:
        df_all = pd.concat(dfs, ignore_index=True)
        tmp_input = os.path.join(DOWNLOAD_DIR, "temp_combined.xlsx")
        df_all.to_excel(tmp_input, index=False, sheet_name="Tong hop san luong")
        ten_file = ten_ky.replace("/", "-").replace(" ", "_")
        output_path = os.path.join(DOWNLOAD_DIR, f"BaoCao_{ten_file}.xlsx")
        tao_bao_cao(tmp_input, output_path)
        with open(output_path, "rb") as f:
            await msg.reply_document(
                document=f,
                filename=os.path.basename(output_path),
                caption=f"Bao cao san luong {ten_ky} ({len(files)} file)"
            )
    except Exception as e:
        logging.error(f"Loi tao bao cao: {e}")
        await msg.reply_text(f"Loi: {str(e)}")


async def tao_bao_cao_doanh_thu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if CHAT_ID and msg.chat_id != CHAT_ID:
        return

    args = context.args or []
    if not args:
        await msg.reply_text(
            "Cu phap:\n"
            "/doanhthu 4 - bao cao doanh thu thang 4\n"
            "/doanhthu 4 2026 - bao cao doanh thu thang 4 nam 2026"
        )
        return

    try:
        thang = int(args[0])
        nam = int(args[1]) if len(args) >= 2 else None
    except ValueError:
        await msg.reply_text("Thang phai la so. Vi du: /doanhthu 4")
        return

    if not (1 <= thang <= 12):
        await msg.reply_text("Thang phai tu 1 den 12.")
        return

    files, tu_ngay, den_ngay = tim_file_theo_thang(thang, nam)
    ten_thang = f"Thang {thang}/{tu_ngay.year}"

    if not files:
        await msg.reply_text(f"Khong tim thay file Excel nao trong {ten_thang}.")
        return

    await msg.reply_text(f"Dang tong hop bao cao doanh thu {ten_thang} ({len(files)} file)...")

    try:
        output_path = os.path.join(DOWNLOAD_DIR, f"DoanhThu_Thang{thang:02d}_{tu_ngay.year}.xlsx")
        tong_hop_doanh_thu(files, output_path, thang=f"{thang}/{tu_ngay.year}")
        with open(output_path, "rb") as f:
            await msg.reply_document(
                document=f,
                filename=os.path.basename(output_path),
                caption=f"Bao cao doanh thu {ten_thang} ({len(files)} file)"
            )
    except Exception as e:
        logging.error(f"Loi tao bao cao doanh thu: {e}")
        await msg.reply_text(f"Loi: {str(e)}")


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, luu_file))
    app.add_handler(CommandHandler("baocao", tao_bao_cao_tuan))
    app.add_handler(CommandHandler("doanhthu", tao_bao_cao_doanh_thu))
    print("Bot dang chay...")
    print("=== SAN LUONG ===")
    print("Go /baocao - tuan truoc")
    print("Go /baocao thang - thang truoc")
    print("Go /baocao 18/5 24/5 - khoang ngay cu the")
    print("=== DOANH THU ===")
    print("Go /doanhthu 4 - bao cao doanh thu thang 4")
    print("Go /doanhthu 4 2026 - bao cao doanh thu thang 4/2026")
    app.run_polling()


if __name__ == "__main__":
    main()
