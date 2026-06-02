from flask import Flask, request, send_file, jsonify, render_template_string
import os, sys, tempfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from doc_file_don_le import doc_file_don_le, them_vao_tong_hop
from tao_bao_cao import tao_bao_cao
from github_storage import upload_file as gh_upload, download_file as gh_download
import pandas as pd
import openpyxl

app = Flask(__name__)
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

MASTER_FILENAME = "tong_hop_san_luong.xlsx"
MASTER_FILE = os.path.join(UPLOAD_DIR, MASTER_FILENAME)

HTML = """
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IndonBot - Gửi đơn hàng</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
  .card { background: white; border-radius: 16px; padding: 40px; max-width: 500px; width: 90%; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
  h1 { color: #1a73e8; font-size: 24px; margin-bottom: 8px; }
  p.sub { color: #666; margin-bottom: 30px; font-size: 14px; }
  .upload-area { border: 2px dashed #1a73e8; border-radius: 12px; padding: 30px; text-align: center; cursor: pointer; transition: background 0.2s; }
  .upload-area:hover { background: #f0f7ff; }
  .upload-area input { display: none; }
  .upload-area label { cursor: pointer; color: #1a73e8; font-weight: bold; }
  .upload-area .icon { font-size: 40px; margin-bottom: 10px; }
  #file-name { margin-top: 10px; color: #333; font-size: 13px; }
  .btn { display: block; width: 100%; padding: 14px; background: #1a73e8; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 20px; transition: background 0.2s; }
  .btn:hover { background: #1557b0; }
  .btn:disabled { background: #aaa; cursor: not-allowed; }
  .result { margin-top: 20px; padding: 15px; border-radius: 8px; display: none; }
  .result.success { background: #e8f5e9; color: #2e7d32; }
  .result.error { background: #ffebee; color: #c62828; }
  .download-btn { display: inline-block; padding: 10px 20px; background: #43a047; color: white; border-radius: 6px; text-decoration: none; font-weight: bold; margin-top: 10px; }
  .tab-bar { display: flex; gap: 10px; margin-bottom: 25px; }
  .tab { padding: 8px 16px; border-radius: 20px; cursor: pointer; font-size: 14px; border: 2px solid #1a73e8; color: #1a73e8; background: white; font-weight: bold; }
  .tab.active { background: #1a73e8; color: white; }
  .report-section { display: none; }
  .report-section.active { display: block; }
  select, input[type=text] { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 8px; font-size: 14px; margin-top: 8px; }
  label.lbl { color: #555; font-size: 13px; font-weight: bold; display: block; margin-top: 15px; }
</style>
</head>
<body>
<div class="card">
  <h1>🗂️ IndonBot</h1>
  <p class="sub">Hệ thống tổng hợp đơn hàng</p>

  <div class="tab-bar">
    <button class="tab active" onclick="showTab('upload')">📤 Gửi đơn</button>
    <button class="tab" onclick="showTab('report')">📊 Báo cáo</button>
  </div>

  <!-- Tab gửi đơn -->
  <div id="tab-upload" class="report-section active">
    <div class="upload-area" onclick="document.getElementById('file-input').click()">
      <div class="icon">📎</div>
      <label>Nhấn để chọn file Excel (.xlsx)</label>
      <input type="file" id="file-input" accept=".xlsx,.xls" onchange="onFileSelect(this)">
      <div id="file-name">Chưa chọn file</div>
    </div>
    <label class="lbl">Tên của bạn (Sale):</label>
    <input type="text" id="sale-name" placeholder="Ví dụ: A Thắng">
    <button class="btn" id="upload-btn" onclick="uploadFile()" disabled>📤 Gửi đơn</button>
    <div class="result" id="upload-result"></div>
  </div>

  <!-- Tab báo cáo -->
  <div id="tab-report" class="report-section">
    <label class="lbl">Chọn kỳ báo cáo:</label>
    <select id="report-type">
      <option value="today">Hôm nay</option>
      <option value="week">Tuần này</option>
      <option value="custom">Khoảng ngày tùy chọn</option>
    </select>
    <div id="date-range" style="display:none">
      <label class="lbl">Từ ngày (DD/MM):</label>
      <input type="text" id="date-from" placeholder="Ví dụ: 01/06">
      <label class="lbl">Đến ngày (DD/MM):</label>
      <input type="text" id="date-to" placeholder="Ví dụ: 07/06">
    </div>
    <button class="btn" onclick="getReport()">📊 Tạo báo cáo</button>
    <div class="result" id="report-result"></div>
  </div>
</div>

<script>
function showTab(tab) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.report-section').forEach(s => s.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  event.target.classList.add('active');
}

function onFileSelect(input) {
  const f = input.files[0];
  document.getElementById('file-name').textContent = f ? f.name : 'Chưa chọn file';
  document.getElementById('upload-btn').disabled = !f;
}

async function uploadFile() {
  const file = document.getElementById('file-input').files[0];
  const sale = document.getElementById('sale-name').value || 'Không rõ';
  const btn = document.getElementById('upload-btn');
  const result = document.getElementById('upload-result');

  btn.disabled = true;
  btn.textContent = '⏳ Đang xử lý...';
  result.style.display = 'none';

  const formData = new FormData();
  formData.append('file', file);
  formData.append('sale', sale);

  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();
    result.style.display = 'block';
    if (data.success) {
      result.className = 'result success';
      result.innerHTML = '✅ ' + data.message + '<br><a class="download-btn" href="' + data.download_url + '" download>⬇️ Tải báo cáo</a>';
    } else {
      result.className = 'result error';
      result.textContent = '❌ ' + data.error;
    }
  } catch(e) {
    result.style.display = 'block';
    result.className = 'result error';
    result.textContent = '❌ Lỗi kết nối';
  }

  btn.disabled = false;
  btn.textContent = '📤 Gửi đơn';
}

document.getElementById('report-type').addEventListener('change', function() {
  document.getElementById('date-range').style.display = this.value === 'custom' ? 'block' : 'none';
});

async function getReport() {
  const type = document.getElementById('report-type').value;
  const result = document.getElementById('report-result');
  result.style.display = 'none';

  let params = 'type=' + type;
  if (type === 'custom') {
    params += '&from=' + document.getElementById('date-from').value;
    params += '&to=' + document.getElementById('date-to').value;
  }

  result.style.display = 'block';
  result.className = 'result success';
  result.textContent = '⏳ Đang tạo báo cáo...';

  try {
    const res = await fetch('/report?' + params);
    const data = await res.json();
    if (data.success) {
      result.innerHTML = '✅ ' + data.message + '<br><a class="download-btn" href="' + data.download_url + '" download>⬇️ Tải báo cáo</a>';
    } else {
      result.className = 'result error';
      result.textContent = '❌ ' + data.error;
    }
  } catch(e) {
    result.className = 'result error';
    result.textContent = '❌ Lỗi kết nối';
  }
}
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "Không có file"})

    f = request.files['file']
    sale = request.form.get('sale', 'Không rõ')

    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"success": False, "error": "Chỉ chấp nhận file Excel (.xlsx)"})

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = os.path.join(UPLOAD_DIR, f"{ts}_{f.filename}")
    f.save(save_path)

    try:
        # Tong hop san luong
        if not os.path.exists(MASTER_FILE):
            gh_download(MASTER_FILENAME, MASTER_FILE)

        du_lieu, ten_kh = doc_file_don_le(save_path, sale)
        ngay_gui = datetime.now()
        for d in du_lieu:
            d["Ngày đưa đơn"] = ngay_gui
        tong = them_vao_tong_hop(du_lieu, MASTER_FILE)
        gh_upload(MASTER_FILE, MASTER_FILENAME)

        # Tao bao cao chi file nay
        tmp = save_path + "_tmp.xlsx"
        df = pd.DataFrame(du_lieu)
        df.to_excel(tmp, sheet_name="Tổng hợp sản lượng", index=False)
        out = save_path.replace(".xlsx", "_baocao.xlsx")
        tao_bao_cao(tmp, out)

        # Them sheet chi tiet
        wb = openpyxl.load_workbook(out)
        ws_raw = wb.create_sheet("Chi tiet", 0)
        ws_raw.append(list(df.columns))
        for _, row in df.iterrows():
            ws_raw.append([row[c] for c in df.columns])
        wb.save(out)

        return jsonify({
            "success": True,
            "message": f"Đã lưu {len(du_lieu)} sản phẩm của {ten_kh}. Tổng cộng: {tong} dòng",
            "download_url": f"/download/{os.path.basename(out)}"
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/report')
def report():
    report_type = request.args.get('type', 'today')
    now = datetime.now()

    if not os.path.exists(MASTER_FILE):
        result = gh_download(MASTER_FILENAME, MASTER_FILE)
        if not result:
            return jsonify({"success": False, "error": "Chưa có dữ liệu. Hãy gửi file đơn hàng trước."})

    try:
        try:
            df = pd.read_excel(MASTER_FILE, sheet_name="Tổng hợp sản lượng")
        except Exception:
            df = pd.read_excel(MASTER_FILE, sheet_name=0)

        df["Ngày đưa đơn"] = pd.to_datetime(df["Ngày đưa đơn"], errors="coerce")

        if report_type == 'today':
            df_filter = df[df["Ngày đưa đơn"].dt.date == now.date()]
            ten_ky = f"hom-nay-{now.strftime('%d%m%Y')}"
        elif report_type == 'week':
            monday = now.date() - __import__('datetime').timedelta(days=now.weekday())
            df_filter = df[df["Ngày đưa đơn"].dt.date >= monday]
            ten_ky = f"tuan-{monday.strftime('%d%m')}"
        else:
            from_str = request.args.get('from', '')
            to_str = request.args.get('to', '')
            try:
                d1 = datetime.strptime(f"{from_str}/{now.year}", "%d/%m/%Y").date()
                d2 = datetime.strptime(f"{to_str}/{now.year}", "%d/%m/%Y").date()
                df_filter = df[(df["Ngày đưa đơn"].dt.date >= d1) & (df["Ngày đưa đơn"].dt.date <= d2)]
                ten_ky = f"{from_str.replace('/','-')}-den-{to_str.replace('/','-')}"
            except Exception:
                return jsonify({"success": False, "error": "Định dạng ngày không đúng. Ví dụ: 01/06"})

        if df_filter.empty:
            return jsonify({"success": False, "error": "Không có dữ liệu trong khoảng thời gian này"})

        tmp = os.path.join(UPLOAD_DIR, "temp_report.xlsx")
        df_filter.to_excel(tmp, sheet_name="Tổng hợp sản lượng", index=False)
        out = os.path.join(UPLOAD_DIR, f"BaoCao_{ten_ky}.xlsx")
        tao_bao_cao(tmp, out)

        wb = openpyxl.load_workbook(out)
        ws_raw = wb.create_sheet("Chi tiet", 0)
        ws_raw.append(list(df_filter.columns))
        for _, row in df_filter.iterrows():
            ws_raw.append([row[c] for c in df_filter.columns])
        wb.save(out)

        return jsonify({
            "success": True,
            "message": f"Báo cáo {len(df_filter)} dòng",
            "download_url": f"/download/{os.path.basename(out)}"
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "File không tìm thấy", 404


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
