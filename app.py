"""
app.py — Flask 主程式
"""
from __future__ import annotations

import csv
import io
import os
import uuid
from pathlib import Path

from flask import (
    Flask,
    jsonify,
    render_template,
    request,
    send_file,
    session,
)

from parser import merge_periods, parse_pdf, build_employee_report
from receipt import generate_receipt_pdf, generate_zip

app = Flask(__name__)
app.secret_key = os.urandom(24)

UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".pdf"}

EMPLOYEE_NAMES_FILE = Path(__file__).parent / "employee_names.csv"


def _load_employee_names():
    # type: () -> dict
    """載入員工編號對應人名對照表。"""
    mapping = {}
    if not EMPLOYEE_NAMES_FILE.exists():
        return mapping
    with open(str(EMPLOYEE_NAMES_FILE), newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            eid = row.get("employee_id", "").strip()
            name = row.get("name", "").strip()
            if eid and name:
                mapping[eid] = name
    return mapping


def _allowed(filename):
    # type: (str) -> bool
    return Path(filename).suffix.lower() in ALLOWED_EXT


# ──────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/debug")
def debug():
    """測試 Tesseract 與 pdf2image 環境是否正常。"""
    import subprocess
    info = {}
    try:
        result = subprocess.run(["tesseract", "--version"], capture_output=True, text=True)
        info["tesseract_version"] = result.stdout or result.stderr
    except Exception as e:
        info["tesseract_version"] = "ERROR: {}".format(e)
    try:
        result = subprocess.run(["tesseract", "--list-langs"], capture_output=True, text=True)
        info["tesseract_langs"] = result.stdout or result.stderr
    except Exception as e:
        info["tesseract_langs"] = "ERROR: {}".format(e)
    info["tessdata_prefix"] = os.environ.get("TESSDATA_PREFIX", "not set")
    try:
        import pytesseract
        info["pytesseract_version"] = str(pytesseract.get_tesseract_version())
    except Exception as e:
        info["pytesseract"] = "ERROR: {}".format(e)
    return jsonify(info)


@app.errorhandler(Exception)
def handle_exception(e):
    import traceback
    return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/upload", methods=["POST"])
def upload():
    """接收上傳的 PDF 檔案，解析後回傳員工費用列表（JSON）。"""
    files = request.files.getlist("pdfs")

    # 新增欄位
    bank_holder = request.form.get("bank_holder", "").strip()
    bank_name = request.form.get("bank_name", "").strip()
    bank_account = request.form.get("bank_account", "").strip()
    year_label = request.form.get("year_label", "").strip()

    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "請選擇至少一個 PDF 檔案"}), 400

    all_period_lists = []
    filenames = []

    for f in files:
        if not _allowed(f.filename):
            return jsonify({"error": "不支援的檔案格式：{}".format(f.filename)}), 400

        safe_name = "{}_{}".format(uuid.uuid4().hex, Path(f.filename).name)
        dest = UPLOAD_DIR / safe_name
        f.save(str(dest))
        filenames.append(f.filename)

        try:
            import traceback
            print("[upload] start parse_pdf:", f.filename, flush=True)
            periods = parse_pdf(str(dest))
            print("[upload] parse_pdf done, periods:", len(periods), flush=True)
            all_period_lists.append(periods)
        except Exception as e:
            print("[upload] ERROR:", traceback.format_exc(), flush=True)
            return jsonify({"error": "解析 {} 時發生錯誤：{}".format(f.filename, e), "trace": traceback.format_exc()}), 500
        finally:
            dest.unlink(missing_ok=True)

    merged = merge_periods(all_period_lists)

    if not merged:
        return jsonify({
            "error": "未能從上傳的 PDF 中解析出任何員工資料，請確認 PDF 格式是否正確。"
        }), 422

    employees = build_employee_report(merged)

    # 填入員工名稱
    name_map = _load_employee_names()
    for emp in employees:
        emp["name"] = name_map.get(emp["employee_id"], "")

    if not employees:
        return jsonify({"error": "未找到任何員工資料。"}), 422

    # 組合回傳給前端的摘要資料
    result_employees = []
    for emp in employees:
        total_bw = sum(p["bw"] for p in emp["periods"])
        total_color = sum(p["color"] for p in emp["periods"])
        result_employees.append({
            "employee_id": emp["employee_id"],
            "name": emp["name"],
            "total_bw": total_bw,
            "total_color": total_color,
            "total": emp["total"],
        })

    # 把完整資料（含期別明細）存入 session
    session_id = uuid.uuid4().hex
    bank_info = {
        "holder": bank_holder,
        "bank": bank_name,
        "account": bank_account,
    }
    app.config.setdefault("_sessions", {})[session_id] = {
        "employees": employees,
        "year_label": year_label,
        "bank_info": bank_info,
    }

    return jsonify({
        "session_id": session_id,
        "year_label": year_label,
        "employees": result_employees,
    })


@app.route("/download/zip/<session_id>")
def download_zip(session_id):
    # type: (str) -> object
    store = app.config.get("_sessions", {}).get(session_id)
    if not store:
        return "Session 已過期，請重新上傳。", 404

    zip_bytes = generate_zip(
        store["employees"],
        store["year_label"],
        store["bank_info"],
    )
    return send_file(
        io.BytesIO(zip_bytes),
        mimetype="application/zip",
        as_attachment=True,
        download_name="receipts_{}.zip".format(store["year_label"] or "all"),
    )


@app.route("/download/single/<session_id>/<employee_id>")
def download_single(session_id, employee_id):
    # type: (str, str) -> object
    store = app.config.get("_sessions", {}).get(session_id)
    if not store:
        return "Session 已過期，請重新上傳。", 404

    emp = next(
        (e for e in store["employees"] if e["employee_id"] == employee_id), None
    )
    if not emp:
        return "找不到員工資料。", 404

    pdf_bytes = generate_receipt_pdf(
        emp,
        store["year_label"],
        store["bank_info"],
    )
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="receipt_{}.pdf".format(employee_id),
    )


if __name__ == "__main__":
    app.run(debug=True, port=8080)
