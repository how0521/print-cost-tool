"""
parser.py — 解析影音機輸出的 PDF 報表（OCR 版）

使用 pdf2image + pytesseract 對掃描 PDF 進行 OCR 解析。
PDF 每頁為直立 JPEG 影像，內容旋轉 90° 逆時針，需旋轉 -90° 才能正常閱讀。
表格欄位很寬，OCR 會把同一列切成多行，改用 image_to_data 依 Y 座標重組每列。
"""
from __future__ import annotations

import re
from collections import defaultdict
from datetime import date, timedelta
from typing import Optional

try:
    from pdf2image import convert_from_path
    import pytesseract
except ImportError as _e:
    raise ImportError(
        "請安裝 pdf2image 與 pytesseract：pip install pdf2image pytesseract"
    ) from _e


# ── 費用設定 ─────────────────────────────────────────────
BW_UNIT_PRICE = 3
COLOR_UNIT_PRICE = 10

# Y 座標分組容差（像素）：同一列內的字元 top 值差異容許範圍
_ROW_TOLERANCE = 22  # DPI=300 時對應約 2mm 高度容差


def _parse_dd_mm_yyyy(text):
    # type: (str) -> Optional[date]
    """將 DD/MM/YYYY 字串解析為 date 物件，失敗回傳 None。"""
    try:
        day, month, year = text.strip().split("/")
        return date(int(year), int(month), int(day))
    except Exception:
        return None


def _ocr_page_text(img):
    # type: (object) -> str
    """對單張 PIL 圖片進行 OCR，回傳純文字（用於偵測報表類型與日期）。"""
    return pytesseract.image_to_string(img, lang="chi_tra+eng", config="--psm 6")


def _ocr_page_words(img):
    # type: (object) -> list
    """
    對單張 PIL 圖片進行 OCR，回傳每個字的位置資訊。
    過濾掉空白字元與置信度過低的結果。
    """
    data = pytesseract.image_to_data(
        img, lang="chi_tra+eng", config="--psm 6",
        output_type=pytesseract.Output.DICT
    )
    words = []
    for i, text in enumerate(data["text"]):
        text = text.strip()
        if not text:
            continue
        conf = int(data["conf"][i])
        if conf < 10:
            continue
        words.append({
            "text": text,
            "left": data["left"][i],
            "top": data["top"][i],
            "width": data["width"][i],
        })
    return words


def _group_words_by_row(words):
    # type: (list) -> list
    """
    將字元依 top 值分組（容差 _ROW_TOLERANCE），
    每組代表表格中的一列，依 left 值排序。
    回傳 list of list[str]（每列的 token 串列）。
    """
    if not words:
        return []

    # 依 top 排序
    sorted_words = sorted(words, key=lambda w: w["top"])

    rows = []
    current_row = [sorted_words[0]]
    current_top = sorted_words[0]["top"]

    for w in sorted_words[1:]:
        if abs(w["top"] - current_top) <= _ROW_TOLERANCE:
            current_row.append(w)
        else:
            # 將目前列依 left 排序後加入
            current_row.sort(key=lambda x: x["left"])
            rows.append([x["text"] for x in current_row])
            current_row = [w]
            current_top = w["top"]

    if current_row:
        current_row.sort(key=lambda x: x["left"])
        rows.append([x["text"] for x in current_row])

    return rows


def _detect_report_type(text):
    # type: (str) -> str
    """偵測報表類型：'print'（列印）或 'copy'（複印）。"""
    header = text[:400]
    if re.search(r"列印", header):
        return "print"
    if re.search(r"複印|複印|複色", header):
        return "copy"
    return "unknown"


def _parse_page(img):
    # type: (object) -> Optional[dict]
    """
    解析單一旋轉後的 PIL 頁面圖片，回傳：
    {
        "report_type": "print" or "copy",
        "init_date": date,
        "print_date": date,
        "employees": {"720629": {"bw": 397, "color": 28}, ...}
    }
    失敗回傳 None。
    """
    # 先用純文字 OCR 取得報表類型和日期
    text = _ocr_page_text(img)
    report_type = _detect_report_type(text)

    m_init = re.search(r"初值化日期\s+(\d{2}/\d{2}/\d{4})", text)
    if not m_init:
        return None
    init_date = _parse_dd_mm_yyyy(m_init.group(1))
    if init_date is None:
        return None

    # 冒號可能被 OCR 漏讀，日期前也可能多一個誤讀字元（如 O05 → 05）
    m_print = re.search(r"報表列印日期\s*[：:]?\s*[^\d]?(\d{2}/\d{2}/\d{4})", text)
    if not m_print:
        return None
    print_date = _parse_dd_mm_yyyy(m_print.group(1))
    if print_date is None:
        return None

    # 用 word-level OCR 重組表格列（避免長列被換行切斷）
    words = _ocr_page_words(img)
    rows = _group_words_by_row(words)

    employees = {}

    for row_tokens in rows:
        try:
            # 從該列取出所有數字 token（保留原始字串以保留前導零）
            nums_str = [t for t in row_tokens if re.match(r"^\d+$", t)]
            if len(nums_str) < 8:
                continue

            nums = [int(n) for n in nums_str]

            # 找最後一個恰好是限制頁數欄的值（9999999，7位數）
            # 使用 9000000 <= n <= 9999999 區間，避免把 8 位員工編號（如 16307809）誤判
            last_limit = -1
            for idx, n in enumerate(nums):
                if 9000000 <= n <= 9999999:
                    last_limit = idx

            if last_limit < 0 or last_limit + 1 >= len(nums):
                continue

            bw = nums[last_limit + 1]
            # color 欄可能因 OCR 漏讀而缺失，設預設值 0
            color = nums[last_limit + 2] if last_limit + 2 < len(nums) else 0

            # user_id：固定取 nums_str 中的第 2 個 token（index=1）
            # 結構：[序號] [使用者名稱=ID] [使用者ID] [卡號] [限制黑白] [限制彩色] [黑白] [彩色] [累積]
            if len(nums_str) < 2:
                continue
            user_id = nums_str[1]

            # 基本驗證
            if len(user_id) < 3 or len(user_id) > 10:
                continue
            # 跳過疑似總計列（全為 0 的編號）
            if re.match(r"^0{3,}$", user_id):
                continue

            # 若黑白彩色都是 0，略過
            if bw == 0 and color == 0:
                continue

            if user_id in employees:
                employees[user_id]["bw"] += bw
                employees[user_id]["color"] += color
            else:
                employees[user_id] = {"bw": bw, "color": color}

        except Exception:
            continue

    return {
        "report_type": report_type,
        "init_date": init_date,
        "print_date": print_date,
        "employees": employees,
    }


def parse_pdf(filepath):
    # type: (str) -> list
    """
    解析單一 PDF 檔案，回傳 period list：
    [
        {
            "start_date": "2025-06-05",
            "label": "06/05-07/01",
            "employees": {
                "720629": {"bw": 403, "color": 28},
                ...
            }
        },
        ...
    ]
    """
    from pdf2image import pdfinfo_from_path
    try:
        page_count = pdfinfo_from_path(filepath)["Pages"]
    except Exception:
        page_count = 999  # fallback: 讓 convert_from_path 自行決定

    period_data = {}  # init_date_iso -> {print_date, employees}

    for page_num in range(1, page_count + 1):
        page_imgs = convert_from_path(filepath, dpi=300, first_page=page_num, last_page=page_num)
        if not page_imgs:
            break
        img = page_imgs[0]
        rotated = img.rotate(-90, expand=True)
        del page_imgs, img  # 釋放記憶體
        try:
            result = _parse_page(rotated)
        except Exception:
            continue

        if result is None:
            continue

        init_iso = result["init_date"].isoformat()

        if init_iso not in period_data:
            period_data[init_iso] = {
                "print_date": result["print_date"],
                "employees": defaultdict(lambda: {"bw": 0, "color": 0}),
            }
        else:
            if result["print_date"] > period_data[init_iso]["print_date"]:
                period_data[init_iso]["print_date"] = result["print_date"]

        for uid, counts in result["employees"].items():
            period_data[init_iso]["employees"][uid]["bw"] += counts["bw"]
            period_data[init_iso]["employees"][uid]["color"] += counts["color"]

    periods = []
    for init_iso, data in sorted(period_data.items()):
        init_date = date.fromisoformat(init_iso)
        end_date = data["print_date"] - timedelta(days=1)
        label = "{:02d}/{:02d}-{:02d}/{:02d}".format(
            init_date.month, init_date.day, end_date.month, end_date.day
        )
        periods.append({
            "start_date": init_iso,
            "label": label,
            "employees": dict(data["employees"]),
        })

    return periods


def merge_periods(all_periods):
    # type: (list) -> list
    """
    合併多份 PDF 解析結果。
    相同 start_date 的 period 合併：
      - 同一 PDF 內的頁面已在 parse_pdf 加總，此處處理跨檔案情況。
      - 各機器報表顯示的是累積總量，跨檔案同一員工取最大值（避免重複計算）。
    依 start_date 排序後回傳。
    """
    merged = {}

    for periods in all_periods:
        for period in periods:
            sd = period["start_date"]
            if sd not in merged:
                merged[sd] = {
                    "start_date": sd,
                    "label": period["label"],
                    "employees": {},
                }
            for uid, counts in period["employees"].items():
                if uid not in merged[sd]["employees"]:
                    merged[sd]["employees"][uid] = {"bw": counts["bw"], "color": counts["color"]}
                else:
                    existing = merged[sd]["employees"][uid]
                    existing["bw"] = max(existing["bw"], counts["bw"])
                    existing["color"] = max(existing["color"], counts["color"])

    result = []
    for sd in sorted(merged.keys()):
        p = merged[sd]
        result.append({
            "start_date": p["start_date"],
            "label": p["label"],
            "employees": dict(p["employees"]),
        })
    return result


def build_employee_report(periods):
    # type: (list) -> list
    """
    以員工為單位，整合各期資料，回傳：
    [
        {
            "employee_id": "720629",
            "name": "",
            "periods": [
                {
                    "label": "06/05-07/01",
                    "bw": 403,
                    "color": 28,
                    "bw_cost": 1209,
                    "color_cost": 280,
                    "subtotal": 1489
                },
                ...
            ],
            "total": 4070
        },
        ...
    ]
    """
    all_uids = set()
    for period in periods:
        all_uids.update(period["employees"].keys())

    employees = []
    for uid in sorted(all_uids):
        emp_periods = []
        total = 0
        for period in periods:
            counts = period["employees"].get(uid, {"bw": 0, "color": 0})
            bw = counts["bw"]
            color = counts["color"]
            bw_cost = bw * BW_UNIT_PRICE
            color_cost = color * COLOR_UNIT_PRICE
            subtotal = bw_cost + color_cost
            total += subtotal
            emp_periods.append({
                "label": period["label"],
                "bw": bw,
                "color": color,
                "bw_cost": bw_cost,
                "color_cost": color_cost,
                "subtotal": subtotal,
            })
        employees.append({
            "employee_id": uid,
            "name": "",
            "periods": emp_periods,
            "total": total,
        })

    return employees


def calculate_cost(bw, color):
    # type: (int, int) -> dict
    bw_cost = bw * BW_UNIT_PRICE
    color_cost = color * COLOR_UNIT_PRICE
    return {"bw_cost": bw_cost, "color_cost": color_cost, "total": bw_cost + color_cost}
