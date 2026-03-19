"""
receipt.py — 產生員工列印費用收款單 PDF（橫式 A4，ReportLab）
"""
from __future__ import annotations

import io
import os
import zipfile
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ── 中文字型 ──────────────────────────────────────────────
_FONT_NAME = "Helvetica"
_FONT_ERRORS = []

_FONT_CANDIDATES = [
    ("/System/Library/Fonts/STHeiti Light.ttc", 0),
    ("/System/Library/Fonts/STHeiti Medium.ttc", 0),
    ("/Library/Fonts/Arial Unicode MS.ttf", None),
    ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 0),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
    ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", 0),
]

for _path, _idx in _FONT_CANDIDATES:
    if os.path.exists(_path):
        try:
            if _idx is not None:
                pdfmetrics.registerFont(TTFont("CJK", _path, subfontIndex=_idx))
            else:
                pdfmetrics.registerFont(TTFont("CJK", _path))
            _FONT_NAME = "CJK"
            break
        except Exception as _e:
            _FONT_ERRORS.append("{}: {}".format(_path, _e))


# ── Styles ─────────────────────────────────────────────────

def _style(name, **kwargs):
    # type: (str, **object) -> ParagraphStyle
    base = getSampleStyleSheet()["Normal"]
    kw = {"fontName": _FONT_NAME}
    kw.update(kwargs)
    return ParagraphStyle(name, parent=base, **kw)


_S_HEADER    = _style("header",    fontSize=11, leading=16)
_S_HEADER_B  = _style("header_b",  fontSize=11, leading=16, fontName=_FONT_NAME)
_S_YEAR      = _style("year",      fontSize=11, leading=16, spaceAfter=4)
_S_TH        = _style("th",        fontSize=9,  textColor=colors.white, alignment=1)
_S_TH_SMALL  = _style("th_small",  fontSize=8,  textColor=colors.white, alignment=1)
_S_TD_C      = _style("tdc",       fontSize=9,  alignment=1)
_S_TD_C_B    = _style("tdc_b",     fontSize=9,  alignment=1)
_S_TD_R      = _style("tdr",       fontSize=9,  alignment=2)
_S_TD_R_B    = _style("tdr_b",     fontSize=9,  alignment=2)
_S_TOTAL     = _style("total",     fontSize=10, textColor=colors.HexColor("#c0392b"), alignment=2)
_S_FOOTER    = _style("footer",    fontSize=8,  textColor=colors.grey, alignment=2)
_S_NOTE      = _style("note",      fontSize=10, textColor=colors.HexColor("#2980b9"))


def _p(text, style):
    # type: (str, ParagraphStyle) -> Paragraph
    return Paragraph(str(text), style)


def generate_receipt_pdf(employee, year_label, bank_info):
    # type: (dict, str, dict) -> bytes
    """
    回傳單一員工收款單的 PDF bytes（橫式 A4）。

    employee: {employee_id, name, periods, total}
    year_label: e.g. "114"
    bank_info: {holder, bank, account}
    """
    buf = io.BytesIO()
    page_w, page_h = landscape(A4)
    margin = 15 * mm
    usable_w = page_w - 2 * margin

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="列印費用收款單 - {}".format(employee["employee_id"]),
    )

    story = []

    # ── 標頭資訊（表格外） ──────────────────────────────────
    display_name = employee.get("name") or employee["employee_id"]
    holder = bank_info.get("holder", "")
    bank = bank_info.get("bank", "")
    account = bank_info.get("account", "")
    total = employee.get("total", 0)

    label_w = usable_w * 0.2
    value_w = usable_w * 0.8
    header_data = [
        [_p("{}：".format(display_name), _S_HEADER), _p("", _S_HEADER)],
        [_p("匯款帳戶：", _S_HEADER), _p(holder, _S_HEADER)],
        [_p("匯款銀行：", _S_HEADER), _p(bank, _S_HEADER)],
        [_p("匯款帳號：", _S_HEADER), _p(account, _S_HEADER)],
        [_p("合計匯款金額：", _S_HEADER), _p("${}".format(total), _S_HEADER)],
    ]
    header_table = Table(header_data, colWidths=[label_w, value_w])
    header_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), _FONT_NAME),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(header_table)
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc")))
    story.append(Spacer(1, 3 * mm))

    # ── 年度標籤 ──────────────────────────────────────────
    story.append(_p("{}年".format(year_label), _S_YEAR))

    # ── 期別表格 ──────────────────────────────────────────
    periods = employee.get("periods", [])
    n = len(periods)

    if n == 0:
        # 無資料時顯示提示
        story.append(_p("（無列印資料）", _S_TD_C))
        doc.build(story)
        buf.seek(0)
        return buf.read()

    num_cols = 2 * n  # 每個期別佔 2 欄（黑白 + 彩色）

    # 動態字型大小：欄數多時縮小
    if n > 6:
        td_size = 7
        th_size = 7
    elif n > 4:
        td_size = 8
        th_size = 8
    else:
        td_size = 9
        th_size = 9

    def _th(text):
        s = _style("_th_{}".format(text), fontSize=th_size, textColor=colors.white, alignment=1)
        return Paragraph(str(text), s)

    def _td(text, bold=False, align="center"):
        amap = {"center": 1, "right": 2, "left": 0}
        s = _style(
            "_td_{}_{}".format(text, bold),
            fontSize=td_size,
            alignment=amap.get(align, 1),
        )
        return Paragraph(str(text), s)

    col_width = usable_w / num_cols

    # Row 0: display name (or employee_id if no name) spanning all columns
    row0 = [_th(display_name)] + [""] * (num_cols - 1)

    # Row 1: period labels, each spanning 2 columns
    row1 = []
    for p in periods:
        row1.append(_th(p["label"]))
        row1.append("")

    # Row 2: "A4:黑白" | "A4:全彩" repeated
    row2 = []
    for _ in periods:
        row2.append(_th("A4:黑白"))
        row2.append(_th("A4:全彩"))

    # Row 3: unit prices
    row3 = []
    for _ in periods:
        row3.append(_th("$3"))
        row3.append(_th("$10"))

    # Row 4: counts
    row4 = []
    for p in periods:
        row4.append(_td(str(p["bw"])))
        row4.append(_td(str(p["color"])))

    # Row 5: costs
    row5 = []
    for p in periods:
        row5.append(_td("${}".format(p["bw_cost"])))
        row5.append(_td("${}".format(p["color_cost"])))

    # Row 6: subtotal spanning (2N-1) cols + total amount in last col
    grand_total = employee.get("total", 0)
    row6 = [""] * (num_cols - 1) + [_td("${}".format(grand_total), align="right")]
    # Put 小計 label in first cell
    row6[0] = _td("小計：", align="right")

    table_data = [row0, row1, row2, row3, row4, row5, row6]

    col_widths = [col_width] * num_cols
    period_table = Table(table_data, colWidths=col_widths, repeatRows=3)

    # Build TableStyle commands
    ts_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), _FONT_NAME),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        # Header rows background
        ("BACKGROUND", (0, 0), (-1, 3), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 3), colors.white),
        # Data rows alternating background
        ("ROWBACKGROUNDS", (0, 4), (-1, 5), [colors.white, colors.HexColor("#f7f9fc")]),
        # Subtotal row
        ("BACKGROUND", (0, 6), (-1, 6), colors.HexColor("#eaf0fb")),
        ("FONTSIZE", (0, 6), (-1, 6), td_size),
        # Grid
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cccccc")),
        ("LINEBELOW", (0, 3), (-1, 3), 1, colors.HexColor("#2c3e50")),
        # Row 0: span all
        ("SPAN", (0, 0), (num_cols - 1, 0)),
        # Row 6: span first (2N-1) cells for 小計 label
        ("SPAN", (0, 6), (num_cols - 2, 6)),
        ("ALIGN", (num_cols - 2, 6), (num_cols - 1, 6), "RIGHT"),
        ("FONTSIZE", (num_cols - 1, 6), (num_cols - 1, 6), td_size + 1),
        ("TEXTCOLOR", (num_cols - 1, 6), (num_cols - 1, 6), colors.HexColor("#c0392b")),
    ]

    # Row 1: span each pair of columns for period labels
    for i, _ in enumerate(periods):
        ts_cmds.append(("SPAN", (2 * i, 1), (2 * i + 1, 1)))

    period_table.setStyle(TableStyle(ts_cmds))
    story.append(period_table)

    story.append(Spacer(1, 4 * mm))
    story.append(_p("黑白一張 $3　彩色一張 $10", _S_NOTE))
    story.append(Spacer(1, 2 * mm))
    today = date.today().strftime("%Y/%m/%d")
    story.append(_p("列印日期：{}".format(today), _S_FOOTER))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def generate_zip(employees, year_label, bank_info):
    # type: (list, str, dict) -> bytes
    """將所有員工收款單打包成 ZIP，回傳 bytes。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for emp in employees:
            pdf_bytes = generate_receipt_pdf(emp, year_label, bank_info)
            filename = "receipt_{}.pdf".format(emp["employee_id"])
            zf.writestr(filename, pdf_bytes)
    buf.seek(0)
    return buf.read()
