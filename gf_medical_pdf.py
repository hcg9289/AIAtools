import io
import math
import os
from datetime import datetime, timezone

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from pypdf import PdfReader, PdfWriter


FONT_NAME = "GFTraditionalChinese"
FONT_CANDIDATES = (
    os.environ.get("GF_PDF_FONT_PATH", ""),
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
)
for _font_path in FONT_CANDIDATES:
    if _font_path and os.path.isfile(_font_path):
        pdfmetrics.registerFont(TTFont(FONT_NAME, _font_path, subfontIndex=0))
        break
else:  # pragma: no cover - Docker image and supported macOS both provide a candidate
    raise RuntimeError("找不到可嵌入 PDF 的繁體中文字型。")
pdfmetrics.registerFontFamily(
    FONT_NAME,
    normal=FONT_NAME,
    bold=FONT_NAME,
    italic=FONT_NAME,
    boldItalic=FONT_NAME,
)

NAVY = colors.HexColor("#123044")
TEAL = colors.HexColor("#0B887C")
PALE_BLUE = colors.HexColor("#DDEBF7")
BLUE = colors.HexColor("#9DC3E6")
PALE_TEAL = colors.HexColor("#E7F4F1")
YELLOW = colors.HexColor("#FFF2A8")
GRID = colors.HexColor("#82918D")
TEXT = colors.HexColor("#17241F")
MUTED = colors.HexColor("#5C6864")
PROJECTION_ROWS_PER_PAGE = 31
ASSET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "gf")
OFFICIAL_PREMIUM_TABLE_PATH = os.path.join(
    ASSET_DIR,
    "aia_vhis_selectwise_premium_table_2025-10-27.pdf",
)
OFFICIAL_PLAN_MARKER = "AIA自願醫保睿選計劃"
DEDUCTIBLE_PAGE_OFFSETS = {
    "0/0": 0,
    "8800/1100": 2,
    "18000/2250": 4,
    "30000/3750": 6,
    "55000/6875": 8,
}


class PdfPayloadError(ValueError):
    pass


def _number(value, name, *, minimum=0):
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise PdfPayloadError(f"{name} 必須是數字。") from exc
    if not math.isfinite(parsed) or parsed < minimum:
        raise PdfPayloadError(f"{name} 不正確。")
    return parsed


def _integer(value, name, *, minimum=0, maximum=999):
    parsed = _number(value, name, minimum=minimum)
    if not parsed.is_integer() or parsed > maximum:
        raise PdfPayloadError(f"{name} 必須是 {minimum} 至 {maximum} 的整數。")
    return int(parsed)


def _text(value, name, *, maximum=160, fallback=""):
    if value is None:
        return fallback
    text = str(value).strip()
    if len(text) > maximum:
        raise PdfPayloadError(f"{name} 過長。")
    return text or fallback


def validate_medical_financing_payload(payload):
    if not isinstance(payload, dict):
        raise PdfPayloadError("PDF 資料格式不正確。")
    raw_result = payload.get("result")
    if not isinstance(raw_result, dict):
        raise PdfPayloadError("欠缺醫療融資結果。")
    raw_input = raw_result.get("input")
    raw_rows = raw_result.get("rows")
    if not isinstance(raw_input, dict) or not isinstance(raw_rows, list):
        raise PdfPayloadError("醫療融資結果不完整。")
    if not 1 <= len(raw_rows) <= 100:
        raise PdfPayloadError("逐年結果必須包含 1 至 100 行。")

    issue_age = _integer(raw_input.get("issueAge"), "投保年齡", maximum=99)
    medical_start_age = _integer(raw_input.get("medicalStartAge"), "醫療保費開始年齡", maximum=99)
    medical_end_age = _integer(raw_input.get("medicalEndAge"), "醫療保費終止年齡", maximum=99)
    if not issue_age < medical_start_age <= medical_end_age:
        raise PdfPayloadError("醫療保費年期必須在投保年齡之後，且開始年齡不可遲於終止年齡。")
    payment_mode = _text(
        raw_input.get("paymentMode"),
        "GF供款方式",
        maximum=20,
        fallback="five_year",
    ).lower()
    if payment_mode not in {"five_year", "single"}:
        raise PdfPayloadError("GF供款方式只可為五年供款或一次性繳費。")
    annual = _number(raw_input.get("annual"), "年繳保費", minimum=0.01)
    total = _number(raw_input.get("total"), "五年總保費", minimum=0.01)
    basic = _number(raw_input.get("basic"), "投保時基本金額", minimum=0)
    rows = []
    for index, raw in enumerate(raw_rows, 1):
        if not isinstance(raw, dict):
            raise PdfPayloadError(f"第 {index} 行格式不正確。")
        row = {
            "age": _integer(raw.get("age"), f"第 {index} 行年齡", maximum=99),
            "policy_year": _integer(raw.get("policy_year"), f"第 {index} 行保單年度", minimum=1, maximum=100),
            "requested_withdrawal": _number(raw.get("requested_withdrawal", 0), f"第 {index} 行醫療保費"),
            "withdrawal_total": _number(raw.get("withdrawal_total", 0), f"第 {index} 行實際提款"),
            "cumulative_withdrawal": _number(raw.get("cumulative_withdrawal", 0), f"第 {index} 行累計提款"),
            "post_guaranteed_cash": _number(raw.get("post_guaranteed_cash", 0), f"第 {index} 行保證現金價值"),
            "post_reversionary_bonus": _number(raw.get("post_reversionary_bonus", 0), f"第 {index} 行復歸紅利"),
            "post_terminal_bonus": _number(raw.get("post_terminal_bonus", 0), f"第 {index} 行終期紅利"),
            "post_surrender_total": _number(raw.get("post_surrender_total", 0), f"第 {index} 行總額"),
            "status": _text(raw.get("status"), f"第 {index} 行狀態", maximum=24, fallback="正常"),
        }
        rows.append(row)
    rows.sort(key=lambda row: row["policy_year"])
    if len({row["policy_year"] for row in rows}) != len(rows):
        raise PdfPayloadError("保單年度不可重複。")
    expected_count = medical_end_age - issue_age
    if len(rows) != expected_count:
        raise PdfPayloadError(f"逐年結果必須完整涵蓋 {issue_age + 1} 至 {medical_end_age} 歲。")
    for index, row in enumerate(rows, 1):
        expected_age = issue_age + index
        if row["age"] != expected_age or row["policy_year"] != index:
            raise PdfPayloadError(f"逐年結果欠缺或錯置 {expected_age} 歲資料。")

    raw_context = payload.get("medicalContext") or {}
    if not isinstance(raw_context, dict):
        raise PdfPayloadError("醫療保費表資料格式不正確。")
    context = {
        "source": _text(raw_context.get("source"), "保費來源", maximum=20, fallback="manual"),
        "planName": _text(raw_context.get("planName"), "計劃名稱", fallback="自訂醫療計劃"),
        "formName": _text(raw_context.get("formName"), "保單形式", fallback="醫療計劃"),
        "deductibleLabel": _text(raw_context.get("deductibleLabel"), "自付費", fallback="自訂自付費"),
        "effectiveDate": _text(raw_context.get("effectiveDate"), "保費表日期", maximum=24, fallback="未提供"),
    }
    premium_by_age = {}
    if context["source"].lower() == "table":
        raw_rates = raw_context.get("premiumByAge")
        if not isinstance(raw_rates, list) or len(raw_rates) != 100:
            raise PdfPayloadError("欠缺 0 至 99 歲的完整官方醫療保費率。")
        for index, raw_rate in enumerate(raw_rates):
            if not isinstance(raw_rate, dict):
                raise PdfPayloadError(f"第 {index + 1} 項官方醫療保費率格式不正確。")
            age = _integer(raw_rate.get("age"), f"第 {index + 1} 項保費年齡", maximum=99)
            if age in premium_by_age:
                raise PdfPayloadError("官方醫療保費率年齡不可重複。")
            premium_by_age[age] = _number(raw_rate.get("premium"), f"{age} 歲官方醫療保費", minimum=0.01)
        if set(premium_by_age) != set(range(100)):
            raise PdfPayloadError("官方醫療保費率必須完整涵蓋 0 至 99 歲。")
    context["premiumByAge"] = premium_by_age
    return {
        "input": {
            "issueAge": issue_age,
            "medicalStartAge": medical_start_age,
            "medicalEndAge": medical_end_age,
            "annual": annual,
            "total": total,
            "basic": basic,
            "paymentMode": payment_mode,
        },
        "rows": rows,
        "context": context,
        "generatedAt": datetime.now(timezone.utc),
    }


def _money(value):
    return f"USD {float(value):,.0f}"


def _cell_text(c, text, x, y, width, height, *, size=8, color=TEXT, align=TA_CENTER, bold=False):
    style = ParagraphStyle(
        "cell",
        fontName=FONT_NAME,
        fontSize=size,
        leading=size * 1.28,
        textColor=color,
        alignment=align,
        spaceAfter=0,
        spaceBefore=0,
    )
    safe = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
    if bold:
        safe = f"<b>{safe}</b>"
    paragraph = Paragraph(safe, style)
    _, ph = paragraph.wrap(width - 5 * mm, height - 2 * mm)
    paragraph.drawOn(c, x + 2.5 * mm, y + max(1 * mm, (height - ph) / 2))


def _page_header(c, page_size, title, subtitle, page_no, total_pages):
    width, height = page_size
    c.setFillColor(NAVY)
    c.rect(0, height - 18 * mm, width, 18 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont(FONT_NAME, 14)
    c.drawString(14 * mm, height - 11.5 * mm, title)
    c.setFont(FONT_NAME, 8)
    c.drawRightString(width - 14 * mm, height - 11.5 * mm, subtitle)
    c.setFillColor(MUTED)
    c.setFont(FONT_NAME, 7)
    c.drawRightString(width - 12 * mm, 7 * mm, f"第 {page_no} / {total_pages} 頁")


def _draw_grid(c, x, top, widths, row_heights, fills=None, line_width=0.45):
    y = top
    for row_index, height in enumerate(row_heights):
        y -= height
        cursor = x
        for col_index, width in enumerate(widths):
            fill = fills.get((row_index, col_index)) if fills else None
            if fill:
                c.setFillColor(fill)
                c.rect(cursor, y, width, height, fill=1, stroke=0)
            cursor += width
    total_width = sum(widths)
    total_height = sum(row_heights)
    c.setStrokeColor(GRID)
    c.setLineWidth(line_width)
    c.rect(x, top - total_height, total_width, total_height, fill=0, stroke=1)
    cursor = x
    for width in widths[:-1]:
        cursor += width
        c.line(cursor, top, cursor, top - total_height)
    cursor_y = top
    for height in row_heights[:-1]:
        cursor_y -= height
        c.line(x, cursor_y, x + total_width, cursor_y)
    return top - total_height


def _draw_projection_pages(c, data, start_page, total_pages):
    page_size = landscape(A4)
    width, height = page_size
    rows = data["rows"]
    rows_per_page = PROJECTION_ROWS_PER_PAGE
    chunks = [rows[index:index + rows_per_page] for index in range(0, len(rows), rows_per_page)]
    col_widths = [22 * mm, 23 * mm, 34 * mm, 33 * mm, 31 * mm, 31 * mm, 31 * mm, 33 * mm, 27 * mm]
    labels = ["年齡", "保單年度", "實際提款\n醫療保費", "保證現金價值", "復歸紅利", "終期紅利", "總額", "狀態", "累計提款"]
    for chunk_index, chunk in enumerate(chunks):
        c.setPageSize(page_size)
        page_no = start_page + chunk_index
        _page_header(c, page_size, "GF 醫療融資逐年列表", "現金提取後之保單價值", page_no, total_pages)
        x = (width - sum(col_widths)) / 2
        top = height - 27 * mm
        header_height = 12 * mm
        row_height = 4.75 * mm
        fills = {(0, index): PALE_BLUE for index in range(len(col_widths))}
        fills[(0, len(col_widths) - 1)] = YELLOW
        _draw_grid(c, x, top, col_widths, [header_height] + [row_height] * len(chunk), fills)
        cursor = x
        for index, (label, col_width) in enumerate(zip(labels, col_widths)):
            _cell_text(c, label, cursor, top - header_height, col_width, header_height, size=7.2, bold=True)
            cursor += col_width
        y = top - header_height
        for row in chunk:
            y -= row_height
            values = [
                row["age"], row["policy_year"], _money(row["withdrawal_total"]),
                _money(row["post_guaranteed_cash"]), _money(row["post_reversionary_bonus"]),
                _money(row["post_terminal_bonus"]), _money(row["post_surrender_total"]),
                row["status"], _money(row["cumulative_withdrawal"]),
            ]
            cursor = x
            for index, (value, col_width) in enumerate(zip(values, col_widths)):
                if index == len(values) - 1:
                    c.setFillColor(YELLOW)
                    c.rect(cursor, y, col_width, row_height, fill=1, stroke=0)
                _cell_text(c, value, cursor, y, col_width, row_height, size=6.1, bold=index in (6, 8))
                cursor += col_width
        c.showPage()
    return len(chunks)


def _premium_rows(data):
    return [row for row in data["rows"] if row["requested_withdrawal"] > 0]


def _compact_label(value):
    return "".join(character for character in str(value) if character.isalnum())


def _deductible_key(label):
    compact = _compact_label(label)
    for key in sorted(DEDUCTIBLE_PAGE_OFFSETS, key=len, reverse=True):
        hkd, usd = key.split("/")
        if hkd in compact and usd in compact:
            return key
    return None


def _official_premium_page_indices(context):
    if context["source"].lower() != "table":
        return ()
    if OFFICIAL_PLAN_MARKER not in _compact_label(context["planName"]):
        raise PdfPayloadError("目前沒有這個醫療計劃的官方保費表。")

    form_name = context["formName"]
    if "附加" in form_name:
        form_offset = 10
    elif "基本" in form_name:
        form_offset = 0
    else:
        raise PdfPayloadError("請先選擇基本計劃或附加契約，才可加入官方保費表。")

    deductible = _deductible_key(context["deductibleLabel"])
    if deductible is None:
        raise PdfPayloadError("目前沒有這個自付費選項的官方保費表。")
    first_page = form_offset + DEDUCTIBLE_PAGE_OFFSETS[deductible]
    return (first_page, first_page + 1)


def _insert_official_premium_pages(generated_pdf, premium_page_indices):
    generated = PdfReader(generated_pdf)
    if not premium_page_indices:
        generated_pdf.seek(0)
        return generated_pdf
    if not os.path.isfile(OFFICIAL_PREMIUM_TABLE_PATH):
        raise PdfPayloadError("官方醫療保費表檔案尚未安裝。")

    official = PdfReader(OFFICIAL_PREMIUM_TABLE_PATH)
    if max(premium_page_indices) >= len(official.pages):
        raise PdfPayloadError("官方醫療保費表頁面不完整。")

    writer = PdfWriter()
    writer.add_metadata({
        "/Title": "GF 醫療融資方案",
        "/Author": "AIAtools",
    })
    writer.append(official, pages=list(premium_page_indices), import_outline=False)
    for page in generated.pages:
        writer.add_page(page)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output


def _value_at_age(rows, age, field):
    candidates = [row for row in rows if row["age"] <= age]
    return candidates[-1][field] if candidates else 0


def _premium_at_age(data, age):
    official_rate = data["context"]["premiumByAge"].get(age)
    if official_rate is not None:
        return official_rate
    row = next((item for item in data["rows"] if item["age"] == age), None)
    return row["requested_withdrawal"] if row else 0


def _medical_paid_to_age(data, age):
    issue_age = data["input"]["issueAge"]
    return sum(_premium_at_age(data, premium_age) for premium_age in range(issue_age + 1, age + 1))


def _combo_paid_to_age(data, age):
    issue_age = data["input"]["issueAge"]
    if data["input"].get("paymentMode") == "single":
        gf_paid = data["input"]["total"] if age >= issue_age else 0
    else:
        completed_years = max(0, min(5, age - issue_age))
        gf_paid = data["input"]["annual"] * completed_years
    medical_paid = _medical_paid_to_age(data, age)
    financed_medical = sum(
        min(_premium_at_age(data, row["age"]), row["withdrawal_total"])
        for row in data["rows"] if row["age"] <= age and row["requested_withdrawal"] > 0
    )
    return gf_paid + max(0, medical_paid - financed_medical)


def _milestone_ages(issue_age, first_standard_age, end_age=99):
    first = first_standard_age if issue_age < first_standard_age else min(end_age, issue_age + 5)
    return sorted({age for age in (first, 65, 85, end_age) if issue_age < age <= end_age})


def _cash_value_milestones(issue_age, end_age=99):
    candidates = [
        ("第 10 個保單年度", issue_age + 10),
        ("第 20 個保單年度", issue_age + 20),
        ("至 65 歲", 65),
        ("至 85 歲", 85),
        (f"至 {end_age} 歲", end_age),
    ]
    milestones = []
    seen_ages = set()
    for label, age in candidates:
        if issue_age < age <= end_age and age not in seen_ages:
            milestones.append((label, age))
            seen_ages.add(age)
    return milestones


def _delayed_financing_start_age(data):
    first_financing_row = next(
        (row for row in data["rows"] if row["requested_withdrawal"] > 0),
        None,
    )
    if first_financing_row is None:
        return None
    immediate_start_year = 2 if data["input"].get("paymentMode") == "single" else 6
    if first_financing_row["policy_year"] <= immediate_start_year:
        return None
    return first_financing_row["age"]


def _draw_summary_page(c, data, page_no, total_pages):
    page_size = A4
    c.setPageSize(page_size)
    width, height = page_size
    issue_age = data["input"]["issueAge"]
    end_age = data["input"]["medicalEndAge"]
    context = data["context"]
    first_medical = _premium_at_age(data, min(99, issue_age + 1))
    plan_label = f"{context['formName']} ({context['deductibleLabel']})"
    _page_header(c, page_size, f"{issue_age} 歲 · 醫療融資方案", context["planName"], page_no, total_pages)

    x = 20 * mm
    table_width = width - 40 * mm
    widths = [42 * mm, (table_width - 42 * mm) / 2, (table_width - 42 * mm) / 2]
    top = height - 28 * mm

    row_heights = [13 * mm, 22 * mm]
    fills = {(0, 1): PALE_BLUE, (0, 2): BLUE, (1, 1): PALE_BLUE, (1, 2): BLUE}
    y = _draw_grid(c, x, top, widths, row_heights, fills, line_width=0.8)
    _cell_text(c, "計劃比較", x, top - row_heights[0], widths[0], row_heights[0], size=9, bold=True)
    _cell_text(c, "基本計劃", x + widths[0], top - row_heights[0], widths[1], row_heights[0], size=10, bold=True)
    _cell_text(c, "組合計劃", x + widths[0] + widths[1], top - row_heights[0], widths[2], row_heights[0], size=10, bold=True)
    _cell_text(c, "內容", x, y, widths[0], row_heights[1], size=8, bold=True)
    _cell_text(c, plan_label, x + widths[0], y, widths[1], row_heights[1], size=8)
    single_payment = data["input"].get("paymentMode") == "single"
    savings_label = (
        f"一次性 {_money(data['input']['total'])} 儲蓄"
        if single_payment
        else f"每年 {_money(data['input']['annual'])} 儲蓄"
    )
    combo_label = f"{savings_label}\n+\n{plan_label}"
    _cell_text(c, combo_label, x + widths[0] + widths[1], y, widths[2], row_heights[1], size=7.5)

    coverage_height = 49 * mm
    coverage_top = y
    coverage_widths = [widths[0], widths[1] + widths[2]]
    y = _draw_grid(c, x, coverage_top, coverage_widths, [coverage_height], {(0, 1): colors.white}, line_width=0.8)
    _cell_text(c, "保障內容", x, y, widths[0], coverage_height, size=8, bold=True)
    coverage = (
        "保障範圍：普通房／半私家房\n"
        "每年入院治療最高保障：HKD 12,000,000\n\n"
        "每日病房及膳食賠償：全數賠償*\n"
        "外科手術費：不論手術分類全數賠償*\n"
        "長期治療、化療（包括標靶治療）及電療：全數賠償*"
    )
    _cell_text(c, coverage, x + widths[0], y, widths[1] + widths[2], coverage_height, size=7.4, align=TA_LEFT)

    contribution_height = 18 * mm
    contribution_top = y
    fills = {(0, 1): YELLOW, (0, 2): YELLOW}
    y = _draw_grid(c, x, contribution_top, widths, [contribution_height], fills, line_width=0.8)
    _cell_text(c, "供款年期\n首年保費", x, y, widths[0], contribution_height, size=8, bold=True)
    _cell_text(c, f"至 {end_age} 歲\n{_money(first_medical)}", x + widths[0], y, widths[1], contribution_height, size=8)
    contribution_label = (
        f"一次性\n{_money(data['input']['total'] + first_medical)}"
        if single_payment
        else f"5 年\n{_money(data['input']['annual'] + first_medical)}"
    )
    contribution_x = x + widths[0] + widths[1]
    delayed_start_age = _delayed_financing_start_age(data)
    if delayed_start_age is None:
        _cell_text(c, contribution_label, contribution_x, y, widths[2], contribution_height, size=8, bold=True)
    else:
        _cell_text(c, contribution_label, contribution_x, y + 5.5 * mm, widths[2], contribution_height - 5.5 * mm, size=8, bold=True)
        _cell_text(
            c,
            f"醫療融資從{delayed_start_age}歲開始",
            contribution_x,
            y + 0.8 * mm,
            widths[2],
            5.2 * mm,
            size=5.8,
            color=colors.HexColor("#C62828"),
        )

    premium_ages = _milestone_ages(issue_age, 25, end_age)
    premium_height = 8.5 * mm
    premium_top = y
    y = _draw_grid(c, x, premium_top, widths, [premium_height] * (len(premium_ages) + 1), {(0, 0): PALE_TEAL, (0, 1): PALE_TEAL, (0, 2): PALE_TEAL}, line_width=0.65)
    for index, label in enumerate(["累計保費", "基本計劃", "組合計劃"]):
        _cell_text(c, label, x + sum(widths[:index]), premium_top - premium_height, widths[index], premium_height, size=8, bold=True)
    row_y = premium_top - premium_height
    for age in premium_ages:
        row_y -= premium_height
        _cell_text(c, f"至 {age} 歲", x, row_y, widths[0], premium_height, size=7.5)
        _cell_text(c, _money(_medical_paid_to_age(data, age)), x + widths[0], row_y, widths[1], premium_height, size=7.5)
        _cell_text(c, _money(_combo_paid_to_age(data, age)), x + widths[0] + widths[1], row_y, widths[2], premium_height, size=7.5, bold=True)

    cash_milestones = _cash_value_milestones(issue_age, end_age)
    cash_top = y
    y = _draw_grid(c, x, cash_top, widths, [premium_height] * (len(cash_milestones) + 1), {(0, 0): PALE_TEAL, (0, 1): PALE_TEAL, (0, 2): PALE_TEAL}, line_width=0.65)
    for index, label in enumerate(["可提取現金價值", "基本計劃", "組合計劃"]):
        _cell_text(c, label, x + sum(widths[:index]), cash_top - premium_height, widths[index], premium_height, size=8, bold=True)
    row_y = cash_top - premium_height
    for label, age in cash_milestones:
        row_y -= premium_height
        _cell_text(c, label, x, row_y, widths[0], premium_height, size=7.5)
        _cell_text(c, "—", x + widths[0], row_y, widths[1], premium_height, size=8)
        _cell_text(c, _money(_value_at_age(data["rows"], age, "post_surrender_total")), x + widths[0] + widths[1], row_y, widths[2], premium_height, size=7.5, bold=True)

    heart_path = os.path.join(ASSET_DIR, "gf_pdf_heart_mascot.png")
    doctor_path = os.path.join(ASSET_DIR, "gf_pdf_doctor.png")
    if os.path.isfile(heart_path):
        c.drawImage(ImageReader(heart_path), 26 * mm, 25 * mm, 48 * mm, 32 * mm, preserveAspectRatio=True, anchor="c", mask="auto")
    if os.path.isfile(doctor_path):
        c.drawImage(ImageReader(doctor_path), 153 * mm, 20 * mm, 29 * mm, 44 * mm, preserveAspectRatio=True, anchor="c", mask="auto")

    c.setFillColor(MUTED)
    c.setFont(FONT_NAME, 6.6)
    c.drawString(x, 15 * mm, "* 保障內容按所提供範本展示，實際保障、保費及續保條款以保險公司最新正式文件為準。")
    c.showPage()


def build_medical_financing_pdf(payload):
    data = validate_medical_financing_payload(payload)
    projection_pages = max(1, math.ceil(len(data["rows"]) / PROJECTION_ROWS_PER_PAGE))
    premium_page_indices = _official_premium_page_indices(data["context"])
    total_pages = projection_pages + len(premium_page_indices) + 1
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=A4, pageCompression=1)
    c.setTitle("GF 醫療融資方案")
    c.setAuthor("AIAtools")
    used = _draw_projection_pages(c, data, len(premium_page_indices) + 1, total_pages)
    _draw_summary_page(c, data, used + len(premium_page_indices) + 1, total_pages)
    c.save()
    output.seek(0)
    return _insert_official_premium_pages(output, premium_page_indices)
