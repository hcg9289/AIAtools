#!/usr/bin/env python3
import io
import sys
from pathlib import Path

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from gf_medical_pdf import (  # noqa: E402
    PdfPayloadError,
    _official_premium_page_indices,
    build_medical_financing_pdf,
)


def sample_payload(row_count=71):
    rows = []
    cumulative = 0
    for year in range(1, row_count + 1):
        age = 28 + year
        requested = 2250 + year * 10 if year >= 6 else 0
        actual = requested
        cumulative += actual
        rows.append({
            "age": age,
            "policy_year": year,
            "requested_withdrawal": requested,
            "withdrawal_total": actual,
            "cumulative_withdrawal": cumulative,
            "post_guaranteed_cash": 12000 + year * 100,
            "post_reversionary_bonus": year * 25,
            "post_terminal_bonus": 5000 + year * 120,
            "post_surrender_total": 17000 + year * 245,
            "status": "已提款" if requested else "正常",
        })
    return {
        "result": {
            "input": {"issueAge": 28, "annual": 4714, "total": 23570, "basic": 47617},
            "rows": rows,
        },
        "medicalContext": {
            "source": "table",
            "planName": "AIA 自願醫保睿選計劃",
            "formName": "基本計劃",
            "deductibleLabel": "HKD 18,000／USD 2,250",
            "effectiveDate": "2025-10-27",
        },
    }


def main():
    deductible_pages = {
        "HKD 0／USD 0": (0, 1),
        "HKD 8,800／USD 1,100": (2, 3),
        "HKD 18,000／USD 2,250": (4, 5),
        "HKD 30,000／USD 3,750": (6, 7),
        "HKD 55,000／USD 6,875": (8, 9),
    }
    for label, basic_pages in deductible_pages.items():
        context = {
            "source": "table",
            "planName": "AIA 自願醫保睿選計劃",
            "formName": "基本計劃",
            "deductibleLabel": label,
        }
        assert _official_premium_page_indices(context) == basic_pages
        context["formName"] = "附加契約"
        assert _official_premium_page_indices(context) == tuple(page + 10 for page in basic_pages)
    print("PASS 十組基本／附加契約及自付費官方頁面映射")

    pdf = build_medical_financing_pdf(sample_payload())
    raw = pdf.getvalue()
    assert raw.startswith(b"%PDF-"), "輸出不是 PDF"
    reader = PdfReader(io.BytesIO(raw))
    assert len(reader.pages) == 6, f"71 行應輸出 3 頁列表 + 2 頁官方保費表 + 總結，共 6 頁，實際 {len(reader.pages)}"
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    for expected in ("GF 醫療融資逐年列表", "基本計劃標準保費表", "18,000港元", "2,250美元自付費", "醫療融資方案", "累計保費", "可提取現金價值"):
        assert expected in text, f"PDF 欠缺：{expected}"
    first_official = reader.pages[3].extract_text() or ""
    continued_official = reader.pages[4].extract_text() or ""
    assert "（續）" not in first_official, "官方保費表首頁錯誤地使用續頁"
    assert "（續）" in continued_official, "官方保費表欠缺 61 至 99+ 歲續頁"
    assert "61" in continued_official and "99+" in continued_official, "官方保費表續頁年齡不完整"
    assert text.count("第 ") >= 4, "系統生成頁欠缺頁碼"
    xobjects = (reader.pages[-1].get("/Resources") or {}).get("/XObject") or {}
    images = [obj.get_object() for obj in xobjects.values() if obj.get_object().get("/Subtype") == "/Image"]
    assert len(images) >= 2, "總結頁欠缺兩個原創醫療公仔"
    print("PASS 動態列表分頁、兩頁官方保費表及總結頁")

    manual = sample_payload()
    manual["medicalContext"] = {"source": "manual"}
    manual_reader = PdfReader(build_medical_financing_pdf(manual))
    assert len(manual_reader.pages) == 4, "手動輸入不應加入不適用的官方保費表"
    manual_text = "\n".join(page.extract_text() or "" for page in manual_reader.pages)
    assert "基本計劃標準保費表" not in manual_text, "手動輸入錯誤地加入官方保費表"
    print("PASS 只有選用官方費率時才加入對應保費表")

    invalid = sample_payload()
    invalid["result"]["rows"] *= 2
    try:
        build_medical_financing_pdf(invalid)
    except PdfPayloadError:
        pass
    else:
        raise AssertionError("沒有拒絕超過 100 行的資料")
    print("PASS PDF payload 行數限制")


if __name__ == "__main__":
    main()
