#!/usr/bin/env python3
import io
import sys
from pathlib import Path

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from gf_medical_pdf import PdfPayloadError, build_medical_financing_pdf  # noqa: E402


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
    pdf = build_medical_financing_pdf(sample_payload())
    raw = pdf.getvalue()
    assert raw.startswith(b"%PDF-"), "輸出不是 PDF"
    reader = PdfReader(io.BytesIO(raw))
    assert len(reader.pages) == 5, f"71 行應輸出 3 頁列表 + 保費表 + 總結，共 5 頁，實際 {len(reader.pages)}"
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    for expected in ("GF 醫療融資逐年列表", "醫療保費表", "醫療融資方案", "累計保費", "可提取現金價值"):
        assert expected in text, f"PDF 欠缺：{expected}"
    assert text.count("第 ") >= 5, "欠缺頁碼"
    xobjects = (reader.pages[-1].get("/Resources") or {}).get("/XObject") or {}
    images = [obj.get_object() for obj in xobjects.values() if obj.get_object().get("/Subtype") == "/Image"]
    assert len(images) >= 2, "總結頁欠缺兩個原創醫療公仔"
    print("PASS 動態列表分頁、保費表及總結頁")

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
