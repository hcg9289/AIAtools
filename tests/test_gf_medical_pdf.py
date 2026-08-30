#!/usr/bin/env python3
import io
import sys
from pathlib import Path

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from gf_medical_pdf import (  # noqa: E402
    PdfPayloadError,
    _cash_value_milestones,
    _combo_paid_to_age,
    _delayed_financing_start_age,
    _medical_paid_to_age,
    _official_premium_page_indices,
    build_medical_financing_pdf,
    validate_medical_financing_payload,
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
            "premiumByAge": [
                {"age": age, "premium": 1970 + age * 10}
                for age in range(100)
            ],
        },
    }


def main():
    assert _cash_value_milestones(28) == [
        ("第 10 個保單年度", 38),
        ("第 20 個保單年度", 48),
        ("至 65 歲", 65),
        ("至 85 歲", 85),
        ("至 99 歲", 99),
    ]
    assert _cash_value_milestones(55) == [
        ("第 10 個保單年度", 65),
        ("第 20 個保單年度", 75),
        ("至 85 歲", 85),
        ("至 99 歲", 99),
    ], "同一歲數的保單年度及固定年齡不應重複顯示"
    print("PASS 可提取現金價值採用第10／20保單年度及65／85／99歲里程碑")

    immediate_five = validate_medical_financing_payload(sample_payload())
    assert _delayed_financing_start_age(immediate_five) is None, "五年供款第6年立即開始不應顯示紅字"
    delayed_five_payload = sample_payload()
    for row in delayed_five_payload["result"]["rows"]:
        if row["policy_year"] < 13:
            row["requested_withdrawal"] = 0
            row["withdrawal_total"] = 0
    delayed_five = validate_medical_financing_payload(delayed_five_payload)
    assert _delayed_financing_start_age(delayed_five) == 41, "五年供款延後至第13年應顯示41歲"

    immediate_single_payload = sample_payload()
    immediate_single_payload["result"]["input"]["paymentMode"] = "single"
    for row in immediate_single_payload["result"]["rows"]:
        row["requested_withdrawal"] = 2250 if row["policy_year"] >= 2 else 0
        row["withdrawal_total"] = row["requested_withdrawal"]
    immediate_single = validate_medical_financing_payload(immediate_single_payload)
    assert _delayed_financing_start_age(immediate_single) is None, "一次性繳費第2年最早開始不應顯示紅字"
    immediate_single_payload["result"]["rows"][1]["requested_withdrawal"] = 0
    immediate_single_payload["result"]["rows"][1]["withdrawal_total"] = 0
    delayed_single = validate_medical_financing_payload(immediate_single_payload)
    assert _delayed_financing_start_age(delayed_single) == 31, "一次性繳費延後至第3年應顯示31歲"
    print("PASS 兩種供款方式只在延後開始醫療融資時顯示年齡")

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
    first_official = reader.pages[0].extract_text() or ""
    continued_official = reader.pages[1].extract_text() or ""
    assert "（續）" not in first_official, "官方保費表首頁錯誤地使用續頁"
    assert "（續）" in continued_official, "官方保費表欠缺 61 至 99+ 歲續頁"
    assert "61" in continued_official and "99+" in continued_official, "官方保費表續頁年齡不完整"
    assert "GF 醫療融資逐年列表" in (reader.pages[2].extract_text() or ""), "逐年列表沒有排在官方保費表之後"
    assert "醫療融資方案" in (reader.pages[-1].extract_text() or ""), "總結表不是最後一部分"
    assert text.count("第 ") >= 4, "系統生成頁欠缺頁碼"
    xobjects = (reader.pages[-1].get("/Resources") or {}).get("/XObject") or {}
    images = [obj.get_object() for obj in xobjects.values() if obj.get_object().get("/Subtype") == "/Image"]
    assert len(images) >= 2, "總結頁欠缺兩個原創醫療公仔"
    print("PASS 動態列表分頁、兩頁官方保費表及總結頁")

    data = validate_medical_financing_payload(sample_payload())
    assert _medical_paid_to_age(data, 33) == 11400, "基本計劃沒有按完整保單年度累計官方醫療保費"
    assert _combo_paid_to_age(data, 33) == 34970, "組合計劃沒有計入提款前需自行支付的醫療保費"
    summary_text = reader.pages[-1].extract_text() or ""
    assert "醫療融資從" not in summary_text, "五年供款第6年立即開始不應顯示紅字"
    assert "第 10 個保單年度" in summary_text and "第 20 個保單年度" in summary_text, "總結頁欠缺新現金價值里程碑"
    assert "至 30 歲" not in summary_text, "總結頁仍顯示舊有30歲現金價值里程碑"
    assert "每年 USD 4,714 儲蓄" in summary_text, "組合計劃仍然顯示每月儲蓄"
    assert "USD 11,400" in summary_text and "USD 34,970" in summary_text, "總結頁累計保費顯示錯誤"
    print("PASS 每年儲蓄及基本／組合計劃累計保費公式")

    delayed_reader = PdfReader(build_medical_financing_pdf(delayed_five_payload))
    delayed_summary = delayed_reader.pages[-1].extract_text() or ""
    assert "醫療融資從41歲開始" in delayed_summary, "延後開始的PDF欠缺細紅字年齡提示"
    print("PASS 延後開始的PDF顯示細紅字醫療融資年齡")

    rider_payload = sample_payload()
    rider_payload["result"]["input"]["annual"] = 4211
    rider_payload["result"]["input"]["total"] = 21055
    rider_payload["medicalContext"]["formName"] = "附加契約"
    rider_rates = {age: 1 for age in range(100)}
    rider_rates.update({29: 551, 30: 556, 31: 576, 32: 584, 33: 592, 34: 608})
    rider_payload["medicalContext"]["premiumByAge"] = [
        {"age": age, "premium": rider_rates[age]} for age in range(100)
    ]
    rider_data = validate_medical_financing_payload(rider_payload)
    assert _medical_paid_to_age(rider_data, 33) == 2859, "附加契約至 33 歲應累計 29 至 33 歲官方保費"
    assert _combo_paid_to_age(rider_data, 33) == 23914, "附加契約組合計劃至 33 歲累計投入錯誤"
    rider_reader = PdfReader(build_medical_financing_pdf(rider_payload))
    rider_summary = rider_reader.pages[-1].extract_text() or ""
    assert "每年 USD 4,211 儲蓄" in rider_summary, "附加契約沒有顯示每年儲蓄"
    assert "USD 2,859" in rider_summary and "USD 23,914" in rider_summary, "附加契約總結頁累計保費錯誤"
    print("PASS 圖中附加契約 18,000 自付費至 33 歲案例")

    single_payload = sample_payload()
    single_payload["result"]["input"].update({
        "annual": 35557.83,
        "total": 35557.83,
        "basic": 35917,
        "paymentMode": "single",
    })
    single_data = validate_medical_financing_payload(single_payload)
    assert single_data["input"]["paymentMode"] == "single", "一次性繳費模式沒有保留"
    assert _combo_paid_to_age(single_data, 33) == 46957.83, "一次性繳費累計投入錯誤"
    single_reader = PdfReader(build_medical_financing_pdf(single_payload))
    single_summary = single_reader.pages[-1].extract_text() or ""
    assert "一次性 USD 35,558 儲蓄" in single_summary, "一次性繳費總結標籤錯誤"
    assert "一次性" in single_summary and "USD 37,818" in single_summary, "一次性首年組合保費錯誤"
    print("PASS 一次性繳費PDF標籤及累計投入")

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
