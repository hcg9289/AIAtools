#!/usr/bin/env python3
"""Compare GF numeric output before/after the year-range fix and build one PDF."""

from __future__ import annotations

import html
import json
import os
import re
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODEL_PAGE = ROOT / "research" / "gf_withdrawal_model.html"
CHROME = Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
OUTPUT_PDF = ROOT / "output" / "pdf" / "gf_29_to_99_numeric_verification.pdf"
SCRIPT_PATHS = (
    "assets/gf/gf_model_minimum_basic_patch.js",
    "assets/gf/gf_model_single_official_data.js",
    "assets/gf/gf_model_single_experiment.js",
    "assets/gf/gf_model_payment_plan.js",
    "assets/gf/gf_model_presentation.js",
)
NUMERIC_FIELDS = (
    "age", "policy_year", "requested_withdrawal", "withdrawal_total",
    "cumulative_withdrawal", "from_guaranteed_cash", "from_reversionary_bonus",
    "from_terminal_bonus", "post_basic_amount", "post_guaranteed_cash",
    "post_reversionary_bonus", "post_terminal_bonus", "post_surrender_total",
    "status",
)


def git_file(path: str) -> bytes:
    return subprocess.check_output(["git", "show", f"HEAD:{path}"], cwd=ROOT)


def build_page(destination: Path, scripts: list[Path]) -> None:
    source = MODEL_PAGE.read_text(encoding="utf-8")
    source = source.replace(
        "  updateMedicalPremiumSourceUi();\n  runFinance();\n",
        "  updateMedicalPremiumSourceUi();\n",
        1,
    )
    script_tags = "".join(f'<script src="{path.as_uri()}"></script>' for path in scripts)
    capture = """
<script>
(() => {
  const output = document.createElement('pre'); output.id = 'numericPayload'; document.body.append(output);
  try {
    document.getElementById('gfPaymentPlan').value = 'five_year';
    document.getElementById('issueAge').value = '29';
    document.getElementById('medicalStartAge').value = '35';
    document.getElementById('medicalEndAge').value = '99';
    document.getElementById('medicalPremiumSource').value = 'table';
    document.getElementById('medicalPlan').value = 'aia_avsw';
    document.getElementById('medicalPolicyForm').value = 'basic';
    document.getElementById('medicalDeductible').value = '30000';
    document.getElementById('financeMode').value = 'fixed';
    document.getElementById('fixedAnnualPremium').value = '20000';
    GFPaymentPlan.runCombinedFinance();
    if (!currentFinanceResult) throw new Error(document.getElementById('financeError').textContent || '沒有結果');
    output.textContent = JSON.stringify({
      result: GFPresentationPatch.visibleResult(currentFinanceResult),
      medicalContext: GFPresentationPatch.buildMedicalContextForPdf(currentMedicalContext),
    });
  } catch (error) {
    output.textContent = JSON.stringify({ error: error instanceof Error ? error.message : String(error) });
  }
})();
</script>
"""
    destination.write_text(source.replace("</body>", f"{script_tags}{capture}</body>", 1), encoding="utf-8")


def chrome_payload(page: Path) -> dict:
    with tempfile.TemporaryDirectory(prefix="gf-numeric-browser-") as profile, tempfile.NamedTemporaryFile() as output:
        process = subprocess.Popen(
            [
                str(CHROME), "--headless=new", "--disable-gpu", "--no-first-run",
                "--no-default-browser-check", "--disable-background-networking",
                "--virtual-time-budget=8000", f"--user-data-dir={profile}",
                "--allow-file-access-from-files", "--dump-dom", page.as_uri(),
            ],
            stdout=output, stderr=subprocess.DEVNULL, start_new_session=True,
        )
        deadline = time.monotonic() + 30
        encoded = ""
        while time.monotonic() < deadline:
            time.sleep(0.25)
            output.flush()
            match = re.search(
                r'<pre id="numericPayload">(.*?)</pre>',
                Path(output.name).read_text(errors="replace"),
                re.S,
            )
            if match:
                encoded = html.unescape(match.group(1))
                break
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
        process.wait(timeout=5)
    if not encoded:
        raise RuntimeError("Chrome沒有輸出GF數字結果。")
    payload = json.loads(encoded)
    if payload.get("error"):
        raise RuntimeError(payload["error"])
    return payload


def comparable(payload: dict) -> dict:
    result = payload["result"]
    return {
        "input": {key: result["input"].get(key) for key in ("annual", "total", "basic", "issueAge")},
        "totals": {key: result.get(key) for key in ("requestedTotal", "actualTotal", "surrendered", "surrenderYear")},
        "rows": [
            {key: row.get(key) for key in NUMERIC_FIELDS}
            for row in result["rows"]
        ],
    }


def main() -> None:
    if not MODEL_PAGE.exists():
        raise SystemExit(f"欠缺現行GF模型：{MODEL_PAGE}")
    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="gf-version-comparison-") as temp_name:
        temp = Path(temp_name)
        old_scripts = []
        for relative in SCRIPT_PATHS:
            old_path = temp / "old" / relative
            old_path.parent.mkdir(parents=True, exist_ok=True)
            old_path.write_bytes(git_file(relative))
            old_scripts.append(old_path)
        old_page = temp / "old.html"
        new_page = temp / "new.html"
        build_page(old_page, old_scripts)
        build_page(new_page, [ROOT / relative for relative in SCRIPT_PATHS])
        old_payload = chrome_payload(old_page)
        new_payload = chrome_payload(new_page)

    old_values = comparable(old_payload)
    new_values = comparable(new_payload)
    if old_values != new_values:
        for index, (old_row, new_row) in enumerate(zip(old_values["rows"], new_values["rows"]), 1):
            if old_row != new_row:
                raise AssertionError(f"第{index}行數字改變：舊={old_row} 新={new_row}")
        raise AssertionError(f"摘要數字改變：舊={old_values} 新={new_values}")

    rows = new_payload["result"]["rows"]
    if rows[0]["age"] != 30 or rows[-1]["age"] != 99 or len(rows) != 70:
        raise AssertionError(f"輸出年期錯誤：首={rows[0]['age']} 尾={rows[-1]['age']} 行數={len(rows)}")

    sys.path.insert(0, str(ROOT))
    from gf_medical_pdf import build_medical_financing_pdf  # noqa: E402

    document = build_medical_financing_pdf(new_payload)
    OUTPUT_PDF.write_bytes(document.getvalue())
    print(json.dumps({
        "case": {
            "paymentMode": "five_year",
            "calculationMode": "fixed",
            "issueAge": 29,
            "medicalStartAge": 35,
            "medicalEndAge": 99,
            "annualPremium": 20000,
            "form": "basic",
            "deductible": "HKD 30,000 / USD 3,750",
        },
        "rowsCompared": len(rows),
        "firstAge": rows[0]["age"],
        "lastAge": rows[-1]["age"],
        "requestedTotal": new_payload["result"]["requestedTotal"],
        "actualTotal": new_payload["result"]["actualTotal"],
        "surrendered": new_payload["result"]["surrendered"],
        "lastRow": {key: rows[-1].get(key) for key in NUMERIC_FIELDS},
        "pdf": str(OUTPUT_PDF),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
