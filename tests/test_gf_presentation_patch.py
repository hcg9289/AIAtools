#!/usr/bin/env python3
"""Run the tracked GF presentation and age-cap browser tests offline."""

from __future__ import annotations

import html
import os
import re
import signal
import subprocess
import tempfile
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parent
PAGE = ROOT / "gf_presentation_patch_harness.html"
CHROME = Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
EXPECTED_TESTS = 9


def main() -> None:
    if not CHROME.exists():
        raise SystemExit(f"Chrome not found: {CHROME}")
    with tempfile.TemporaryDirectory(prefix="gf-presentation-test-") as profile, tempfile.NamedTemporaryFile() as output:
        process = subprocess.Popen(
            [
                str(CHROME),
                "--headless=new",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-background-networking",
                "--window-size=390,1000",
                "--virtual-time-budget=1000",
                f"--user-data-dir={profile}",
                "--allow-file-access-from-files",
                "--dump-dom",
                PAGE.as_uri(),
            ],
            stdout=output,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
        deadline = time.monotonic() + 20
        report = ""
        while time.monotonic() < deadline:
            time.sleep(0.25)
            output.flush()
            source = Path(output.name).read_text(errors="replace")
            match = re.search(r'<pre id="report">(.*?)</pre>', source, re.S)
            if match:
                report = html.unescape(match.group(1))
                break
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
        process.wait(timeout=5)

    if not report:
        raise SystemExit("FAIL browser did not emit the GF presentation test report")
    print(report)
    lines = [line for line in report.splitlines() if line.strip()]
    if len(lines) != EXPECTED_TESTS or not all(line.startswith("PASS") for line in lines):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
