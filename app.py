import html
import json
import multiprocessing as mp
import os
import re
import secrets
import base64
import time
from pathlib import Path
from typing import Any
from datetime import datetime, timedelta
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import fitz
import requests
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from PIL import Image, ImageChops, ImageDraw
from pydantic import BaseModel


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

OPENROUTER_ENV_CANDIDATES = [
    BASE_DIR / ".env",
    Path("/app/.env"),
    Path("/app/data/1008_ppt_generator/.env"),
    Path(r"C:\Users\user\Desktop\VScode\VPS\finance_vault\data\1008_ppt_generator\.env"),
    Path(r"C:\Users\user\Desktop\VScode\VPS\ppt_generator\.env"),
]

DEFAULT_MODEL = "google/gemini-3.1-flash-lite"
DEFAULT_VISION_CROP_MODEL = "google/gemini-3.1-flash-lite"
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
OPENROUTER_CONNECT_TIMEOUT = 15
OPENROUTER_READ_TIMEOUT = 95
OPENROUTER_OUTLINE_TIMEOUT = 120
OPENROUTER_REPAIR_TIMEOUT = 75
APP_PORT = int(os.getenv("PORT", "5008"))
APP_INTERNAL_BASE_URL = os.getenv("APP_INTERNAL_BASE_URL", f"http://127.0.0.1:{APP_PORT}").rstrip("/")

VAULT_AUTH_URL = os.environ.get("VAULT_AUTH_URL", "http://wa-vault-1006:5001/api/v1/token/validate")
SESSION_TTL_SECONDS = int(os.environ.get("SESSION_TTL_SECONDS", str(20 * 60)))
AUTH_SESSIONS: dict[str, dict[str, Any]] = {}
PREVIEW_BOTS = ["WhatsApp", "facebookexternalhit", "Twitterbot", "LinkedInBot", "Slackbot"]
VAULT_AUTH_TIMEOUT_SECONDS = int(os.environ.get("VAULT_AUTH_TIMEOUT_SECONDS", "15"))
VAULT_AUTH_RETRY_ATTEMPTS = int(os.environ.get("VAULT_AUTH_RETRY_ATTEMPTS", "2"))
VAULT_AUTH_RETRY_BACKOFF_SECONDS = float(os.environ.get("VAULT_AUTH_RETRY_BACKOFF_SECONDS", "0.4"))
AUTH_RETRY_PAGE_SECONDS = int(os.environ.get("AUTH_RETRY_PAGE_SECONDS", "2"))


def find_chrome_executable() -> Path:
    env_path = os.getenv("CHROME_EXECUTABLE_PATH")
    candidates = [
        Path(env_path) if env_path else None,
        Path("/usr/bin/chromium"),
        Path("/usr/bin/chromium-browser"),
        Path("/usr/bin/google-chrome"),
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    ]
    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate
    return Path(env_path or "/usr/bin/chromium")


LOCAL_CHROME = find_chrome_executable()


app = FastAPI(title="1008 中文簡報生成器")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/outputs", StaticFiles(directory=str(OUTPUT_DIR)), name="outputs")


def get_client_ip(request: Request) -> str:
    cf_ip = request.headers.get("cf-connecting-ip")
    if cf_ip:
        return cf_ip.strip()
    forwarded = request.headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[0].strip()
    real_ip = request.headers.get("x-real-ip")
    if real_ip:
        return real_ip.strip()
    return request.client.host if request.client else "127.0.0.1"


def expects_json_response(request: Request) -> bool:
    accept = request.headers.get("accept", "").lower()
    return request.url.path.startswith("/api/") or "application/json" in accept or request.method != "GET"


def auth_forbidden_response(request: Request, message: str = "Unauthorized") -> Response:
    if expects_json_response(request):
        return JSONResponse({"success": False, "error": message}, status_code=403)
    html_text = (
        '<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><title>授權失敗</title></head>'
        '<body style="font-family:sans-serif;text-align:center;padding-top:50px">'
        f"<h2>授權失敗 (1008)</h2><p>{html.escape(message)}</p><p>請回 WhatsApp 重新取得工具連結。</p>"
        "</body></html>"
    )
    return HTMLResponse(html_text, status_code=403)


def auth_retry_response(request: Request, message: str) -> Response:
    if expects_json_response(request):
        return JSONResponse({"success": False, "error": {"code": "auth_uncertain", "message": message}}, status_code=503)
    html_text = (
        '<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">'
        f'<meta http-equiv="refresh" content="{AUTH_RETRY_PAGE_SECONDS}">'
        "<title>OTT 驗證中</title></head>"
        '<body style="font-family:sans-serif;text-align:center;padding-top:50px">'
        "<h2>OTT 驗證暫時未確認 (1008)</h2>"
        f"<p>{html.escape(message)}</p><p>系統會自動重試，或你可以重新整理同一條連結。</p>"
        "</body></html>"
    )
    return HTMLResponse(
        html_text,
        status_code=503,
        headers={"Retry-After": str(AUTH_RETRY_PAGE_SECONDS), "Cache-Control": "no-store, max-age=0"},
    )


def is_vault_timeout_error(exc: Exception) -> bool:
    return isinstance(exc, requests.exceptions.Timeout) or "timeout" in str(exc).lower() or "timed out" in str(exc).lower()


def validate_vault_ott(ott: str, request: Request) -> tuple[str, dict[str, Any] | None, int]:
    token_prefix = (ott or "")[:8]
    real_ip = get_client_ip(request)
    request_id = secrets.token_hex(16)
    saw_timeout = False

    for attempt in range(1, VAULT_AUTH_RETRY_ATTEMPTS + 1):
        try:
            resp = requests.get(
                VAULT_AUTH_URL,
                params={"token": ott, "request_id": request_id},
                headers={"CF-Connecting-IP": real_ip},
                timeout=VAULT_AUTH_TIMEOUT_SECONDS,
            )
            if resp.status_code == 200:
                data = resp.json()
                if data.get("valid"):
                    print(f"OTT validation result=valid token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} vault_code=200", flush=True)
                    return "valid", data, attempt
                result = "auth_uncertain" if saw_timeout else "invalid"
                print(f"OTT validation result={result} token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} vault_code=200", flush=True)
                return result, data, attempt
            if resp.status_code in (401, 403):
                result = "auth_uncertain" if saw_timeout else "invalid"
                print(f"OTT validation result={result} token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} vault_code={resp.status_code}", flush=True)
                return result, None, attempt
            if resp.status_code >= 500 and attempt < VAULT_AUTH_RETRY_ATTEMPTS:
                print(f"OTT validation result=temporary_error_retry token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} vault_code={resp.status_code}", flush=True)
                time.sleep(VAULT_AUTH_RETRY_BACKOFF_SECONDS * attempt)
                continue
            print(f"OTT validation result=temporary_error token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} vault_code={resp.status_code}", flush=True)
            return "temporary_error", None, attempt
        except Exception as exc:
            if is_vault_timeout_error(exc):
                saw_timeout = True
                result = "timeout_retry" if attempt < VAULT_AUTH_RETRY_ATTEMPTS else "timeout"
                print(f"OTT validation result={result} token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} exception={type(exc).__name__}", flush=True)
                if attempt < VAULT_AUTH_RETRY_ATTEMPTS:
                    time.sleep(VAULT_AUTH_RETRY_BACKOFF_SECONDS * attempt)
                    continue
                return "timeout", None, attempt
            result = "auth_uncertain" if saw_timeout else "temporary_error"
            print(f"OTT validation result={result} token_prefix={token_prefix} attempt={attempt} real_ip={real_ip} exception={type(exc).__name__}", flush=True)
            return result, None, attempt
    return ("auth_uncertain" if saw_timeout else "temporary_error"), None, VAULT_AUTH_RETRY_ATTEMPTS


def clean_ott_url(request: Request) -> str:
    parts = urlsplit(str(request.url))
    query = [(key, value) for key, value in parse_qsl(parts.query, keep_blank_values=True) if key.lower() != "ott"]
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


@app.middleware("http")
async def verify_ott_access(request: Request, call_next):
    path = request.url.path
    if path in ("/health", "/favicon.ico") or path.startswith("/outputs/"):
        return await call_next(request)

    ua = request.headers.get("user-agent", "")
    if any(bot in ua for bot in PREVIEW_BOTS):
        if expects_json_response(request):
            return JSONResponse({"success": False, "error": "Unauthorized"}, status_code=403)
        return HTMLResponse(
            '<!DOCTYPE html><html><head><meta charset="utf-8">'
            "<title>1008 中文簡報生成器</title>"
            '<meta property="og:title" content="1008 中文簡報生成器"/>'
            '<meta property="og:description" content="請透過 WhatsApp 取得專屬連結後使用。"/>'
            "</head><body></body></html>",
            status_code=200,
        )

    now = datetime.now()
    sid = request.cookies.get("auth_sid")
    session = AUTH_SESSIONS.get(sid or "")
    if session and session["expiry"] > now:
        if path == "/api/ping":
            session["expiry"] = now + timedelta(seconds=SESSION_TTL_SECONDS)
        return await call_next(request)
    if sid:
        AUTH_SESSIONS.pop(sid, None)

    ott = request.query_params.get("ott", "")
    if not ott:
        return auth_forbidden_response(request, "缺少或已過期的 OTT。請回 WhatsApp 重新取得工具連結。")

    result, data, _attempts = validate_vault_ott(ott, request)
    if result == "valid":
        new_sid = secrets.token_urlsafe(32)
        AUTH_SESSIONS[new_sid] = {
            "uid": str((data or {}).get("user_id") or (data or {}).get("uid") or "user"),
            "expiry": now + timedelta(seconds=SESSION_TTL_SECONDS),
        }
        response = RedirectResponse(clean_ott_url(request), status_code=302)
        response.set_cookie(
            "auth_sid",
            new_sid,
            max_age=SESSION_TTL_SECONDS,
            httponly=True,
            samesite="lax",
            path="/",
        )
        return response
    if result == "invalid":
        return auth_forbidden_response(request, "OTT 已失效或不存在。請回 WhatsApp 重新取得工具連結。")
    return auth_retry_response(request, "驗證服務暫時未確認這條連結是否有效，請重試同一條連結。")


class RenderPayload(BaseModel):
    prompt: str
    slide_count: int = 7
    style: str = "finance_clean"
    outline: list[dict[str, Any]]
    source_id: str | None = None


class ExportPayload(BaseModel):
    deck_id: str


def load_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def get_openrouter_config() -> tuple[str | None, str, str]:
    key = os.getenv("OPENROUTER_API_KEY")
    model = os.getenv("FRONTEND_SLIDES_TEXT_MODEL") or os.getenv("OPENROUTER_LAB_TEXT_MODEL") or DEFAULT_MODEL
    base_url = os.getenv("OPENROUTER_BASE_URL") or OPENROUTER_BASE_URL
    for path in OPENROUTER_ENV_CANDIDATES:
        env = load_env_file(path)
        key = key or env.get("OPENROUTER_API_KEY")
        if path == BASE_DIR / ".env":
            model = env.get("FRONTEND_SLIDES_TEXT_MODEL") or env.get("OPENROUTER_LAB_TEXT_MODEL") or model
        base_url = env.get("OPENROUTER_BASE_URL") or base_url
        if key:
            break
    return key, model, base_url.rstrip("/")


def get_openrouter_vision_config() -> tuple[str | None, str, str]:
    key, _, base_url = get_openrouter_config()
    model = os.getenv("FRONTEND_SLIDES_VISION_MODEL") or os.getenv("OPENROUTER_LAB_VISION_MODEL") or DEFAULT_VISION_CROP_MODEL
    for path in OPENROUTER_ENV_CANDIDATES:
        env = load_env_file(path)
        if path == BASE_DIR / ".env":
            model = env.get("FRONTEND_SLIDES_VISION_MODEL") or env.get("OPENROUTER_LAB_VISION_MODEL") or model
        if key:
            break
    return key, model, base_url


def extract_pdf_text(data: bytes, name: str) -> str:
    if not name.lower().endswith(".pdf"):
        return ""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"PDF 解析失敗：{name}") from exc
    pages = []
    for index, page in enumerate(doc, start=1):
        text = page.get_text("text").strip()
        if text:
            pages.append(f"[第 {index} 頁]\n{text}")
    return "\n\n".join(pages)


def save_source_pdfs(uploaded: list[tuple[str, bytes]]) -> str | None:
    pdfs = [(name, data) for name, data in uploaded if name.lower().endswith(".pdf") and data]
    if not pdfs:
        return None
    source_id = secrets.token_hex(6)
    source_dir = OUTPUT_DIR / "sources" / source_id
    source_dir.mkdir(parents=True, exist_ok=True)
    manifest = []
    for idx, (name, data) in enumerate(pdfs, start=1):
        pdf_path = source_dir / f"source_{idx:02d}.pdf"
        pdf_path.write_bytes(data)
        manifest.append({"file": pdf_path.name, "original_name": name})
    (source_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return source_id


def text_fragments(value: str) -> list[str]:
    value = re.sub(r"[#*_`<>]+", " ", str(value or ""))
    chunks = re.split(r"[\s，。；：、,.!?()\[\]（）「」『』]+", value)
    terms: list[str] = []
    stop = {"PDF", "ppt", "html", "客戶", "產品", "方案", "說明", "重點", "文件", "證據"}
    for chunk in chunks:
        chunk = chunk.strip()
        if len(chunk) < 2 or chunk in stop:
            continue
        if len(chunk) > 18 and re.search(r"[\u4e00-\u9fff]", chunk):
            for start in range(0, min(len(chunk), 18), 5):
                part = chunk[start : start + 8]
                if len(part) >= 3:
                    terms.append(part)
        else:
            terms.append(chunk)
    numbers = re.findall(r"(?:[$]|USD|HKD|US\$|HK\$)?\s?\d[\d,]*(?:\.\d+)?%?", value)
    terms.extend([num.strip() for num in numbers if len(num.strip()) >= 2])
    seen: list[str] = []
    for term in terms:
        if term not in seen:
            seen.append(term)
    return seen[:35]


def slide_search_terms(slide: dict[str, Any]) -> list[str]:
    raw: list[str] = []
    for key in ["title", "takeaway"]:
        raw.append(str(slide.get(key) or ""))
    for key in ["points", "evidence"]:
        value = slide.get(key) or []
        if isinstance(value, list):
            raw.extend(str(item) for item in value)
        else:
            raw.append(str(value))
    table = slide.get("table") or []
    if isinstance(table, list):
        for row in table:
            if isinstance(row, list):
                raw.extend(str(cell) for cell in row)
            else:
                raw.append(str(row))
    terms: list[str] = []
    for item in raw:
        terms.extend(text_fragments(item))
    seen: list[str] = []
    for term in terms:
        if term and term not in seen:
            seen.append(term)
    return seen[:45]


def block_text_and_rects(page: fitz.Page) -> list[tuple[str, fitz.Rect]]:
    blocks: list[tuple[str, fitz.Rect]] = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        lines = []
        for line in block.get("lines", []):
            line_text = "".join(span.get("text", "") for span in line.get("spans", [])).strip()
            if line_text:
                lines.append(line_text)
        text = "\n".join(lines).strip()
        if text:
            blocks.append((text, fitz.Rect(block["bbox"])))
    return blocks


def score_text(text: str, terms: list[str]) -> int:
    lowered = text.lower()
    score = 0
    for term in terms:
        term_l = term.lower()
        if not term_l:
            continue
        if term_l in lowered:
            score += 5 if re.search(r"\d", term_l) else 3
    return score


def page_to_png_data_url(page: fitz.Page, scale: float = 1.35) -> str:
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    encoded = base64.b64encode(pix.tobytes("png")).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def slide_context_for_crop(slide: dict[str, Any], candidate_text: str) -> str:
    lines = [
        f"Slide title: {slide.get('title') or ''}",
        f"Slide takeaway: {slide.get('takeaway') or ''}",
        "Slide points:",
    ]
    for item in slide.get("points") or []:
        lines.append(f"- {item}")
    lines.append("Evidence/search hints:")
    for item in slide.get("evidence") or []:
        lines.append(f"- {item}")
    table = slide.get("table") or []
    if table:
        lines.append("Table data:")
        for row in table[:6]:
            if isinstance(row, list):
                lines.append(" | ".join(str(cell) for cell in row[:5]))
            else:
                lines.append(str(row))
    if candidate_text:
        lines.append("Text extracted near likely match:")
        lines.append(candidate_text[:1800])
    return "\n".join(lines)[:4200]


def parse_vision_crop(raw: str) -> tuple[dict[str, float] | None, float, str]:
    parsed = safe_json_from_text(raw)
    if isinstance(parsed, list) and parsed:
        parsed = parsed[0]
    if not isinstance(parsed, dict):
        return None, 0.0, "vision JSON is not an object"
    box = parsed.get("crop_box") or parsed.get("box") or parsed.get("bbox")
    if box is None:
        return None, float(parsed.get("confidence") or 0), str(parsed.get("reason") or "no crop_box")
    if isinstance(box, list) and len(box) >= 4:
        x1, y1, x2, y2 = box[:4]
        box = {"x1": x1, "y1": y1, "x2": x2, "y2": y2}
    if not isinstance(box, dict):
        return None, float(parsed.get("confidence") or 0), str(parsed.get("reason") or "crop_box is invalid")
    try:
        coords = {key: float(box[key]) for key in ["x1", "y1", "x2", "y2"]}
    except Exception:
        return None, float(parsed.get("confidence") or 0), str(parsed.get("reason") or "crop_box has non numeric coords")
    confidence = float(parsed.get("confidence") or 0)
    reason = str(parsed.get("reason") or "")
    return coords, confidence, reason


def normalized_box_to_page_rect(page: fitz.Page, box: dict[str, float]) -> fitz.Rect | None:
    values = dict(box)
    if any(value > 1.0 for value in values.values()):
        for key, value in list(values.items()):
            if value > 100:
                values[key] = value / 1000.0
            elif value > 1:
                values[key] = value / 100.0
    x1 = max(0.0, min(1.0, values["x1"]))
    y1 = max(0.0, min(1.0, values["y1"]))
    x2 = max(0.0, min(1.0, values["x2"]))
    y2 = max(0.0, min(1.0, values["y2"]))
    if x2 <= x1 or y2 <= y1:
        return None
    page_rect = page.rect
    clip = fitz.Rect(
        page_rect.x0 + page_rect.width * x1,
        page_rect.y0 + page_rect.height * y1,
        page_rect.x0 + page_rect.width * x2,
        page_rect.y0 + page_rect.height * y2,
    )
    if clip.width < page_rect.width * 0.20 or clip.height < page_rect.height * 0.26:
        center_x = (clip.x0 + clip.x1) / 2
        center_y = (clip.y0 + clip.y1) / 2
        half_w = max(clip.width / 2, page_rect.width * 0.18)
        half_h = max(clip.height / 2, page_rect.height * 0.13)
        clip.x0 = max(page_rect.x0, center_x - half_w)
        clip.x1 = min(page_rect.x1, center_x + half_w)
        clip.y0 = max(page_rect.y0, center_y - half_h)
        clip.y1 = min(page_rect.y1, center_y + half_h)
    if clip.width > page_rect.width * 0.98 and clip.height > page_rect.height * 0.92:
        return None
    return clip


def call_vision_crop(page: fitz.Page, slide: dict[str, Any], candidate_text: str) -> tuple[fitz.Rect | None, dict[str, Any]]:
    key, model, base_url = get_openrouter_vision_config()
    debug: dict[str, Any] = {"model": model, "used": False}
    if not key:
        debug["error"] = "missing OPENROUTER_API_KEY"
        return None, debug
    system = (
        "你是 PDF 原文截圖裁切助手。你的任務是看一頁 PDF 圖片，根據 slide 內容，"
        "選出最應該放進簡報的原文區域。只回 JSON，不要 Markdown。"
        "crop_box 必須是 normalized page coordinates，x1/y1/x2/y2 介乎 0 到 1。"
        "裁切要包含可讀標題或行列脈絡，不要只框一個數字；也不要截整頁。"
        "如果 slide 提到年份、年期、提款、保費、戶口價值或百分比，必須截到相關行與欄，"
        "例如第6年、第15年、第20年、第25年或第30年要包含該年份所在行及數字，"
        "不要只截表格最上方或說明文字。"
        "如果是表格證據，裁切高度要足夠顯示表頭、目標行與左右欄位脈絡。"
        "截圖會放在簡報右半頁，請優先選擇接近 4:3 或 3:2 的可讀區域；"
        "避免過闊的整行橫條，必要時只保留最相關的欄位與行。"
        "如果這頁和 slide 無關，crop_box 回 null，confidence 回 0。"
    )
    user_text = (
        slide_context_for_crop(slide, candidate_text)
        + "\n\n請只輸出："
        + '{"crop_box":{"x1":0.0,"y1":0.0,"x2":1.0,"y2":1.0},"confidence":0.0,"reason":"..."}'
    )
    try:
        response = requests.post(
            f"{base_url}/chat/completions",
            headers={
                "Authorization": f"Bearer {key}",
                "Content-Type": "application/json",
                "HTTP-Referer": APP_INTERNAL_BASE_URL,
                "X-Title": "1008 Chinese Presentation Generator Vision Crop",
            },
            json={
                "model": model,
                "temperature": 0.05,
                "max_tokens": 420,
                "messages": [
                    {"role": "system", "content": system},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": user_text},
                            {"type": "image_url", "image_url": {"url": page_to_png_data_url(page)}},
                        ],
                    },
                ],
            },
            timeout=(OPENROUTER_CONNECT_TIMEOUT, OPENROUTER_READ_TIMEOUT),
        )
        if response.status_code >= 400:
            debug["error"] = response.text[:500]
            return None, debug
        raw = response.json()["choices"][0]["message"]["content"]
        box, confidence, reason = parse_vision_crop(str(raw))
        debug.update({"raw": str(raw)[:700], "confidence": confidence, "reason": reason})
        if not box or confidence < 0.35:
            debug["error"] = "low confidence or empty crop"
            return None, debug
        clip = normalized_box_to_page_rect(page, box)
        if clip is None:
            debug["error"] = "invalid crop dimensions"
            return None, debug
        debug.update({"used": True, "crop_box": box})
        return clip, debug
    except Exception as exc:
        debug["error"] = f"{type(exc).__name__}: {str(exc)[:300]}"
        return None, debug


def text_layer_crop(page: fitz.Page, scored: list[tuple[int, str, fitz.Rect]]) -> fitz.Rect:
    rects = [item[2] for item in scored[:3]]
    clip = fitz.Rect(rects[0])
    for rect in rects[1:]:
        clip |= rect
    page_rect = page.rect
    pad_x = page_rect.width * 0.055
    pad_y = page_rect.height * 0.045
    clip = fitz.Rect(
        max(page_rect.x0, clip.x0 - pad_x),
        max(page_rect.y0, clip.y0 - pad_y),
        min(page_rect.x1, clip.x1 + pad_x),
        min(page_rect.y1, clip.y1 + pad_y),
    )
    if clip.width < page_rect.width * 0.42:
        center = (clip.x0 + clip.x1) / 2
        width = page_rect.width * 0.62
        clip.x0 = max(page_rect.x0, center - width / 2)
        clip.x1 = min(page_rect.x1, center + width / 2)
    if clip.height < page_rect.height * 0.18:
        center_y = (clip.y0 + clip.y1) / 2
        height = page_rect.height * 0.26
        clip.y0 = max(page_rect.y0, center_y - height / 2)
        clip.y1 = min(page_rect.y1, center_y + height / 2)
    return clip


def trim_image_whitespace(path: Path, padding: int = 22) -> None:
    try:
        image = Image.open(path).convert("RGB")
        white = Image.new("RGB", image.size, "white")
        diff = ImageChops.difference(image, white).convert("L")
        mask = diff.point(lambda value: 255 if value > 18 else 0)
        bbox = mask.getbbox()
        if not bbox:
            return
        x0, y0, x1, y1 = bbox
        x0 = max(0, x0 - padding)
        y0 = max(0, y0 - padding)
        x1 = min(image.width, x1 + padding)
        y1 = min(image.height, y1 + padding)
        crop_area = (x1 - x0) * (y1 - y0)
        full_area = image.width * image.height
        if crop_area >= full_area * 0.96:
            return
        image.crop((x0, y0, x1, y1)).save(path)
    except Exception:
        return


def crop_pdf_evidence(source_id: str | None, outline: list[dict[str, Any]], deck_dir: Path) -> dict[int, str]:
    if not source_id:
        return {}
    source_dir = OUTPUT_DIR / "sources" / source_id
    if not source_dir.exists():
        return {}
    pdf_paths = sorted(source_dir.glob("*.pdf"))
    if not pdf_paths:
        return {}
    evidence_dir = deck_dir / "evidence"
    evidence_dir.mkdir(exist_ok=True)
    results: dict[int, str] = {}
    debug_rows: list[dict[str, Any]] = []
    opened = []
    try:
        opened = [fitz.open(path) for path in pdf_paths]
        for slide_idx, slide in enumerate(outline, start=1):
            terms = slide_search_terms(slide)
            best: tuple[int, str, int, fitz.Page, list[tuple[int, str, fitz.Rect]]] | None = None
            for doc in opened:
                doc_name = Path(doc.name).name if doc.name else "source.pdf"
                for page_idx, page in enumerate(doc, start=1):
                    blocks = block_text_and_rects(page)
                    scored_blocks = []
                    page_score = 0
                    for text, rect in blocks:
                        score = score_text(text, terms)
                        if score:
                            scored_blocks.append((score, text, rect))
                            page_score += score
                    if scored_blocks and (best is None or page_score > best[0]):
                        best = (page_score, doc_name, page_idx, page, sorted(scored_blocks, key=lambda item: item[0], reverse=True))
            if best is None:
                debug_rows.append({"slide": slide_idx, "status": "no_text_candidate", "terms": terms[:12]})
                continue
            score, doc_name, page_idx, page, scored = best
            candidate_text = "\n".join(item[1] for item in scored[:6])
            fallback_clip = text_layer_crop(page, scored)
            vision_clip, vision_debug = call_vision_crop(page, slide, candidate_text)
            clip = vision_clip or fallback_clip
            pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), clip=clip, alpha=False)
            out_path = evidence_dir / f"slide_{slide_idx:02d}.png"
            pix.save(out_path)
            trim_image_whitespace(out_path)
            results[slide_idx] = f"/outputs/{deck_dir.name}/evidence/{out_path.name}"
            debug_rows.append(
                {
                    "slide": slide_idx,
                    "status": "vision_crop" if vision_clip else "text_fallback",
                    "doc": doc_name,
                    "page": page_idx,
                    "score": score,
                    "terms": terms[:12],
                    "vision": vision_debug,
                    "output": out_path.name,
                }
            )
    finally:
        for doc in opened:
            doc.close()
    if debug_rows:
        (evidence_dir / "crop_debug.json").write_text(json.dumps(debug_rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return results


def strip_markdown_noise(value: str) -> str:
    value = re.sub(r"<br\s*/?>", "\n", value, flags=re.I)
    value = re.sub(r"[*_`#>]+", "", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def compress_material_text(text: str, limit: int = 3200) -> str:
    keywords = [
        "提款",
        "提取",
        "戶口價值",
        "退保",
        "供款",
        "保費",
        "第6年",
        "第 6 年",
        "第15年",
        "第20年",
        "第25年",
        "第30年",
        "現金價值",
        "保單",
        "Guaranteed",
        "Withdrawal",
        "Account Value",
        "Surrender",
        "Premium",
    ]
    cleaned: list[str] = []
    current_page = ""
    page_blocks: list[tuple[str, list[str]]] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if re.match(r"^\[第\s+\d+\s+頁\]$", line):
            current_page = line
            page_blocks.append((current_page, []))
            continue
        if len(line) > 220 and re.fullmatch(r"[A-Za-z0-9+/=_\- .]+", line):
            continue
        if re.search(r"[A-Za-z0-9+/]{40,}=*", line) and not re.search(r"USD|HKD|MOP|RMB|\$|\d+\s?年|提款|供款|保單", line):
            continue
        cleaned.append(line)
        if page_blocks:
            page_blocks[-1][1].append(line)

    joined = "\n".join(cleaned)
    scored: list[tuple[int, int, str]] = []
    for idx, line in enumerate(cleaned):
        score = sum(1 for key in keywords if key.lower() in line.lower())
        if re.search(r"(USD|HKD|US\$|HK\$|\$)\s?\d|\d[\d,]+\.\d|\d+\s?年|第\s?\d+\s?年", line):
            score += 1
        if score:
            scored.append((-score, idx, line))
    important = [line for _, _, line in sorted(scored)[:55]]
    summaries: list[str] = []
    for page, lines in page_blocks:
        hits = []
        for line in lines:
            if any(key.lower() in line.lower() for key in keywords) or re.search(r"(USD|HKD|US\$|HK\$|\$)\s?\d|\d[\d,]+|\d+\s?年|第\s?\d+\s?年", line):
                hits.append(line)
        if hits:
            summaries.append(page + "\n" + "\n".join(hits[:18]))
    intro = "\n".join(cleaned[:35])
    compressed = intro + "\n\n[按頁關鍵資料]\n" + "\n\n".join(summaries[:10]) + "\n\n[重要候選資料]\n" + "\n".join(important)
    if len(compressed) < 1200:
        compressed = joined
    return compressed[:limit]


def safe_json_from_text(text: str) -> Any:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text, flags=re.I).strip()
        text = re.sub(r"```$", "", text).strip()
    match = re.search(r"(\{.*\}|\[.*\])", text, flags=re.S)
    if match:
        text = match.group(1)
    return json.loads(text)


def call_openrouter_http(messages: list[dict[str, str]]) -> str:
    key, model, base_url = get_openrouter_config()
    if not key:
        raise HTTPException(
            status_code=400,
            detail="OpenRouter API Key 未設定。請先在現有 PPT 工具環境或 1008 .env 設定 OPENROUTER_API_KEY。",
        )
    response = requests.post(
        f"{base_url}/chat/completions",
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "HTTP-Referer": APP_INTERNAL_BASE_URL,
            "X-Title": "1008 Chinese Presentation Generator",
        },
        json={
            "model": model,
            "temperature": 0.35,
            "max_tokens": 1800,
            "messages": messages,
        },
        timeout=(OPENROUTER_CONNECT_TIMEOUT, OPENROUTER_READ_TIMEOUT),
    )
    if response.status_code >= 400:
        detail = response.text[:700]
        raise HTTPException(status_code=response.status_code, detail=f"OpenRouter 回應錯誤：{detail}")
    payload = response.json()
    return payload["choices"][0]["message"]["content"]


def openrouter_worker(messages: list[dict[str, str]], queue: Any) -> None:
    try:
        queue.put({"ok": True, "content": call_openrouter_http(messages)})
    except HTTPException as exc:
        queue.put({"ok": False, "status": exc.status_code, "detail": exc.detail})
    except Exception as exc:
        queue.put({"ok": False, "status": 502, "detail": f"{type(exc).__name__}: {str(exc)[:500]}"})


def call_openrouter(messages: list[dict[str, str]], hard_timeout: int = OPENROUTER_OUTLINE_TIMEOUT) -> str:
    ctx = mp.get_context("spawn")
    queue = ctx.Queue()
    proc = ctx.Process(target=openrouter_worker, args=(messages, queue))
    proc.start()
    proc.join(hard_timeout)
    if proc.is_alive():
        proc.terminate()
        proc.join(5)
        raise TimeoutError(f"OpenRouter 超過 {hard_timeout} 秒未回應")
    if queue.empty():
        raise RuntimeError("OpenRouter worker 沒有返回結果")
    result = queue.get()
    if result.get("ok"):
        return str(result.get("content") or "")
    raise HTTPException(status_code=int(result.get("status") or 502), detail=str(result.get("detail") or "OpenRouter error"))


def normalize_outline(raw: Any, slide_count: int) -> list[dict[str, Any]]:
    if isinstance(raw, dict):
        slides = raw.get("slides") or raw.get("outline") or []
    else:
        slides = raw
    if not isinstance(slides, list):
        raise HTTPException(status_code=502, detail="AI 大綱格式錯誤：slides 不是列表。")
    normalized: list[dict[str, Any]] = []
    for idx, item in enumerate(slides[: max(slide_count, 1)], start=1):
        if not isinstance(item, dict):
            continue
        title = strip_markdown_noise(str(item.get("title") or f"第 {idx} 頁"))
        layout = str(item.get("layout") or "key_points")
        takeaway = strip_markdown_noise(str(item.get("takeaway") or ""))
        points = item.get("points") or item.get("bullets") or []
        if isinstance(points, str):
            points = [points]
        points = [strip_markdown_noise(str(point)) for point in points if strip_markdown_noise(str(point))]
        evidence = item.get("evidence") or []
        if isinstance(evidence, str):
            evidence = [evidence]
        evidence = [strip_markdown_noise(str(point)) for point in evidence if strip_markdown_noise(str(point))]
        table = item.get("table") if isinstance(item.get("table"), list) else []
        normalized.append(
            {
                "title": title[:42],
                "layout": layout,
                "takeaway": takeaway,
                "points": points[:6],
                "evidence": evidence[:5],
                "table": table[:8],
            }
        )
    return normalized


def expand_outline_to_count(prompt: str, slide_count: int, material_text: str, slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if len(slides) == slide_count:
        return slides
    system = (
        "你是中文銷售簡報大綱編輯。你要把現有大綱改成剛好指定頁數。"
        "只輸出 JSON，不要 Markdown。不要輸出 ##、<br>、星號或排版符號。"
        "每頁都要有 title、layout、takeaway、points、evidence、table。"
        "不得刪走重要數字；如需要新增頁，請按提示詞與素材補出有用的銷售說明頁。"
        "layout 只能是 cover, metrics, table, timeline, three_steps, compare, summary, key_points。"
    )
    user = (
        f"提示詞：{prompt}\n\n"
        f"目標頁數：{slide_count}\n\n"
        f"素材摘要：\n{material_text[:2600]}\n\n"
        f"現有大綱：\n{json.dumps({'slides': slides}, ensure_ascii=False)}\n\n"
        f"請輸出剛好 {slide_count} 頁的 JSON：{{\"slides\":[...]}}"
    )
    try:
        raw = call_openrouter([{"role": "system", "content": system}, {"role": "user", "content": user}], hard_timeout=OPENROUTER_REPAIR_TIMEOUT)
        return normalize_outline(safe_json_from_text(raw), slide_count)
    except Exception:
        return slides


def build_outline(prompt: str, slide_count: int, material_text: str) -> list[dict[str, Any]]:
    compact_material = compress_material_text(material_text)
    system = (
        "你是中文銷售簡報策劃師。你要先理解用戶提示詞，再根據 PDF 內容找證據，生成面向客戶的銷售說明大綱。"
        "這不是教學課件；每頁都要回答客戶得到什麼好處、產品特點如何配合客戶需要、為什麼現在值得了解。"
        "只輸出 JSON，不要 Markdown。不要輸出 ##、<br>、星號粗體或任何排版符號。"
        f"必須剛好輸出 {slide_count} 頁，不可多也不可少。"
        "每頁要短而清楚：title 不超過 18 字，points 最多 3 點，evidence 最多 3 點。"
        "每頁要有 title、layout、takeaway、points、evidence、table。"
        "layout 只能是 cover, metrics, table, timeline, three_steps, compare, summary, key_points。"
        "如果內容是金融/保險建議書，必須抽取年份、供款、提款、戶口價值等實際數字；不要空泛，也不要像教科書。"
        "不得加入 PDF 沒有提供的市場利率、銀行、債券或其他產品比較；所有數字必須可從素材或用戶提示詞找到。"
    )
    user = (
        f"提示詞：{prompt}\n\n"
        f"頁數：{slide_count}\n\n"
        f"PDF/素材文字：\n{compact_material}\n\n"
        "請輸出 JSON：{\"slides\":[...]}"
    )
    try:
        raw = call_openrouter([{"role": "system", "content": system}, {"role": "user", "content": user}])
        try:
            parsed = safe_json_from_text(raw)
        except Exception:
            repair_raw = call_openrouter(
                [
                    {
                        "role": "system",
                        "content": "你只負責把輸入內容轉成合法 JSON。不得新增內容。只輸出 JSON。",
                    },
                    {
                        "role": "user",
                        "content": "請把以下內容修成 {\"slides\":[...]} JSON。每頁保留 title/layout/takeaway/points/evidence/table：\n"
                        + raw[:6000],
                    },
                ],
                hard_timeout=OPENROUTER_REPAIR_TIMEOUT,
            )
            parsed = safe_json_from_text(repair_raw)
        slides = normalize_outline(parsed, slide_count)
        if len(slides) != slide_count:
            slides = expand_outline_to_count(prompt, slide_count, compact_material, slides)
        if len(slides) != slide_count:
            raise ValueError(f"AI returned {len(slides)} slides, expected {slide_count}")
        return slides
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=504,
            detail=f"AI 生成失敗，已停止，不會使用硬編碼假大綱。原因：{type(exc).__name__}: {str(exc)[:240]}",
        ) from exc


def esc(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def evidence_panel(slide: dict[str, Any], evidence: list[Any]) -> str:
    image_url = slide.get("_evidence_image_url")
    if image_url:
        return f"""
        <div class="source-shot">
          <img src="{esc(image_url)}" alt="PDF 原文截圖">
        </div>"""
    facts = "".join(f"<span>{esc(item)}</span>" for item in evidence[:6])
    return f"""<div class="fact-box"><strong>文件線索</strong>{facts}</div>"""


def normalize_table_rows(rows: Any, max_cols: int = 4, max_rows: int = 7) -> list[list[Any]]:
    if not isinstance(rows, list) or not rows:
        return []
    if all(isinstance(row, dict) for row in rows):
        keys: list[str] = []
        for row in rows:
            for key in row.keys():
                if key not in keys:
                    keys.append(str(key))
        keys = keys[:max_cols]
        normalized: list[list[Any]] = [keys]
        for row in rows[: max_rows - 1]:
            normalized.append([row.get(key, "") for key in keys])
        return normalized
    normalized = []
    for row in rows[:max_rows]:
        if isinstance(row, dict):
            normalized.append([f"{key}：{value}" for key, value in list(row.items())[:max_cols]])
        elif isinstance(row, list):
            normalized.append(row[:max_cols])
        else:
            normalized.append([row])
    return normalized


def table_html(rows: Any, compact: bool = False) -> str:
    normalized = normalize_table_rows(rows, max_cols=4, max_rows=7)
    body = ""
    for row_idx, row in enumerate(normalized):
        tag = "th" if row_idx == 0 and len(normalized) > 1 else "td"
        body += "<tr>" + "".join(f"<{tag}>{esc(cell)}</{tag}>" for cell in row) + "</tr>"
    cls = "data-table compact" if compact else "data-table"
    return f"<table class=\"{cls}\">{body}</table>"


def render_slide(slide: dict[str, Any], index: int, total: int) -> str:
    title = esc(slide.get("title"))
    takeaway = esc(slide.get("takeaway"))
    points = slide.get("points") or []
    evidence = slide.get("evidence") or []
    layout = slide.get("layout") or "key_points"
    page_no = f"{index:02d}/{total:02d}"

    if layout == "cover":
        chips = "".join(f"<span>{esc(item)}</span>" for item in evidence[:4])
        return f"""
        <section class="slide cover">
          <div class="slide-no">{page_no}</div>
          <div class="cover-grid">
            <div>
              <p class="eyebrow">客戶方案說明</p>
              <h1>{title}</h1>
              <p class="lead">{takeaway}</p>
              <div class="chips">{chips}</div>
            </div>
            <div class="signal-card">
              <strong>重點</strong>
              <span>{esc(points[0] if points else "根據文件與提示詞整理")}</span>
            </div>
          </div>
        </section>"""

    if layout == "metrics":
        table = slide.get("table") or []
        if table and len(table) > 1:
            bullets = "".join(f"<li>{esc(point)}</li>" for point in points[:3])
            return f"""
            <section class="slide">
              <div class="slide-no">{page_no}</div>
              <h2>{title}</h2>
              <p class="takeaway">{takeaway}</p>
              <div class="table-with-notes">
                {table_html(table, compact=True)}
                <ul>{bullets}</ul>
              </div>
            </section>"""
        cards = "".join(
            f"<div class=\"metric\"><span>證據 {i}</span><strong>{esc(item)}</strong></div>"
            for i, item in enumerate((evidence or points)[:6], start=1)
        )
        return f"""
        <section class="slide">
          <div class="slide-no">{page_no}</div>
          <h2>{title}</h2>
          <p class="takeaway">{takeaway}</p>
          <div class="metric-grid">{cards}</div>
        </section>"""

    if layout == "table":
        rows = slide.get("table") or [[p, e] for p, e in zip(points[:5], (evidence + [""] * 5)[:5])]
        return f"""
        <section class="slide">
          <div class="slide-no">{page_no}</div>
          <h2>{title}</h2>
          <p class="takeaway">{takeaway}</p>
          {table_html(rows)}
        </section>"""

    if layout == "timeline":
        items = evidence or points
        nodes = "".join(
            f"<div class=\"time-node\"><span>{esc(item)}</span><p>{esc(points[(i-1) % len(points)] if points else takeaway)}</p></div>"
            for i, item in enumerate(items[:5], start=1)
        )
        return f"""
        <section class="slide">
          <div class="slide-no">{page_no}</div>
          <h2>{title}</h2>
          <p class="takeaway">{takeaway}</p>
          <div class="timeline">{nodes}</div>
        </section>"""

    if layout == "three_steps":
        step_cards = []
        for i, point in enumerate(points[:3], start=1):
            if "：" in point:
                head, detail = point.split("：", 1)
            else:
                head, detail = point[:16], point
            step_cards.append(f"<div class=\"step\"><b>{i}</b><h3>{esc(head)}</h3><p>{esc(detail)}</p></div>")
        steps = "".join(step_cards)
        return f"""
        <section class="slide">
          <div class="slide-no">{page_no}</div>
          <h2>{title}</h2>
          <p class="takeaway">{takeaway}</p>
          <div class="steps">{steps}</div>
        </section>"""

    bullets = "".join(f"<li>{esc(point)}</li>" for point in points[:6])
    panel = evidence_panel(slide, evidence)
    section_class = "slide with-source-shot" if slide.get("_evidence_image_url") else "slide"
    return f"""
    <section class="{section_class}">
      <div class="slide-no">{page_no}</div>
      <h2>{title}</h2>
      <p class="takeaway">{takeaway}</p>
      <div class="content-grid">
        <ul>{bullets}</ul>
        {panel}
      </div>
    </section>"""


def render_html_deck(prompt: str, outline: list[dict[str, Any]], deck_id: str, evidence_assets: dict[int, str] | None = None) -> str:
    evidence_assets = evidence_assets or {}
    slides_html = []
    for i, slide in enumerate(outline, start=1):
        slide_copy = dict(slide)
        if evidence_assets.get(i):
            slide_copy["_evidence_image_url"] = evidence_assets[i]
        slides_html.append(render_slide(slide_copy, i, len(outline)))
    slides = "\n".join(slides_html)
    return f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>1008 中文簡報</title>
<style>
:root {{
  --ink:#17211c; --muted:#66736d; --paper:#f7f2e8; --panel:#fffaf0;
  --green:#116b55; --gold:#d8a640; --line:#d9ccb5; --black:#111;
}}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:#201f1a; color:var(--ink); font-family:"Microsoft JhengHei","Noto Sans TC",sans-serif; }}
.deck {{ width:100vw; min-height:100vh; display:grid; place-items:center; padding:28px; gap:28px; }}
.slide {{
  position:relative; width:min(96vw, 1280px); aspect-ratio:16/9; background:var(--paper);
  border-radius:18px; overflow:hidden; padding:72px 84px; box-shadow:0 28px 70px rgba(0,0,0,.32);
  page-break-after:always; break-after:page;
}}
.slide:before {{ content:""; position:absolute; inset:0; background:linear-gradient(135deg,rgba(17,107,85,.10),transparent 42%), radial-gradient(circle at 86% 14%,rgba(216,166,64,.22),transparent 22%); pointer-events:none; }}
.slide-no {{ position:absolute; right:42px; top:34px; font-weight:700; color:var(--green); letter-spacing:.06em; }}
.eyebrow {{ color:var(--green); font-weight:800; letter-spacing:.12em; font-size:18px; margin:0 0 20px; }}
h1,h2,h3,p,li,td,strong,span {{ word-break:keep-all; overflow-wrap:break-word; }}
h1 {{ font-size:64px; line-height:1.08; letter-spacing:0; margin:0 0 28px; max-width:820px; }}
h2 {{ font-size:54px; line-height:1.08; margin:0 0 18px; max-width:900px; }}
h3 {{ font-size:25px; line-height:1.2; margin:0 0 14px; }}
.lead,.takeaway {{ font-size:28px; line-height:1.45; color:#33433b; max-width:1050px; margin:0 0 36px; }}
.cover-grid,.content-grid {{ position:relative; display:grid; grid-template-columns:1.5fr .8fr; gap:54px; align-items:end; z-index:1; }}
.signal-card,.fact-box,.source-shot,.metric,.step {{ background:rgba(255,250,240,.86); border:2px solid var(--line); border-radius:14px; padding:28px; }}
.signal-card strong {{ display:block; color:var(--green); font-size:24px; margin-bottom:18px; }}
.signal-card span {{ display:block; font-size:27px; font-weight:800; line-height:1.3; }}
.chips {{ display:flex; flex-wrap:wrap; gap:14px; margin-top:34px; }}
.chips span,.fact-box span {{ display:inline-flex; border:1px solid var(--line); background:#fff; border-radius:999px; padding:10px 16px; font-size:20px; margin:6px; }}
ul {{ margin:0; padding-left:30px; font-size:28px; line-height:1.55; }}
li {{ margin:0 0 14px; }}
.fact-box strong {{ display:block; font-size:26px; color:var(--green); margin-bottom:16px; }}
.source-shot img {{ display:block; width:100%; max-height:330px; object-fit:contain; background:#fff; border:1px solid var(--line); border-radius:10px; }}
.slide.with-source-shot .content-grid {{ grid-template-columns:1fr 1fr; gap:42px; align-items:stretch; min-height:430px; }}
.slide.with-source-shot ul {{ font-size:27px; line-height:1.5; align-self:center; }}
.slide.with-source-shot .source-shot {{ background:#fff; border:2px solid var(--line); border-radius:12px; padding:0; min-height:430px; display:flex; align-items:stretch; justify-content:center; overflow:hidden; box-shadow:0 12px 26px rgba(31,24,12,.08); }}
.slide.with-source-shot .source-shot img {{ width:100%; height:100%; max-height:500px; object-fit:cover; object-position:left center; border:0; border-radius:0; box-shadow:none; }}
.metric-grid {{ position:relative; z-index:1; display:grid; grid-template-columns:repeat(auto-fit,minmax(270px,1fr)); gap:22px; margin-top:36px; }}
.metric span {{ display:block; color:var(--muted); font-size:18px; margin-bottom:14px; }}
.metric strong {{ display:block; font-size:28px; line-height:1.22; color:#10271f; }}
.data-table {{ position:relative; z-index:1; width:100%; border-collapse:collapse; background:#fffaf2; font-size:25px; margin-top:34px; }}
.data-table td,.data-table th {{ border:2px solid var(--line); padding:18px 22px; line-height:1.25; vertical-align:top; }}
.data-table th {{ background:rgba(17,107,85,.10); color:var(--green); font-weight:900; text-align:left; }}
.data-table.compact {{ margin-top:0; font-size:23px; }}
.table-with-notes {{ position:relative; z-index:1; display:grid; grid-template-columns:1.2fr .8fr; gap:34px; align-items:start; margin-top:30px; }}
.table-with-notes ul {{ font-size:24px; }}
.timeline {{ position:relative; z-index:1; display:grid; grid-template-columns:repeat(auto-fit,minmax(280px,1fr)); gap:18px; margin-top:46px; }}
.time-node {{ border-top:8px solid var(--green); background:#fffaf2; padding:22px 18px; min-height:220px; }}
.time-node span {{ color:var(--green); font-weight:900; font-size:30px; display:block; margin-bottom:18px; }}
.time-node p {{ font-size:22px; line-height:1.35; margin:0; }}
.steps {{ position:relative; z-index:1; display:grid; grid-template-columns:repeat(3,1fr); gap:26px; margin-top:48px; }}
.step b {{ display:grid; place-items:center; width:58px; height:58px; background:var(--green); color:white; border-radius:50%; font-size:26px; margin-bottom:22px; }}
.step p {{ font-size:23px; line-height:1.45; margin:0; }}
@media print {{
  body {{ background:white; }}
  .deck {{ display:block; padding:0; }}
  .slide {{ width:1920px; height:1080px; border-radius:0; box-shadow:none; margin:0; transform-origin:top left; }}
}}
</style>
</head>
<body>
<main class="deck" data-deck-id="{esc(deck_id)}" data-prompt="{esc(prompt)}">
{slides}
</main>
</body>
</html>"""


HOME_HTML = """<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>1008 中文簡報生成器</title>
<style>
:root{--bg:#f4efe5;--panel:#fffaf0;--ink:#17211c;--muted:#68736d;--green:#0f735d;--line:#d8cab4;--soft:#eee5d6}
*{box-sizing:border-box} body{margin:0;background:var(--bg);font-family:"Microsoft JhengHei","Noto Sans TC",sans-serif;color:var(--ink)}
.wrap{max-width:1180px;margin:0 auto;padding:32px 28px 60px}.top{display:flex;justify-content:space-between;gap:20px;align-items:flex-start;border-bottom:1px solid var(--line);padding-bottom:24px}
h1{font-size:44px;margin:0 0 10px;letter-spacing:0}p{line-height:1.6}.badge{border:1px solid var(--line);border-radius:999px;padding:10px 16px;background:#fffdf8;color:var(--green);font-weight:700}
.grid{display:grid;grid-template-columns:360px 1fr;gap:24px;margin-top:24px}.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:22px;box-shadow:0 12px 24px rgba(31,24,12,.05)}
label{display:block;font-weight:800;margin:16px 0 8px}textarea,input,select{width:100%;border:1px solid #cdbfa8;border-radius:8px;padding:12px 14px;background:#fffdf8;font:inherit}
textarea{min-height:180px;resize:vertical}.row{display:grid;grid-template-columns:1fr 1fr;gap:14px}.btns{display:flex;flex-wrap:wrap;gap:12px;margin-top:18px}
button,a.button{border:0;border-radius:9px;padding:13px 18px;background:var(--green);color:white;font-weight:900;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center}
button.secondary{background:var(--soft);color:var(--ink)}button.small{padding:8px 10px;font-size:13px}.danger{color:#8a2d1d!important}button:disabled{opacity:.5;cursor:not-allowed}.hint{color:var(--muted);font-size:14px}
.status{white-space:pre-wrap;background:#f1eadc;border-radius:9px;padding:14px;margin-top:14px;color:#34443c}.outline{display:grid;gap:14px}.slide-card{border:1px solid var(--line);border-radius:10px;background:#fffdf8;padding:16px}
.slide-card-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:8px}.slide-actions{display:flex;gap:8px;flex-wrap:wrap}.slide-card h3{margin:0;color:var(--green)}.slide-card input{font-weight:800}.points{min-height:92px}.links{display:flex;gap:12px;flex-wrap:wrap;margin-top:14px}
.result-box{display:none;margin-top:14px;padding:14px;border:1px solid var(--line);border-radius:10px;background:#fffdf8}.result-box strong{display:block;color:var(--green);margin-bottom:10px}.result-box .links{margin-top:0}
iframe{width:100%;height:520px;border:1px solid var(--line);border-radius:10px;background:white;margin-top:16px}
</style>
</head>
<body>
<div class="wrap">
  <div class="top">
    <div>
      <h1>1008 中文簡報生成器</h1>
      <p>把 PDF + 提示詞轉成中文銷售簡報，先講客戶好處，再用文件數字支持產品如何配合需要。</p>
    </div>
    <div class="badge" id="keyStatus">檢查中...</div>
  </div>
  <div class="grid">
    <section class="card">
      <h2>1. 內容</h2>
      <label>你想做成什麼簡報？</label>
      <textarea id="prompt">說明回報好，善用提款例子，第6年起每年提款6%本金。</textarea>
      <div class="row">
        <div><label>頁數</label><input id="slideCount" type="number" min="1" max="30" value="7"></div>
        <div><label>取向</label><select id="density"><option selected>銷售說明型</option><option>閱讀報告型</option></select></div>
      </div>
      <label>上傳 PDF</label>
      <input id="files" type="file" accept=".pdf" multiple>
      <p class="hint">會先抽取 PDF 文字層；第一版不做 OCR。</p>
      <div class="btns">
        <button id="outlineBtn">生成大綱</button>
        <button id="renderBtn" class="secondary" disabled>生成 HTML 簡報</button>
        <button id="pdfBtn" class="secondary" disabled>匯出 PDF</button>
      </div>
      <div class="status" id="status">等待開始。</div>
      <div class="result-box" id="resultBox">
        <strong>生成結果</strong>
        <div class="links" id="resultLinks"></div>
      </div>
    </section>
    <main class="card">
      <h2>2. 大綱校對</h2>
      <p class="hint">大綱是給人看的，可直接修改。生成時會使用你畫面上的版本。</p>
      <div class="btns">
        <button id="addSlideBtn" class="secondary" disabled>新增頁</button>
        <span class="hint" id="outlineCount">目前 0 頁</span>
      </div>
      <div id="outline" class="outline"></div>
      <div class="links" id="links"></div>
      <iframe id="preview" title="HTML 簡報預覽"></iframe>
    </main>
  </div>
</div>
<script>
let outline = [];
let deckId = "";
let sourceId = "";
const $ = (id) => document.getElementById(id);
function setStatus(text){ $("status").textContent = text; }
function setResultLinks(html){
  $("resultBox").style.display = html ? "block" : "none";
  $("resultLinks").innerHTML = html || "";
  $("links").innerHTML = html || "";
}
async function api(path, options={}){
  const res = await fetch(path, options);
  const data = await res.json().catch(() => ({}));
  if(!res.ok) throw new Error(data.detail || `HTTP ${res.status}`);
  return data;
}
function escapeHtml(value){
  return String(value || "").replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}
function clampSlideCount(value){
  const parsed = Number.parseInt(value, 10);
  if(!Number.isFinite(parsed)) return 7;
  return Math.max(1, Math.min(30, parsed));
}
function blankSlide(index){
  return {
    title: `第 ${index + 1} 頁`,
    layout: "key_points",
    takeaway: "",
    points: ["請輸入這頁要說明的重點"],
    evidence: [],
    table: []
  };
}
function syncOutlineCount(){
  $("outlineCount").textContent = `目前 ${outline.length} 頁`;
  if(outline.length > 0) $("slideCount").value = outline.length;
  $("addSlideBtn").disabled = false;
}
function addSlide(afterIndex = outline.length - 1){
  outline.splice(afterIndex + 1, 0, blankSlide(afterIndex + 1));
  renderOutline();
  $("renderBtn").disabled = outline.length === 0;
  setStatus(`已新增 1 頁，目前共 ${outline.length} 頁。`);
}
function deleteSlide(index){
  if(outline.length <= 1){
    alert("至少要保留 1 頁。");
    return;
  }
  outline.splice(index, 1);
  renderOutline();
  setStatus(`已刪除 1 頁，目前共 ${outline.length} 頁。`);
}
function renderOutline(){
  const root = $("outline");
  root.innerHTML = "";
  outline.forEach((s, i) => {
    const card = document.createElement("div");
    card.className = "slide-card";
    card.innerHTML = `
      <div class="slide-card-head">
        <h3>第 ${i+1} 頁</h3>
        <div class="slide-actions">
          <button class="secondary small" data-action="add" data-i="${i}">在後面新增</button>
          <button class="secondary small danger" data-action="delete" data-i="${i}">刪除</button>
        </div>
      </div>
      <label>標題</label><input data-i="${i}" data-k="title" value="${escapeHtml(s.title||"")}">
      <label>一句重點</label><textarea data-i="${i}" data-k="takeaway">${escapeHtml(s.takeaway||"")}</textarea>
      <label>內容要點（每行一點）</label><textarea class="points" data-i="${i}" data-k="points">${escapeHtml((s.points||[]).join("\\n"))}</textarea>
      <label>截圖定位線索（每行一點）</label><textarea class="points" data-i="${i}" data-k="evidence">${escapeHtml((s.evidence||[]).join("\\n"))}</textarea>`;
    root.appendChild(card);
  });
  root.querySelectorAll("input,textarea").forEach(el => el.addEventListener("input", () => {
    const i = Number(el.dataset.i), k = el.dataset.k;
    outline[i][k] = (k === "points" || k === "evidence") ? el.value.split("\\n").map(x=>x.trim()).filter(Boolean) : el.value;
  }));
  root.querySelectorAll("[data-action]").forEach(btn => btn.addEventListener("click", () => {
    const index = Number(btn.dataset.i);
    if(btn.dataset.action === "add") addSlide(index);
    if(btn.dataset.action === "delete") deleteSlide(index);
  }));
  syncOutlineCount();
}
async function refreshStatus(){
  const data = await api("/api/status");
  $("keyStatus").textContent = data.key_available ? `OpenRouter 已連接 · 文字 ${data.model} · 截圖 ${data.vision_model}` : "未找到 OpenRouter Key";
}
$("outlineBtn").onclick = async () => {
  try{
    setStatus("正在抽取 PDF 並生成中文大綱...");
    const fd = new FormData();
    fd.append("prompt", $("prompt").value);
    fd.append("slide_count", clampSlideCount($("slideCount").value));
    fd.append("density", $("density").value);
    [...$("files").files].forEach(f => fd.append("files", f));
    const data = await api("/api/outline", {method:"POST", body:fd});
    outline = data.slides;
    sourceId = data.source_id || "";
    renderOutline();
    setResultLinks("");
    $("renderBtn").disabled = false;
    setStatus(`已生成 ${outline.length} 頁大綱。請先檢查，再生成 HTML。`);
  }catch(e){ setStatus("失敗：" + e.message); }
};
$("renderBtn").onclick = async () => {
  try{
    setStatus("正在生成 HTML 簡報...");
    const data = await api("/api/render", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({prompt:$("prompt").value, slide_count:outline.length || clampSlideCount($("slideCount").value), style:"finance_clean", outline, source_id: sourceId})});
    deckId = data.deck_id;
    $("preview").src = data.html_url + "?t=" + Date.now();
    setResultLinks(`<a class="button" href="${data.html_url}" target="_blank">打開 HTML</a>`);
    $("pdfBtn").disabled = false;
    setStatus("HTML 已生成。你可以預覽或匯出 PDF。");
  }catch(e){ setStatus("失敗：" + e.message); }
};
$("pdfBtn").onclick = async () => {
  try{
    setStatus("正在用伺服器 Chromium 匯出 PDF...");
    const data = await api("/api/export_pdf", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({deck_id:deckId})});
    setResultLinks(`${$("resultLinks").innerHTML}<a class="button" href="${data.pdf_url}" target="_blank">下載 PDF</a><a class="button" href="${data.contact_sheet_url}" target="_blank">查看縮圖檢查</a>`);
    setStatus("PDF 已匯出，並已產生縮圖檢查。");
  }catch(e){ setStatus("失敗：" + e.message); }
};
$("addSlideBtn").onclick = () => addSlide();
refreshStatus().catch(e => $("keyStatus").textContent = "狀態檢查失敗");
</script>
</body>
</html>"""


@app.get("/", response_class=HTMLResponse)
def home() -> str:
    return HOME_HTML


@app.get("/health")
def health() -> dict[str, Any]:
    return {"status": "healthy", "service": "1008-presentation-generator"}


@app.get("/api/ping")
def ping() -> dict[str, Any]:
    return {"ok": True}


@app.get("/api/status")
def status() -> dict[str, Any]:
    key, model, _ = get_openrouter_config()
    _, vision_model, _ = get_openrouter_vision_config()
    return {"ok": True, "key_available": bool(key), "model": model, "vision_model": vision_model}


@app.post("/api/outline")
def outline(
    prompt: str = Form(...),
    slide_count: int = Form(7),
    density: str = Form("閱讀型"),
    files: list[UploadFile] = File(default=[]),
) -> dict[str, Any]:
    texts = []
    uploaded: list[tuple[str, bytes]] = []
    for file in files:
        data = file.file.read()
        uploaded.append((file.filename or "upload.pdf", data))
        texts.append(extract_pdf_text(data, file.filename or "upload.pdf"))
    material = "\n\n".join(texts)
    full_prompt = f"{prompt}\n密度：{density}"
    requested_count = max(1, min(slide_count, 30))
    slides = build_outline(full_prompt, requested_count, material)
    source_id = save_source_pdfs(uploaded)
    return {"ok": True, "slides": slides, "source_id": source_id}


@app.post("/api/render")
def render(payload: RenderPayload) -> dict[str, Any]:
    deck_id = secrets.token_hex(6)
    deck_dir = OUTPUT_DIR / deck_id
    deck_dir.mkdir(parents=True, exist_ok=True)
    target_count = max(1, min(len(payload.outline) or payload.slide_count, 30))
    outline = normalize_outline(payload.outline, target_count)
    evidence_assets = crop_pdf_evidence(payload.source_id, outline, deck_dir)
    html_text = render_html_deck(payload.prompt, outline, deck_id, evidence_assets)
    html_path = deck_dir / "deck.html"
    html_path.write_text(html_text, encoding="utf-8")
    (deck_dir / "outline.json").write_text(json.dumps(outline, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"ok": True, "deck_id": deck_id, "html_url": f"/outputs/{deck_id}/deck.html"}


@app.post("/api/export_pdf")
def export_pdf(payload: ExportPayload) -> dict[str, Any]:
    deck_dir = OUTPUT_DIR / payload.deck_id
    html_path = deck_dir / "deck.html"
    if not html_path.exists():
        raise HTTPException(status_code=404, detail="找不到 HTML 簡報。")
    if not LOCAL_CHROME.exists():
        raise HTTPException(status_code=500, detail="找不到伺服器 Chrome/Chromium，無法匯出 PDF。")
    pdf_path = deck_dir / "deck.pdf"
    preview_dir = deck_dir / "preview"
    preview_dir.mkdir(exist_ok=True)
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, executable_path=str(LOCAL_CHROME))
        page = browser.new_page(viewport={"width": 1920, "height": 1080}, device_scale_factor=1)
        page.goto(f"{APP_INTERNAL_BASE_URL}/outputs/{payload.deck_id}/deck.html", wait_until="networkidle")
        slide_count = page.locator(".slide").count()
        preview_paths = []
        for idx in range(slide_count):
            preview_path = preview_dir / f"slide_{idx + 1:02d}.png"
            page.locator(".slide").nth(idx).screenshot(path=str(preview_path))
            preview_paths.append(preview_path)
        contact_sheet = preview_dir / "contact_sheet.png"
        make_contact_sheet(preview_paths, contact_sheet)
        page.pdf(path=str(pdf_path), width="1920px", height="1080px", print_background=True, margin={"top": "0", "right": "0", "bottom": "0", "left": "0"})
        browser.close()
    return {"ok": True, "pdf_url": f"/outputs/{payload.deck_id}/deck.pdf", "contact_sheet_url": f"/outputs/{payload.deck_id}/preview/contact_sheet.png"}


def make_contact_sheet(image_paths: list[Path], output_path: Path) -> None:
    if not image_paths:
        return
    thumbs = []
    target_w = 560
    for path in image_paths:
        image = Image.open(path).convert("RGB")
        ratio = target_w / image.width
        thumbs.append(image.resize((target_w, int(image.height * ratio))))
    gap = 26
    cols = 2
    rows = (len(thumbs) + cols - 1) // cols
    cell_h = max(img.height for img in thumbs) + 42
    sheet = Image.new("RGB", (cols * target_w + (cols + 1) * gap, rows * cell_h + (rows + 1) * gap), "#f4efe5")
    draw = ImageDraw.Draw(sheet)
    for idx, img in enumerate(thumbs):
        row, col = divmod(idx, cols)
        x = gap + col * (target_w + gap)
        y = gap + row * (cell_h + gap)
        sheet.paste(img, (x, y + 34))
        draw.text((x, y), f"Slide {idx + 1}", fill="#0f735d")
    sheet.save(output_path)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=APP_PORT)
