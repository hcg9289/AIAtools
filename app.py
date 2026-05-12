import os
import io
import json
import base64
import uuid
import threading
import subprocess
import requests
import fitz  # PyMuPDF
from PIL import Image
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify, send_file, redirect
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
os.makedirs('uploads', exist_ok=True)
os.makedirs('outputs', exist_ok=True)

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

VAULT_AI_URL = os.environ.get('VAULT_AI_URL', 'http://wa-vault-1006:5001')
PPT_AI_ENDPOINT = f"{VAULT_AI_URL}/api/v1/ai/ppt/1008"
PPT_GENERATE_ENDPOINT = f"{VAULT_AI_URL}/api/v1/ai/ppt/generate"
VAULT_AUTH_URL = os.environ.get('VAULT_AUTH_URL', 'http://wa-vault-1006:5001/api/v1/token/validate')
SESSION_TTL_SECONDS = int(os.environ.get('SESSION_TTL_SECONDS', str(20 * 60)))
AUTH_SESSIONS = {}  # {sid: {"uid": str, "expiry": datetime, "ott": str}}
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf', 'ppt', 'pptx'}
PREVIEW_BOTS = ['WhatsApp', 'facebookexternalhit', 'Twitterbot', 'LinkedInBot', 'Slackbot']

# Async task store — {task_id: {"status": "processing"|"done"|"error", "result": dict, "created": datetime}}
TASKS = {}
TASKS_LOCK = threading.Lock()

TASK_TTL_SECONDS = 1800  # 30 minutes
PPT_FONT_FAMILY = os.environ.get('PPT_FONT_FAMILY', 'Noto Sans CJK TC')
PPT_FALLBACK_FONT_FAMILY = os.environ.get('PPT_FALLBACK_FONT_FAMILY', 'Microsoft JhengHei')
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

LAYOUT_BOXES = {
    'left': (0.72, 0.70, 5.45, 5.92),
    'right': (7.18, 0.70, 5.45, 5.92),
    'center': (1.38, 1.04, 10.56, 5.38),
    'top': (0.90, 0.60, 11.55, 3.08),
    'bottom': (0.90, 3.38, 11.55, 3.30),
    'full': (0.78, 0.70, 11.78, 5.92),
}


def _cleanup_old_tasks():
    now = datetime.now()
    with TASKS_LOCK:
        stale = [k for k, v in TASKS.items() if (now - v['created']).total_seconds() > TASK_TTL_SECONDS]
        for k in stale:
            del TASKS[k]


def get_client_ip():
    cf_ip = request.headers.get('cf-connecting-ip')
    if cf_ip:
        return cf_ip.strip()
    forwarded = request.headers.get('x-forwarded-for')
    if forwarded:
        return forwarded.split(',')[0].strip()
    real_ip = request.headers.get('x-real-ip')
    if real_ip:
        return real_ip.strip()
    return request.remote_addr or '127.0.0.1'


@app.before_request
def verify_ott_access():
    is_api_request = request.path.startswith('/api/')
    if request.path.startswith('/static') or request.path in ('/health',):
        return
    if request.path.endswith('.css') or request.path.endswith('.js') or request.path.endswith('.png') or request.path.endswith('.jpg') or request.path.endswith('.ico'):
        return

    ua = request.headers.get('User-Agent', '')
    if any(bot in ua for bot in PREVIEW_BOTS):
        if is_api_request:
            return jsonify({'success': False, 'error': 'Unauthorized'}), 403
        return (
            '<!DOCTYPE html><html><head><meta charset="utf-8">'
            '<title>PPT Generator</title>'
            '<meta property="og:title" content="PPT Generator"/>'
            '<meta property="og:description" content="請透過 WhatsApp 取得專屬連結後使用。"/>'
            '</head><body></body></html>',
            200,
            {'Content-Type': 'text/html; charset=utf-8'}
        )

    sid = request.cookies.get('auth_sid')
    if sid:
        session = AUTH_SESSIONS.get(sid)
        if session and session['expiry'] > datetime.now():
            return

    ott = request.args.get('ott')
    if ott:
        try:
            real_ip = get_client_ip()
            resp = requests.get(
                VAULT_AUTH_URL,
                params={'token': ott},
                headers={'CF-Connecting-IP': real_ip},
                timeout=5
            )
            data = resp.json() if resp.ok else {}
            if resp.status_code == 200 and data.get('valid'):
                new_sid = str(uuid.uuid4())
                AUTH_SESSIONS[new_sid] = {
                    'uid': data.get('uid'),
                    'expiry': datetime.now() + timedelta(seconds=SESSION_TTL_SECONDS),
                    'ott': ott
                }
                clean_url = request.path
                out = redirect(clean_url)
                out.set_cookie('auth_sid', new_sid, max_age=SESSION_TTL_SECONDS, httponly=True)
                return out
        except Exception:
            return jsonify({'success': False, 'error': '安全服務連線異常，請稍後再試'}), 503

    return jsonify({'success': False, 'error': 'Unauthorized'}), 403


# ── AI 背景圖模式映射（1006 新端點用）───────────────────────
ASPECT_RATIO_MAP = {
    '16:9': (16, 9),
    '4:3': (4, 3),
    '1:1': (1, 1),
    '3:4': (3, 4),
    '9:16': (9, 16),
}


def _filename_extension(filename):
    return os.path.splitext(filename or '')[1].lower().lstrip('.')


def allowed_file(filename):
    return _filename_extension(filename) in ALLOWED_EXTENSIONS


def _safe_upload_filename(original_filename):
    ext = _filename_extension(original_filename)
    if ext not in ALLOWED_EXTENSIONS:
        raise ValueError(f"不支援的檔案格式：{original_filename}")
    return f"upload_{uuid.uuid4()}.{ext}"


def pdf_to_base64_images(pdf_path):
    images = []
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        buffered = io.BytesIO()
        img.save(buffered, format="JPEG")
        img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
        images.append(f"data:image/jpeg;base64,{img_str}")
    return images


def convert_ppt_to_pdf(ppt_path, output_dir):
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', ppt_path, '--outdir', output_dir],
            check=True, capture_output=True
        )
        filename = os.path.basename(ppt_path)
        pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
        return os.path.join(output_dir, pdf_filename)
    except Exception as e:
        print(f"Error converting PPT to PDF: {e}")
        return None


def process_file_to_images(filepath):
    ext = _filename_extension(filepath)
    if ext not in ALLOWED_EXTENSIONS:
        return []
    if ext in ['png', 'jpg', 'jpeg']:
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode("utf-8")
            mime_ext = 'jpeg' if ext == 'jpg' else ext
            return [f"data:image/{mime_ext};base64,{encoded_string}"]
    elif ext == 'pdf':
        return pdf_to_base64_images(filepath)
    elif ext in ['ppt', 'pptx']:
        pdf_path = convert_ppt_to_pdf(filepath, app.config['UPLOAD_FOLDER'])
        if pdf_path:
            images = pdf_to_base64_images(pdf_path)
            os.remove(pdf_path)
            return images
        return []
    return []


def _safe_hex(hex_code, fallback='#111827'):
    value = str(hex_code or fallback).strip()
    if not value.startswith('#'):
        value = f'#{value}'
    if len(value) == 4:
        value = '#' + ''.join(ch * 2 for ch in value[1:])
    if len(value) != 7:
        return fallback
    try:
        int(value[1:], 16)
    except ValueError:
        return fallback
    return value.upper()


def hex_to_rgb(hex_code):
    hex_code = _safe_hex(hex_code).lstrip('#')
    return tuple(int(hex_code[i:i+2], 16) for i in (0, 2, 4))


def _rgb(hex_code):
    return RGBColor(*hex_to_rgb(hex_code))


def _set_run_font(run, size_pt, color_hex, bold=False, font_name=None):
    font = run.font
    font.name = font_name or PPT_FONT_FAMILY
    font.size = Pt(size_pt)
    font.bold = bold
    font.color.rgb = _rgb(color_hex)

    resolved_font = font_name or PPT_FONT_FAMILY
    r_pr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        font_elem = r_pr.find(qn(tag))
        if font_elem is None:
            font_elem = OxmlElement(tag)
            r_pr.append(font_elem)
        font_elem.set('typeface', resolved_font)


def _add_textbox(slide, box, vertical_anchor=MSO_ANCHOR.TOP):
    x, y, w, h = box
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.vertical_anchor = vertical_anchor
    tf.margin_left = Inches(0.06)
    tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    return shape


def _paragraph(tf, text, size, color, bold=False, align=PP_ALIGN.LEFT, space_after=8):
    p = tf.paragraphs[0] if len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = str(text or '').strip()
    _set_run_font(run, size, color, bold=bold)
    return p


def _set_slide_background(slide, color_hex):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(color_hex)


def _move_shape_to_back(slide, shape):
    sp_tree = slide.shapes._spTree
    elem = shape._element
    sp_tree.remove(elem)
    sp_tree.insert(2, elem)


def _add_background_image(prs, slide, base64_img_str):
    """Add a full-slide cover image without stretching the source aspect ratio."""
    try:
        raw = _strip_data_uri_prefix(base64_img_str)
        img_bytes = base64.b64decode(raw)
        pil_img = Image.open(io.BytesIO(img_bytes))
        src_w, src_h = pil_img.size
        if src_w <= 0 or src_h <= 0:
            return False

        slide_aspect = float(prs.slide_width) / float(prs.slide_height)
        img_aspect = src_w / src_h

        img_stream = io.BytesIO(img_bytes)
        pic = slide.shapes.add_picture(
            img_stream,
            left=0,
            top=0,
            width=prs.slide_width,
            height=prs.slide_height
        )

        if img_aspect > slide_aspect:
            crop = max(0, min(0.49, (1 - (slide_aspect / img_aspect)) / 2))
            pic.crop_left = crop
            pic.crop_right = crop
        elif img_aspect < slide_aspect:
            crop = max(0, min(0.49, (1 - (img_aspect / slide_aspect)) / 2))
            pic.crop_top = crop
            pic.crop_bottom = crop

        _move_shape_to_back(slide, pic)
        return True

    except Exception as e:
        print(f"Background image failed: {e}")
        return False


def _coerce_layout(layout_value, index):
    if isinstance(layout_value, dict):
        layout_value = layout_value.get('text_position') or layout_value.get('type') or layout_value.get('name')
    layout = str(layout_value or '').strip().lower().replace('-', '_')
    aliases = {
        'left_text': 'left',
        'right_text': 'right',
        'center_text': 'center',
        'title': 'center',
        'cover': 'center',
        'closing': 'center',
    }
    layout = aliases.get(layout, layout)
    if layout not in LAYOUT_BOXES:
        layout = 'center' if index == 0 else 'left'
    return layout


def _coerce_body_blocks(slide_data):
    blocks = slide_data.get('body_blocks')
    if isinstance(blocks, list) and blocks:
        return blocks

    content = slide_data.get('content', [])
    if isinstance(content, str):
        content = [content]
    if isinstance(content, list):
        return [{'type': 'bullet', 'text': item} for item in content if str(item).strip()]
    return []


def _normalize_slide(slide_data, index, deck_defaults):
    layout = _coerce_layout(slide_data.get('layout') or slide_data.get('layout_type'), index)
    title = str(slide_data.get('title') or f'Slide {index + 1}').strip()
    body_blocks = _coerce_body_blocks(slide_data)

    colors = deck_defaults.get('colors', {})
    bg_color = _safe_hex(slide_data.get('bg_color') or colors.get('background'), '#111827')
    title_color = _safe_hex(slide_data.get('title_color') or colors.get('title'), '#FFFFFF')
    content_color = _safe_hex(slide_data.get('content_color') or colors.get('content'), '#E5E7EB')
    accent_color = _safe_hex(slide_data.get('accent_color') or colors.get('accent'), '#60A5FA')

    overlay = slide_data.get('overlay') if isinstance(slide_data.get('overlay'), dict) else {}
    overlay_color = _safe_hex(overlay.get('color') or slide_data.get('overlay_color') or '#050816', '#050816')
    overlay_opacity = overlay.get('opacity', slide_data.get('overlay_opacity', 0.46))
    try:
        overlay_opacity = float(overlay_opacity)
    except (TypeError, ValueError):
        overlay_opacity = 0.46
    overlay_opacity = max(0.0, min(0.82, overlay_opacity))

    return {
        'title': title,
        'role': slide_data.get('role') or ('cover' if index == 0 else 'content'),
        'layout': layout,
        'body_blocks': body_blocks,
        'bg_color': bg_color,
        'title_color': title_color,
        'content_color': content_color,
        'accent_color': accent_color,
        'overlay_color': overlay_color,
        'overlay_opacity': overlay_opacity,
        'background_image': slide_data.get('background_image'),
        'footer': slide_data.get('footer') or deck_defaults.get('footer') or '',
    }


def _add_overlay(slide, box, color_hex, opacity):
    x, y, w, h = box
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(color_hex)
    try:
        shape.fill.transparency = opacity
    except Exception:
        shape.fill.transparency = int(opacity * 100)
    return shape


def _estimate_body_size(blocks):
    total = sum(len(str(block.get('text') or '')) + sum(len(str(i)) for i in block.get('items', []))
                for block in blocks if isinstance(block, dict))
    if total > 520:
        return 15
    if total > 340:
        return 17
    return 19


def _render_body_blocks(slide, box, slide_info, align):
    body_box = (box[0], box[1] + 1.28, box[2], max(0.7, box[3] - 1.42))
    body_shape = _add_textbox(slide, body_box)
    tf = body_shape.text_frame
    body_size = _estimate_body_size(slide_info['body_blocks'])

    for block in slide_info['body_blocks'][:7]:
        if not isinstance(block, dict):
            block = {'type': 'bullet', 'text': block}
        block_type = str(block.get('type') or 'bullet').lower()
        text = str(block.get('text') or '').strip()
        items = block.get('items') if isinstance(block.get('items'), list) else []

        if block_type in ('kicker', 'eyebrow'):
            _paragraph(tf, text, 13, slide_info['accent_color'], bold=True, align=align, space_after=6)
        elif block_type in ('stat', 'metric'):
            value = str(block.get('value') or text).strip()
            label = str(block.get('label') or '').strip()
            _paragraph(tf, value, min(30, body_size + 10), slide_info['accent_color'], bold=True, align=align, space_after=2)
            if label:
                _paragraph(tf, label, max(12, body_size - 2), slide_info['content_color'], align=align, space_after=10)
        elif items:
            if text:
                _paragraph(tf, text, body_size, slide_info['content_color'], bold=True, align=align, space_after=5)
            for item in items[:5]:
                _paragraph(tf, f'- {item}', body_size - 1, slide_info['content_color'], align=align, space_after=4)
        else:
            prefix = '- ' if block_type in ('bullet', 'point') else ''
            _paragraph(tf, f'{prefix}{text}', body_size, slide_info['content_color'], align=align, space_after=7)


def _render_slide(slide, prs, slide_info, index, total, hybrid_mode):
    _set_slide_background(slide, slide_info['bg_color'])
    has_bg_image = bool(hybrid_mode and slide_info.get('background_image') and _add_background_image(prs, slide, slide_info['background_image']))

    box = LAYOUT_BOXES[slide_info['layout']]
    align = PP_ALIGN.CENTER if slide_info['layout'] == 'center' else PP_ALIGN.LEFT
    if has_bg_image or slide_info['layout'] in ('center', 'full'):
        _add_overlay(slide, box, slide_info['overlay_color'], slide_info['overlay_opacity'])

    title_shape = _add_textbox(slide, (box[0], box[1], box[2], 1.10), vertical_anchor=MSO_ANCHOR.MIDDLE)
    title_len = len(slide_info['title'])
    title_size = 42 if index == 0 and title_len < 24 else 34
    if title_len > 42:
        title_size = 28
    _paragraph(title_shape.text_frame, slide_info['title'], title_size, slide_info['title_color'], bold=True, align=align, space_after=4)

    _render_body_blocks(slide, box, slide_info, align)

    marker = f'{index + 1:02d}/{total:02d}'
    footer_text = slide_info['footer']
    if footer_text:
        marker = f'{footer_text}  |  {marker}'
    footer_shape = _add_textbox(slide, (0.78, 6.92, 11.80, 0.28), vertical_anchor=MSO_ANCHOR.MIDDLE)
    _paragraph(footer_shape.text_frame, marker, 9, slide_info['content_color'], align=PP_ALIGN.RIGHT, space_after=0)


def create_pptx_from_json(json_data, output_path, hybrid_mode=False):
    """
    Render an editable PPTX using blank slides, cover backgrounds, and CJK-safe text boxes.
    Supports both the new design schema and the legacy title/content/colors schema.
    """
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    blank_layout = prs.slide_layouts[6]

    slides_data = json_data.get('slides') or []
    if not slides_data:
        slides_data = [{
            'title': '簡報生成結果',
            'content': ['目前沒有可用頁面資料，請重新提交提示詞或素材。'],
            'layout': 'center',
            'bg_color': '#111827',
            'title_color': '#FFFFFF',
            'content_color': '#E5E7EB',
        }]

    design_system = json_data.get('design_system') if isinstance(json_data.get('design_system'), dict) else {}
    deck_brief = json_data.get('deck_brief') if isinstance(json_data.get('deck_brief'), dict) else {}
    deck_defaults = {
        'colors': design_system.get('colors') if isinstance(design_system.get('colors'), dict) else {},
        'footer': deck_brief.get('audience') or deck_brief.get('purpose') or ''
    }

    total = len(slides_data)
    for i, slide_data in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)
        slide_info = _normalize_slide(slide_data, i, deck_defaults)
        _render_slide(slide, prs, slide_info, i, total, hybrid_mode)

    prs.save(output_path)


def _strip_data_uri_prefix(img_str):
    """剝離 data:image/xxx;base64, 前綴，只留 base64 本體"""
    if ',' in img_str:
        return img_str.split(',', 1)[1]
    return img_str


def _collect_uploaded_images(files):
    all_base64_images = []
    processed_any_file = False

    for file in files:
        if not file or not file.filename:
            continue
        if not allowed_file(file.filename):
            continue
        processed_any_file = True
        safe_filename = _safe_upload_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)

        try:
            file.save(filepath)
            images = process_file_to_images(filepath)
            for img in images:
                all_base64_images.append(_strip_data_uri_prefix(img))
        except Exception as e:
            original_name = secure_filename(file.filename) or file.filename
            raise ValueError(f"素材解析失敗：{original_name}") from e
        finally:
            try:
                os.remove(filepath)
            except Exception:
                pass

    return all_base64_images, processed_any_file


def _format_vault_error(resp, fallback):
    try:
        data = resp.json()
        detail = data.get('detail') or data.get('error') or data.get('message')
        if detail:
            return str(detail)
    except Exception:
        pass
    text = (resp.text or '').replace('\n', ' ').strip()
    return text[:240] if text else fallback


# ── Background task workers ──────────────────────────────────

def _run_standard_task(task_id, all_base64_images, prompt, language):
    """Standard (non-hybrid) PPT generation — runs in background thread."""
    try:
        vault_resp = requests.post(
            PPT_AI_ENDPOINT,
            json={
                "prompt": prompt,
                "language": language,
                "images": all_base64_images
            },
            timeout=180
        )

        if vault_resp.status_code == 503:
            raise ValueError('尚未設定 API Key，請聯絡管理員。')
        if vault_resp.status_code >= 400:
            raise ValueError(_format_vault_error(vault_resp, '1006 AI 生成失敗，請檢查 API Key、JSON schema 或素材格式。'))

        vault_data = vault_resp.json()
        ai_content = vault_data.get('content', '')
        try:
            json_data = json.loads(ai_content)
        except json.JSONDecodeError as e:
            raise ValueError(f'AI JSON 格式錯誤：{e}')

        output_filename = f"generated_presentation_{uuid.uuid4()}.pptx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        create_pptx_from_json(json_data, output_path, hybrid_mode=False)

        with TASKS_LOCK:
            TASKS[task_id]['status'] = 'done'
            TASKS[task_id]['result'] = {'download_url': f'/api/download/{output_filename}'}

    except Exception as e:
        print(f"[Task {task_id}] Standard task error: {e}")
        with TASKS_LOCK:
            TASKS[task_id]['status'] = 'error'
            TASKS[task_id]['result'] = {'error': str(e) or 'AI 生成失敗，請稍後重試。'}


def _run_hybrid_task(task_id, all_base64_images, prompt, language, aspect_ratio):
    """Hybrid (AI background) PPT generation — runs in background thread."""
    try:
        vault_resp = requests.post(
            PPT_GENERATE_ENDPOINT,
            json={
                "prompt": prompt,
                "language": language,
                "images": all_base64_images,
                "aspect_ratio": aspect_ratio
            },
            timeout=300
        )

        if vault_resp.status_code == 503:
            raise ValueError('尚未設定 API Key，請聯絡管理員。')
        if vault_resp.status_code >= 400:
            raise ValueError(_format_vault_error(vault_resp, '1006 AI 背景簡報生成失敗，請檢查 API Key、JSON schema 或素材格式。'))

        vault_data = vault_resp.json()
        ai_content = vault_data.get('content', '')
        try:
            json_data = json.loads(ai_content)
        except json.JSONDecodeError as e:
            raise ValueError(f'AI JSON 格式錯誤：{e}')

        output_filename = f"hybrid_presentation_{uuid.uuid4()}.pptx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        create_pptx_from_json(json_data, output_path, hybrid_mode=True)

        with TASKS_LOCK:
            TASKS[task_id]['status'] = 'done'
            TASKS[task_id]['result'] = {'download_url': f'/api/download/{output_filename}'}

    except Exception as e:
        print(f"[Task {task_id}] Hybrid task error: {e}")
        with TASKS_LOCK:
            TASKS[task_id]['status'] = 'error'
            TASKS[task_id]['result'] = {'error': str(e) or 'AI 簡報生成失敗，請稍後重試。'}


# ── Routes ────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/ping')
def ping():
    """Session keepalive — called by frontend every 60s to refresh session TTL."""
    sid = request.cookies.get('auth_sid')
    if sid and sid in AUTH_SESSIONS:
        AUTH_SESSIONS[sid]['expiry'] = datetime.now() + timedelta(seconds=SESSION_TTL_SECONDS)
    return jsonify({'ok': True})


@app.route('/api/task/<task_id>')
def get_task_status(task_id):
    """Poll for async task result."""
    _cleanup_old_tasks()
    with TASKS_LOCK:
        task = TASKS.get(task_id)
    if not task:
        return jsonify({'status': 'not_found'}), 404
    resp = {'status': task['status']}
    if task['result']:
        resp.update(task['result'])
    return jsonify(resp)


@app.route('/api/generate-ppt', methods=['POST'])
def generate_ppt():
    """
    Standard mode: prompt is required; files are optional source material.
    """
    files = request.files.getlist('files') if 'files' in request.files else []
    prompt = request.form.get('prompt', '').strip()
    language = request.form.get('language', 'Traditional Chinese')

    if not prompt:
        return jsonify({'error': '請輸入核心提示詞，說明你想做成什麼 PPT。'}), 400

    try:
        all_base64_images, processed_any_file = _collect_uploaded_images(files)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    if processed_any_file and not all_base64_images:
        return jsonify({'error': '素材解析失敗。請確認檔案為有效 PDF、PPT、PPTX、PNG 或 JPG。'}), 400

    task_id = str(uuid.uuid4())
    with TASKS_LOCK:
        TASKS[task_id] = {'status': 'processing', 'result': None, 'created': datetime.now()}

    t = threading.Thread(
        target=_run_standard_task,
        args=(task_id, all_base64_images, prompt, language),
        daemon=True
    )
    t.start()

    return jsonify({'task_id': task_id, 'status': 'processing'})


@app.route('/api/generate-ppt/hybrid', methods=['POST'])
def generate_ppt_hybrid():
    """
    Hybrid mode: prompt is required; files are optional source material.
    """
    files = request.files.getlist('files') if 'files' in request.files else []
    prompt = request.form.get('prompt', '').strip()
    language = request.form.get('language', 'Traditional Chinese')
    aspect_ratio = request.form.get('aspect_ratio', '16:9')

    if aspect_ratio not in ASPECT_RATIO_MAP:
        aspect_ratio = '16:9'

    if not prompt:
        return jsonify({'error': '請輸入核心提示詞，說明你想做成什麼 PPT。'}), 400

    try:
        all_base64_images, processed_any_file = _collect_uploaded_images(files)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    if processed_any_file and not all_base64_images:
        return jsonify({'error': '素材解析失敗。請確認檔案為有效 PDF、PPT、PPTX、PNG 或 JPG。'}), 400

    task_id = str(uuid.uuid4())
    with TASKS_LOCK:
        TASKS[task_id] = {'status': 'processing', 'result': None, 'created': datetime.now()}

    t = threading.Thread(
        target=_run_hybrid_task,
        args=(task_id, all_base64_images, prompt, language, aspect_ratio),
        daemon=True
    )
    t.start()

    return jsonify({'task_id': task_id, 'status': 'processing'})


@app.route('/api/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return jsonify({'error': 'File not found'}), 404


@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'service': '1008-ppt-generator'}), 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5008, debug=True)
