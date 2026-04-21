import os
import io
import json
import base64
import uuid
import subprocess
import requests
import fitz  # PyMuPDF
from PIL import Image
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify, send_file, redirect
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
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

# AI 背景圖模式映射（1006 新端點用）
ASPECT_RATIO_MAP = {
    '16:9': (16, 9),
    '4:3': (4, 3),
    '1:1': (1, 1),
    '3:4': (3, 4),
    '9:16': (9, 16),
}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


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
    ext = filepath.rsplit('.', 1)[1].lower()
    if ext in ['png', 'jpg', 'jpeg']:
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode("utf-8")
            return [f"data:image/{ext};base64,{encoded_string}"]
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


def hex_to_rgb(hex_code):
    hex_code = hex_code.lstrip('#')
    return tuple(int(hex_code[i:i+2], 16) for i in (0, 2, 4))


def _add_background_image(prs, slide, base64_img_str):
    """將 base64 PNG 圖片設為幻燈片背景（填充整頁）"""
    try:
        img_bytes = base64.b64decode(base64_img_str)
        img_stream = io.BytesIO(img_bytes)

        # 用 shape 新增圖片，撐滿整個 slide
        slide.shapes.add_picture(
            img_stream,
            left=0,
            top=0,
            width=prs.slide_width,
            height=prs.slide_height
        )

        # 將圖片 shape 移到最底層（z-order）
        # 找到最後加入的 shape（就是我們剛加的圖片），移到最前，再移到底
        shapes = slide.shapes
        if len(shapes) > 0:
            pic_shape = shapes[-1]
            sp_tree = slide.shapes._spTree
            # 取得該 shape 的 XML element
            pic_elem = pic_shape._element
            # 先移除
            sp_tree.remove(pic_elem)
            # 再插入到最前面（z-order 最低）
            sp_tree.insert(2, pic_elem)  # 1=nvGrpSpPr, 2= первый shape 通常在這之後

    except Exception as e:
        print(f"背景圖添加失敗: {e}")


def create_pptx_from_json(json_data, output_path, hybrid_mode=False):
    """
    根據 AI 回傳的 slides JSON 生成 PPTX。
    hybrid_mode=True 時：若 slide 有 background_image，先加圖片背景，再疊標題/內容文字。
    """
    prs = Presentation()

    title_slide_layout = prs.slide_layouts[0]
    bullet_slide_layout = prs.slide_layouts[1]

    slides_data = json_data.get('slides', [])

    for i, slide_data in enumerate(slides_data):
        layout = title_slide_layout if i == 0 else bullet_slide_layout
        slide = prs.slides.add_slide(layout)

        # 混合模式：優先用 AI 生成的背景圖
        bg_img_b64 = slide_data.get('background_image') if hybrid_mode else None

        if bg_img_b64:
            # AI 背景圖模式：加圖片背景 + 透明文字
            _add_background_image(prs, slide, bg_img_b64)
            # 覆蓋一層半透明黑色薄圖層（增加文字可讀性）
            # python-pptx 不支援直接設 fill 透明度，改用色塊 shape
            # 此處略過，保持圖片原樣

        else:
            # 純色背景模式
            bg_color_hex = slide_data.get('bg_color', '#FFFFFF')
            if bg_color_hex:
                bg = slide.background
                fill = bg.fill
                fill.solid()
                r, g, b = hex_to_rgb(bg_color_hex)
                fill.fore_color.rgb = RGBColor(r, g, b)

        # 設定標題
        title_shape = slide.shapes.title
        if title_shape and 'title' in slide_data:
            title_shape.text = slide_data['title']
            title_color_hex = slide_data.get('title_color', '#000000')
            for p in title_shape.text_frame.paragraphs:
                for run in p.runs:
                    r, g, b = hex_to_rgb(title_color_hex)
                    run.font.color.rgb = RGBColor(r, g, b)

        # 設定內容
        if i > 0 and 'content' in slide_data:
            body_shape = slide.shapes.placeholders[1]
            tf = body_shape.text_frame
            tf.text = ""

            content_color_hex = slide_data.get('content_color', '#000000')
            r, g, b = hex_to_rgb(content_color_hex)

            if isinstance(slide_data['content'], list):
                for point in slide_data['content']:
                    p = tf.add_paragraph()
                    p.text = point
                    p.level = 0
                    for run in p.runs:
                        run.font.color.rgb = RGBColor(r, g, b)
            else:
                p = tf.add_paragraph()
                p.text = slide_data['content']
                for run in p.runs:
                    run.font.color.rgb = RGBColor(r, g, b)

    prs.save(output_path)


def _strip_data_uri_prefix(img_str):
    """剝離 data:image/xxx;base64, 前綴，只留 base64 本體"""
    if ',' in img_str:
        return img_str.split(',', 1)[1]
    return img_str


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/generate-ppt', methods=['POST'])
def generate_ppt():
    """
    純文字版（向後相容）：AI 生成文字簡報內容，無背景圖。
    對應 1006 舊端點 POST /api/v1/ai/ppt/1008
    """
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400

    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')
    language = request.form.get('language', 'Traditional Chinese')

    if not files or files[0].filename == '':
        return jsonify({'error': 'No selected files'}), 400

    all_base64_images = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            safe_filename = f"{uuid.uuid4()}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
            file.save(filepath)

            images = process_file_to_images(filepath)
            for img in images:
                all_base64_images.append(_strip_data_uri_prefix(img))

            try:
                os.remove(filepath)
            except Exception:
                pass

    if not all_base64_images:
        return jsonify({'error': 'Failed to process files. Ensure they are valid PDF, PPT, or Image files.'}), 400

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
            return jsonify({'error': '尚未設定 API Key，請聯絡管理員。'}), 503
        vault_resp.raise_for_status()

        vault_data = vault_resp.json()
        ai_content = vault_data.get('content', '')

        json_data = json.loads(ai_content)

        output_filename = f"generated_presentation_{uuid.uuid4()}.pptx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        create_pptx_from_json(json_data, output_path, hybrid_mode=False)

        return jsonify({
            'success': True,
            'download_url': f"/api/download/{output_filename}"
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'error': 'Failed to communicate with AI service or generate PPT.'}), 500


@app.route('/api/generate-ppt/hybrid', methods=['POST'])
def generate_ppt_hybrid():
    """
    混合模式簡報生成（AI 背景圖 + 可編輯文字）。
    對應 1006 新端點 POST /api/v1/ai/ppt/generate
    - 支援 aspect_ratio 參數（預設 16:9）
    - 回傳的 slides 含 background_image 時，PPTX 該頁使用背景圖
    - 無 background_image 時降級為純色背景
    """
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400

    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')
    language = request.form.get('language', 'Traditional Chinese')
    aspect_ratio = request.form.get('aspect_ratio', '16:9')

    if aspect_ratio not in ASPECT_RATIO_MAP:
        aspect_ratio = '16:9'

    if not files or files[0].filename == '':
        return jsonify({'error': 'No selected files'}), 400

    all_base64_images = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            safe_filename = f"{uuid.uuid4()}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
            file.save(filepath)

            images = process_file_to_images(filepath)
            for img in images:
                all_base64_images.append(_strip_data_uri_prefix(img))

            try:
                os.remove(filepath)
            except Exception:
                pass

    if not all_base64_images:
        return jsonify({'error': 'Failed to process files. Ensure they are valid PDF, PPT, or Image files.'}), 400

    try:
        vault_resp = requests.post(
            PPT_GENERATE_ENDPOINT,
            json={
                "prompt": prompt,
                "language": language,
                "images": all_base64_images,
                "aspect_ratio": aspect_ratio
            },
            timeout=300  # Imagen 生成較慢，給更多時間
        )

        if vault_resp.status_code == 503:
            return jsonify({'error': '尚未設定 API Key，請聯絡管理員。'}), 503
        vault_resp.raise_for_status()

        vault_data = vault_resp.json()
        ai_content = vault_data.get('content', '')

        json_data = json.loads(ai_content)

        output_filename = f"hybrid_presentation_{uuid.uuid4()}.pptx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        create_pptx_from_json(json_data, output_path, hybrid_mode=True)

        return jsonify({
            'success': True,
            'download_url': f"/api/download/{output_filename}"
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'error': 'AI 簡報生成失敗。'}), 500


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
