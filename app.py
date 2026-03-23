import os
import io
import json
import base64
import uuid
import subprocess
import requests
import fitz  # PyMuPDF
from PIL import Image
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
# 建立上傳與輸出的暫存資料夾
os.makedirs('uploads', exist_ok=True)
os.makedirs('outputs', exist_ok=True)

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 最大上傳限制 50MB

VAULT_AI_URL = os.environ.get('VAULT_AI_URL', 'http://wa-vault-1006:5001')
PPT_AI_ENDPOINT = f"{VAULT_AI_URL}/api/v1/ai/ppt/1008"
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf', 'ppt', 'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def pdf_to_base64_images(pdf_path):
    images = []
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # 提高解析度
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        buffered = io.BytesIO()
        img.save(buffered, format="JPEG")
        img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
        images.append(f"data:image/jpeg;base64,{img_str}")
    return images

def convert_ppt_to_pdf(ppt_path, output_dir):
    try:
        # 使用 LibreOffice headless 模式將 ppt/pptx 轉為 pdf
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', ppt_path, '--outdir', output_dir], check=True)
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
            os.remove(pdf_path) # 清理暫存的 PDF
            return images
        return []
    return []

def hex_to_rgb(hex_code):
    hex_code = hex_code.lstrip('#')
    return tuple(int(hex_code[i:i+2], 16) for i in (0, 2, 4))

def create_pptx_from_json(json_data, output_path):
    prs = Presentation()

    title_slide_layout = prs.slide_layouts[0]
    bullet_slide_layout = prs.slide_layouts[1]

    slides_data = json_data.get('slides', [])

    for i, slide_data in enumerate(slides_data):
        layout = title_slide_layout if i == 0 else bullet_slide_layout
        slide = prs.slides.add_slide(layout)

        # 設定背景顏色
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
            tf.text = "" # 清除預設文字

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


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/generate-ppt', methods=['POST'])
def generate_ppt():
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
            unique_id = str(uuid.uuid4())
            safe_filename = f"{unique_id}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
            file.save(filepath)

            images = process_file_to_images(filepath)
            # 剝離 data URI 前綴，只傳 base64 內容給 1006
            for img in images:
                if ',' in img:
                    all_base64_images.append(img.split(',', 1)[1])
                else:
                    all_base64_images.append(img)

            try:
                os.remove(filepath)
            except Exception:
                pass

    if not all_base64_images:
        return jsonify({'error': 'Failed to process files. Ensure they are valid PDF, PPT, or Image files.'}), 400

    try:
        # 把 AI 呼叫轉發給 1006（key 統一由 1006 中央管理）
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

        create_pptx_from_json(json_data, output_path)

        return jsonify({
            'success': True,
            'download_url': f"/api/download/{output_filename}"
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'error': 'Failed to communicate with AI service or generate PPT.'}), 500

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
