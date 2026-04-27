"""
Flask web app for QCP to CNPE conversion.
Vercel serverless - api/index.py IS the entry point.
"""

import os
import uuid
import sys
from flask import Flask, request, render_template, send_file, jsonify, session
from functools import wraps
from werkzeug.utils import secure_filename

# Vercel serverless path setup
VERCEL_APP_DIR = os.environ.get('VERCEL_APP_DIR', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, VERCEL_APP_DIR)

app = Flask(__name__,
            template_folder=os.path.join(VERCEL_APP_DIR, 'templates'),
            static_folder=os.path.join(VERCEL_APP_DIR, 'static'))
app.secret_key = os.environ.get('SECRET_KEY', 'qcp-cnpe-secret-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = '/tmp/qcp-uploads'

APP_PASSWORD = os.environ.get('QCP_PASSWORD', 'KSB2026')
ALLOWED_EXTENSIONS = {'pdf'}


def password_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get('authenticated') != True:
            return jsonify({'error': '请先通过密码验证'}), 401
        return f(*args, **kwargs)
    return decorated_function


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    try:
        if session.get('authenticated'):
            return render_template('index.html', authenticated=True)
        return render_template('index.html', authenticated=False)
    except Exception as e:
        return f"Template error: {str(e)}", 500


@app.route('/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        if data.get('password') == APP_PASSWORD:
            session['authenticated'] = True
            return jsonify({'status': 'ok'})
        return jsonify({'error': '密码错误'}), 401
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/upload', methods=['POST'])
@password_required
def upload():
    try:
        item_code_19 = request.form.get('item_code_19', '').strip()
        supplier_item_code = request.form.get('supplier_item_code', '').strip()
        part_no = request.form.get('part_no', '').strip()

        if not item_code_19 or not supplier_item_code:
            return jsonify({'error': '19位编码和厂家物项编号为必填项'}), 400

        file = request.files.get('file')
        if not file or file.filename == '':
            return jsonify({'error': '请上传PDF文件'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '只支持PDF格式文件'}), 400

        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{uuid.uuid4().hex}_{secure_filename(file.filename)}")
        file.save(pdf_path)

        from utils.pdf_parser import detect_pdf_type, extract_qcp_data
        qcp_data = extract_qcp_data(pdf_path)

        if not qcp_data:
            os.remove(pdf_path)
            return jsonify({'error': '未能从PDF提取工序数据'}), 422

        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"CNPE_转换_{item_code_19}.xlsx")
        template_path = os.path.join(VERCEL_APP_DIR, 'templates', 'CNPE_质量计划导入Excel模板.xlsx')

        if not os.path.exists(template_path):
            os.remove(pdf_path)
            return jsonify({'error': 'Excel模板不存在'}), 500

        from utils.excel_filler import fill_cnpe_template
        fill_cnpe_template(template_path, output_path, qcp_data, item_code_19, part_no, supplier_item_code)
        os.remove(pdf_path)

        return send_file(output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"CNPE_转换_{item_code_19}.xlsx")

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'dir': VERCEL_APP_DIR})


# Vercel serverless handler
handler = app
