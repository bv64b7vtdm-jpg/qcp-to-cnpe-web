"""
Flask web application for QCP to CNPE Excel conversion.
Vercel serverless version.
"""

import os
import uuid
import sys
from flask import Flask, request, render_template, send_file, jsonify, session
from functools import wraps
from werkzeug.utils import secure_filename

# Determine base directory - works in both local and Vercel serverless
VERCEL_APP_DIR = os.environ.get('VERCEL_APP_DIR', '')
if VERCEL_APP_DIR:
    BASE_DIR = VERCEL_APP_DIR
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.secret_key = os.environ.get('SECRET_KEY', 'qcp-cnpe-secret-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = '/tmp/qcp-uploads'

APP_PASSWORD = os.environ.get('QCP_PASSWORD', 'KSB2026')



def _is_pump_casing_qcp(pdf_path):
    """检查 PDF 是否是泵壳相关 QCP（按名称识别）"""
    import pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ''
                if '泵壳' in text:
                    return True
    except Exception:
        pass
    return False


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
        password = data.get('password', '')
        if password == APP_PASSWORD:
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
        part_no = request.form.get('part_no', '').strip()
        supplier_item_code = request.form.get('supplier_item_code', '').strip()

        if not item_code_19:
            return jsonify({'error': '19位编码为必填项'}), 400
        if not supplier_item_code:
            return jsonify({'error': '厂家物项编号为必填项'}), 400

        if 'file' not in request.files:
            return jsonify({'error': '请上传PDF文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择PDF文件'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '只支持PDF格式文件'}), 400

        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4().hex}_{filename}"
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(pdf_path)

        # 泵壳及子件特殊规则（2026-08-19 陈老师反馈）：
        # 泵壳本体及所有泵壳相关子件（底脚/安全端/焊接见证件/母材见证件/焊材）
        # 共用零件号 101.01 → 物项标识码固定为 1907RCP10101
        # 识别方法：PDF 题目或任意页含"泵壳" → 强制覆盖（按名称）
        if _is_pump_casing_qcp(pdf_path):
            item_code_19 = '1907RCP10101'
            app.logger.info(f'泵壳相关 QCP（按名称识别）→ 强制使用 1907RCP10101')

        # 调用 v3.7 一站式转换
        template_path = os.path.join(BASE_DIR, 'templates', 'CNPE_质量计划导入Excel模板.xlsx')

        if not os.path.exists(template_path):
            os.remove(pdf_path)
            return jsonify({'error': f'Excel模板文件不存在: {template_path}'}), 500

        output_filename = f"CNPE_转换_{item_code_19}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        # v3.7 接口：convert_qcp_to_cnpe(pdf, tmpl, item_code, supplier_count=None, out_path=None)
        # sys.path 加 BASE_DIR 让 qcp_converter 可被导入
        if BASE_DIR not in sys.path:
            sys.path.insert(0, BASE_DIR)
        from qcp_converter import convert_qcp_to_cnpe
        out_path, n_procedures, supplier_count = convert_qcp_to_cnpe(
            pdf_path=pdf_path,
            tmpl_path=template_path,
            item_code=item_code_19,
            out_path=output_path
        )

        os.remove(pdf_path)

        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'base_dir': BASE_DIR})


handler = app
