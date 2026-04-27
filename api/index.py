"""
Vercel serverless handler for QCP→CNPE Flask app.
"""

import os
import sys
import io
import json

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from flask import Flask, request, render_template, send_file, jsonify, session
from werkzeug.utils import secure_filename
import uuid

ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.secret_key = os.environ.get('SECRET_KEY', 'qcp-cnpe-secret-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = '/tmp/qcp-uploads'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login', methods=['POST'])
def login():
    session['authenticated'] = True
    return jsonify({'status': 'ok'})


@app.route('/upload', methods=['POST'])
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
        template_path = os.path.join(BASE_DIR, 'templates', 'CNPE_质量计划导入Excel模板.xlsx')

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
    return jsonify({'status': 'ok'})


def handler(event, context=None):
    """Vercel Python serverless entry point."""
    path = event.get('path', '/')
    http_method = event.get('httpMethod', 'GET')
    headers = event.get('headers', {})
    query = event.get('query', {})
    body = event.get('body', '')

    if http_method == 'OPTIONS':
        return {'statusCode': 200, 'headers': {'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type'}, 'body': ''}

    # Build URL for test client
    scheme = headers.get('x-forwarded-proto', 'https')
    host = headers.get('host', 'localhost')
    url = f'{scheme}://{host}{path}'
    if query:
        url += '?' + '&'.join(f'{k}={v}' for k, v in query.items())

    headers_lower = {k.lower(): v for k, v in headers.items()}

    with app.test_client() as client:
        if http_method == 'GET':
            rv = client.get(path, headers=headers_lower)
        elif http_method == 'POST':
            content_type = headers.get('content-type', 'application/json')
            if 'application/json' in content_type:
                rv = client.post(path, data=body, headers=headers_lower, content_type=content_type)
            elif 'multipart/form-data' in content_type:
                rv = client.post(path, data=body, headers=headers_lower)
            else:
                rv = client.post(path, data=body, headers=headers_lower)
        else:
            rv = client.open(path, method=http_method, headers=headers_lower)

        response_headers = {}
        for k, v in rv.headers:
            k_lower = k.lower()
            if k_lower not in ('content-encoding', 'transfer-encoding', 'connection'):
                response_headers[k.title()] = v

        response_body = rv.get_data(as_text=True)
        if not isinstance(response_body, str):
            response_body = str(response_body)

        return {
            'statusCode': rv.status_code,
            'headers': response_headers,
            'body': response_body,
            'isBase64Encoded': False,
        }
