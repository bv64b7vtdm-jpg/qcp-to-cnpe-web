"""
Flask web app for QCP to CNPE conversion - generic version.
Parses any QCP PDF table using pdfplumber.
"""

import os
import re
import uuid
import tempfile
from datetime import datetime

import pdfplumber
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.secret_key = os.environ.get('SECRET_KEY', 'qcp-cnpe-secret-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/health')
def health():
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

        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{uuid.uuid4().hex}_{secure_filename(file.filename)}")
        file.save(pdf_path)

        steps = extract_qcp_data(pdf_path)

        if not steps:
            os.remove(pdf_path)
            return jsonify({'error': '未能从PDF中识别到QCP工序数据。请确认PDF为文本型（不是扫描件）。'}), 422

        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"CNPE_转换_{item_code_19}.xlsx")
        generate_cnpe_excel(output_path, steps, item_code_19, part_no, supplier_item_code)
        os.remove(pdf_path)

        return send_file(output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"CNPE_转换_{item_code_19}.xlsx")

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


STEP_RE = re.compile(r'^([A-Z](?:\d+(?:\.\d+)*))$')

def extract_qcp_data(pdf_path):
    """Extract QCP steps from any QCP PDF using table detection."""
    all_steps = []
    seen_keys = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 2:
                        continue
                    cells = [str(c).strip() if c else '' for c in row]
                    for idx, cell in enumerate(cells):
                        m = STEP_RE.match(cell)
                        if not m:
                            continue
                        step_num = m.group(1)
                        # Skip year patterns like A2026, B2025
                        inner = step_num.lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
                        if inner.startswith('20') or inner.startswith('19'):
                            continue
                        step_name = None
                        for offset in range(1, 4):
                            ni = idx + offset
                            if ni < len(cells):
                                nc = cells[ni].strip()
                                if not nc:
                                    continue
                                if nc in ('序号', 'No.', '工序名称', 'Process Name', '报告编号', 'Remark', '备注', '版', 'Rev.', 'S', 'C', 'O', 'W', 'H', 'R', '图号', 'Specification', '零件名称', 'Part Name', 'Part No.', '零件编号', '材质', 'Material', 'Pending', 'No report', '查', 'NA', 'R', 'W', 'H'):
                                    continue
                                if re.match(r'^\d{2,4}[-./]', nc):
                                    continue
                                if len(nc) < 2:
                                    continue
                                step_name = nc.split('\n')[0].strip()
                                break
                        if step_name and len(step_name) >= 2:
                            # Skip section headers where step == name
                            if step_num == step_name:
                                continue
                            # For A-series without sub-number, skip BOM items
                            letter = step_num[0]
                            if letter == 'A' and '.' not in step_num:
                                continue
                            key = f"{step_num}|{step_name}"
                            if key not in seen_keys:
                                seen_keys.add(key)
                                all_steps.append({
                                    'step': step_num,
                                    'name': step_name,
                                    'content': ''
                                })

    def sort_key(s):
        m = re.match(r'^([A-Z])(\d+(?:\.\d+)*)', s['step'])
        if m:
            return (m.group(1), [int(x) for x in m.group(2).split('.')])
        return (s['step'], [])

    all_steps.sort(key=sort_key)
    return all_steps


def generate_cnpe_excel(output_path, steps, item_code_19, part_no, supplier_item_code):
    """Generate CNPE Excel from QCP data."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment

        wb = Workbook()
        ws = wb.active
        ws.title = 'QCP工序'

        # Title
        ws['A1'] = 'QCP转CNPE质量计划'
        ws['A1'].font = Font(bold=True, size=14)
        ws.merge_cells('A1:F1')

        # Info
        ws['A3'] = '19位编码'
        ws['B3'] = item_code_19
        ws['A4'] = '零件号'
        ws['B4'] = part_no or ''
        ws['A5'] = '厂家物项编号'
        ws['B5'] = supplier_item_code
        ws['A6'] = '生成日期'
        ws['B6'] = datetime.now().strftime('%Y-%m-%d')
        ws['A7'] = '工序数量'
        ws['B7'] = len(steps)

        # Header row
        hdr_fill = PatternFill(start_color='667EEA', end_color='667EEA', fill_type='solid')
        hdr_font = Font(bold=True, color='FFFFFF', size=11)
        headers = ['序号', '工序号', '工序名称', '质量控制点', '报告编号', '备注']
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=9, column=col, value=h)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = Alignment(horizontal='center', vertical='center')

        # Data rows
        for i, step in enumerate(steps):
            r = 10 + i
            ws.cell(row=r, column=1, value=i + 1)
            ws.cell(row=r, column=2, value=step['step'])
            ws.cell(row=r, column=3, value=step['name'])
            ws.cell(row=r, column=4, value='')
            ws.cell(row=r, column=5, value='')
            ws.cell(row=r, column=6, value=step.get('content', ''))
            if i % 2 == 0:
                alt_fill = PatternFill(start_color='F8F6FF', end_color='F8F6FF', fill_type='solid')
                for col in range(1, 7):
                    ws.cell(row=r, column=col).fill = alt_fill

        # Column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 35

        # Sheet 2: 填写说明
        ws2 = wb.create_sheet('填写说明')
        notes = [
            ['QCP转CNPE填写说明'],
            [],
            ['字段', '说明', '示例'],
            ['19位编码', '采购订单中的19位物料编码', '1P12345678901234567'],
            ['零件号', '可选，零件编号', 'A19'],
            ['厂家物项编号', '供应商提供的物项编号', 'SK93303-B1'],
            ['序号', '自动生成的序号', '1, 2, 3...'],
            ['工序号', 'QCP中的编号', 'B1.0, B1.1, C1.1.1...'],
            ['工序名称', '该工序的名称', '水压试验、性能试验...'],
            ['质量控制点', 'R/W/H', 'W'],
            ['报告编号', '对应的检验程序编号', 'SKR80004540...'],
            ['备注', '检验内容或备注', ''],
            [],
            ['支持的QCP格式', '适用于KSB SEC-KSB、三门、漳州等多种QCP表格格式', ''],
        ]
        for ri, row_data in enumerate(notes, 1):
            for ci, val in enumerate(row_data, 1):
                c = ws2.cell(row=ri, column=ci, value=val)
                if ri == 3:
                    c.font = Font(bold=True)
        ws2.column_dimensions['A'].width = 18
        ws2.column_dimensions['B'].width = 40
        ws2.column_dimensions['C'].width = 30

        wb.save(output_path)

    except ImportError:
        # Fallback to CSV if openpyxl not available
        import csv
        with open(output_path.replace('.xlsx', '.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['QCP转CNPE质量计划'])
            writer.writerow(['19位编码', item_code_19])
            writer.writerow(['零件号', part_no or ''])
            writer.writerow(['厂家物项编号', supplier_item_code])
            writer.writerow(['生成日期', datetime.now().strftime('%Y-%m-%d')])
            writer.writerow([])
            writer.writerow(['序号', '工序号', '工序名称', '备注'])
            for i, s in enumerate(steps, 1):
                writer.writerow([i, s['step'], s['name'], s.get('content', '')])
        os.rename(output_path.replace('.xlsx', '.csv'), output_path)
