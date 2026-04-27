"""
Flask web application for QCP to CNPE Excel conversion.
"""

import os
import uuid
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

from utils.pdf_parser import detect_pdf_type, extract_qcp_data, is_scanned
from utils.excel_filler import fill_cnpe_template

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = '/tmp/qcp-uploads'

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Render the main upload page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    """Handle file upload and processing."""
    try:
        # Get form data
        item_code_19 = request.form.get('item_code_19', '').strip()
        part_no = request.form.get('part_no', '').strip()
        supplier_item_code = request.form.get('supplier_item_code', '').strip()
        
        # Validate required fields
        if not item_code_19:
            return jsonify({'error': '19位编码为必填项'}), 400
        if not supplier_item_code:
            return jsonify({'error': '厂家物项编号为必填项'}), 400
        
        # Check file
        if 'file' not in request.files:
            return jsonify({'error': '请上传PDF文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择PDF文件'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': '只支持PDF格式文件'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4().hex}_{filename}"
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(pdf_path)
        
        # Detect PDF type
        pdf_type = detect_pdf_type(pdf_path)
        print(f"PDF type detected: {pdf_type}")
        
        # Extract QCP data
        qcp_data = extract_qcp_data(pdf_path)
        
        if not qcp_data:
            # Clean up temp file
            os.remove(pdf_path)
            return jsonify({
                'error': '未能从PDF中提取到工序数据。请确认PDF为QCP质量计划格式。',
                'pdf_type': pdf_type
            }), 422
        
        print(f"Extracted {len(qcp_data)} process steps")
        
        # Generate output file
        output_filename = f"CNPE_转换_{item_code_19}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # Template path
        template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            'templates',
            'CNPE_质量计划导入Excel模板.xlsx'
        )
        
        # Check if template exists
        if not os.path.exists(template_path):
            os.remove(pdf_path)
            return jsonify({'error': 'Excel模板文件不存在'}), 500
        
        # Fill template
        fill_cnpe_template(
            template_path=template_path,
            output_path=output_path,
            qcp_data=qcp_data,
            item_code_19=item_code_19,
            part_no=part_no,
            supplier_item_code=supplier_item_code
        )
        
        # Clean up PDF
        os.remove(pdf_path)
        
        # Return file for download
        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint."""
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)