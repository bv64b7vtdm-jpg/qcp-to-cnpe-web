"""
Flask web app for QCP to CNPE Excel conversion.
Surgical XML approach: keep template rows 1-9 intact, only replace data row values.
"""

import os
import re
import shutil
import uuid
import zipfile
import tempfile

import pdfplumber
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'template', 'CNPE_质量计划导入Excel模板.xlsx')

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

        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'],
                                f"{uuid.uuid4().hex}_{secure_filename(file.filename)}")
        file.save(pdf_path)

        steps = extract_qcp_with_whr(pdf_path)

        if not steps:
            os.remove(pdf_path)
            return jsonify({'error': '未能从PDF中识别到QCP工序数据。请确认PDF为文本型（不是扫描件）。'}), 422

        output_name = f"CNPE_转换_{item_code_19}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_name)
        fill_template_surgical(output_path, steps, item_code_19, supplier_item_code)
        os.remove(pdf_path)

        return send_file(output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_name)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ─── PDF 解析 ────────────────────────────────────────────────────────────

STEP_RE = re.compile(r'^([A-Z]\d+(?:\.\d+)*\*?)$')
WHR_RE = re.compile(r'\b([WHR])\b')
DOC_RE = re.compile(r'^([A-Z]{2,}\d+)')


def extract_qcp_with_whr(pdf_path):
    """从 QCP PDF 提取工序数据。

    19 列 QCP 表结构（三门泵装配 PDF）：
      col 0  → 工序号
      col 1  → 工序名称
      col 3  → 依据文件编号
      col 8  → S列选点（H/W/R）
      col 17 → 备注
    """
    all_steps = []
    seen_keys = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if len(table) == 0:
                    continue
                if len(table[0]) not in (18, 19, 20):
                    continue

                for row in table:
                    if not row:
                        continue
                    cells = [str(c).strip() if c else '' for c in row]

                    for idx, cell in enumerate(cells):
                        m = STEP_RE.match(cell)
                        if not m:
                            continue
                        step_num = m.group(1).rstrip('*')
                        inner = step_num.lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
                        if inner.startswith(('20', '19')):
                            continue

                        # S列选点（col 8）
                        s_col_idx = idx + 8
                        whr_s = ''
                        if s_col_idx < len(cells):
                            raw = cells[s_col_idx].replace('点', '').strip()
                            if raw in ('W', 'H', 'R'):
                                whr_s = raw + '点'

                        # col 11 = A列选点
                        a_col_idx = idx + 11
                        whr_a = ''
                        if a_col_idx < len(cells):
                            raw = cells[a_col_idx].replace('点', '').strip()
                            if raw in ('W', 'H', 'R'):
                                whr_a = raw + '点'

                        # 名称
                        name = ''
                        for offset in range(1, len(cells)):
                            nc = cells[idx + offset]
                            if not nc or nc in (
                                '序号', 'No.', '工序名称', 'Process Name',
                                '报告编号', 'Remark', '备注', '版', 'Rev.',
                                'S', 'C', 'O', 'W', 'H', 'R',
                                '图号', 'Specification', '零件名称', 'Part Name',
                                'Part No.', '零件编号', '材质', 'Material',
                                'Pending', 'No report', '查', 'NA'
                            ):
                                continue
                            if re.match(r'^\d{2,4}[-./]', nc):
                                continue
                            name = nc.split('\n')[0].strip()
                            break

                        if not name or name == step_num:
                            continue

                        # BOM 过滤
                        if '.' not in step_num:
                            if re.search(r'\d', name):
                                break
                            # A1 类保留

                        # 依据文件编号（col 3）
                        doc_no = ''
                        doc_col_idx = idx + 3
                        if doc_col_idx < len(cells):
                            raw = cells[doc_col_idx].strip()
                            if DOC_RE.match(raw):
                                doc_no = raw.split('\n')[0].strip()

                        # 备注（col 17）
                        remark = ''
                        rem_col_idx = idx + 17
                        if rem_col_idx < len(cells):
                            raw = cells[rem_col_idx].strip()
                            if raw == 'No report':
                                remark = 'No report'
                            elif raw and raw not in ('No report', 'NA', '查'):
                                remark = raw.split('\n')[0].strip()

                        key = f"{step_num}|{name}"
                        if key not in seen_keys:
                            seen_keys.add(key)
                            all_steps.append({
                                'step': step_num,
                                'name': name,
                                'whr_s': whr_s,
                                'whr_a': whr_a,
                                'doc_no': doc_no,
                                'remark': remark,
                            })
                        break

    def sort_key(s):
        m = re.match(r'^([A-Z])(\d+(?:\.\d+)*)', s['step'])
        if m:
            return (m.group(1), [int(x) for x in m.group(2).split('.')])
        return (s['step'], [])

    all_steps.sort(key=sort_key)
    return all_steps


def q_val(step):
    """Q列（是否产生报告）判断逻辑。优先级：
    1. 先决条件检查 / 计划关闭 → N
    2. 备注含 'No report' → N
    3. 有 WHR 点 → Y
    4. 其他 → N
    """
    n = step['name']
    if '先决条件检查' in n or '计划关闭' in n:
        return 'N'
    if 'No report' in step['remark']:
        return 'N'
    if step['whr_s']:
        return 'Y'
    return 'N'


def fill_template_surgical(output_path, steps, item_code_19, supplier_item_code):
    """手术式 XML 修改：
    1. 复制模板 xlsx
    2. 替换 xl/sharedStrings.xml（追加新字符串）
    3. 替换 xl/worksheets/sheet2.xml（保留 rows 1-9，改写 rows 10+ 的 cell 值）
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"模板文件不存在: {TEMPLATE_PATH}")

    if os.path.exists(output_path):
        os.remove(output_path)
    shutil.copy(TEMPLATE_PATH, output_path)

    pid = os.getpid()
    work_dir = f'/tmp/pdf_work/xlsx_{pid}'
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)

    with zipfile.ZipFile(output_path, 'r') as z:
        z.extractall(work_dir)

    # ── 读取原始 XML ────────────────────────────────────────────────
    sheet_path = os.path.join(work_dir, 'xl', 'worksheets', 'sheet2.xml')
    with open(sheet_path, 'r', encoding='utf-8') as f:
        sheet_content = f.read()

    ss_path = os.path.join(work_dir, 'xl', 'sharedStrings.xml')
    with open(ss_path, 'r', encoding='utf-8') as f:
        ss_content = f.read()

    # ── 追加新字符串到 sharedStrings ──────────────────────────────
    existing = re.findall(r'<si>(.*?)</si>', ss_content, re.DOTALL)
    exist_map = {}
    for i, m in enumerate(existing):
        txt = re.sub(r'<[^>]+>', '', m).strip()
        exist_map[txt] = i

    total_count = int(re.search(r'count="(\d+)"', ss_content).group(1))
    total_unique = int(re.search(r'uniqueCount="(\d+)"', ss_content).group(1))

    new_strs = []

    def get_idx(text):
        if text in exist_map:
            return exist_map[text]
        if text not in new_strs:
            new_strs.append(text)
        return len(existing) + new_strs.index(text)

    get_idx(item_code_19)
    get_idx(supplier_item_code)
    for s in steps:
        get_idx(s['step'])
        get_idx(s['name'])
        if s['whr_s']:
            get_idx(s['whr_s'])
        if s['whr_a']:
            get_idx(s['whr_a'])
        if s['doc_no']:
            get_idx(s['doc_no'])
        if s['remark']:
            get_idx(s['remark'])
        get_idx('否')
        get_idx(q_val(s))

    for txt in new_strs:
        ss_content = ss_content.replace(
            '</sst>',
            f'<si><t xml:space="preserve">{txt}</t></si></sst>'
        )
    new_total = total_count + sum(1 for t in new_strs if t not in exist_map)
    new_unique = total_unique + sum(1 for t in new_strs if t not in exist_map)
    ss_content = re.sub(r'count="\d+"', f'count="{new_total}"', ss_content, count=1)
    ss_content = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{new_unique}"', ss_content, count=1)

    all_si = re.findall(r'<si>(.*?)</si>', ss_content, re.DOTALL)
    str_map = {}
    for i, m in enumerate(all_si):
        txt = re.sub(r'<[^>]+>', '', m).strip()
        str_map[txt] = i

    # ── 手术式替换 sheet2.xml ─────────────────────────────────────
    # 策略：找到所有 <row r="N">...</row>，替换 row 10+ 的行内容
    # 保留 row 1-9 完全原样（包括 hidden 属性）

    def replace_row_content(m):
        row_tag = m.group(0)
        rnum_match = re.search(r'<row r="(\d+)"', row_tag)
        if not rnum_match:
            return row_tag
        rnum = int(rnum_match.group(1))

        # 保留 rows 1-9 原样
        if rnum <= 9:
            return row_tag

        # rows 10+：找到对应的 step 数据
        step_idx = rnum - 10
        if step_idx >= len(steps):
            # 超出 step 数量，删除该行（返回空）
            return ''

        s = steps[step_idx]
        seq = (step_idx + 1) * 10

        # 提取原始 row 的属性（spans, ht, customHeight 等）
        row_attrs = re.search(r'<row r="\d+"(.*?)>', row_tag)
        attrs = row_attrs.group(1) if row_attrs else ''

        # 生成 cell 内容
        cells_xml = ''

        # D列（排序号，直接数值，无 t 属性）
        cells_xml += f'<c r="D{rnum}"{attrs}><v>{seq}</v></c>'

        # F列（厂家物项编码，s="31" 与模板一致）
        cells_xml += f'<c r="F{rnum}" s="31" t="s"><v>{str_map[supplier_item_code]}</v></c>'

        # H列（工序编号）
        cells_xml += f'<c r="H{rnum}" t="s"><v>{str_map[s["step"]]}</v></c>'

        # I列（工序名称）
        cells_xml += f'<c r="I{rnum}" t="s"><v>{str_map[s["name"]]}</v></c>'

        # K列（选点S）
        if s['whr_s']:
            cells_xml += f'<c r="K{rnum}" t="s"><v>{str_map[s["whr_s"]]}</v></c>'

        # M列（依据文件编号）
        if s['doc_no']:
            cells_xml += f'<c r="M{rnum}" t="s"><v>{str_map[s["doc_no"]]}</v></c>'

        # P列（实物是否须有编号）
        cells_xml += f'<c r="P{rnum}" t="s"><v>{str_map["否"]}</v></c>'

        # Q列（是否产生报告）
        cells_xml += f'<c r="Q{rnum}" t="s"><v>{str_map[q_val(s)]}</v></c>'

        # R列（选点A1）
        if s['whr_a']:
            cells_xml += f'<c r="R{rnum}" t="s"><v>{str_map[s["whr_a"]]}</v></c>'

        # AA列（备注）
        if s['remark']:
            cells_xml += f'<c r="AA{rnum}" t="s"><v>{str_map[s["remark"]]}</v></c>'

        return f'<row r="{rnum}"{attrs}>{cells_xml}</row>'

    # 匹配所有 <row ...>...</row>
    sheet_content = re.sub(r'<row r="\d+".*?</row>', replace_row_content, sheet_content, flags=re.DOTALL)

    with open(sheet_path, 'w', encoding='utf-8') as f:
        f.write(sheet_content)
    with open(ss_path, 'w', encoding='utf-8') as f:
        f.write(ss_content)

    # ── 重新打包 xlsx ──────────────────────────────────────────────
    if os.path.exists(output_path):
        os.remove(output_path)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(work_dir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, work_dir)
                zout.write(fpath, arcname)

    shutil.rmtree(work_dir)
