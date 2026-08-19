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
        item_code_19_input = request.form.get('item_code_19', '').strip()
        supplier_item_code = request.form.get('supplier_item_code', '').strip()
        part_no = request.form.get('part_no', '').strip()

        if not supplier_item_code:
            return jsonify({'error': '厂家物项编号为必填项'}), 400

        file = request.files.get('file')
        if not file or file.filename == '':
            return jsonify({'error': '请上传PDF文件'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '只支持PDF格式文件'}), 400

        # 保存上传的 PDF
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'],
                                f"{uuid.uuid4().hex}_{secure_filename(file.filename)}")
        file.save(pdf_path)

        # 19位编码：优先用用户输入，其次从 part_no 动态计算（v2.0 修复）
        if item_code_19_input:
            item_code_19 = item_code_19_input
        elif part_no:
            # 公式：1907RCP + 零件号去小数点
            item_code_19 = f"1907RCP{part_no.replace('.', '')}"
        else:
            item_code_19 = "1907RCP10101"  # 默认值，实际场景下应该报错

        # 泵壳及子件特殊规则（2026-08-19 陈老师反馈）：
        # 泵壳本体及所有泵壳相关子件（底脚/安全端/焊接见证件/母材见证件/焊材）
        # 共用零件号 101.01 → 物项标识码固定为 1907RCP10101
        # 识别方法：PDF 题目或任意页含"泵壳" → 强制覆盖（按名称）
        if _is_pump_casing_qcp(pdf_path):
            item_code_19 = '1907RCP10101'
            app.logger.info(f'泵壳相关 QCP（按名称识别）→ 强制使用 1907RCP10101')

        # 调用 v3.7 一站式转换
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'templates', 'CNPE_质量计划导入Excel模板.xlsx'
        )

        output_name = f"{item_code_19}_{part_no}_QCP导入数据.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_name)

        # v3.7 接口：convert_qcp_to_cnpe(pdf, tmpl, item_code, supplier_count=None, out_path=None)
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
            download_name=output_name
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ─── PDF 解析 ────────────────────────────────────────────────────────────

STEP_RE = re.compile(r'^([A-Z]\d+(?:\.\d+)*\*?)$')
# 二重装备格式步骤号：0, 1, 2, 6-1, B-1, B-5-a, B-12-1 等
STEP_NUM_RE = re.compile(r'^(\d+[A-Z]?(?:-\d+)*|[A-Z]-\d+(?:-\d+)*|[A-Z]-\d+-[a-z])$')
WHR_RE = re.compile(r'\b([WHR])\b')
DOC_RE = re.compile(r'^([A-Z]{2,}\d+)')
CODE19_RE = re.compile(r'(SMX[0-9A-Z]{16,20}|(?:SMX)?44400[A-Z0-9]{10,14}GN?)')


def extract_code19_from_filename(filename):
    """从文件名提取19位编码。"""
    m = CODE19_RE.search(filename.upper())
    if m:
        code = m.group(0)
        # 标准化：统一为SMX开头20位
        if code.startswith('SMX') and len(code) >= 16:
            return code[:20] if len(code) > 20 else code
    return None


def extract_code19_from_pdf(pdf_path):
    """从PDF内容页提取19位编码（通常在第1页或表格表头）。"""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:3]:  # 只查前3页
            text = page.extract_text() or ''
            # 匹配 SMX44400Z... 格式
            matches = re.findall(r'(SMX44400[A-Z0-9]{12}GN?)', text.upper())
            if matches:
                return matches[0][:20] if len(matches[0]) > 20 else matches[0]
            # 也匹配 SMX 开头 + 16位
            matches2 = re.findall(r'(SMX[0-9A-Z]{16})', text.upper())
            if matches2:
                return matches2[0]
    return None


def validate_code19(filename, pdf_path):
    """核对文件名编码与文件内编码，返回准确实码。"""
    code_from_name = extract_code19_from_filename(filename)
    code_from_pdf = extract_code19_from_pdf(pdf_path)

    if code_from_name and code_from_pdf:
        if code_from_name == code_from_pdf:
            return code_from_pdf, f'{code_from_pdf}'
        else:
            # 不一致，以文件内为准
            return code_from_pdf, f'{code_from_name}→{code_from_pdf}'
    elif code_from_pdf:
        return code_from_pdf, f'{code_from_pdf}'
    elif code_from_name:
        return code_from_name, f'{code_from_name}'
    return 'UNKNOWN', '未能识别'

def _is_valid_step(s):
    if not s: return False
    if s.isdigit(): return True
    if re.match(r'^[A-Z]\d+(?:\.\d+)*\*?$', s): return True
    if re.match(r'^\d+-\d+$', s): return True
    if re.match(r'^[A-Z]-\d+(?:-\d+)*$', s): return True
    if re.match(r'^[A-Z]-\d+-[a-z]$', s): return True
    return False


def extract_qcp_with_whr(pdf_path):
    """从 QCP PDF 提取工序数据。支持两种格式：

    【SEC-KSB 19列格式】（三门泵装配 PDF）
      col 0  → 工序号（如 A1, B1.1, E1.5）
      col 1  → 工序名称
      col 3  → 依据文件编号
      col 8  → S列选点（H/W/R）
      col 11 → A列选点
      col 17 → 备注

    【二重装备 24列格式】（泵壳底脚锻件 PDF）
      col 0  → 工序号（如 0, 1, 6-1, B-1, B-5-a）
      col 1  → 工序名称
      col 4  → 依据文件编号
      col 12 → A列选点
      col 15 → S列选点（H/W/R）
      col 22 → 备注
    """
    all_steps = []
    seen_keys = set()

    SKIP_CELLS = {
        '序号', 'No.', '工序名称', 'Process Name',
        '报告编号', 'Remark', '备注', '版', 'Rev.',
        'S', 'C', 'O', 'W', 'H', 'R',
        '图号', 'Specification', '零件名称', 'Part Name',
        'Part No.', '零件编号', '材质', 'Material',
        'Pending', 'No report', '查', 'NA'
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if len(table) == 0:
                    continue
                n_cols = len(table[0])

                # 判断格式：19列=SEC-KSB，24列=二重装备
                if n_cols >= 20:
                    fmt = 24  # 二重装备
                elif n_cols >= 18:
                    fmt = 19  # SEC-KSB
                else:
                    continue

                for row in table:
                    if not row:
                        continue
                    cells = [str(c).strip() if c else '' for c in row]

                    for idx, cell in enumerate(cells):
                        # 工序号只允许在前5列（col 0~4），避免误识表头单元格
                        if idx > 4:
                            continue

                        step_num_raw = cell.strip()
                        if not _is_valid_step(step_num_raw):
                            continue

                        step_num = step_num_raw
                        inner = step_num.lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/-')
                        if inner.startswith(('20', '19')):
                            continue

                        # ── 名称 ──────────────────────────
                        name = ''
                        for offset in range(1, min(len(cells), 25)):
                            nc = cells[idx + offset] if idx + offset < len(cells) else ''
                            if not nc or nc in SKIP_CELLS:
                                continue
                            if re.match(r'^\d{2,4}[-./]', nc):
                                continue
                            name = nc.split('\n')[0].strip()
                            break
                        # 过滤无效名称：日期、PRE、仅大写字母等
                        if not name or name == step_num:
                            continue
                        if name in ('PRE', 'N/A', 'NA', 'Pending'):
                            continue
                        # 日期格式（如 2024-08-30）
                        if re.match(r'^\d{4}[-/]\d{2}[-/]\d{2}$', name):
                            continue

                        # BOM 过滤
                        if '.' not in step_num and '-' not in step_num:
                            if re.search(r'\d', name):
                                break

                        # ── WHR 提取 ──────────────────────
                        if fmt == 24:
                            # 二重装备：col 15=S列
                            s_idx = idx + 15
                            whr_s = ''
                            if s_idx < len(cells):
                                raw = cells[s_idx].replace('点', '').strip()
                                if raw in ('W', 'H', 'R'):
                                    whr_s = raw + '点'
                            # col 12=A列
                            a_idx = idx + 12
                            whr_a = ''
                            if a_idx < len(cells):
                                raw = cells[a_idx].replace('点', '').strip()
                                if raw in ('W', 'H', 'R'):
                                    whr_a = raw + '点'
                            # col 17=C列
                            c_idx = idx + 17
                            whr_c = ''
                            if c_idx < len(cells):
                                raw = cells[c_idx].replace('点', '').strip()
                                if raw in ('W', 'H', 'R'):
                                    whr_c = raw + '点'
                            # col 4=文件编号
                            doc_no = ''
                            doc_idx = idx + 4
                            if doc_idx < len(cells):
                                raw = cells[doc_idx].strip()
                                if DOC_RE.match(raw):
                                    doc_no = raw.split('\n')[0].strip()
                            # col 22=备注
                            remark = ''
                            rem_idx = idx + 22
                            if rem_idx < len(cells):
                                raw = cells[rem_idx].strip()
                                if raw == 'No report':
                                    remark = 'No report'
                                elif raw and raw not in ('No report', 'NA', '查'):
                                    remark = raw.split('\n')[0].strip()
                        else:
                            # SEC-KSB：col 8=S列，col 11=A列，col 3=文件，col 17=备注
                            s_idx = idx + 8
                            whr_s = ''
                            if s_idx < len(cells):
                                raw = cells[s_idx].replace('点', '').strip()
                                if raw in ('W', 'H', 'R'):
                                    whr_s = raw + '点'
                            a_idx = idx + 11
                            whr_a = ''
                            if a_idx < len(cells):
                                raw = cells[a_idx].replace('点', '').strip()
                                if raw in ('W', 'H', 'R'):
                                    whr_a = raw + '点'
                            # col 11=C列（SEC-KSB 19列格式）
                            c_idx = idx + 11
                            whr_c = ''
                            if c_idx < len(cells):
                                raw = cells[c_idx].replace('点', '').strip()
                                if raw in ('W', 'H', 'R'):
                                    whr_c = raw + '点'
                            doc_no = ''
                            doc_idx = idx + 3
                            if doc_idx < len(cells):
                                raw = cells[doc_idx].strip()
                                if DOC_RE.match(raw):
                                    doc_no = raw.split('\n')[0].strip()
                            remark = ''
                            rem_idx = idx + 17
                            if rem_idx < len(cells):
                                raw = cells[rem_idx].strip()
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
                                'whr_c': whr_c,
                                'doc_no': doc_no,
                                'remark': remark,
                            })
                        break

    def sort_key(s):
        num = s['step']
        # 统一排序：数字部分 + 字母前缀
        m = re.match(r'^([A-Z]?)(-?)(\d+)(.*)$', num)
        if m:
            prefix = m.group(1) or ''
            dash = m.group(2) or ''
            nums = [int(x) for x in m.group(3).split('.')]
            suffix = m.group(4) or ''
            return (prefix, dash, nums, suffix)
        return (num, '', [], '')

    all_steps.sort(key=sort_key)
    return all_steps


def q_val(step):
    """Q列（是否产生报告）判断逻辑。优先级：
    1. 先决条件检查 / 计划关闭 → N
    2. 备注含 'No report' → N
    3. K列（S列）或R列（A列）有选点（H点/W点/R点）→ Y
    4. K列和R列都无选点时 → N
    """
    n = step['name']
    if '先决条件检查' in n or '计划关闭' in n:
        return 'N'
    if 'No report' in step['remark']:
        return 'N'
    if step['whr_s'] or step['whr_a']:
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
            get_idx(s.get('whr_c'))
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

        # K列（选点S）填 S列 WHR 值
        if s['whr_s']:
            cells_xml += f'<c r="K{rnum}" t="s"><v>{str_map[s["whr_s"]]}</v></c>'
        # X列（选点C）填 C列 WHR 值
        if s.get('whr_c'):
            cells_xml += f'<c r="X{rnum}" t="s"><v>{str_map[s["whr_c"]]}</v></c>'

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
