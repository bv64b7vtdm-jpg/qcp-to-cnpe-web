"""
Batch process 3 QCP PDFs → CNPE Excel.
Extract core logic without Flask imports.
"""
import os
import re
import shutil
import zipfile
import pdfplumber

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'template', 'CNPE_质量计划导入Excel模板.xlsx')
OUT_DIR = '/mnt/d/SynologyDrive/QCP to Excel/'

PDFS = [
    {
        'path': '/mnt/d/SynologyDrive/NAS temp/SMX44400Z59101A04GN Rev.A1 PRE 三门核电项目5、6号机组反应堆冷却剂泵泵壳塞孔不锈钢堆焊质量计划（11232010400315_ _ _）.pdf',
        'part_no': '泵壳塞孔不锈钢堆焊',
    },
    {
        'path': '/mnt/d/SynologyDrive/NAS temp/SMX44400Z61101A04GN Rev.A1 PRE 三门核电项目5、6号机组反应堆冷却剂泵泵壳产品堆焊见证件制造质量计划（11232010400317_ _ _）.pdf',
        'part_no': '泵壳产品堆焊见证件制造',
    },
    {
        'path': '/mnt/d/SynologyDrive/NAS temp/SMX44400Z64101A04GN Rev.A1 PRE 三门核电项目5、6号机组反应堆冷却剂泵泵壳镍基合金气体保护焊焊丝复验质量计划（1123201040031001）.pdf',
        'part_no': '泵壳镍基合金气体保护焊焊丝复验',
    },
]

# ─── Replicated core functions from api/index.py ─────────────────────────────────

STEP_RE = re.compile(r'^([A-Z]\d+(?:\.\d+)*\*?)$')
STEP_NUM_RE = re.compile(r'^(\d+[A-Z]?(?:-\d+)*|[A-Z]-\d+(?:-\d+)*|[A-Z]-\d+-[a-z])$')
DOC_RE = re.compile(r'^([A-Z]{2,}\d+)')
CODE19_RE = re.compile(r'(SMX[0-9A-Z]{16,20}|(?:SMX)?44400[A-Z0-9]{10,14}GN?)')

def _is_valid_step(s):
    if not s: return False
    if s.isdigit(): return True
    if re.match(r'^[A-Z]\d+(?:\.\d+)*\*?$', s): return True
    if re.match(r'^\d+-\d+$', s): return True
    if re.match(r'^[A-Z]-\d+(?:-\d+)*$', s): return True
    if re.match(r'^[A-Z]-\d+-[a-z]$', s): return True
    return False

def extract_code19_from_filename(filename):
    m = CODE19_RE.search(filename.upper())
    if m:
        code = m.group(0)
        if code.startswith('SMX') and len(code) >= 16:
            return code[:20] if len(code) > 20 else code
    return None

def extract_code19_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:3]:
            text = page.extract_text() or ''
            matches = re.findall(r'(SMX44400[A-Z0-9]{12}GN?)', text.upper())
            if matches:
                return matches[0][:20] if len(matches[0]) > 20 else matches[0]
            matches2 = re.findall(r'(SMX[0-9A-Z]{16})', text.upper())
            if matches2:
                return matches2[0]
    return None

def extract_qcp_with_whr(pdf_path):
    """从 QCP PDF 提取工序数据"""
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
                if n_cols >= 20:
                    fmt = 24
                elif n_cols >= 18:
                    fmt = 19
                else:
                    continue
                for row in table:
                    if not row:
                        continue
                    cells = [str(c).strip() if c else '' for c in row]
                    for idx, cell in enumerate(cells):
                        if idx > 4:
                            continue
                        step_num_raw = cell.strip()
                        if not _is_valid_step(step_num_raw):
                            continue
                        step_num = step_num_raw
                        inner = step_num.lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/-')
                        if inner.startswith(('20', '19')):
                            continue
                        name = ''
                        for offset in range(1, min(len(cells), 25)):
                            nc = cells[idx + offset] if idx + offset < len(cells) else ''
                            if not nc or nc in SKIP_CELLS:
                                continue
                            if re.match(r'^\d{2,4}[-./]', nc):
                                continue
                            name = nc.split('\n')[0].strip()
                            break
                        if not name or name == step_num:
                            continue
                        if name in ('PRE', 'N/A', 'NA', 'Pending'):
                            continue
                        if re.match(r'^\d{4}[-/]\d{2}[-/]\d{2}$', name):
                            continue
                        if '.' not in step_num and '-' not in step_num:
                            if re.search(r'\d', name):
                                break
                        if fmt == 24:
                            s_idx = idx + 15
                            whr_s = cells[s_idx].replace('点', '').strip() if s_idx < len(cells) else ''
                            whr_s = whr_s if whr_s in ('W', 'H', 'R') else ''
                            a_idx = idx + 12
                            whr_a = cells[a_idx].replace('点', '').strip() if a_idx < len(cells) else ''
                            whr_a = whr_a if whr_a in ('W', 'H', 'R') else ''
                            c_idx = idx + 17
                            whr_c = cells[c_idx].replace('点', '').strip() if c_idx < len(cells) else ''
                            whr_c = whr_c if whr_c in ('W', 'H', 'R') else ''
                            doc_no = ''
                            doc_idx = idx + 4
                            if doc_idx < len(cells):
                                raw = cells[doc_idx].strip()
                                if DOC_RE.match(raw):
                                    doc_no = raw.split('\n')[0].strip()
                            remark = ''
                            rem_idx = idx + 22
                            if rem_idx < len(cells):
                                raw = cells[rem_idx].strip()
                                if raw == 'No report':
                                    remark = 'No report'
                                elif raw and raw not in ('No report', 'NA', '查'):
                                    remark = raw.split('\n')[0].strip()
                        else:
                            s_idx = idx + 8
                            whr_s = cells[s_idx].replace('点', '').strip() if s_idx < len(cells) else ''
                            whr_s = whr_s if whr_s in ('W', 'H', 'R') else ''
                            a_idx = idx + 11
                            whr_a = cells[a_idx].replace('点', '').strip() if a_idx < len(cells) else ''
                            whr_a = whr_a if whr_a in ('W', 'H', 'R') else ''
                            c_idx = idx + 11
                            whr_c = cells[c_idx].replace('点', '').strip() if c_idx < len(cells) else ''
                            whr_c = whr_c if whr_c in ('W', 'H', 'R') else ''
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
    n = step['name']
    if '先决条件检查' in n or '计划关闭' in n:
        return 'N'
    if 'No report' in step['remark']:
        return 'N'
    if step['whr_s'] or step['whr_a']:
        return 'Y'
    return 'N'

def extract_supplier_code_from_pdf(pdf_path):
    """提取厂家物项编码"""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:3]:
            text = page.extract_text() or ''
            # 找 44400 或 SMX 开头但不是19位编码的字符串（通常是厂家编码）
            matches = re.findall(r'(44400[A-Z0-9]{8,14})', text.upper())
            for m in matches:
                if len(m) >= 8:
                    return m
    return ''

def fill_template_surgical(output_path, steps, item_code_19, supplier_item_code):
    """手术式 XML 修改"""
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
    sheet_path = os.path.join(work_dir, 'xl', 'worksheets', 'sheet2.xml')
    with open(sheet_path, 'r', encoding='utf-8') as f:
        sheet_content = f.read()
    ss_path = os.path.join(work_dir, 'xl', 'sharedStrings.xml')
    with open(ss_path, 'r', encoding='utf-8') as f:
        ss_content = f.read()
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
    def replace_row_content(m):
        row_tag = m.group(0)
        rnum_match = re.search(r'<row r="(\d+)"', row_tag)
        if not rnum_match:
            return row_tag
        rnum = int(rnum_match.group(1))
        if rnum <= 9:
            return row_tag
        step_idx = rnum - 10
        if step_idx >= len(steps):
            return ''
        s = steps[step_idx]
        seq = (step_idx + 1) * 10
        row_attrs = re.search(r'<row r="\d+"(.*?)>', row_tag)
        attrs = row_attrs.group(1) if row_attrs else ''
        cells_xml = ''
        cells_xml += f'<c r="D{rnum}"{attrs}><v>{seq}</v></c>'
        cells_xml += f'<c r="F{rnum}" s="31" t="s"><v>{str_map[supplier_item_code]}</v></c>'
        cells_xml += f'<c r="H{rnum}" t="s"><v>{str_map[s["step"]]}</v></c>'
        cells_xml += f'<c r="I{rnum}" t="s"><v>{str_map[s["name"]]}</v></c>'
        if s['whr_s']:
            cells_xml += f'<c r="K{rnum}" t="s"><v>{str_map[s["whr_s"]]}</v></c>'
        if s.get('whr_c'):
            cells_xml += f'<c r="X{rnum}" t="s"><v>{str_map[s["whr_c"]]}</v></c>'
        if s['doc_no']:
            cells_xml += f'<c r="M{rnum}" t="s"><v>{str_map[s["doc_no"]]}</v></c>'
        cells_xml += f'<c r="P{rnum}" t="s"><v>{str_map["否"]}</v></c>'
        cells_xml += f'<c r="Q{rnum}" t="s"><v>{str_map[q_val(s)]}</v></c>'
        if s['whr_a']:
            cells_xml += f'<c r="R{rnum}" t="s"><v>{str_map[s["whr_a"]]}</v></c>'
        if s['remark']:
            cells_xml += f'<c r="AA{rnum}" t="s"><v>{str_map[s["remark"]]}</v></c>'
        return f'<row r="{rnum}"{attrs}>{cells_xml}</row>'
    sheet_content = re.sub(r'<row r="\d+".*?</row>', replace_row_content, sheet_content, flags=re.DOTALL)
    with open(sheet_path, 'w', encoding='utf-8') as f:
        f.write(sheet_content)
    with open(ss_path, 'w', encoding='utf-8') as f:
        f.write(ss_content)
    if os.path.exists(output_path):
        os.remove(output_path)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(work_dir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, work_dir)
                zout.write(fpath, arcname)
    shutil.rmtree(work_dir)

# ─── Main batch loop ────────────────────────────────────────────────

os.makedirs(OUT_DIR, exist_ok=True)

results = []

for info in PDFS:
    pdf_path = info['path']
    part_no = info['part_no']

    print(f"\n{'='*60}")
    print(f"处理：{part_no}")
    print(f"路径：{pdf_path}")

    fname = os.path.basename(pdf_path)
    code19 = extract_code19_from_filename(fname)
    if not code19:
        code19 = extract_code19_from_pdf(pdf_path)
    print(f"19位编码：{code19}")

    # 提取厂家物项编码
    supplier_code = extract_supplier_code_from_pdf(pdf_path)
    print(f"厂家物项编码：{supplier_code or '(未找到)'}")
    if not supplier_code:
        supplier_code = code19  # fallback

    steps = extract_qcp_with_whr(pdf_path)
    print(f"识别工序数：{len(steps)}")
    for s in steps[:5]:
        print(f"  {s['step']} | {s['name']} | S={s['whr_s']} A={s['whr_a']} C={s.get('whr_c','')}")
    if len(steps) > 5:
        print(f"  ... 共{len(steps)}项")

    output_name = f"{code19}_{part_no}_QCP导入数据.xlsx"
    output_path = os.path.join(OUT_DIR, output_name)

    try:
        fill_template_surgical(output_path, steps, code19, supplier_code)
        print(f"✅ 输出：{output_path}")
        results.append({
            'file': output_name,
            'steps': len(steps),
            'llm': False,
            'supplier_code': supplier_code,
            'note': ''
        })
    except Exception as e:
        print(f"❌ 错误：{e}")
        import traceback
        traceback.print_exc()
        results.append({
            'file': output_name,
            'steps': len(steps),
            'llm': False,
            'supplier_code': supplier_code,
            'note': f'异常: {e}'
        })

print("\n\n" + "="*60)
print("汇总：")
for r in results:
    print(f"  {r['file']} | 工序:{r['steps']} | LLM:否 | 厂家编码:{r['supplier_code']} {r['note']}")
