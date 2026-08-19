"""
QCP → CNPE 通用转换工具
适用多种QCP格式：KSB SEC-KSB / 三门 / 漳州 等
使用方法：双击 run.bat 或 python qcp_converter.py
"""

import os
import sys
import re
import csv
import tempfile
import shutil
from datetime import datetime

# 尝试导入需要的库
try:
    import pdfplumber
except ImportError:
    print("缺少 pdfplumber，正在安装...")
    os.system(f"{sys.executable} -m pip install pdfplumber -q")
    import pdfplumber

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("缺少 openpyxl，正在安装...")
    os.system(f"{sys.executable} -m pip install openpyxl -q")
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment


def extract_qcp_steps(pdf_path):
    """从任意QCP PDF中提取工序数据（表格方式）"""
    steps = []
    seen = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 2:
                        continue

                    cells = [str(c).strip() if c else '' for c in row]

                    # 在所有单元格中找工序号格式 (字母开头+数字)
                    step_num = None
                    step_name = None

                    for idx, cell in enumerate(cells):
                        cell_clean = cell.strip()
                        # 匹配工序号: A1, B2.3, C1.1.2, A10, B1.0 等
                        m = re.match(r'^([A-Z]\d+(?:\.\d+)*)$', cell_clean)
                        if m:
                            step_num = m.group(1)
                            # 尝试找配套的工序名称
                            # 通常工序名称在相邻的单元格
                            for offset in [1, 2, 3, -1, -2]:
                                ni = idx + offset
                                if 0 <= ni < len(cells) and cells[ni]:
                                    name_candidate = cells[ni].strip()
                                    if len(name_candidate) >= 2:
                                        # 排除表头
                                        if name_candidate not in ('序号', 'No.', '工序名称', 'Process Name', '报告编号', 'Remark', '备注', '质量控制点', 'Quality Control Point', '版', 'Rev.', 'S', 'C', 'O', 'W', 'H', 'R'):
                                            step_name = name_candidate
                                            break
                            break

                    if step_num and step_name:
                        key = f"{step_num}|{step_name}"
                        if key not in seen and len(step_name) >= 2:
                            seen.add(key)
                            steps.append({
                                'step': step_num,
                                'name': step_name.split('\n')[0].strip(),
                                'content': ''
                            })

    # 排序
    def sort_key(s):
        m = re.match(r'^([A-Z])(\d+(?:\.\d+)*)', s['step'])
        if m:
            letter = m.group(1)
            nums = m.group(2).split('.')
            return (letter, [int(x) for x in nums])
        return (s['step'], [])

    steps.sort(key=sort_key)
    return steps


def generate_excel(steps, item_code_19, part_no, supplier_item_code, output_path):
    """生成CNPE格式Excel"""
    wb = Workbook()

    # Sheet 1
    ws = wb.active
    ws.title = 'QCP工序'

    # 标题
    ws['A1'] = 'QCP转CNPE质量计划'
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:F1')

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

    # 表头
    headers = ['序号', '工序号', '工序名称', '质量控制点', '报告编号', '备注']
    hdr_fill = PatternFill(start_color='667EEA', end_color='667EEA', fill_type='solid')
    hdr_font = Font(bold=True, color='FFFFFF', size=11)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=9, column=col, value=h)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal='center', vertical='center')

    # 数据行
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

    # 列宽
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 35

    # Sheet 2
    ws2 = wb.create_sheet('填写说明')
    notes_data = [
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
        ['识别的工序号格式', 'A1, B2.3, C1.1.2 等字母+数字格式', ''],
    ]
    for ri, row_data in enumerate(notes_data, 1):
        for ci, val in enumerate(row_data, 1):
            c = ws2.cell(row=ri, column=ci, value=val)
            if ri == 3:
                c.font = Font(bold=True)
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 40
    ws2.column_dimensions['C'].width = 30

    wb.save(output_path)
    print(f"Excel已生成: {output_path}")


def main():
    print("=" * 50)
    print("  QCP → CNPE 通用转换工具")
    print("  适用多种QCP格式 | 双击运行 | 无需联网")
    print("=" * 50)
    print()

    # 获取用户输入
    item_code_19 = input("请输入19位编码: ").strip()
    while not item_code_19 or len(item_code_19) < 5:
        if not item_code_19:
            print("19位编码不能为空")
        else:
            print(f"编码长度不够，当前: {len(item_code_19)} 位")
        item_code_19 = input("请输入19位编码: ").strip()

    supplier_item_code = input("请输入厂家物项编号: ").strip()
    while not supplier_item_code:
        print("厂家物项编号不能为空")
        supplier_item_code = input("请输入厂家物项编号: ").strip()

    part_no = input("请输入零件号（直接回车跳过）: ").strip()

    print()
    print("请将QCP PDF文件拖拽到窗口内，或输入PDF文件完整路径:")
    print()
    pdf_path = input("PDF路径: ").strip().strip('"')

    while not pdf_path:
        print("路径不能为空")
        pdf_path = input("PDF路径: ").strip().strip('"')

    if not os.path.exists(pdf_path):
        print(f"文件不存在: {pdf_path}")
        input("按回车退出...")
        sys.exit(1)

    print(f"\n正在解析: {pdf_path}")
    print("请稍候...")

    try:
        steps = extract_qcp_steps(pdf_path)
        if not steps:
            print("\n⚠ 未能在PDF中识别到工序数据")
            print("可能原因:")
            print("  1. PDF是扫描件（图片），不是文本型")
            print("  2. PDF中的表格格式不标准")
            print("  3. 工序号不在表格中")
            print()
            input("按回车退出...")
            sys.exit(1)

        print(f"\n✓ 成功识别 {len(steps)} 道工序")

        # 生成输出文件名
        safe_name = re.sub(r'[^\w\-_.]', '_', item_code_19)[:50]
        output_dir = os.path.dirname(pdf_path) or os.path.expanduser('~/Downloads')
        output_path = os.path.join(output_dir, f"CNPE_转换_{safe_name}.xlsx")

        # 如果文件已存在，追加编号
        counter = 1
        base_output = output_path
        while os.path.exists(output_path):
            name, ext = os.path.splitext(base_output)
            output_path = f"{name}_{counter}{ext}"
            counter += 1

        generate_excel(steps, item_code_19, part_no, supplier_item_code, output_path)

        print(f"\n✅ 转换完成！")
        print(f"   工序数量: {len(steps)} 道")
        print(f"   输出文件: {output_path}")
        print()

    except Exception as e:
        print(f"\n❌ 处理失败: {e}")
        import traceback
        traceback.print_exc()

    input("\n按回车退出...")


if __name__ == '__main__':
    main()
