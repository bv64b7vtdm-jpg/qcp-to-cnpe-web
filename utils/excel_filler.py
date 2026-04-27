"""
Excel filler using zipfile XML manipulation to preserve hidden columns and formatting.
"""

import zipfile
import shutil
import os
import re
from typing import List

# Column mapping (1-indexed for openpyxl, but XML uses 0-indexed)
# D=4, F=6, H=8, I=9, K=11, P=16, Q=17, R=18
COLUMN_LETTERS = {
    "D": 4,   # Sort number
    "F": 6,   # Supplier item code
    "H": 8,   # Process number
    "I": 9,   # Process name
    "K": 11,  # Point S
    "P": 16,  # Report
    "Q": 17,  # Has point?
    "R": 18,  # Point A1
}

# Special processes that always get Q=N (no witness point)
SPECIAL_PROCESSES = ["先决条件检查", "质量计划关闭", "NCR", "不符合项", "关闭"]


def _col_letter_to_index(col: str) -> int:
    """Convert column letter to 1-based index."""
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def _row_has_point(point_a: str, point_s: str) -> bool:
    """Check if a row has any witness points."""
    return point_a not in ["-", "", None] or point_s not in ["-", "", None]


def _determine_q_value(process_name: str, point_a: str, point_s: str) -> str:
    """Determine Q column value (Y/N)."""
    # Special processes never have witness points
    for sp in SPECIAL_PROCESSES:
        if sp in process_name:
            return "N"
    
    # Check if there are any actual witness points
    if _row_has_point(point_a, point_s):
        return "Y"
    
    return "N"


def _make_cell_ref(col_letter: str, row: int) -> str:
    """Make a cell reference like D10."""
    return f"{col_letter}{row}"


def _get_or_create_cell(xml_content: str, cell_ref: str, row_num: int) -> str:
    """Get existing cell XML or create new cell XML."""
    # Try to find existing cell
    cell_pattern = re.compile(
        rf'<c r="{re.escape(cell_ref)}"[^>]*>.*?</c>',
        re.DOTALL
    )
    match = cell_pattern.search(xml_content)
    
    if match:
        return match.group(0)
    
    return f'<c r="{cell_ref}" t="inlineStr"><is><t></t></is></c>'


def _set_cell_value(xml_content: str, cell_ref: str, value: str, is_number: bool = False) -> str:
    """Set value in a cell, preserving styles."""
    # Find existing cell
    cell_pattern = re.compile(
        rf'(<c r="{re.escape(cell_ref)}"[^>]*>)(.*?)(</c>)',
        re.DOTALL
    )
    
    if is_number:
        replacement = rf'\1<v>{value}</v>\3'
    else:
        # Handle inline string
        cell_match = cell_pattern.search(xml_content)
        if cell_match:
            full_cell = cell_match.group(0)
            # Check if it's an inline string type
            if 't="inlineStr"' in full_cell or 't="s"' in full_cell:
                # Replace the value portion
                new_cell = re.sub(r'<is><t>.*?</t></is>', f'<is><t>{value}</t></is>', full_cell)
                new_cell = re.sub(r'<v>.*?</v>', f'<v>{value}</v>', new_cell)
                return xml_content.replace(full_cell, new_cell)
            else:
                # Add v element
                new_cell = full_cell.replace('</c>', f'<v>{value}</v></c>')
                return xml_content.replace(full_cell, new_cell)
        replacement = rf'\1<v>{value}</v>\3'
    
    return cell_pattern.sub(replacement, xml_content)


def fill_cnpe_template(
    template_path: str,
    output_path: str,
    qcp_data: list[dict],
    item_code_19: str,
    part_no: str,
    supplier_item_code: str
) -> str:
    """
    Fill the CNPE Excel template using zipfile XML manipulation.
    Preserves all original formatting, hidden columns, and styles.
    
    Columns to fill:
    - D: Sort number (10, 20, 30...)
    - F: Supplier item code (厂家物项编码)
    - H: Process number (工序编号)
    - I: Process name (工序名称)
    - K: Point S (选点S: H点/W点/R点/-)
    - P: Report (是否产生报告: 否)
    - Q: Has point? (Y/N)
    - R: Point A1 (选点A1: H点/W点/R点/-)
    
    Returns: output file path.
    """
    # Copy template to output
    shutil.copy2(template_path, output_path)
    
    # Find the actual sheet file - check workbook for first sheet
    with zipfile.ZipFile(output_path, 'r') as zf:
        rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        
        # Find first sheet relationship
        sheet_rel_match = re.search(r'<Relationship Id="(rId\d+)" Type="[^"]*worksheet[^"]*" Target="([^"]+)"', rels_xml)
        if not sheet_rel_match:
            # Try alternate format
            sheet_rel_match = re.search(r'Id="(rId\d+)"[^>]*Target="([^"]+)"[^>]*worksheet', rels_xml)
        
        if sheet_rel_match:
            sheet_file = "xl/" + sheet_rel_match.group(2)
        else:
            # Default to sheet1.xml
            sheet_file = "xl/worksheets/sheet1.xml"
        
        sheet_xml = zf.read(sheet_file).decode('utf-8')
        all_files = {name: zf.read(name) for name in zf.namelist()}
    
    # Determine starting row - look for existing data rows
    # Typically data starts around row 5 or 6 after headers
    start_row = 5
    
    # Fill in the data rows
    current_row = start_row
    
    for idx, item in enumerate(qcp_data):
        row_num = current_row + idx
        sort_num = (idx + 1) * 10
        
        # D column - Sort number
        cell_d = _make_cell_ref("D", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_d, str(sort_num), is_number=True)
        
        # F column - Supplier item code
        cell_f = _make_cell_ref("F", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_f, supplier_item_code)
        
        # H column - Process number
        cell_h = _make_cell_ref("H", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_h, item.get("process_no", ""))
        
        # I column - Process name
        cell_i = _make_cell_ref("I", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_i, item.get("process_name", ""))
        
        # K column - Point S
        cell_k = _make_cell_ref("K", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_k, item.get("point_s", "-"))
        
        # P column - Report (always "否")
        cell_p = _make_cell_ref("P", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_p, "否")
        
        # Q column - Has point? (Y/N)
        point_a = item.get("point_a", "-")
        point_s = item.get("point_s", "-")
        process_name = item.get("process_name", "")
        q_value = _determine_q_value(process_name, point_a, point_s)
        cell_q = _make_cell_ref("Q", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_q, q_value)
        
        # R column - Point A1
        cell_r = _make_cell_ref("R", row_num)
        sheet_xml = _set_cell_value(sheet_xml, cell_r, point_a)
    
    # Write back to zip
    all_files[sheet_file] = sheet_xml.encode('utf-8')
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in all_files.items():
            zf.writestr(name, data)
    
    return output_path