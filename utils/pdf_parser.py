"""
PDF type detection and QCP data extraction.
"""

import pdfplumber
from typing import Literal


def is_scanned(pdf_path: str) -> bool:
    """Quick check: if first 3 pages have little text, it's scanned."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_chars = 0
            pages_to_check = min(3, len(pdf.pages))
            for page in pdf.pages[:pages_to_check]:
                text = page.extract_text() or ""
                total_chars += len(text.strip())
            return total_chars < 200
    except Exception:
        return True


def detect_pdf_type(pdf_path: str) -> Literal["digital", "scanned"]:
    """Detect if PDF is digital (Word-converted) or scanned."""
    if is_scanned(pdf_path):
        return "scanned"
    return "digital"


def _normalize_point(value: str | None) -> str:
    """Normalize H/W/R to 点 format."""
    if value is None:
        return "-"
    v = value.strip().upper()
    if v == "H":
        return "H点"
    elif v == "W":
        return "W点"
    elif v == "R":
        return "R点"
    return "-"


def _extract_table_row_data(row_cells: list) -> dict | None:
    """Extract process data from a table row."""
    if not row_cells or len(row_cells) < 2:
        return None
    
    # Join all cells to search for process info
    row_text = " | ".join([str(c).strip() if c else "" for c in row_cells])
    
    # Look for process number pattern like A19.0, B1.2, C3.0 etc.
    import re
    # Pattern: letter + number(s) + optional . + digit
    process_pattern = re.compile(r'([A-Z]\d+(?:\.\d+)?)\s*[^\|]*\|\s*([^\|]+)')
    match = process_pattern.search(row_text)
    
    if not match:
        return None
    
    process_no = match.group(1).strip()
    process_name = match.group(2).strip()
    
    # Find A and S columns - typically at positions that contain H/W/R
    point_a = None
    point_s = None
    
    # Search within the row text for H/W/R markers
    # Common column positions in QCP tables
    for cell in row_cells:
        cell_str = str(cell).strip() if cell else ""
        if cell_str.upper() in ["H", "W", "R"]:
            # This could be in A column or S column
            pass
    
    # More robust: check for H/W/R in specific column positions
    # QCP tables usually have: col A (point A), various columns, col S (point S)
    # We'll identify by column index relative to the row
    
    # Simple heuristic: find H/W/R values in the row
    hwr_positions = []
    for i, cell in enumerate(row_cells):
        if cell and str(cell).strip().upper() in ["H", "W", "R"]:
            hwr_positions.append((i, str(cell).strip().upper()))
    
    if len(hwr_positions) >= 2:
        # Assume first is point A, last is point S
        point_a = hwr_positions[0][1]
        point_s = hwr_positions[-1][1]
    elif len(hwr_positions) == 1:
        # Only one HWR marker - decide based on position
        if hwr_positions[0][0] < len(row_cells) / 2:
            point_a = hwr_positions[0][1]
        else:
            point_s = hwr_positions[0][1]
    
    # Filter out non-process rows (headers, footers, etc.)
    if len(process_no) <= 5 and process_no[0].isupper():
        return {
            "process_no": process_no,
            "process_name": process_name,
            "point_a": _normalize_point(point_a),
            "point_s": _normalize_point(point_s)
        }
    
    return None


def extract_qcp_data(pdf_path: str) -> list[dict]:
    """
    Extract QCP process steps from PDF.
    Returns list of dicts with keys:
    - process_no: str (e.g. 'A19.0')
    - process_name: str (e.g. '原材料检验')
    - point_a: str (H点/W点/R点/-)
    - point_s: str (H点/W点/R点/-)
    """
    qcp_data = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                
                for table in tables:
                    for row in table:
                        result = _extract_table_row_data(row)
                        if result:
                            # Avoid duplicates
                            if not any(d["process_no"] == result["process_no"] for d in qcp_data):
                                qcp_data.append(result)
    
    except Exception as e:
        print(f"PDF extraction error: {e}")
    
    return qcp_data