"""
bom_reader.py
-------------
Reads an Excel (.xlsx) file:
  1. Scans ROW 1 left-to-right to locate Description, Part Type, Qty, Unit columns
     using the alias list in component_dictionary.json.
  2. Scans all data rows (row 2 onward) and returns structured dicts.
"""

import json
import os
import openpyxl

# --- HOTFIX FOR ERP EXPORTS ---
# Many ERP systems (like HPCC DB) generate malformed styles.xml that crash openpyxl.
# Since we only care about data, we bypass the entire stylesheet parsing.
import openpyxl.reader.excel
openpyxl.reader.excel.apply_stylesheet = lambda archive, wb: None
# ------------------------------

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(_BASE_DIR, "component_dictionary.json"), encoding="utf-8") as f:
    _DICT = json.load(f)

_ALIASES = _DICT["column_header_aliases"]

# Build reverse lookup: normalized header text → role name
_HEADER_ROLE: dict[str, str] = {}
for role, aliases in _ALIASES.items():
    if role.startswith("_"):
        continue
    for alias in aliases:
        _HEADER_ROLE[alias.strip().lower()] = role


def detect_columns(headers: list) -> dict[str, int]:
    """
    Scan header list (row 1) left-to-right.
    Returns dict: role → column index (0-based).
    E.g. {"description_col": 3, "part_type_col": 2, ...}
    """
    found: dict[str, int] = {}
    for col_idx, cell_value in enumerate(headers):
        if cell_value is None:
            continue
        normalized = str(cell_value).strip().lower()
        role = _HEADER_ROLE.get(normalized)
        if role and role not in found:          # first match wins
            found[role] = col_idx
    return found


def read_bom_file(filepath: str) -> tuple[list[dict], dict, str | None]:
    """
    Open xlsx file. For each sheet, try to detect BOM columns in row 1.
    Returns:
      (rows, col_map, sheet_name)
        rows     : list of dicts per data row
        col_map  : {role: col_index}
        sheet_name: name of the sheet used
    Returns ([], {}, None) if no suitable sheet found.

    Each row dict has keys:
      raw_row      : tuple of all cell values
      description  : str or None
      part_type    : str or None
      quantity     : float or None
      unit         : str or None
      partnumber   : str or None
      row_number   : int (1-based excel row)
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)

    best_sheet = None
    best_col_map: dict[str, int] = {}
    best_score = 0
    best_data_start_row = 2

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row is not None and ws.max_row < 2:
            continue
        
        # Scan the first 30 rows to dynamically find the header row
        for row_idx, row_cells in enumerate(ws.iter_rows(min_row=1, max_row=30), start=1):
            row_values = [cell.value for cell in row_cells]
            col_map = detect_columns(row_values)
            score = len(col_map)
            
            # Bonus: must have description col to be useful
            if "description_col" in col_map:
                score += 10
                
            if score > best_score:
                best_score = score
                best_sheet = sheet_name
                best_col_map = col_map
                best_data_start_row = row_idx + 1

    if not best_sheet or "description_col" not in best_col_map:
        return [], {}, None

    ws = wb[best_sheet]
    rows_out: list[dict] = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=best_data_start_row, values_only=True), start=best_data_start_row):
        def _get(role):
            idx = best_col_map.get(role)
            if idx is None or idx >= len(row):
                return None
            val = row[idx]
            return val

        desc = _get("description_col")
        if desc is None:
            continue                   # skip blank description rows
        desc_str = str(desc).strip()
        if not desc_str:
            continue

        pt_raw = _get("part_type_col")
        qty_raw = _get("quantity_col")
        unit_raw = _get("unit_col")
        pn_raw = _get("partnumber_col")

        # Parse qty safely
        qty = None
        if qty_raw is not None:
            try:
                qty = float(qty_raw)
            except (ValueError, TypeError):
                qty = None

        rows_out.append({
            "raw_row":     row,
            "description": desc_str,
            "part_type":   str(pt_raw).strip() if pt_raw else None,
            "quantity":    qty,
            "unit":        str(unit_raw).strip() if unit_raw else None,
            "partnumber":  str(pn_raw).strip() if pn_raw else None,
            "row_number":  row_idx,
        })

    return rows_out, best_col_map, best_sheet


def get_available_sheets(filepath: str) -> list[str]:
    """Return list of sheet names in the workbook."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names
