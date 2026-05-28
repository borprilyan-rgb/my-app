import io
import re

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font


def generate_exact_portfolio_excel(port_meta, port_data, port_assumptions):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio Summary"

    # --- 1. Styling Definitions ---
    blue_fill = PatternFill(start_color="005A9C", end_color="005A9C", fill_type="solid")
    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    
    white_font = Font(color="FFFFFF", bold=True, name='Calibri', size=11)
    bold_font = Font(bold=True, name='Calibri', size=10)
    reg_font = Font(bold=False, name='Calibri', size=10)
    small_font = Font(name='Calibri', size=9)
    
    black_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                          top=Side(style='thin'), bottom=Side(style='thin'))
    
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center')

    # --- 2. Blue Header Section ---
    headers = [
        ["ASG GROUP PROPERTY DEVELOPMENT", f"VERSION : {port_meta.get('version', '')}"],
        ["QS & PROCUREMENT DIVISION", f"UPDATED : {port_meta.get('updated', '')}"],
        [port_meta.get('title', ''), f"CREATED : {port_meta.get('created', '')}"],
        [port_meta.get('ref', ''), ""]
    ]

    for r_idx, (text_left, text_right) in enumerate(headers, 1):
        for c_idx in range(1, 12):
            c = ws.cell(row=r_idx, column=c_idx)
            c.fill = blue_fill
            c.font = white_font
            if c_idx == 1: c.value = text_left
            if c_idx == 11: 
                c.value = text_right
                c.alignment = Alignment(horizontal='right')
        ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=10)

    # --- 3. Table Header (Rows 6 & 7) ---
    ws.merge_cells(start_row=6, start_column=1, end_row=7, end_column=1) # SN
    ws.merge_cells(start_row=6, start_column=2, end_row=7, end_column=2) # AREA
    ws.merge_cells(start_row=6, start_column=3, end_row=6, end_column=5) # BLDG AREA
    ws.merge_cells(start_row=6, start_column=6, end_row=6, end_column=7) # UNIT
    ws.merge_cells(start_row=6, start_column=8, end_row=7, end_column=8) # BUDGET
    ws.merge_cells(start_row=6, start_column=9, end_row=6, end_column=11) # COST RATIO

    header_labels = {
        (6, 1): "SN", (6, 2): "AREA", (6, 3): "BUILDING AREA (M2)", 
        (6, 6): "UNIT", (6, 8): "BUDGET ESTIMATE\nRP", (6, 9): "COST RATIO RP/M2",
        (7, 3): "GBA", (7, 4): "GFA", (7, 5): "SGFA", (7, 9): "GBA", (7, 10): "GFA", (7, 11): "SGFA"
    }

    for (r, c), text in header_labels.items():
        cell = ws.cell(row=r, column=c, value=text)
        cell.alignment = center_align
        cell.font = bold_font
        cell.fill = gray_fill

    for r in range(6, 8):
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = black_border

    # --- 4. Data Rows ---
    current_row = 8
    for p_row in port_data:
        is_total = p_row.get("AREA") == "TOTAL"
        is_parking = "PARKING" in str(p_row.get("AREA", "")).upper()
        
        # Set SN and Area
        ws.cell(row=current_row, column=1, value=p_row.get("SN", "")).alignment = center_align
        
        area_cell = ws.cell(row=current_row, column=2, value=p_row.get("AREA", ""))
        area_cell.alignment = Alignment(horizontal='left', vertical='top' if is_parking else 'center', wrap_text=True)
        
        # Set Values
        cols = ["GBA", "GFA", "SGFA", "QTY", "UNIT", "BUDGET", "R_GBA", "R_GFA", "R_SGFA"]
        for i, key in enumerate(cols, 3):
            val = p_row.get(key, "")
            cell = ws.cell(row=current_row, column=i, value=val)
            cell.alignment = Alignment(vertical='top' if is_parking else 'center', horizontal='right' if i != 7 else 'center')
            
            if key in ["GBA", "GFA", "SGFA"]: cell.number_format = "#,##0.00"
            if key in ["BUDGET", "R_GBA", "R_GFA", "R_SGFA"]: cell.number_format = "#,##0"

        # Apply Row Styles
        for c in range(1, 12):
            ws.cell(row=current_row, column=c).border = black_border
            ws.cell(row=current_row, column=c).font = bold_font if (p_row.get("SN") or is_total) else reg_font
            if is_total: ws.cell(row=current_row, column=c).fill = gray_fill

        if is_total:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            ws.cell(row=current_row, column=6, value="TOTAL").alignment = center_align

        # Adjust height for multi-line rows (Parking)
        if is_parking:
            ws.row_dimensions[current_row].height = 100 

        current_row += 1

    # --- 5. Assumptions Section ---
    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
    cell_assump = ws.cell(row=current_row, column=1, value="I.  ASSUMPTIONS")
    cell_assump.font = bold_font
    for c in range(1, 12):
        ws.cell(row=current_row, column=c).fill = yellow_fill
        ws.cell(row=current_row, column=c).border = black_border

    current_row += 1
    for _, a_row in port_assumptions.iterrows():
        ws.cell(row=current_row, column=1, value=a_row.get("No.", "")).alignment = center_align
        
        desc_cell = ws.cell(row=current_row, column=2, value=a_row.get("Assumption Description", ""))
        desc_cell.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=11)
        
        for c in range(1, 12):
            ws.cell(row=current_row, column=c).border = black_border
        current_row += 1

    # Column Widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 35
    for c in ['C', 'D', 'E', 'H', 'I', 'J', 'K']:
        ws.column_dimensions[c].width = 16

    wb.save(output)
    return output.getvalue()


def _normalize_header_token(value, trailing_dot=False):
    import re

    text = str(value or "").strip().upper()
    text = re.sub(r"(?<=[A-Z])\s+(?=\d)", "", text)
    text = re.sub(r"\s+", ".", text)

    if trailing_dot and text and not text.endswith("."):
        text = f"{text}."

    return text


def build_portfolio_meta_from_inputs(header_inputs):
    project_location = _normalize_header_token(
        header_inputs.get("project_location", ""),
        trailing_dot=True,
    )
    project_name = _normalize_header_token(header_inputs.get("project_name", ""))
    option_number = str(header_inputs.get("option_number", "")).strip()
    revision_number = str(header_inputs.get("revision_number", "")).strip()
    drawing_date = str(header_inputs.get("drawing_date", "")).strip()
    updated_date = str(header_inputs.get("updated_date", "")).strip()
    created_date = str(header_inputs.get("created_date", "")).strip()

    title = (
        f"PROJECT PORTFOLIO | "
        f"{project_location}{project_name} "
        f"OPT.{option_number} "
        f"R({revision_number})"
    )

    ref = (
        f"REF. DATA R({revision_number}) | "
        f"CONCEPT DWG {drawing_date}.DPA"
    )

    version = f"R ({revision_number}) OPT{option_number}"

    return {
        "title": title,
        "ref": ref,
        "version": version,
        "updated": updated_date,
        "created": created_date,
    }
