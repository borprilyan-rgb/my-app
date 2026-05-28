import io
import re

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

from project_database import PROJECT_DATABASE


def _safe_float(v, default=0.0):
    if v is None:
        return default

    try:
        import pandas as pd
        if pd.isna(v):
            return default
    except Exception:
        pass

    if isinstance(v, str):
        v = (
            v.strip()
            .replace("Rp", "")
            .replace("rp", "")
            .replace(",", "")
            .replace("%", "")
        )

        if v == "" or v.lower() in ["none", "nan", "null", "-"]:
            return default

    try:
        return float(v)
    except Exception:
        return default


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

def get_recap_values(pdata):
    d = pdata.get("data", {})
    curr_type = pdata.get("type", "Hotel")
    pt_data = PROJECT_DATABASE.get(curr_type, {})
    
    def get_val(key, default_db_key, default_val=0.0):
        val = d.get(key)
        if val is not None and val != "":
            try: return _safe_float(val)
            except: pass
        if default_db_key and default_db_key in pt_data:
            try: return _safe_float(pt_data[default_db_key])
            except: pass
        return _safe_float(default_val)
        
    gba = get_val("m_gba", None, 0); gfa = get_val("m_gfa", None, 0)
    struc_earth = get_val("u_earth", "struc_earth", 0)
    struc_found = get_val("u_found", "struc_found", 0)
    struc_work = get_val("u_struc", "struc_work", 0)
    arch_base = get_val("u_arch", "arch_base", 0)
    facade = get_val("m_facade", None, 0)
    facade_precast_pct = get_val("r_fac_pre", "facade_precast_pct", 0)
    facade_precast_rate = get_val("u_f_pre", "facade_precast_rate", 0)
    facade_window_pct = get_val("r_fac_win", "facade_window_pct", 0)
    facade_window_rate = get_val("u_f_win", "facade_window_rate", 0)
    facade_double_pct = get_val("r_fac_doub", "facade_double_pct", 0)
    facade_double_rate = get_val("u_f_doub", "facade_double_rate", 0)
    wooden_door = get_val("m_door_w", None, 0); door_wood = get_val("u_d_wood", "door_wood", 0)
    glass_door = get_val("m_door_g", None, 0); door_glass = get_val("u_d_glass", "door_glass", 0)
    steel_door = get_val("m_door_s", None, 0); door_steel = get_val("u_d_steel", "door_steel", 0)
    lobby_interior = get_val("m_lobby", None, 0); lobby_rate = get_val("u_lobby", "lobby", 0)
    gondola_unit = get_val("m_gondola", None, 0); gondola_rate = get_val("u_gondola", "gondola", 0)
    rooms = get_val("m_rooms", None, 0)
    san_qty_room = get_val("r_san_qty", "san_room_qty", 0); san_room_rate = get_val("u_s_room", "san_room_rate", 0)
    toilet_male = get_val("m_toil_m", None, 0); san_pub_m = get_val("u_s_pub_m", "san_pub_m", 0)
    toilet_female = get_val("m_toil_f", None, 0); san_pub_f = get_val("u_s_pub_f", "san_pub_f", 0)
    disabled_toil = get_val("m_toil_d", None, 0); san_dis = get_val("u_s_dis", "san_dis", 0)
    mushola_unit = get_val("m_mushola", None, 0); san_mushola = get_val("u_s_mushola", "san_mushola", 0)
    kitchen_rate = get_val("u_kit", "kitchen", 0)
    hw_wood = get_val("u_hw_wood", "hw_wood", 0); hw_steel = get_val("u_hw_steel", "hw_steel", 0)
    fl_waste = get_val("w_floor", "fl_waste", 10); fl_skirt = get_val("s_floor", "fl_skirt", 20)
    fl_ht_pct = get_val("r_fl_ht", "fl_ht_pct", 0); fl_ht_rate = get_val("u_fl_ht", None, 0) 
    fl_vinyl_pct = get_val("r_fl_vin", "fl_vinyl_pct", 0); fl_vinyl_rate = get_val("u_fl_vin", None, 0)
    fl_marmer_pct = get_val("r_fl_mar", "fl_marmer_pct", 0); fl_marmer_rate = get_val("u_fl_mar", None, 0)
    carpet_m2 = get_val("m_carpet", None, 0); carpet_rate = get_val("u_carpet", "carpet", 0)
    glass_m2 = get_val("m_glass", None, 0); glass_rate = get_val("u_glass", "glass", 0)
    ffe_rate = get_val("u_ffe", "ffe", 0); misc_rate = get_val("u_misc", "misc", 0); misc_switch = get_val("misc_switch", None, 0)
    mep_rate = get_val("u_mep", "mep", 0); utility_rate = get_val("u_util", "utility", 0)
    railing_qty = get_val("r_rail_qty", "railing_qty", 0); railing_rate = get_val("u_rail", "railing_rate", 0)
    skylight_area = get_val("m_skylight", None, 0); skylight_rate = get_val("u_sky", "skylight_rate", 0)
    land_m2 = get_val("m_land_m2", None, 0); ext_land_rate = get_val("u_ext", "ext_land", 0)
    pub_fac_m2 = get_val("m_fac_pub", None, 0); fac_pub_rate = get_val("u_fac_p", "fac_pub", 0)
    res_fac_m2 = get_val("m_fac_res", None, 0); fac_res_rate = get_val("u_fac_r", "fac_res", 0)
    proj_fac_u = get_val("m_fac_proj", None, 0); fac_proj_rate = get_val("u_fac_pr", "fac_proj", 0)
    consultancy_rate = get_val("sc_cons", "cons", 0); qs_months = get_val("sc_qs_m", None, 0); qs_rate = get_val("sc_qs_r", None, 0)
    pm_months = get_val("sc_pm_m", None, 0); pm_rate = get_val("sc_pm_r", None, 0); insurance_pct = get_val("sc_ins", None, 0.12)
    smart_custom_costs = sum(_safe_float(i.get("Rate (Rp)", 0)) * _safe_float(i.get("Quantity", 1)) for i in d.get("smart_custom_costs", []) if isinstance(i, dict))

    t_earth = gba * struc_earth; t_found = gba * struc_found; t_struc = gba * struc_work; t_arch_base = gfa * arch_base
    t_precast = facade * (facade_precast_pct / 100) * facade_precast_rate; t_window = facade * (facade_window_pct / 100) * facade_window_rate
    t_double = facade * (facade_double_pct / 100) * facade_double_rate; t_w_door = wooden_door * door_wood
    t_g_door = glass_door * door_glass; t_s_door = steel_door * door_steel; t_lobby = lobby_interior * lobby_rate
    t_gondola = gondola_unit * gondola_rate; t_unit_san = rooms * san_qty_room * san_room_rate
    t_t_male = toilet_male * san_pub_m; t_t_female = toilet_female * san_pub_f; t_t_dis = disabled_toil * san_dis
    t_mushola = mushola_unit * san_mushola; t_kitchen = rooms * kitchen_rate; t_hw_w = wooden_door * hw_wood
    t_hw_s = steel_door * hw_steel; f_mult = (1 + (fl_waste/100)) * (1 + (fl_skirt/100))
    t_ht = gfa * (fl_ht_pct / 100) * fl_ht_rate * f_mult; t_vinyl = gfa * (fl_vinyl_pct / 100) * fl_vinyl_rate * f_mult
    t_marmer = gfa * (fl_marmer_pct / 100) * fl_marmer_rate * f_mult; t_carpet = carpet_m2 * carpet_rate
    t_glass_work = glass_m2 * glass_rate; t_ffe = rooms * ffe_rate; t_misc = misc_rate * misc_switch
    t_mep = gba * mep_rate; t_utility = gba * utility_rate; t_railing = (rooms * railing_qty) * railing_rate
    t_skylight = skylight_area * skylight_rate; t_external = land_m2 * ext_land_rate
    t_pub_fac = pub_fac_m2 * fac_pub_rate; t_res_fac = res_fac_m2 * fac_res_rate; t_proj_fac = proj_fac_u * fac_proj_rate

    construction_subtotal = sum([
        t_earth, t_found, t_struc, t_arch_base, t_precast, t_window, t_double, t_w_door, t_g_door, t_s_door, 
        t_lobby, t_gondola, t_unit_san, t_t_male, t_t_female, t_t_dis, t_mushola, t_kitchen, t_hw_w, t_hw_s, 
        t_ht, t_vinyl, t_marmer, t_carpet, t_glass_work, t_ffe, t_misc, t_mep, t_utility, t_railing, t_skylight, 
        t_external, t_pub_fac, t_res_fac, t_proj_fac, smart_custom_costs
    ])

    t_preliminary = construction_subtotal * 0.05
    t_contingency = (construction_subtotal + t_preliminary) * 0.03
    grand_total_hc = construction_subtotal + t_preliminary + t_contingency

    t_consultancy = gfa * consultancy_rate
    t_qs = qs_months * qs_rate
    t_pm = pm_months * pm_rate
    t_insurance = grand_total_hc * (insurance_pct / 100.0)

    total_soft_cost = t_consultancy + t_qs + t_pm + t_insurance
    grand_total_project = grand_total_hc + total_soft_cost

    group_arch = (t_arch_base + t_lobby + t_carpet + t_gondola + t_glass_work + t_kitchen + t_railing + t_skylight + 
                  (t_precast + t_window + t_double) + (t_unit_san + t_t_male + t_t_female + t_t_dis + t_mushola) + 
                  (t_ht + t_vinyl + t_marmer) + (t_w_door + t_g_door + t_s_door + t_hw_w + t_hw_s) + smart_custom_costs)
    
    return {
        "EARTHWORKS": t_earth, "FOUNDATIONS": t_found, "STRUCTURAL WORKS": t_struc,
        "ARCHITECTURAL WORKS": group_arch, "FF & E": t_ffe + t_misc, "M.E.P WORKS": t_mep,
        "UTILITY CONNECTION": t_utility, "EXTERNAL WORKS": t_external, "FACILITY": t_pub_fac + t_res_fac + t_proj_fac,
        "PRELIMINARIES WORKS": t_preliminary, "CONTINGENCIES": t_contingency, "HARDCOST": grand_total_hc,
        "CONSULTANCY SERVICES FEE": t_consultancy, "QS SERVICES": t_qs, 
        "PROJECT MANAGEMENT SERVICES": t_pm, "INSURANCE COVERAGE": t_insurance,
        "SOFTCOST": total_soft_cost, "TOTAL, EXCLD PPN": grand_total_project
    }

def generate_recap_excel(port_meta, projects):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Recap Cost"

    

    # --- Generate Global Totals ---
    tot_cache = {}; global_cost = {}; tot_gba = tot_gfa = tot_sgfa = 0
    for pid, pdata in projects.items():
        vals = get_recap_values(pdata)
        tot_cache[pid] = vals
        for k, v in vals.items(): global_cost[k] = global_cost.get(k, 0) + v
        d = pdata.get("data", {})
        tot_gba += _safe_float(d.get("m_gba", 0)); tot_gfa += _safe_float(d.get("m_gfa", 0)); tot_sgfa += _safe_float(d.get("m_sgfa", 0))

    # Calculate global % divisors
    global_hc = global_cost.get("HARDCOST", 0)
    global_sc = global_cost.get("SOFTCOST", 0)
    safe_hc = global_hc if global_hc > 0 else 1
    safe_sc = global_sc if global_sc > 0 else 1

    # --- 1. Styling Definitions ---
    blue_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    white_font, bold_font, reg_font = Font(color="FFFFFF", bold=True, size=10), Font(bold=True, size=9), Font(size=9)
    black_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align, left_align = Alignment(horizontal='center', vertical='center', wrap_text=True), Alignment(horizontal='left', vertical='center')

    # --- 2. Blue Metadata Header ---
    headers = [["ASG GROUP PROPERTY DEVELOPMENT", f"VERSION : {port_meta.get('version', '')}"], ["QS & PROCUREMENT DIVISION", f"UPDATED : {port_meta.get('updated', '')}"], [port_meta.get('title', ''), f"CREATED : {port_meta.get('created', '')}"], [port_meta.get('ref', ''), ""]]
    for r_idx, (text_left, text_right) in enumerate(headers, 1):
        for c in range(1, 100): ws.cell(row=r_idx, column=c).fill = blue_fill
        ws.cell(row=r_idx, column=1, value=text_left).font = white_font
        ws.cell(row=r_idx, column=15, value=text_right).font = white_font
        ws.cell(row=r_idx, column=15).alignment = Alignment(horizontal='right', vertical='center')

    # --- 3. Static Table Headers ---
    static_cols = [("SN", 1), ("DESCRIPTION", 2), ("COA", 3), ("%", 4)]
    for name, col_idx in static_cols:
        ws.merge_cells(start_row=6, start_column=col_idx, end_row=9, end_column=col_idx)
        c = ws.cell(row=6, column=col_idx, value=name)
        c.alignment, c.font, c.fill = center_align, bold_font, PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        for r in range(6, 10): ws.cell(row=r, column=col_idx).border = black_border

    # --- 4. Dynamic Project Headers ---
    bg_colors = ["EAEAEA", "FCE4D6", "F2DCDB", "E1D5E7", "DDEBF7", "E2EFDA", "D9E1F2", "F4B084", "FFF2CC"]
    project_list = [("TOTAL", {"name": "TOTAL"})] + list(projects.items())
    current_col = 5
    
    for i, (pid, pdata) in enumerate(project_list):
        group_fill = PatternFill(start_color=bg_colors[i % len(bg_colors)], end_color=bg_colors[i % len(bg_colors)], fill_type="solid")
        ws.merge_cells(start_row=6, start_column=current_col, end_row=6, end_column=current_col+4)
        c = ws.cell(row=6, column=current_col, value=pdata.get('name', 'PROJECT').upper())
        c.alignment, c.font, c.fill = center_align, bold_font, group_fill
        
        ws.cell(row=7, column=current_col, value="ESTIMATE").alignment = center_align
        ws.merge_cells(start_row=7, start_column=current_col+1, end_row=7, end_column=current_col+4)
        ws.cell(row=7, column=current_col+1, value="Cost Ratio (Rp/m2)").alignment = center_align
        
        for j, lbl in enumerate(["TOTAL", "GBA", "GFA", "SGFA", "NFA"]):
            ws.cell(row=8, column=current_col+j, value=lbl).alignment = center_align
            
        if pid == "TOTAL":
            gba, gfa, sgfa, nfa = tot_gba, tot_gfa, tot_sgfa, tot_gfa * 0.82
        else:
            d = pdata.get("data", {})
            gba, gfa, sgfa = _safe_float(d.get("m_gba", 0)), _safe_float(d.get("m_gfa", 0)), _safe_float(d.get("m_sgfa", 0))
            nfa = gfa * 0.82
        
        gba_f = gba if gba > 0 else 1; gfa_f = gfa if gfa > 0 else 1; sgfa_f = sgfa if sgfa > 0 else 1; nfa_f = nfa if nfa > 0 else 1
            
        for j, val in enumerate(["Rp", gba, gfa, sgfa, nfa]):
            c = ws.cell(row=9, column=current_col+j, value=val)
            c.alignment = center_align
            if j > 0: c.number_format = '#,##0'

        for r in range(6, 10):
            for c_idx in range(current_col, current_col+5):
                ws.cell(row=r, column=c_idx).fill = group_fill
                ws.cell(row=r, column=c_idx).border = black_border
        current_col += 5

    # --- 5. Data Rows Map (With Categories for %) ---
    row_mapping = [
        ("I", "HARDCOST", "118-14-000", True, "HC"), ("1", "PRELIMINARIES WORKS", "118-14-100", False, "HC"), 
        ("2", "EARTHWORKS", "118-14-200", False, "HC"), ("3", "FOUNDATIONS", "118-14-300", False, "HC"), 
        ("4", "STRUCTURAL WORKS", "118-14-500", False, "HC"), ("5", "ARCHITECTURAL WORKS", "118-14-600", False, "HC"),
        ("6", "FF & E", "118-14-700", False, "HC"), ("7", "M.E.P WORKS", "118-14-800", False, "HC"), 
        ("8", "UTILITY CONNECTION", "118-13-900", False, "HC"), ("9", "EXTERNAL WORKS", "118-14-930", False, "HC"), 
        ("10", "FACILITY", "118-14-960", False, "HC"), ("11", "CONTINGENCIES", "", False, "HC"),
        ("II", "SOFTCOST", "118-13-000", True, "SC_TOTAL"), ("1", "CONSULTANCY SERVICES FEE", "118-13-202", False, "SC"), 
        ("2", "QS SERVICES", "118-13-201", False, "SC"), ("3", "PROJECT MANAGEMENT SERVICES", "118-13-203", False, "SC"), 
        ("4", "INSURANCE COVERAGE", "118-13-300", False, "SC"), ("IV", "TOTAL, EXCLD PPN", "", True, "TOTAL")
    ]

    r_idx = 10
    for sn, desc, coa, is_bold, cat in row_mapping:
        # Calculate % based on Global Totals and Category
        global_val = global_cost.get(desc, 0)
        if cat == "HC": pct = global_val / safe_hc
        elif cat == "SC": pct = global_val / safe_sc
        elif cat == "SC_TOTAL": pct = global_val / safe_hc # Softcost total vs Hardcost
        elif cat == "TOTAL": pct = global_val / safe_hc # Grand total vs Hardcost
        else: pct = 0

        ws.cell(row=r_idx, column=1, value=sn).alignment = center_align
        ws.cell(row=r_idx, column=2, value=desc).alignment = left_align
        ws.cell(row=r_idx, column=3, value=coa).alignment = center_align
        
        # Write Percentage
        pct_cell = ws.cell(row=r_idx, column=4, value=pct)
        pct_cell.alignment = center_align
        pct_cell.number_format = '0.00%'

        for c in range(1, 5):
            ws.cell(row=r_idx, column=c).border = black_border
            ws.cell(row=r_idx, column=c).font = bold_font if is_bold else reg_font

        col_offset = 5
        for pid, pdata in project_list:
            val = global_cost.get(desc, 0) if pid == "TOTAL" else tot_cache[pid].get(desc, 0)
            
            if pid == "TOTAL":
                gba_f = tot_gba if tot_gba > 0 else 1; gfa_f = tot_gfa if tot_gfa > 0 else 1
                sgfa_f = tot_sgfa if tot_sgfa > 0 else 1; nfa_f = (tot_gfa*0.82) if tot_gfa > 0 else 1
            else:
                d = pdata.get("data", {})
                gba_f = _safe_float(d.get("m_gba", 1) if d.get("m_gba", 0) > 0 else 1)
                gfa_f = _safe_float(d.get("m_gfa", 1) if d.get("m_gfa", 0) > 0 else 1)
                sgfa_f = _safe_float(d.get("m_sgfa", 1) if d.get("m_sgfa", 0) > 0 else 1)
                nfa_f = gfa_f * 0.82
            
            for j, v in enumerate([val, val/gba_f, val/gfa_f, val/sgfa_f, val/nfa_f]):
                c = ws.cell(row=r_idx, column=col_offset+j, value=v)
                c.number_format = '#,##0'
                c.border = black_border
                c.font = bold_font if is_bold else reg_font
                if is_bold: c.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            col_offset += 5
        r_idx += 1

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 10
    for c in range(5, col_offset): ws.column_dimensions[get_column_letter(c)].width = 16
    ws.freeze_panes = 'E10'
    wb.save(output)
    return output.getvalue()

