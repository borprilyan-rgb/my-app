#region --- LIBRARY AND SUCH ---
import streamlit as st
import pandas as pd
from area_helpers import (
    calculate_area_totals_from_table,
    generate_area_rows,
    guess_area_f2f_height,
    normalize_none_records,
    clean_area_records,
    calculate_area_dataframe,
    clean_door_records,
    area_records_to_input_view,
    input_view_to_area_records,
    validate_consultant_summary_input,
)
from area_helpers import safe_float as _safe_float

from excel_helpers import (
    ExcelImportError,
    create_area_excel_form_bytes,
    read_area_input_sheet,
    read_architectural_facade_inputs,
    read_architectural_sheet,
    read_consultancy_sheet,
    read_earthworks_sheet,
    read_external_sheet,
    read_ffe_sheet,
    read_foundation_sheet,
    read_mep_sheet,
    read_pintu_sheet,
    read_residential_area_sheet,
    read_structural_sheet,
    read_utility_sheet,
)
from report_excel_helpers import (
    build_portfolio_meta_from_inputs,
    generate_exact_portfolio_excel,
    generate_recap_excel,
    get_recap_values,
    normalize_header_token,
)

import altair as alt
import plotly.graph_objects as go
import num2words as n2w
import ast
import numpy as np
from io import BytesIO
import json as _json

APP_VERSION = "1.1.0" #app version for future compatibility check
AREA_UNIT = "m2"
MULTIPLY_SIGN = "*"
st.set_page_config(page_title="Project Feasibility Study - Agung Sedayu Group",
                    layout="wide", page_icon="Agung-Sedayu.png",)

st.logo("Agung-Sedayu-Group.png")
#endregion

#region 
import copy

# ==================================================
# CENTRAL APP CONFIG DEFAULTS
# ==================================================
DEFAULT_REPORT_CONFIG = {
    "port_meta": {
        "title": "PROJECT PORTFOLIO | PIK2.D2.GINZA.MIDTOWN OPT.2 R(1)",
        "ref": "REF. DATA R(0) | CONCEPT DWG 2026-02-02.DPA",
        "version": "R (1) OPT2",
        "updated": "02-02-2026",
        "created": "02-02-2026"
    },
    "export_settings": {
        "prepared_by": "",
        "checked_by": ""
    },
    "port_assumptions": [
        {"No.": str(i), "Assumption Description": desc}
        for i, desc in enumerate([
            "Include Vacuum Project + Urugan kembali asumsi 1m",
            "Foundation System standard pilecaps.",
            "No Basement and No Parking Podium.",
            "Parking provison limited to ON STREET LEVEL parking; Floor Hardener finish",
            "Floor to Floor Height at 3.5M",
            "Facade Alumunium Window Wall - + Grill Outdoor AC",
            "External Facade Precast, No double skin for parking podium if any.",
            "Ground Lobby Finishes completed with Artificial stone & HT.",
            "Typical Corridor | Floor finishes : HT | Wall Finishes : Cement Sand Plaster c/w Emulsion Paint.",
            "Aircon System | Apartement : AC Split | Hotel : VRF SYSTEM",
            "SBO Rebars @ Rp. 10.000/kg",
            "Excluded Smarthome",
            "Lift : Luxury Apartment : 8 Private Lift + 2 Services Lift | Hotel 3* : 3 Passenger Lift + 1 Services Lift\nTerrace Village : 16 Private Lift + 8 Services Lift | Retail : No Elevator + Escalator 12 units\nApartment 2 : 4 Passenger Lift + 2 Services Lift\nPodium Village : 10 Private Lift + 5 Services Lift",
            "Exclude Wardrobe",
            "FFE : Kitchen cabinet, Hob & Hood, Refrigerator & Washing Machine",
            "Water Heater : Installation only",
            "Based on Resume Calculation DP dated on 2026.02.02"
        ], 1)
    ],
}

def get_valid_token():
    """Refresh the session token if needed before any Supabase write."""
    try:
        session = supabase.auth.get_session()
        if session and session.session:
            # Refresh token in session state
            st.session_state.access_token = session.session.access_token
            return session.session.access_token
    except Exception:
        pass
    
    # Fall back to stored token
    return st.session_state.get("access_token")

def clean_for_json(obj):
    """
    Convert Streamlit / pandas / numpy objects into JSON-safe Python objects.
    This prevents save_data() from crashing when session_state contains
    numpy, pandas, NaN, inf, Timestamp, DataFrame, Series, etc.
    """

    import math
    import datetime
    from decimal import Decimal

    # None
    if obj is None:
        return None

    # Pandas DataFrame / Series
    if isinstance(obj, pd.DataFrame):
        return clean_for_json(obj.to_dict(orient="records"))

    if isinstance(obj, pd.Series):
        return clean_for_json(obj.to_dict())

    # Pandas / datetime objects
    if isinstance(obj, (pd.Timestamp, datetime.datetime, datetime.date)):
        return obj.isoformat()

    # Numpy scalar types
    if isinstance(obj, np.integer):
        return int(obj)

    if isinstance(obj, np.floating):
        val = float(obj)
        if math.isnan(val) or math.isinf(val):
            return None
        return val

    if isinstance(obj, np.bool_):
        return bool(obj)

    # Normal Python scalar types
    if isinstance(obj, bool):
        return obj

    if isinstance(obj, int):
        return obj

    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj

    if isinstance(obj, Decimal):
        val = float(obj)
        if math.isnan(val) or math.isinf(val):
            return None
        return val

    if isinstance(obj, str):
        return obj

    # Dictionary
    if isinstance(obj, dict):
        return {
            str(k): clean_for_json(v)
            for k, v in obj.items()
        }

    # List / tuple / set
    if isinstance(obj, (list, tuple, set)):
        return [clean_for_json(v) for v in obj]

    # Fallback for unknown objects
    try:
        return str(obj)
    except Exception:
        return None

def init_report_config():
    """
    Creates one global/report-level config bucket.
    Old port_meta / port_assumptions are only migrated ONCE
    when report_config does not exist yet.
    """
    if "report_config" not in st.session_state:
        st.session_state.report_config = copy.deepcopy(DEFAULT_REPORT_CONFIG)

        # One-time migration from old keys
        if "port_meta" in st.session_state:
            st.session_state.report_config["port_meta"] = st.session_state.port_meta

        if "port_assumptions" in st.session_state:
            if isinstance(st.session_state.port_assumptions, pd.DataFrame):
                st.session_state.report_config["port_assumptions"] = (
                    st.session_state.port_assumptions.to_dict("records")
                )
            else:
                st.session_state.report_config["port_assumptions"] = st.session_state.port_assumptions

    st.session_state.report_config.setdefault(
        "port_meta",
        copy.deepcopy(DEFAULT_REPORT_CONFIG["port_meta"])
    )

    st.session_state.report_config.setdefault(
        "export_settings",
        copy.deepcopy(DEFAULT_REPORT_CONFIG["export_settings"])
    )

    st.session_state.report_config.setdefault(
        "port_assumptions",
        copy.deepcopy(DEFAULT_REPORT_CONFIG["port_assumptions"])
    )

def get_report_config():
    init_report_config()
    return st.session_state.report_config

def get_port_meta():
    cfg = get_report_config()
    cfg.setdefault("port_meta", copy.deepcopy(DEFAULT_REPORT_CONFIG["port_meta"]))
    return cfg["port_meta"]

def get_port_assumptions_df():
    cfg = get_report_config()
    cfg.setdefault(
        "port_assumptions",
        copy.deepcopy(DEFAULT_REPORT_CONFIG["port_assumptions"])
    )

    assumptions = cfg["port_assumptions"]

    if isinstance(assumptions, pd.DataFrame):
        return assumptions.copy()

    return pd.DataFrame(assumptions)

def set_port_assumptions_df(df):
    cfg = get_report_config()
    cfg["port_assumptions"] = df.to_dict("records")

def build_app_payload():
    """
    Single source of truth for saving.
    save_data() and save_snapshot() should both use this.
    """
    init_report_config()

    curr_id, _ = repair_projects_state(save=False)

    payload = {
        "app_version": APP_VERSION,
        "projects": st.session_state.get("projects", make_default_projects()),
        "current_proj_id": curr_id,
        "proj_counter": st.session_state.get("proj_counter", 1),
        "current_study_name": st.session_state.get("current_study_name"),

        "loaded_snapshot_id": st.session_state.get("loaded_snapshot_id"),
        "loaded_snapshot_name": st.session_state.get("loaded_snapshot_name"),

        "report_config": st.session_state.get(
            "report_config",
            copy.deepcopy(DEFAULT_REPORT_CONFIG)
        )
    }

    return clean_for_json(payload)

# ==================================================
# CENTRAL STATE KEY REGISTRY
# ==================================================

APP_KEYS = {
    # Core app state
    "projects": "projects",
    "current_proj_id": "current_proj_id",
    "proj_counter": "proj_counter",
    "current_study_name": "current_study_name",

    # Auth / cloud
    "logged_in": "logged_in",
    "storage_loaded": "storage_loaded",
    "access_token": "access_token",
    "user": "user",
    "session_fingerprint": "session_fingerprint",

    # Archive / snapshot
    "loaded_snapshot_id": "loaded_snapshot_id",
    "loaded_snapshot_name": "loaded_snapshot_name",

    # Report config
    "report_config": "report_config",
    "port_meta": "port_meta",
    "port_assumptions": "port_assumptions",

    # UI mode
    "sidebar_component_mode": "sidebar_component_mode",
    "current_page": "current_page",
}


PROJECT_DATA_KEYS = {
    # Area data
    "area_table": "area_table_data",
    "gba": "m_gba",
    "gfa": "m_gfa",
    "sgfa": "m_sgfa",
    "nfa": "m_nfa",
    "rooms": "m_rooms",
    "facade": "m_facade",
    "lobby": "m_lobby",

    # Door data
    "door_table": "area_door_table_data",
    "door_wood_calc": "area_door_wood_calc",
    "door_steel_calc": "area_door_steel_calc",
    "door_glass_calc": "area_door_glass_calc",
    "door_wood_qty": "m_door_wood",
    "door_steel_qty": "m_door_steel",
    "door_glass_qty": "m_door_glass",

    # Custom cost data
    "smart_custom_costs": "smart_custom_costs",
    "external_custom_costs": "external_custom_costs",

    # Common cost quantity/rate keys
    "earth_rate": "u_earth",
    "foundation_rate": "u_found",
    "structure_rate": "u_struc",
    "architecture_rate": "u_arch",
    "mep_rate": "u_mep",
    "utility_rate": "u_util",
}


PROJECT_DATA_GROUPS = {
    "area": [
        "area_table",
        "gba",
        "gfa",
        "sgfa",
        "nfa",
        "rooms",
        "facade",
        "lobby",
    ],
    "door": [
        "door_table",
        "door_wood_calc",
        "door_steel_calc",
        "door_glass_calc",
        "door_wood_qty",
        "door_steel_qty",
        "door_glass_qty",
    ],
    "cost": [
        "smart_custom_costs",
        "external_custom_costs",
        "earth_rate",
        "foundation_rate",
        "structure_rate",
        "architecture_rate",
        "mep_rate",
        "utility_rate",
    ],
}


UI_CACHE_PREFIXES = (
    "base_table_",
    "m_",
    "r_",
    "u_",
    "sc_",
    "misc_sw_",
    "wid_",
    "temp_spec_",
    "area_committed_",
    "area_draft_",
    "door_committed_",
    "door_draft_",
    "area_page_",
    "area_input_mode_",
    "area_keliling_facade_",
    "area_panjang_railing_",
    "area_tinggi_railing_",
    "area_facade_tolerance_pct_",
    "excel_form_",
    "area_excel_upload_",
    "import_area_excel_",
    "generate_area_table_",
    "reset_area_draft_",
    "save_area_table_",
    "save_door_table_to_cloud_",
    "save_external_",
    "save_res_fac_",
    "use_area_analysis_",
    "save_smart_custom_to_cloud_",
    "save_cost_analysis_",
    "architectural_detail_",
    "consultancy_detail_",
    "earthwork_detail_",
    "ffe_detail_",
    "foundation_detail_",
    "mep_detail_",
    "structural_detail_",
    "utility_detail_",
    "area_editor_",
    "edit_smart_cc_",
    "external_table_",
    "other_external_editor_",
    "landscape_qty_",
    "hardscape_pct_",
    "softscape_pct_",
    "softscape_pct_display_",
    "hardscape_rate_",
    "softscape_rate_",
    "res_fac_table_",
    "res_fac_editor_",
    "door_table_",
    "door_editor_",
    "smart_custom_data_",
)

def make_default_projects():
    return {
        "proj_1": {
            "name": "New Project 1",
            "type": "Hotel",
            "data": {}
        }
    }

def resolve_current_study_name(projects=None, fallback="Untitled Study"):
    name = st.session_state.get("current_study_name")
    if isinstance(name, str) and name.strip():
        return name.strip()

    archive_name = st.session_state.get("loaded_snapshot_name")
    if isinstance(archive_name, str) and archive_name.strip():
        return archive_name.strip()

    if projects is None:
        projects = st.session_state.get("projects", {})

    if isinstance(projects, dict) and projects:
        first_project = next(iter(projects.values()), {})
        if isinstance(first_project, dict):
            project_name = first_project.get("name")
            if isinstance(project_name, str) and project_name.strip():
                return project_name.strip()

    return fallback

def repair_projects_state(save=False):
    projects = st.session_state.get("projects", {})

    # Guard 1: projects must be a non-empty dict
    if not isinstance(projects, dict) or len(projects) == 0:
        st.session_state.projects = make_default_projects()
        st.session_state.current_proj_id = "proj_1"
        st.session_state.proj_counter = 1
        if save:
            save_data()
        return "proj_1", st.session_state.projects["proj_1"]

    # Guard 2: remove any corrupt project entries (non-dict values)
    corrupt_keys = [k for k, v in projects.items() if not isinstance(v, dict)]
    for k in corrupt_keys:
        del projects[k]

    # If all entries were corrupt, rebuild from scratch
    if len(projects) == 0:
        st.session_state.projects = make_default_projects()
        st.session_state.current_proj_id = "proj_1"
        st.session_state.proj_counter = 1
        if save:
            save_data()
        return "proj_1", st.session_state.projects["proj_1"]

    # Guard 3: current_proj_id must be a valid string key in projects
    curr_id = st.session_state.get("current_proj_id")

    if not isinstance(curr_id, str) or curr_id not in projects:
        curr_id = list(projects.keys())[0]
        st.session_state.current_proj_id = curr_id
        if save:
            save_data()

    return curr_id, projects[curr_id]

def get_current_project():
    return repair_projects_state(save=True)

# ==================================================
# PROJECT DATA HELPERS
# ==================================================

def resolve_project_data_key(key_alias):
    """
    Converts a readable alias like 'gba' into the actual stored key 'm_gba'.
    If the alias is not registered, return it unchanged.
    """
    return PROJECT_DATA_KEYS.get(key_alias, key_alias)


def get_project_data_map():
    rows = []

    for group, aliases in PROJECT_DATA_GROUPS.items():
        for alias in aliases:
            rows.append({
                "Group": group,
                "Alias": alias,
                "Stored Key": PROJECT_DATA_KEYS.get(alias, alias)
            })

    return pd.DataFrame(rows)


def get_project_data():
    """
    Returns current project id, project object, and its saved data dict.
    This is for actual saved project data, not temporary widget state.
    """
    curr_id, curr_proj = get_current_project()

    if "data" not in curr_proj or not isinstance(curr_proj.get("data"), dict):
        curr_proj["data"] = {}

    return curr_id, curr_proj, curr_proj["data"]


def get_data(key_alias, default=None):
    """
    Read from current project's saved data.
    Example:
        get_data("gba", 0)
        same as d.get("m_gba", 0)
    """
    _, _, d = get_project_data()
    real_key = resolve_project_data_key(key_alias)
    return d.get(real_key, default)


def set_data(key_alias, value):
    """
    Write to current project's saved data.
    Example:
        set_data("gba", 10000)
        same as d["m_gba"] = 10000
    """
    curr_id, curr_proj, d = get_project_data()
    real_key = resolve_project_data_key(key_alias)

    d[real_key] = value

    try:
        if isinstance(value, list):
            summary = f"list with {len(value)} rows"
        elif isinstance(value, dict):
            summary = f"dict with {len(value)} keys"
        else:
            summary = str(value)

        debug_log(
            f"set_data: project={curr_id}, alias={key_alias}, key={real_key}, value={summary}"
        )
    except Exception:
        pass

    return value


def update_data(values):
    """
    Write multiple saved project data values at once.

    Example:
        update_data({
            "gba": gba,
            "gfa": gfa,
            "sgfa": sgfa,
        })
    """
    curr_id, curr_proj, d = get_project_data()

    changed_keys = []

    for key_alias, value in values.items():
        real_key = resolve_project_data_key(key_alias)
        d[real_key] = value
        changed_keys.append(f"{key_alias}->{real_key}")

    try:
        debug_log(
            f"update_data: project={curr_id}, keys={', '.join(changed_keys)}"
        )
    except Exception:
        pass

    return d

def is_duplicate_project_name(clean_name, current_id=None):
    clean_name = str(clean_name).strip().lower()

    if clean_name == "":
        return False

    for pid, pdata in st.session_state.get("projects", {}).items():
        if current_id is not None and pid == current_id:
            continue

        existing_name = str(pdata.get("name", "")).strip().lower()

        if existing_name == clean_name:
            return True

    return False

def clear_project_ui_cache():
    """
    Clear Streamlit widget/session cache that can override freshly loaded project data.
    Do not delete actual project payload keys.
    """
    safe_keys = {
        APP_KEYS["projects"],
        APP_KEYS["current_proj_id"],
        APP_KEYS["proj_counter"],
        APP_KEYS["current_study_name"],
        APP_KEYS["logged_in"],
        APP_KEYS["access_token"],
        APP_KEYS["user"],
        APP_KEYS["storage_loaded"],
        APP_KEYS["session_fingerprint"],
        APP_KEYS["loaded_snapshot_id"],
        APP_KEYS["loaded_snapshot_name"],
        APP_KEYS["report_config"],
    }

    keys_to_delete = [
        k for k in list(st.session_state.keys())
        if (
            "proj_" in str(k)
            or any(str(k).startswith(prefix) for prefix in UI_CACHE_PREFIXES)
        )
        and k not in safe_keys
    ]

    for k in keys_to_delete:
        del st.session_state[k]

def clear_project_ui_cache_for_ids(project_ids):
    """
    Clear Streamlit widget/session cache for specific deleted project/component IDs.
    This prevents deleted component widget values from surviving invisibly.
    """
    if not project_ids:
        return

    project_ids = [str(pid) for pid in project_ids]

    safe_keys = {
        APP_KEYS["projects"],
        APP_KEYS["current_proj_id"],
        APP_KEYS["proj_counter"],
        APP_KEYS["current_study_name"],
        APP_KEYS["logged_in"],
        APP_KEYS["access_token"],
        APP_KEYS["user"],
        APP_KEYS["storage_loaded"],
        APP_KEYS["session_fingerprint"],
        APP_KEYS["loaded_snapshot_id"],
        APP_KEYS["loaded_snapshot_name"],
        APP_KEYS["report_config"],
    }

    keys_to_delete = []

    for key in list(st.session_state.keys()):
        key_str = str(key)

        if any(pid in key_str for pid in project_ids) and key not in safe_keys:
            keys_to_delete.append(key)

    for key in keys_to_delete:
        del st.session_state[key]

def restore_app_payload(data):
    """
    Single source of truth for loading.
    Snapshot load and startup load should both use this.
    """
    # WIPE UI CACHE to prevent old widgets from overriding the newly loaded data!
    clear_project_ui_cache()

    if not data:
        data = {}

    projects = data.get("projects", make_default_projects())

    # Critical guard: prevent empty project dictionary
    if not isinstance(projects, dict) or len(projects) == 0:
        projects = make_default_projects()

    st.session_state.projects = projects

    saved_curr_id = data.get("current_proj_id")

    if isinstance(saved_curr_id, str) and saved_curr_id in st.session_state.projects:
        st.session_state.current_proj_id = saved_curr_id
    else:
        st.session_state.current_proj_id = list(st.session_state.projects.keys())[0]

    st.session_state.proj_counter = data.get(
        "proj_counter",
        len(st.session_state.projects)
    )

    # New format
    report_config = copy.deepcopy(DEFAULT_REPORT_CONFIG)
    report_config.update(data.get("report_config", {}))

    # Backward compatibility with old snapshots
    if "port_meta" in data:
        report_config["port_meta"] = data["port_meta"]

    if "port_assumptions" in data:
        report_config["port_assumptions"] = data["port_assumptions"]

    st.session_state.report_config = report_config

    # Active archive reference
    st.session_state.loaded_snapshot_id = data.get("loaded_snapshot_id")
    st.session_state.loaded_snapshot_name = data.get("loaded_snapshot_name")

    current_study_name = data.get("current_study_name")
    if isinstance(current_study_name, str) and current_study_name.strip():
        st.session_state.current_study_name = current_study_name.strip()
    else:
        if "current_study_name" in st.session_state:
            del st.session_state["current_study_name"]
        st.session_state.current_study_name = resolve_current_study_name(
            st.session_state.projects
        )

    # Optional backward compatibility aliases
    st.session_state.port_meta = st.session_state.report_config["port_meta"]
    st.session_state.port_assumptions = pd.DataFrame(
        st.session_state.report_config["port_assumptions"]
    )
    #endregion 

#region --- DO NOT CHANGE (OR I WILL KICK YOUR BUTT)---
from supabase import create_client, Client
url: str = st.secrets["SUPABASE_URL"]
key: str = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(url, key)

def calculate_project_totals(pdata, curr_type):
    """Calculates all totals dynamically for a given project."""
    d = pdata.get("data", {})
    # Default to empty dict if project type isn't in database yet
    pt_data = PROJECT_DATABASE.get(curr_type, {})

    def get_val(key, default=0.0):
        val = d.get(key, default)
        if isinstance(val, list): return val
        try: return _safe_float(val)
        except: return val

    # --- Area Calculations ---
    area_table = get_val("area_table_data", [])
    table_totals = {
        "gba": 0.0,
        "gfa": 0.0,
        "sgfa": 0.0,
        "nfa": 0.0,
    }

    if isinstance(area_table, list) and len(area_table) > 0:
        unit_area_col = "Unit Area"
        breakdown_cols = ["Parkir", "Roof/Deck", "MEP Outdoor", "Koridor/Lobby", "Stair, MEP, Etc", unit_area_col, "Office"]
        
        for row in area_table:
            row_total = sum(_safe_float(row.get(c, 0)) for c in breakdown_cols)
            if unit_area_col not in row and "Unit" in row:
                row_total += _safe_float(row.get("Unit", 0))
            table_totals["gba"] += row_total
            table_totals["gfa"] += row_total - sum(_safe_float(row.get(c, 0)) for c in ["Parkir", "Roof/Deck", "MEP Outdoor"])
            unit_area = _safe_float(row.get(unit_area_col, row.get("Unit", 0)))
            table_totals["sgfa"] += unit_area + sum(_safe_float(row.get(c, 0)) for c in ["Office", "Koridor/Lobby"])
            table_totals["nfa"] += unit_area + _safe_float(row.get("Office", 0))

    def resolve_area_value(manual_key, table_key):
        manual_value = _safe_float(get_val(manual_key, 0.0))
        return manual_value if manual_value > 0 else table_totals[table_key]

    calc_gba = resolve_area_value("m_gba", "gba")
    calc_gfa = resolve_area_value("m_gfa", "gfa")
    calc_sgfa = resolve_area_value("m_sgfa", "sgfa")
    calc_nfa = resolve_area_value("m_nfa", "nfa")

    # --- Cost Calculations ---
    struc_earth = get_val("u_earth", pt_data.get("struc_earth", 0))
    struc_found = get_val("u_found", pt_data.get("struc_found", 0))
    struc_work = get_val("u_struc", pt_data.get("struc_work", 0))
    arch_base = get_val("u_arch", pt_data.get("arch_base", 0))
    
    facade = get_val("m_facade", 0.0)
    fac_precast_pct = get_val("r_fac_pre", pt_data.get("facade_precast_pct", 0))
    fac_precast_rate = get_val("u_f_pre", pt_data.get("facade_precast_rate", 0))
    
    rooms = get_val("m_rooms", 0.0)
    ffe_rate = get_val("u_ffe", pt_data.get("ffe", 0))
    mep_rate = get_val("u_mep", pt_data.get("mep", 0))
    utility_rate = get_val("u_util", pt_data.get("utility", 0))
    
    consultancy_rate = get_val("sc_cons", pt_data.get("cons", 0))
    insurance_pct = get_val("sc_ins", 0.12)
    
    smart_custom_costs = sum(_safe_float(item.get("Rate (Rp)", 0)) * _safe_float(item.get("Quantity", 1)) for item in get_val("smart_custom_costs", []))

    t_earth = calc_gba * struc_earth
    t_found = calc_gba * struc_found
    t_struc = calc_gba * struc_work
    t_arch_base = calc_gfa * arch_base
    t_precast = facade * (fac_precast_pct / 100) * fac_precast_rate
    t_ffe = rooms * ffe_rate
    t_mep = calc_gba * mep_rate
    t_utility = calc_gba * utility_rate
    
    construction_subtotal = sum([t_earth, t_found, t_struc, t_arch_base, t_precast, t_ffe, t_mep, t_utility, smart_custom_costs])

    t_preliminary = construction_subtotal * 0.05
    t_contingency = (construction_subtotal + t_preliminary) * 0.03
    grand_total_hc = construction_subtotal + t_preliminary + t_contingency

    total_soft_cost = (calc_gfa * consultancy_rate) + (grand_total_hc * (insurance_pct / 100.0)) 
    
    calc_budget = grand_total_hc + total_soft_cost

    return calc_gba, calc_gfa, calc_sgfa, calc_budget, rooms

from datetime import date, datetime, timedelta

def _get_authed_snapshot_client(show_error=True):
    token = st.session_state.get("access_token")
    user = st.session_state.get("user")

    # Handle both object and dict representations of user
    user_id = None
    if user is not None:
        user_id = getattr(user, "id", None)
        if user_id is None and isinstance(user, dict):
            user_id = user.get("id")

    if not token:
        if show_error:
            st.error("Not authenticated: missing access token.")
        return None, None

    if not user_id:
        if show_error:
            st.error("Not authenticated: missing user ID.")
        return None, None

    try:
        authed_client = create_client(url, key)
        authed_client.postgrest.auth(token)
        return authed_client, user_id
    except Exception as e:
        if show_error:
            st.error(f"Failed to create authenticated client: {e}")
        return None, None

import hashlib

def generate_session_fingerprint(token):
    """Create a unique fingerprint from the access token."""
    try:
        return hashlib.sha256(token.encode()).hexdigest()
    except Exception:
        return None

def register_device_session():
    """
    Called on login. Stores this device's session fingerprint in Supabase.
    If another device is already logged in, it gets kicked out.
    """
    token = st.session_state.get("access_token")
    user = st.session_state.get("user")
    user_id = getattr(user, "id", None)
    if user_id is None and isinstance(user, dict):
        user_id = user.get("id")

    if not token or not user_id:
        return False

    fingerprint = generate_session_fingerprint(token)
    if not fingerprint:
        return False

    st.session_state["session_fingerprint"] = fingerprint

    try:
        supabase.postgrest.auth(token)

        supabase.table("user_sessions").upsert({
            "user_id": user_id,
            "session_token": fingerprint,
            "device_info": f"Session registered at login",
            "logged_in_at": datetime.utcnow().isoformat()
        }).execute()
        return True
    except Exception as e:
        st.error(f"Session registration error: {e}")
        return False

def is_session_valid():
    """
    Check if this device still holds the active session.
    Returns True if valid, False if another device has taken over.
    """
    token = st.session_state.get("access_token")
    user = st.session_state.get("user")
    user_id = getattr(user, "id", None)
    if user_id is None and isinstance(user, dict):
        user_id = user.get("id")

    if not token or not user_id:
        return False

    local_fingerprint = st.session_state.get("session_fingerprint")
    if not local_fingerprint:
        # Regenerate from current token if missing
        local_fingerprint = generate_session_fingerprint(token)
        st.session_state["session_fingerprint"] = local_fingerprint

    try:
        supabase.postgrest.auth(token)

        response = supabase.table("user_sessions") \
            .select("session_token") \
            .eq("user_id", user_id) \
            .execute()

        if not response.data:
            return False

        stored_fingerprint = response.data[0].get("session_token")
        return stored_fingerprint == local_fingerprint

    except Exception as e:
        st.error(f"Session check error: {e}")
        return False

def enforce_single_device():
    """
    Call before any write operation.
    Hard stops the app if another session has taken over.
    """
    
    if not is_session_valid():
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.warning("Session Expired", icon=icon_safe("lock"))
            st.error(
                "**Your session is over** \n\n"
                "Cause: Log in from another device.\n\n"
                "Only one active session is allowed at a time. \n\n"
                "**Please log out and log back in to continue.**"
            )
            if st.button("Logout", type="primary", key="force_logout_btn"):
                st.session_state.logged_in = False
                st.session_state.access_token = None
                st.session_state.user = None
                st.session_state.session_fingerprint = None
                for key in ["projects", "storage_loaded", "report_config"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
        st.stop()
        return False
    return True

def format_snapshot_time(created_at):
    if not created_at:
        return ""

    try:
        created_utc = datetime.fromisoformat(created_at.replace("Z", "+00:00"))
        created_local = created_utc + timedelta(hours=7)
        return created_local.strftime("%d %b %Y, %H:%M WIB")
    except Exception:
        return ""

def save_snapshot(snapshot_name):
    if not enforce_single_device():
        return False

    authed_client, user_id = _get_authed_snapshot_client()

    if not authed_client or not user_id:
        return False

    st.session_state.current_study_name = snapshot_name
    payload = build_app_payload()

    try:
        response = authed_client.table("project_snapshots").insert({
            "user_id": user_id,
            "snapshot_name": snapshot_name,
            "data": payload
        }).execute()

        #debugcode
        if not response.data:
            st.error("Project Save failed - Supabase returned no inserted row.")
            return False
        #enddebugcode

        if response.data:
            st.session_state.loaded_snapshot_id = response.data[0].get("id")
            st.session_state.loaded_snapshot_name = response.data[0].get("snapshot_name", snapshot_name)
            st.session_state.current_study_name = st.session_state.loaded_snapshot_name

        save_data()
        return True

    except Exception as e:
        st.error(f"Project Save Error: {e}")
        return False

def overwrite_current_snapshot():
    if not enforce_single_device():
        return False

    snapshot_id = st.session_state.get("loaded_snapshot_id")
    snapshot_name = st.session_state.get("loaded_snapshot_name")

    if not snapshot_id:
        st.error("No archive linked. Use Archive > Online Backup > Save first.")
        return False

    authed_client, user_id = _get_authed_snapshot_client()
    if not authed_client or not user_id:
        return False

    payload = build_app_payload()

    try:
        response = authed_client.table("project_snapshots") \
            .update({
                "snapshot_name": snapshot_name,
                "data": payload
            }) \
            .eq("id", snapshot_id) \
            .eq("user_id", user_id) \
            .execute()

        # Supabase returns empty data if the row wasn't found/matched
        if not response.data:
            st.error(
                f"Snapshot not saved - no matching row found. "
                f"ID: {snapshot_id}, user: {user_id}. "
                "The archive may belong to a different account."
            )
            return False

        # Keep project_storage aligned with the active snapshot,
        # because login restores from project_storage.
        storage_ok = save_data()
        if not storage_ok:
            st.warning("Snapshot saved, but cloud storage was not updated. Login may restore older data.")
            return False

        return True

    except Exception as e:
        st.error(f"Quick Save Error: {e}")
        return False

def rename_snapshot(snapshot_id, new_name):
    token = st.session_state.get("access_token")
    user = st.session_state.get("user")
    user_id = getattr(user, "id", None)

    if not token or not user_id:
        st.error("Not authenticated.")
        return False

    if not snapshot_id or not str(new_name).strip():
        st.error("Invalid project name.")
        return False

    authed_client = create_client(url, key)
    authed_client.postgrest.auth(token)

    clean_name = str(new_name).strip()

    try:
        authed_client.table("project_snapshots") \
            .update({
                "snapshot_name": clean_name
            }) \
            .eq("id", snapshot_id) \
            .eq("user_id", user_id) \
            .execute()

        if st.session_state.get("loaded_snapshot_id") == snapshot_id:
            st.session_state.loaded_snapshot_name = clean_name
            st.session_state.current_study_name = clean_name
            save_data()

        return True

    except Exception as e:
        st.error(f"Project Rename Error: {e}")
        return False

def overwrite_snapshot(snapshot_id, snapshot_name=None):
    authed_client, user_id = _get_authed_snapshot_client()

    if not authed_client or not user_id:
        return False

    if not snapshot_id:
        st.error("No saved project selected.")
        return False

    payload = build_app_payload()

    update_payload = {
        "data": payload
    }

    # Optional rename while overwriting
    if snapshot_name is not None and str(snapshot_name).strip() != "":
        update_payload["snapshot_name"] = str(snapshot_name).strip()

    try:
        authed_client.table("project_snapshots") \
            .update(update_payload) \
            .eq("id", snapshot_id) \
            .eq("user_id", user_id) \
            .execute()

        st.session_state.loaded_snapshot_id = snapshot_id

        if snapshot_name is not None and str(snapshot_name).strip() != "":
            st.session_state.loaded_snapshot_name = str(snapshot_name).strip()
            st.session_state.current_study_name = st.session_state.loaded_snapshot_name

        return True

    except Exception as e:
        st.error(f"Project Overwrite Error: {e}")
        return False

def load_snapshots():
    authed_client, user_id = _get_authed_snapshot_client(show_error=False)

    if not authed_client or not user_id:
        return []

    try:
        response = authed_client.table("project_snapshots") \
            .select("id, snapshot_name, created_at") \
            .eq("user_id", user_id) \
            .order("created_at", desc=True) \
            .execute()

        return response.data if response.data else []

    except Exception as e:
        st.error(f"Project Load Error: {e}")
        return []

def load_snapshot_data(snapshot_id):
    authed_client, user_id = _get_authed_snapshot_client()

    if not authed_client or not user_id:
        return None

    try:
        response = authed_client.table("project_snapshots") \
            .select("data, snapshot_name") \
            .eq("id", snapshot_id) \
            .eq("user_id", user_id) \
            .execute()

        if response.data:
            st.session_state.loaded_snapshot_id = snapshot_id
            st.session_state.loaded_snapshot_name = response.data[0].get("snapshot_name", "")
            st.session_state.current_study_name = st.session_state.loaded_snapshot_name
            return response.data[0]["data"]

    except Exception as e:
        st.error(f"Saved Project Fetch Error: {e}")

    return None

def delete_snapshot(snapshot_id):
    if not enforce_single_device():
        return False

    authed_client, user_id = _get_authed_snapshot_client()

    if not authed_client or not user_id:
        return False

    try:
        #debugcode
        response = authed_client.table("project_snapshots") \
            .delete() \
            .eq("id", snapshot_id) \
            .eq("user_id", user_id) \
            .execute()

        if not response.data:
            st.error("Project Delete failed - no matching snapshot row was deleted.")
            return False
        #enddebugcode

        if st.session_state.get("loaded_snapshot_id") == snapshot_id:
            st.session_state.loaded_snapshot_id = None
            st.session_state.loaded_snapshot_name = None
            if not st.session_state.get("current_study_name"):
                st.session_state.current_study_name = resolve_current_study_name()

        return True

    except Exception as e:
        st.error(f"Project Delete Error: {e}")
        return False

def save_data_force():
    """
    Save current app state to project_storage even if storage_loaded is not set yet.
    Use only after explicit user actions like create/edit/delete component.
    """
    if not st.session_state.get("logged_in", False):
        st.error("Cloud Save Error: not logged in.")
        return False

    if not enforce_single_device():
        return False

    token = st.session_state.get("access_token")
    user = st.session_state.get("user")

    user_id = getattr(user, "id", None)
    if user_id is None and isinstance(user, dict):
        user_id = user.get("id")

    if not token or not user_id:
        st.error("Cloud Save Error: missing token or user ID.")
        return False

    try:
        supabase.postgrest.auth(token)

        payload = build_app_payload()

        response = supabase.table("project_storage").upsert({
            "id": f"storage_{user_id}",
            "user_id": user_id,
            "data": payload
        }).execute()

        if not response.data:
            st.error("Cloud Save failed - Supabase returned no saved row.")
            return False

        return True

    except Exception as e:
        st.error(f"Cloud Save Error: {e}")
        return False

def save_data():
    if not st.session_state.get("storage_loaded", False):
        return False
    if not st.session_state.get("logged_in", False):
        return False

    if not enforce_single_device():
        return False

    token = st.session_state.get("access_token")
    user = st.session_state.get("user")
    user_id = getattr(user, "id", None)
    if user_id is None and isinstance(user, dict):
        user_id = user.get("id")

    if not token or not user_id:
        return False

    # Reuse global client
    try:
        #debugcode
        debug_log("save_data started")
        #enddebugcode
        supabase.postgrest.auth(token)

        payload = build_app_payload()

        #debugcode
        response = supabase.table("project_storage").upsert({
            "id": f"storage_{user_id}",
            "user_id": user_id,
            "data": payload
        }).execute()

        if not response.data:
            st.error("Cloud Save failed - Supabase returned no saved row.")
            debug_log("save_data failed: Supabase returned no saved row.", level="ERROR")
            return False

        debug_log("save_data completed successfully")
        #enddebugcode
        return True
    except Exception as e:
        #debugcode
        debug_log(f"save_data failed: {e}", level="ERROR")
        #enddebugcode
        st.error(f"Cloud Save Error: {e}")
        return False
        
def load_data():
    token = st.session_state.get("access_token")
    user = st.session_state.get("user")

    if not token or not user:
        return None

    user_id = getattr(user, "id", None)
    if user_id is None and isinstance(user, dict):
        user_id = user.get("id")

    if not user_id:
        st.error("Cloud Load Error: missing user ID.")
        return None

    try:
        authed_client = create_client(url, key)
        authed_client.postgrest.auth(token)

        response = authed_client.table("project_storage") \
            .select("data") \
            .eq("id", f"storage_{user_id}") \
            .execute()

        if response.data:
            return response.data[0]["data"]
    except Exception as e:
        st.error(f"Cloud Load Error: {e}")
    return None

#debugcode
# ==================================================
# APP DEBUGGER / HEALTH CHECK SYSTEM
# ==================================================

def _debug_user_id():
    """Safely resolve current Supabase user_id from session_state."""
    user = st.session_state.get("user")

    if user is None:
        return None

    user_id = getattr(user, "id", None)

    if user_id is None and isinstance(user, dict):
        user_id = user.get("id")

    return user_id

def _debug_mask(value, visible=6):
    """Hide sensitive token/session values."""
    if not value:
        return None

    value = str(value)

    if len(value) <= visible * 2:
        return "***"

    return f"{value[:visible]}...{value[-visible:]}"

def _debug_result(name, status, detail="", fix=""):
    """
    Standard debug result object.
    status options:
    - PASS
    - WARN
    - FAIL
    - INFO
    """
    return {
        "Check": name,
        "Status": status,
        "Detail": detail,
        "Suggested Fix": fix
    }

def debug_log(message, level="INFO"):
    if "debug_logs" not in st.session_state:
        st.session_state.debug_logs = []

    st.session_state.debug_logs.append({
        "time": datetime.utcnow().isoformat(),
        "level": level,
        "message": str(message)
    })

    st.session_state.debug_logs = st.session_state.debug_logs[-200:]

def render_debug_log():
    logs = st.session_state.get("debug_logs", [])

    if not logs:
        st.info("No debug logs yet.")
        return

    df = pd.DataFrame(logs)
    st.dataframe(df.sort_values("time", ascending=False), width="stretch", hide_index=True)

def run_local_state_debug_checks():
    """Check Streamlit session_state and local app structure."""
    results = []

    logged_in = st.session_state.get("logged_in", False)
    results.append(
        _debug_result(
            "Logged in flag",
            "PASS" if logged_in else "FAIL",
            f"logged_in={logged_in}",
            "Log in again if this is False."
        )
    )

    token = None
    try:
        token = get_valid_token()
    except Exception as e:
        results.append(
            _debug_result(
                "get_valid_token()",
                "FAIL",
                f"Function crashed: {e}",
                "Fix get_valid_token() before testing cloud save/load."
            )
        )

    if token:
        results.append(
            _debug_result(
                "Access token",
                "PASS",
                f"Token found: {_debug_mask(token)}",
                ""
            )
        )
    else:
        results.append(
            _debug_result(
                "Access token",
                "FAIL",
                "No token found.",
                "Check login flow. Make sure st.session_state.access_token is set after login."
            )
        )

    user_id = _debug_user_id()
    results.append(
        _debug_result(
            "User ID",
            "PASS" if user_id else "FAIL",
            f"user_id={user_id}",
            "Make sure st.session_state.user is stored after login."
        )
    )

    storage_loaded = st.session_state.get("storage_loaded", False)
    results.append(
        _debug_result(
            "Storage loaded flag",
            "PASS" if storage_loaded else "WARN",
            f"storage_loaded={storage_loaded}",
            "save_data() will skip saving if this is False. This is okay before initial load, bad after login."
        )
    )

    projects = st.session_state.get("projects")
    if not isinstance(projects, dict):
        results.append(
            _debug_result(
                "Projects state",
                "FAIL",
                f"projects type is {type(projects).__name__}",
                "Run repair_projects_state(save=False)."
            )
        )
    elif len(projects) == 0:
        results.append(
            _debug_result(
                "Projects state",
                "FAIL",
                "projects is an empty dict.",
                "Reset to make_default_projects()."
            )
        )
    else:
        results.append(
            _debug_result(
                "Projects state",
                "PASS",
                f"{len(projects)} project(s) found.",
                ""
            )
        )

    curr_id = st.session_state.get("current_proj_id")
    if isinstance(projects, dict) and curr_id in projects:
        results.append(
            _debug_result(
                "Current project ID",
                "PASS",
                f"current_proj_id={curr_id}",
                ""
            )
        )
    else:
        results.append(
            _debug_result(
                "Current project ID",
                "FAIL",
                f"current_proj_id={curr_id}; not found in projects.",
                "Run repair_projects_state(save=False)."
            )
        )

    report_config = st.session_state.get("report_config")
    results.append(
        _debug_result(
            "Report config",
            "PASS" if isinstance(report_config, dict) else "WARN",
            f"report_config type={type(report_config).__name__}",
            "Run init_report_config() if missing."
        )
    )

    snapshot_id = st.session_state.get("loaded_snapshot_id")
    snapshot_name = st.session_state.get("loaded_snapshot_name")

    if snapshot_id:
        results.append(
            _debug_result(
                "Loaded snapshot ID",
                "PASS",
                f"loaded_snapshot_id={snapshot_id}; name={snapshot_name}",
                ""
            )
        )
    else:
        results.append(
            _debug_result(
                "Loaded snapshot ID",
                "WARN",
                "No loaded snapshot linked.",
                "Quick Save to archive will fail until user uses Archive > Online Backup > Save or loads an archive."
            )
        )

    return results

def run_payload_debug_checks():
    """Check build_app_payload() and JSON serialization safety."""
    results = []

    try:
        payload = build_app_payload()
        results.append(
            _debug_result(
                "build_app_payload()",
                "PASS",
                "Payload built successfully.",
                ""
            )
        )
    except Exception as e:
        results.append(
            _debug_result(
                "build_app_payload()",
                "FAIL",
                f"Payload build crashed: {e}",
                "Check repair_projects_state(), report_config, and project data structure."
            )
        )
        return results, None

    try:
        encoded = _json.dumps(payload, ensure_ascii=False, default=str)
        size_kb = len(encoded.encode("utf-8")) / 1024

        results.append(
            _debug_result(
                "Payload JSON serialization",
                "PASS",
                f"Payload is JSON-safe. Size: {size_kb:,.1f} KB",
                ""
            )
        )

        if size_kb > 5000:
            results.append(
                _debug_result(
                    "Payload size",
                    "WARN",
                    f"Payload is large: {size_kb:,.1f} KB",
                    "Large payloads can slow down Supabase save/load. Check if widget cache or unnecessary data is being saved."
                )
            )
        else:
            results.append(
                _debug_result(
                    "Payload size",
                    "PASS",
                    f"{size_kb:,.1f} KB",
                    ""
                )
            )

    except Exception as e:
        results.append(
            _debug_result(
                "Payload JSON serialization",
                "FAIL",
                f"JSON dump failed: {e}",
                "clean_for_json() missed an object type. Add conversion for the failing type."
            )
        )

    required_keys = [
        "app_version",
        "projects",
        "current_proj_id",
        "proj_counter",
        "current_study_name",
        "loaded_snapshot_id",
        "loaded_snapshot_name",
        "report_config"
    ]

    for key in required_keys:
        if key in payload:
            results.append(
                _debug_result(
                    f"Payload key: {key}",
                    "PASS",
                    "Exists.",
                    ""
                )
            )
        else:
            results.append(
                _debug_result(
                    f"Payload key: {key}",
                    "FAIL",
                    "Missing.",
                    f"Add {key} to build_app_payload()."
                )
            )

    return results, payload

def run_supabase_debug_checks():
    """Check Supabase auth, storage row, snapshot row, and session row."""
    results = []

    token = None
    try:
        token = get_valid_token()
    except Exception as e:
        results.append(
            _debug_result(
                "Supabase token check",
                "FAIL",
                f"get_valid_token() crashed: {e}",
                "Fix token refresh first."
            )
        )
        return results

    user_id = _debug_user_id()

    if not token or not user_id:
        results.append(
            _debug_result(
                "Supabase prerequisites",
                "FAIL",
                f"token={bool(token)}, user_id={user_id}",
                "Login again. Token and user_id are required."
            )
        )
        return results

    try:
        supabase.postgrest.auth(token)
        results.append(
            _debug_result(
                "Supabase auth binding",
                "PASS",
                "Token applied to global Supabase client.",
                ""
            )
        )
    except Exception as e:
        results.append(
            _debug_result(
                "Supabase auth binding",
                "FAIL",
                f"supabase.postgrest.auth(token) failed: {e}",
                "Check token validity and Supabase client initialization."
            )
        )
        return results

    try:
        response = supabase.table("user_sessions") \
            .select("session_token, logged_in_at") \
            .eq("user_id", user_id) \
            .execute()

        if response.data:
            stored_fingerprint = response.data[0].get("session_token")
            local_fingerprint = st.session_state.get("session_fingerprint")

            if stored_fingerprint == local_fingerprint:
                results.append(
                    _debug_result(
                        "Single-device session",
                        "PASS",
                        "Current device session matches Supabase.",
                        ""
                    )
                )
            else:
                results.append(
                    _debug_result(
                        "Single-device session",
                        "FAIL",
                        "Local session fingerprint does not match Supabase.",
                        "Another device/browser likely logged in. Log out and log back in."
                    )
                )
        else:
            results.append(
                _debug_result(
                    "Single-device session",
                    "FAIL",
                    "No user_sessions row found.",
                    "Run register_device_session() after login."
                )
            )

    except Exception as e:
        results.append(
            _debug_result(
                "Single-device session query",
                "FAIL",
                f"user_sessions query failed: {e}",
                "Check user_sessions table and RLS policies."
            )
        )

    try:
        storage_id = f"storage_{user_id}"

        response = supabase.table("project_storage") \
            .select("id, user_id, data") \
            .eq("id", storage_id) \
            .execute()

        if response.data:
            data = response.data[0].get("data", {})
            results.append(
                _debug_result(
                    "Project storage row",
                    "PASS",
                    f"Found project_storage row: {storage_id}. Data keys: {list(data.keys()) if isinstance(data, dict) else type(data).__name__}",
                    ""
                )
            )
        else:
            results.append(
                _debug_result(
                    "Project storage row",
                    "WARN",
                    f"No project_storage row found for {storage_id}.",
                    "This is normal for first login, but bad if user already saved before."
                )
            )

    except Exception as e:
        results.append(
            _debug_result(
                "Project storage query",
                "FAIL",
                f"project_storage query failed: {e}",
                "Check project_storage table and RLS policies."
            )
        )

    snapshot_id = st.session_state.get("loaded_snapshot_id")

    if snapshot_id:
        try:
            response = supabase.table("project_snapshots") \
                .select("id, user_id, snapshot_name, data") \
                .eq("id", snapshot_id) \
                .eq("user_id", user_id) \
                .execute()

            if response.data:
                data = response.data[0].get("data", {})
                results.append(
                    _debug_result(
                        "Loaded snapshot row",
                        "PASS",
                        f"Found loaded snapshot: {response.data[0].get('snapshot_name')}. Data keys: {list(data.keys()) if isinstance(data, dict) else type(data).__name__}",
                        ""
                    )
                )
            else:
                results.append(
                    _debug_result(
                        "Loaded snapshot row",
                        "FAIL",
                        f"No project_snapshots row matched id={snapshot_id}, user_id={user_id}.",
                        "Snapshot may belong to a different account, was deleted, or RLS blocked SELECT."
                    )
                )

        except Exception as e:
            results.append(
                _debug_result(
                    "Loaded snapshot query",
                    "FAIL",
                    f"project_snapshots query failed: {e}",
                    "Check project_snapshots table and RLS policies."
                )
            )
    else:
        results.append(
            _debug_result(
                "Loaded snapshot row",
                "WARN",
                "No loaded_snapshot_id in session_state.",
                "Archive Quick Save cannot work until a saved file is linked."
            )
        )

    return results

def run_save_function_debug_checks():
    """
    Non-destructive function availability checks.
    Does not write unless user explicitly clicks write-test button elsewhere.
    """
    results = []

    required_functions = [
        "get_valid_token",
        "_get_authed_snapshot_client",
        "build_app_payload",
        "clean_for_json",
        "repair_projects_state",
        "restore_app_payload",
        "save_data",
        "load_data",
        "save_snapshot",
        "overwrite_current_snapshot",
        "overwrite_snapshot",
        "rename_snapshot",
        "load_snapshots",
        "load_snapshot_data",
        "delete_snapshot",
        "enforce_single_device",
        "is_session_valid",
        "register_device_session",
    ]

    for fn_name in required_functions:
        fn = globals().get(fn_name)

        if callable(fn):
            results.append(
                _debug_result(
                    f"Function exists: {fn_name}",
                    "PASS",
                    "Callable.",
                    ""
                )
            )
        else:
            results.append(
                _debug_result(
                    f"Function exists: {fn_name}",
                    "FAIL",
                    "Missing or not callable.",
                    f"Define {fn_name} before it is used."
                )
            )

    return results

def run_quick_save_write_test():
    """
    Real write test.
    This intentionally writes a debug marker into the loaded snapshot and storage.
    It then reads back both rows to confirm Supabase actually changed.
    """
    results = []

    if not st.session_state.get("logged_in", False):
        results.append(
            _debug_result(
                "Write test prerequisite",
                "FAIL",
                "User is not logged in.",
                "Login first."
            )
        )
        return results

    if not st.session_state.get("storage_loaded", False):
        results.append(
            _debug_result(
                "Write test prerequisite",
                "FAIL",
                "storage_loaded=False, so save_data() will skip.",
                "Run ensure_app_state_loaded() after login first."
            )
        )
        return results

    snapshot_id = st.session_state.get("loaded_snapshot_id")
    if not snapshot_id:
        results.append(
            _debug_result(
                "Write test prerequisite",
                "FAIL",
                "No loaded_snapshot_id.",
                "Load or use Archive > Online Backup > Save for a feasibility study first."
            )
        )
        return results

    token = get_valid_token()
    user_id = _debug_user_id()

    if not token or not user_id:
        results.append(
            _debug_result(
                "Write test prerequisite",
                "FAIL",
                f"token={bool(token)}, user_id={user_id}",
                "Login again."
            )
        )
        return results

    debug_marker = datetime.utcnow().isoformat()

    try:
        init_report_config()
        st.session_state.report_config.setdefault("_debug", {})
        st.session_state.report_config["_debug"]["last_write_test_utc"] = debug_marker
    except Exception as e:
        results.append(
            _debug_result(
                "Prepare write marker",
                "FAIL",
                f"Could not add debug marker: {e}",
                "Check report_config structure."
            )
        )
        return results

    try:
        snapshot_ok = overwrite_current_snapshot()
        storage_ok = save_data()

        results.append(
            _debug_result(
                "overwrite_current_snapshot()",
                "PASS" if snapshot_ok else "FAIL",
                f"Returned {snapshot_ok}",
                "Check error above if False."
            )
        )

        results.append(
            _debug_result(
                "save_data()",
                "PASS" if storage_ok else "FAIL",
                f"Returned {storage_ok}",
                "Check error above if False."
            )
        )

    except Exception as e:
        results.append(
            _debug_result(
                "Quick Save write execution",
                "FAIL",
                f"Write function crashed: {e}",
                "Check traceback and the function that crashed."
            )
        )
        return results

    try:
        supabase.postgrest.auth(token)

        snapshot_response = supabase.table("project_snapshots") \
            .select("data") \
            .eq("id", snapshot_id) \
            .eq("user_id", user_id) \
            .execute()

        if snapshot_response.data:
            snap_data = snapshot_response.data[0].get("data", {})
            snap_marker = (
                snap_data.get("report_config", {})
                .get("_debug", {})
                .get("last_write_test_utc")
                if isinstance(snap_data, dict)
                else None
            )

            if snap_marker == debug_marker:
                results.append(
                    _debug_result(
                        "Snapshot write verification",
                        "PASS",
                        f"Snapshot marker matched: {snap_marker}",
                        ""
                    )
                )
            else:
                results.append(
                    _debug_result(
                        "Snapshot write verification",
                        "FAIL",
                        f"Expected marker {debug_marker}, got {snap_marker}",
                        "Update executed but data did not persist correctly."
                    )
                )
        else:
            results.append(
                _debug_result(
                    "Snapshot write verification",
                    "FAIL",
                    "Could not read snapshot after write.",
                    "Check RLS SELECT policy and snapshot ownership."
                )
            )

    except Exception as e:
        results.append(
            _debug_result(
                "Snapshot write verification",
                "FAIL",
                f"Readback failed: {e}",
                "Check Supabase connection/RLS."
            )
        )

    try:
        storage_response = supabase.table("project_storage") \
            .select("data") \
            .eq("id", f"storage_{user_id}") \
            .execute()

        if storage_response.data:
            storage_data = storage_response.data[0].get("data", {})
            storage_marker = (
                storage_data.get("report_config", {})
                .get("_debug", {})
                .get("last_write_test_utc")
                if isinstance(storage_data, dict)
                else None
            )

            if storage_marker == debug_marker:
                results.append(
                    _debug_result(
                        "Storage write verification",
                        "PASS",
                        f"Storage marker matched: {storage_marker}",
                        ""
                    )
                )
            else:
                results.append(
                    _debug_result(
                        "Storage write verification",
                        "FAIL",
                        f"Expected marker {debug_marker}, got {storage_marker}",
                        "project_storage did not persist the latest save."
                    )
                )
        else:
            results.append(
                _debug_result(
                    "Storage write verification",
                    "FAIL",
                    "Could not read project_storage after write.",
                    "Check RLS SELECT policy and storage row."
                )
            )

    except Exception as e:
        results.append(
            _debug_result(
                "Storage write verification",
                "FAIL",
                f"Readback failed: {e}",
                "Check Supabase connection/RLS."
            )
        )

    return results

def run_calculation_debug_checks():
    results = []

    projects = st.session_state.get("projects", {})

    if not isinstance(projects, dict) or not projects:
        return [
            _debug_result(
                "Projects available",
                "FAIL",
                "No valid projects found.",
                "Fix project state before checking calculations."
            )
        ]

    for proj_id, pdata in projects.items():
        pname = pdata.get("name", proj_id)
        ptype = pdata.get("type", "")

        try:
            gba, gfa, sgfa, budget, rooms = calculate_project_totals(pdata, ptype)

            results.append(
                _debug_result(
                    f"{pname} calculation runs",
                    "PASS",
                    f"GBA={gba:,.2f}, GFA={gfa:,.2f}, SGFA={sgfa:,.2f}, Budget={budget:,.0f}",
                    ""
                )
            )

            if gba < 0 or gfa < 0 or sgfa < 0 or budget < 0:
                results.append(
                    _debug_result(
                        f"{pname} negative value check",
                        "FAIL",
                        "One or more calculated values are negative.",
                        "Check input values and formula signs."
                    )
                )
            else:
                results.append(
                    _debug_result(
                        f"{pname} negative value check",
                        "PASS",
                        "No negative major totals.",
                        ""
                    )
                )

            if gfa > gba:
                results.append(
                    _debug_result(
                        f"{pname} GFA <= GBA",
                        "FAIL",
                        f"GFA {gfa:,.2f} is greater than GBA {gba:,.2f}.",
                        "Check area classification rules."
                    )
                )
            else:
                results.append(
                    _debug_result(
                        f"{pname} GFA <= GBA",
                        "PASS",
                        f"GFA {gfa:,.2f}, GBA {gba:,.2f}.",
                        ""
                    )
                )

            if sgfa > gfa:
                results.append(
                    _debug_result(
                        f"{pname} SGFA <= GFA",
                        "WARN",
                        f"SGFA {sgfa:,.2f} is greater than GFA {gfa:,.2f}.",
                        "This may be valid only if your classification intentionally allows it."
                    )
                )
            else:
                results.append(
                    _debug_result(
                        f"{pname} SGFA <= GFA",
                        "PASS",
                        f"SGFA {sgfa:,.2f}, GFA {gfa:,.2f}.",
                        ""
                    )
                )

        except Exception as e:
            results.append(
                _debug_result(
                    f"{pname} calculation runs",
                    "FAIL",
                    f"Calculation crashed: {e}",
                    "Check calculate_project_totals() and required input keys."
                )
            )

    return results

def render_debug_results(results):
    """Render debug results as dataframe + status summary."""
    if not results:
        st.info("No debug results.")
        return

    df = pd.DataFrame(results)

    status_order = {
        "FAIL": 0,
        "WARN": 1,
        "INFO": 2,
        "PASS": 3
    }

    df["_sort"] = df["Status"].map(status_order).fillna(9)
    df = df.sort_values(["_sort", "Check"]).drop(columns=["_sort"])

    fail_count = int((df["Status"] == "FAIL").sum())
    warn_count = int((df["Status"] == "WARN").sum())
    pass_count = int((df["Status"] == "PASS").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("FAIL", fail_count)
    c2.metric("WARN", warn_count)
    c3.metric("PASS", pass_count)

    st.dataframe(df, width="stretch", hide_index=True)

    if fail_count > 0:
        st.error("Debugger found blocking issues.")
    elif warn_count > 0:
        st.warning("Debugger found warnings. App may still work, but check them.")
    else:
        st.success("No obvious issues found.")

def debug_project_data_snapshot(label="Project Data Snapshot"):
    """
    Shows current project saved data keys and selected important values.
    Use this to confirm whether data exists in curr_proj["data"].
    """
    curr_id, curr_proj = get_current_project()
    d = curr_proj.get("data", {})

    important_keys = [
        "area_table_data",
        "m_gba",
        "m_gfa",
        "m_sgfa",
        "m_nfa",
        "area_door_table_data",
        "area_door_wood_calc",
        "area_door_steel_calc",
        "area_door_glass_calc",
        "smart_custom_costs",
        "external_custom_costs",
    ]

    rows = []

    for key in important_keys:
        value = d.get(key, None)

        if isinstance(value, list):
            summary = f"list with {len(value)} rows"
        elif isinstance(value, dict):
            summary = f"dict with {len(value)} keys"
        else:
            summary = value

        rows.append({
            "Key": key,
            "Exists": key in d,
            "Summary": summary,
        })

    st.subheader(label)
    st.caption(f"Current project: {curr_id} | {curr_proj.get('name', '')}")
    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)

    with st.expander("Raw current project data"):
        st.json(d)


def debug_payload_current_project_snapshot():
    """
    Confirms whether build_app_payload() contains the current project data.
    """
    payload = build_app_payload()

    curr_id = payload.get("current_proj_id")
    projects = payload.get("projects", {})
    curr_proj = projects.get(curr_id, {})
    d = curr_proj.get("data", {})

    st.subheader("Payload Current Project Check")
    st.caption(f"Payload current project: {curr_id}")

    rows = []

    for key in [
        "area_table_data",
        "m_gba",
        "m_gfa",
        "m_sgfa",
        "m_nfa",
        "area_door_table_data",
        "smart_custom_costs",
    ]:
        value = d.get(key)

        if isinstance(value, list):
            summary = f"list with {len(value)} rows"
        elif isinstance(value, dict):
            summary = f"dict with {len(value)} keys"
        else:
            summary = value

        rows.append({
            "Key": key,
            "Exists in Payload": key in d,
            "Summary": summary,
        })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)

    with st.expander("Raw payload current project data"):
        st.json(d)


def _debug_safe_preview(value, max_chars=120):
    if isinstance(value, pd.DataFrame):
        preview = f"DataFrame rows={len(value)}, cols={len(value.columns)}"
    elif isinstance(value, list):
        preview = f"list with {len(value)} rows"
    elif isinstance(value, dict):
        preview = f"dict with {len(value)} keys"
    else:
        preview = str(value)

    if len(preview) > max_chars:
        preview = preview[:max_chars] + "..."

    return preview


def _debug_sample_keys(value, limit=12):
    if isinstance(value, dict):
        return ", ".join(list(value.keys())[:limit])
    return ""


def render_debug_component_overview():
    projects = st.session_state.get("projects", {})

    if not isinstance(projects, dict) or not projects:
        st.info("No valid projects found.")
        return

    rows = []

    for project_id, pdata in projects.items():
        pdata = pdata if isinstance(pdata, dict) else {}
        data = pdata.get("data", {})

        rows.append({
            "project_id": project_id,
            "name": pdata.get("name", ""),
            "type": pdata.get("type", ""),
            "data_key_count": len(data) if isinstance(data, dict) else None,
            "has_inputs": isinstance(pdata.get("inputs"), dict),
            "has_calculations": isinstance(pdata.get("calculations"), dict),
            "has_ui": isinstance(pdata.get("ui"), dict),
        })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)


def render_debug_current_project_buckets():
    curr_id, curr_proj = get_current_project()

    st.caption(
        f"Current project: {curr_id} | "
        f"{curr_proj.get('name', '')} ({curr_proj.get('type', '')})"
    )
    st.write("Top-level keys:", list(curr_proj.keys()))

    rows = []

    for bucket in ["data", "inputs", "calculations", "ui"]:
        value = curr_proj.get(bucket)
        rows.append({
            "bucket": bucket,
            "type": type(value).__name__,
            "key_count": len(value) if isinstance(value, dict) else None,
            "sample_keys": _debug_sample_keys(value),
        })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)


def render_debug_area_source_check():
    curr_id, curr_proj = get_current_project()
    data = curr_proj.get("data", {}) if isinstance(curr_proj.get("data"), dict) else {}
    calculations = (
        curr_proj.get("calculations", {})
        if isinstance(curr_proj.get("calculations"), dict)
        else {}
    )
    area_calc = calculations.get("area", {}) if isinstance(calculations.get("area"), dict) else {}
    clean_totals = area_calc.get("totals", {}) if isinstance(area_calc.get("totals"), dict) else {}
    table_totals = calculate_area_totals_from_table(data.get("area_table_data", []))

    rows = []

    for metric, legacy_key in [
        ("gba", "m_gba"),
        ("gfa", "m_gfa"),
        ("sgfa", "m_sgfa"),
        ("nfa", "m_nfa"),
    ]:
        legacy_value = _safe_float(data.get(legacy_key, 0.0))
        clean_value = _safe_float(clean_totals.get(metric, 0.0))
        table_value = _safe_float(table_totals.get(metric, 0.0))

        rows.append({
            "metric": metric,
            "legacy_data": legacy_value,
            "clean_calculation": clean_value,
            "recalculated_from_table": table_value,
            "legacy_vs_clean_diff": legacy_value - clean_value,
            "legacy_vs_table_diff": legacy_value - table_value,
        })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)


def render_debug_widget_cache_keys():
    curr_id = str(st.session_state.get("current_proj_id", ""))
    rows = []

    for key in sorted(list(st.session_state.keys()), key=lambda x: str(x)):
        key_str = str(key)

        if (
            (curr_id and curr_id in key_str)
            or any(key_str.startswith(prefix) for prefix in UI_CACHE_PREFIXES)
        ):
            value = st.session_state.get(key)
            rows.append({
                "key": key_str,
                "type": type(value).__name__,
                "preview": _debug_safe_preview(value),
            })

    if rows:
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    else:
        st.info("No current-project widget/cache keys found.")


def render_debug_cost_key_summary():
    _, curr_proj = get_current_project()
    data = curr_proj.get("data", {}) if isinstance(curr_proj.get("data"), dict) else {}

    rows = []

    for prefix in ["m_", "u_", "r_", "sc_"]:
        keys = sorted([key for key in data.keys() if str(key).startswith(prefix)])
        rows.append({
            "prefix": prefix,
            "count": len(keys),
            "sample_keys": ", ".join(keys[:15]),
        })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)

    smart_custom = data.get("smart_custom_costs")
    info = {
        "grand_total_project": data.get("grand_total_project"),
        "smart_custom_costs_exists": "smart_custom_costs" in data,
        "smart_custom_costs_type": type(smart_custom).__name__,
        "smart_custom_costs_count": len(smart_custom) if isinstance(smart_custom, list) else None,
    }
    st.json(info)


def render_debug_summary_calculations_check():
    value = st.session_state.get("summary_calculations")
    exists = "summary_calculations" in st.session_state

    info = {
        "exists": exists,
        "type": type(value).__name__ if exists else None,
        "key_count": len(value) if isinstance(value, dict) else None,
        "row_count": len(value) if isinstance(value, (list, pd.DataFrame)) else None,
        "summary": _debug_safe_preview(value) if exists else None,
    }

    st.json(info)


def render_app_debugger():
    """
    Streamlit UI page/panel for debugging the full app save/load system.
    Add this as a sidebar navigation page or call it temporarily from main_app().
    """
    st.title("App Debugger")
    st.caption("Checks app state, Supabase auth, payload integrity, snapshot linkage, and cloud save verification.")

    st.warning(
        "The write test intentionally writes a small _debug marker into report_config. "
        "This is safe, but it changes the saved data timestamp/marker."
    )

    tab_local, tab_project_data, tab_payload, tab_supabase, tab_functions, tab_write, tab_calculations = st.tabs([
        "Local State",
        "Project Data",
        "Payload",
        "Supabase",
        "Functions",
        "Write Test",
        "Calculation Checks"
    ])

    with tab_local:
        st.subheader("Local Streamlit State Checks")
        if st.button("Run Local State Debug", key="debug_run_local"):
            results = run_local_state_debug_checks()
            render_debug_results(results)

        with st.expander("Raw key session summary"):
            safe_summary = {
                "logged_in": st.session_state.get("logged_in"),
                "storage_loaded": st.session_state.get("storage_loaded"),
                "user_id": _debug_user_id(),
                "access_token": _debug_mask(st.session_state.get("access_token")),
                "session_fingerprint": _debug_mask(st.session_state.get("session_fingerprint")),
                "current_study_name": st.session_state.get("current_study_name"),
                "loaded_snapshot_id": st.session_state.get("loaded_snapshot_id"),
                "loaded_snapshot_name": st.session_state.get("loaded_snapshot_name"),
                "current_proj_id": st.session_state.get("current_proj_id"),
                "project_count": len(st.session_state.get("projects", {})) if isinstance(st.session_state.get("projects"), dict) else None,
            }
            st.json(safe_summary)

        with st.expander("Component Overview", expanded=True):
            render_debug_component_overview()

        with st.expander("Summary Calculations Check"):
            render_debug_summary_calculations_check()

    with tab_project_data:
        st.subheader("Current Project Saved Data")
        st.write(
            "This checks the actual saved project data bucket: "
            "st.session_state.projects[current_proj_id]['data']"
        )

        if st.button("Check Current Project Data", key="debug_check_project_data"):
            debug_project_data_snapshot()

        with st.expander("Current Project Buckets", expanded=True):
            render_debug_current_project_buckets()

        with st.expander("Area Source Check", expanded=True):
            render_debug_area_source_check()

        with st.expander("Cost Key Summary", expanded=True):
            render_debug_cost_key_summary()

        with st.expander("Widget Cache Keys For Current Project"):
            render_debug_widget_cache_keys()

        with st.expander("Project data key map"):
            st.dataframe(get_project_data_map(), width="stretch", hide_index=True)

    with tab_payload:
        st.subheader("Payload / JSON Safety Checks")
        if st.button("Run Payload Debug", key="debug_run_payload"):
            results, payload = run_payload_debug_checks()
            render_debug_results(results)

            if payload is not None:
                with st.expander("Payload preview"):
                    st.json(payload)

        st.divider()

        if st.button("Check Current Project Inside Payload", key="debug_check_payload_current_project"):
            debug_payload_current_project_snapshot()

    with tab_supabase:
        st.subheader("Supabase Read/Auth Checks")
        if st.button("Run Supabase Debug", key="debug_run_supabase"):
            results = run_supabase_debug_checks()
            render_debug_results(results)

    with tab_functions:
        st.subheader("Function Availability Checks")
        if st.button("Run Function Debug", key="debug_run_functions"):
            results = run_save_function_debug_checks()
            render_debug_results(results)

    with tab_write:
        st.subheader("Real Quick Save Write Verification")
        st.write(
            "This runs the actual save path, then reads Supabase back to confirm "
            "whether the marker was actually written."
        )

        confirm = st.checkbox(
            "I understand this will write a small debug marker to the active saved file.",
            key="debug_write_confirm"
        )

        if st.button(
            "Run Real Write Test",
            key="debug_run_write_test",
            type="primary",
            disabled=not confirm
        ):
            results = run_quick_save_write_test()
            render_debug_results(results)

    with tab_calculations:
        st.subheader("Calculation Checks")
        if st.button("Run Calculation Debug", key="debug_calculations"):
            render_debug_results(run_calculation_debug_checks())

    st.divider()
    st.subheader("Debug Log")
    render_debug_log()

def debug_call(label, func, *args, **kwargs):
    """
    Safely run a function and show detailed error if it crashes.
    Useful for callbacks.
    """
    try:
        return func(*args, **kwargs)
    except Exception as e:
        st.error(f"{label} crashed: {e}")

        with st.expander(f"Debug details: {label}"):
            st.write("Function:", getattr(func, "__name__", str(func)))
            st.write("Args:", args)
            st.write("Kwargs:", kwargs)
            st.write("Current user_id:", _debug_user_id())
            st.write("Logged in:", st.session_state.get("logged_in"))
            st.write("Storage loaded:", st.session_state.get("storage_loaded"))
            st.write("Current project:", st.session_state.get("current_proj_id"))
            st.write("Loaded snapshot:", st.session_state.get("loaded_snapshot_id"))

        return None
#enddebugcode

def ensure_app_state_loaded():
    """
    Load cloud state only after login.
    This prevents Streamlit from creating default projects before user authentication.
    """
    if st.session_state.get("storage_loaded", False):
        repair_projects_state(save=False)
        return

    stored_data = load_data()

    if stored_data:
        restore_app_payload(stored_data)
        repair_projects_state(save=False)
        st.session_state.storage_loaded = True
        return

    # First-time user fallback only.
    # Do not immediately save default data unless explicitly creating a new file.
    restore_app_payload({
        "app_version": APP_VERSION,
        "projects": make_default_projects(),
        "current_proj_id": "proj_1",
        "proj_counter": 1,
        "report_config": copy.deepcopy(DEFAULT_REPORT_CONFIG)
    })

    repair_projects_state(save=False)
    st.session_state.storage_loaded = True

def n2w(amount):
    try:
        amount = _safe_float(amount)
        if amount >= 1_000_000_000_000: # Trillion
            return f"{amount / 1_000_000_000_000:,.2f} Triliun"
        elif amount >= 1_000_000_000: # Billion (Miliar)
            return f"{amount / 1_000_000_000:,.2f} Miliar"
        elif amount >= 1_000_000: # Million (Juta)
            return f"{amount / 1_000_000:,.2f} Juta"
        else:
            return f"{amount:,.0f}"
    except:
        return "0"
#endregion

def get_default_earthwork_detail_rows():
    return [
        {"code": "1.2.1", "description": "Cut Fill", "unit": "m2", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "1.2.2", "description": "Dewatering", "unit": "ls", "quantity": 1.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "1.2.3", "description": "Soil Improvement", "unit": "m2", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "1.2.4", "description": "Shoring System", "unit": "ls", "quantity": 1.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "1.2.5", "description": "Others", "unit": "ls", "quantity": 1.0, "unit_price": 0.0, "amount": 0.0},
    ]

def clean_earthwork_detail_rows(rows):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_earthwork_detail_rows()

    cleaned_rows = []

    default_rows = get_default_earthwork_detail_rows()

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}
        unit = str(row.get("unit", default_row["unit"]) or default_row["unit"]).strip().lower()

        raw_quantity = row.get("quantity", default_row["quantity"])
        if unit == "ls" and (raw_quantity is None or str(raw_quantity).strip() == ""):
            quantity = 1.0
        else:
            quantity = _safe_float(raw_quantity, default_row["quantity"])

        unit_price = _safe_float(row.get("unit_price", 0.0))
        amount = quantity * unit_price

        cleaned_rows.append({
            "code": str(row.get("code", default_row["code"]) or default_row["code"]).strip(),
            "description": str(row.get("description", default_row["description"]) or default_row["description"]).strip(),
            "unit": unit,
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_earthwork_detail(rows, gba):
    cleaned_rows = clean_earthwork_detail_rows(rows)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gba = _safe_float(gba)
    derived_unit_price = detail_total / gba if gba > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_default_foundation_detail_rows():
    return [
        {"code": "1", "description": "Supply Tiang Pancang", "unit": "m'", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2", "description": "Install Tiang Pancang", "unit": "m'", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def clean_foundation_detail_rows(rows):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_foundation_detail_rows()

    cleaned_rows = []
    default_rows = get_default_foundation_detail_rows()

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}
        quantity = _safe_float(row.get("quantity", 0.0))
        unit_price = _safe_float(row.get("unit_price", 0.0))
        amount = quantity * unit_price

        cleaned_rows.append({
            "code": default_row["code"],
            "description": default_row["description"],
            "unit": default_row["unit"],
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_foundation_detail(rows, gba):
    cleaned_rows = clean_foundation_detail_rows(rows)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gba = _safe_float(gba)
    derived_unit_price = detail_total / gba if gba > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_default_ffe_detail_rows():
    return [
        {"code": "1", "description": "Seater & Chair", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2", "description": "Beds & Linen", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "3", "description": "Kitchen Cabinet, Drawer", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "4", "description": "Electronic: TV 32\", Minibar, Kettle, SDB", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "5", "description": "Housewares", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "6", "description": "Stove with 2 burner + Hoods", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "7", "description": "Microwave, Refrigerator, Washing Machine", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "8", "description": "Others: Artworks", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "9", "description": "Misc (Linen/Gym)", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def clean_ffe_detail_rows(rows):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_ffe_detail_rows()

    cleaned_rows = []
    default_rows = get_default_ffe_detail_rows()

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}
        quantity = _safe_float(row.get("quantity", 0.0))
        unit_price = _safe_float(row.get("unit_price", 0.0))
        amount = quantity * unit_price

        cleaned_rows.append({
            "code": default_row["code"],
            "description": default_row["description"],
            "unit": default_row["unit"],
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_ffe_detail(rows, rooms):
    cleaned_rows = clean_ffe_detail_rows(rows)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    rooms = _safe_float(rooms)
    derived_unit_price = detail_total / rooms if rooms > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_default_mep_detail_rows():
    return [
        {"code": "1", "description": "STP & WTP system", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2", "description": "Plumbing Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "3", "description": "Fire Fighting & Protection Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "4", "description": "Electrical Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "5", "description": "Genset Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "6", "description": "MVAC Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "7", "description": "Lifts / Escalator / Travelators", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "8", "description": "Electronic Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "9", "description": "System Data", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "10", "description": "Gas Installation", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "11", "description": "Special Lighting", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "12", "description": "SBO : Pompa Pemadam", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13", "description": "SBO : Chillers, AHU, FCU", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "14", "description": "SBO : Lighting Fixtures", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "15", "description": "SBO : Genset", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "16", "description": "Heat Pump", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "17", "description": "Cooling Towers", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "18", "description": "SBO : Water Heater", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "19", "description": "Swimming Pool", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "20", "description": "Deep Well", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "21", "description": "Check Point", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "22", "description": "SBO : AC Unit", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "23", "description": "SBO : AC VRV / Split", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "24", "description": "SBO : Unit Fan", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "25", "description": "Other - Pek. M.E.P.", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def clean_mep_detail_rows(rows):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_mep_detail_rows()

    cleaned_rows = []
    default_rows = get_default_mep_detail_rows()

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}
        quantity = _safe_float(row.get("quantity", 0.0))
        unit_price = _safe_float(row.get("unit_price", 0.0))
        amount = quantity * unit_price

        cleaned_rows.append({
            "code": default_row["code"],
            "description": default_row["description"],
            "unit": default_row["unit"],
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_mep_detail(rows, gba):
    cleaned_rows = clean_mep_detail_rows(rows)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gba = _safe_float(gba)
    derived_unit_price = detail_total / gba if gba > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_default_utility_detail_rows():
    return [
        {"code": "1", "description": "Listrik", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2", "description": "Air Bersih", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "3", "description": "Telkom", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "4", "description": "Gas", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "5", "description": "Air Limbah", "unit": "unit", "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def clean_utility_detail_rows(rows):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_utility_detail_rows()

    cleaned_rows = []
    default_rows = get_default_utility_detail_rows()

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}
        quantity = _safe_float(row.get("quantity", 0.0))
        unit_price = _safe_float(row.get("unit_price", 0.0))
        amount = quantity * unit_price

        cleaned_rows.append({
            "code": default_row["code"],
            "description": default_row["description"],
            "unit": default_row["unit"],
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_utility_detail(rows, gba):
    cleaned_rows = clean_utility_detail_rows(rows)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gba = _safe_float(gba)
    derived_unit_price = detail_total / gba if gba > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_default_consultancy_detail_rows():
    return [
        {"code": "1", "description": "Master Plan", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2", "description": "Feasibility Study", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "3", "description": "Quantity Surveyor", "unit": "bln", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "4", "description": "Project Management Fee", "unit": "bln", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "5", "description": "Architectural Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "6", "description": "Structural Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "7", "description": "M.E.P Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "8", "description": "Interior Designer", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "9", "description": "Landscaping Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "10", "description": "Soil Investigation Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "11", "description": "Signage Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "12", "description": "Special Lighting", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13", "description": "Infrastructure Consultant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "14", "description": "Amdal", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "15", "description": "Traffic Analysis", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "16", "description": "Technical Assistant", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "17", "description": "Topografi", "unit": "m2", "manual_quantity": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def get_consultancy_detail_base_values(data, area_df=None):
    data = data if isinstance(data, dict) else {}

    gfa = _safe_float(data.get("m_gfa", 0.0))
    if gfa <= 0 and area_df is not None and "GFA" in getattr(area_df, "columns", []):
        gfa = _safe_float(area_df["GFA"].sum())

    koridor_lobby = _safe_float(data.get("area_lobby_interior_calc", 0.0))
    if koridor_lobby <= 0:
        koridor_lobby = _safe_float(data.get("m_lobby", 0.0))
    if koridor_lobby <= 0 and area_df is not None and "Koridor/Lobby" in getattr(area_df, "columns", []):
        koridor_lobby = _safe_float(area_df["Koridor/Lobby"].sum())

    landscape_qty = _safe_float(data.get("area_landscape_qty_calc", 0.0))
    if landscape_qty <= 0:
        landscape_qty = _safe_float(data.get("m_land_m2", 0.0))

    return {
        "gfa": gfa,
        "koridor_lobby": koridor_lobby,
        "landscape_qty": landscape_qty,
    }

def clean_consultancy_detail_rows(rows, base_values):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_consultancy_detail_rows()

    base_values = base_values if isinstance(base_values, dict) else {}
    default_rows = get_default_consultancy_detail_rows()
    cleaned_rows = []

    gfa = _safe_float(base_values.get("gfa", 0.0))
    koridor_lobby = _safe_float(base_values.get("koridor_lobby", 0.0))
    landscape_qty = _safe_float(base_values.get("landscape_qty", 0.0))

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}

        code = default_row["code"]
        manual_quantity = _safe_float(row.get("manual_quantity", row.get("quantity", 0.0)))
        unit_price = _safe_float(row.get("unit_price", 0.0))

        if code in ["3", "4"]:
            quantity = manual_quantity
        elif code == "8":
            quantity = koridor_lobby
        elif code == "9":
            quantity = landscape_qty
        else:
            quantity = gfa

        amount = quantity * unit_price

        cleaned_rows.append({
            "code": code,
            "description": default_row["description"],
            "unit": default_row["unit"],
            "manual_quantity": manual_quantity if code in ["3", "4"] else 0.0,
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_consultancy_detail(rows, base_values):
    base_values = base_values if isinstance(base_values, dict) else {}
    cleaned_rows = clean_consultancy_detail_rows(rows, base_values)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gfa = _safe_float(base_values.get("gfa", 0.0))
    consultant_subtotal_excl_qs_pm = sum(
        _safe_float(row.get("amount", 0.0))
        for row in cleaned_rows
        if str(row.get("code", "")) not in ["3", "4"]
    )
    derived_unit_price = consultant_subtotal_excl_qs_pm / gfa if gfa > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_consultancy_detail_outputs(rows, base_values):
    base_values = base_values if isinstance(base_values, dict) else {}
    cleaned_rows, detail_total, consultant_rate_excl_qs_pm = calculate_consultancy_detail(
        rows,
        base_values,
    )

    qs_row = next((row for row in cleaned_rows if str(row.get("code", "")) == "3"), {})
    pm_row = next((row for row in cleaned_rows if str(row.get("code", "")) == "4"), {})
    consultant_subtotal_excl_qs_pm = sum(
        _safe_float(row.get("amount", 0.0))
        for row in cleaned_rows
        if str(row.get("code", "")) not in ["3", "4"]
    )

    return {
        "consultancy_detail_rows": cleaned_rows,
        "consultancy_detail_total": detail_total,
        "consultancy_derived_unit_price": consultant_rate_excl_qs_pm,
        "qs_duration_from_consultancy": _safe_float(qs_row.get("quantity", 0.0)),
        "qs_rate_from_consultancy": _safe_float(qs_row.get("unit_price", 0.0)),
        "pm_duration_from_consultancy": _safe_float(pm_row.get("quantity", 0.0)),
        "pm_rate_from_consultancy": _safe_float(pm_row.get("unit_price", 0.0)),
        "consultant_subtotal_excl_qs_pm": consultant_subtotal_excl_qs_pm,
        "consultant_rate_excl_qs_pm": consultant_rate_excl_qs_pm,
    }

def store_consultancy_detail_outputs(data, outputs):
    data["consultancy_detail_rows"] = outputs["consultancy_detail_rows"]
    data["consultancy_detail_total"] = outputs["consultancy_detail_total"]
    data["consultancy_derived_unit_price"] = outputs["consultancy_derived_unit_price"]
    data["qs_duration_from_consultancy"] = outputs["qs_duration_from_consultancy"]
    data["qs_rate_from_consultancy"] = outputs["qs_rate_from_consultancy"]
    data["pm_duration_from_consultancy"] = outputs["pm_duration_from_consultancy"]
    data["pm_rate_from_consultancy"] = outputs["pm_rate_from_consultancy"]
    data["consultant_subtotal_excl_qs_pm"] = outputs["consultant_subtotal_excl_qs_pm"]
    data["consultant_rate_excl_qs_pm"] = outputs["consultant_rate_excl_qs_pm"]

def get_default_architectural_detail_rows():
    return [
        {"code": "1", "description": "Basic Finishes Work", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2.1", "description": "Aluminium Facade / Window Wall", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2.2", "description": "Kisi2 Facade / Double Skin", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2.3", "description": "Precast Facade", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "3", "description": "Pintu Kaca Dalam Ruangan", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "4", "description": "Railing Balkon", "unit": "m'", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "5", "description": "Pintu Kayu", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "6", "description": "Pintu Besi", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "7", "description": "Shower Screen", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "8", "description": "Marble / Door Jamb Lift", "unit": "m'", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "9", "description": "Interior - Main Lobby & Typical Lobby", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "10", "description": "Signage / Fixtures", "unit": "ls", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "11", "description": "Gondola", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "12", "description": "Roof - Skylight", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13.1", "description": "Sanitary Fittings - T. Wanita", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13.2", "description": "Sanitary Fittings - T. Pria", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13.3", "description": "Sanitary Fittings - T. Disable", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13.4", "description": "Sanitary Fittings - Musholla", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "13.5", "description": "Sanitary Fittings - Toilet Unit", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "14", "description": "Kitchen Equipment", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "15.1", "description": "Ironmongeries - Pintu Kayu", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "15.2", "description": "Ironmongeries - Pintu Besi", "unit": "unit", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "16.1", "description": "Keramik & HT", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "16.2", "description": "Marmer", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "16.3", "description": "Vinyl", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "17", "description": "Carpet", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "18", "description": "Kaca", "unit": "m2", "factor": 0.0, "overlap": 0.0, "waste": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def get_architectural_detail_base_values(data, area_df=None):
    data = data if isinstance(data, dict) else {}

    def data_value(*keys):
        for key in keys:
            value = _safe_float(data.get(key, 0.0))
            if value > 0:
                return value
        return 0.0

    gfa = data_value("m_gfa")
    if gfa <= 0 and area_df is not None and "GFA" in getattr(area_df, "columns", []):
        gfa = _safe_float(area_df["GFA"].sum())

    return {
        "gfa": gfa,
        "facade": data_value("m_facade", "area_facade_calc"),
        "rooms": data_value("m_rooms", "area_rooms_calc", "area_typical_units_total_calc"),
        "glass_door": data_value("m_door_g", "area_door_glass_calc"),
        "wooden_door": data_value("m_door_w", "area_door_wood_calc"),
        "steel_door": data_value("m_door_s", "area_door_steel_calc"),
        "lobby": data_value("m_lobby", "area_lobby_interior_calc"),
    }

def clean_architectural_detail_rows(rows, base_values):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_architectural_detail_rows()

    base_values = base_values if isinstance(base_values, dict) else {}
    default_rows = get_default_architectural_detail_rows()
    cleaned_rows = []

    gfa = _safe_float(base_values.get("gfa", 0.0))
    facade = _safe_float(base_values.get("facade", 0.0))
    rooms = _safe_float(base_values.get("rooms", 0.0))
    glass_door = _safe_float(base_values.get("glass_door", 0.0))
    wooden_door = _safe_float(base_values.get("wooden_door", 0.0))
    steel_door = _safe_float(base_values.get("steel_door", 0.0))
    lobby = _safe_float(base_values.get("lobby", 0.0))

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}

        code = default_row["code"]
        factor = _safe_float(row.get("factor", 0.0))
        overlap = _safe_float(row.get("overlap", 0.0))
        waste = _safe_float(row.get("waste", 0.0))
        manual_quantity = _safe_float(row.get("quantity", 0.0))
        unit_price = _safe_float(row.get("unit_price", 0.0))

        if code == "1":
            quantity = gfa
        elif code in ["2.1", "2.2", "2.3"]:
            quantity = facade * (factor / 100.0)
        elif code == "3":
            quantity = glass_door
        elif code == "4":
            quantity = rooms * factor
        elif code == "5":
            quantity = wooden_door
        elif code == "6":
            quantity = steel_door
        elif code in ["7", "8", "10", "11", "12", "13.1", "13.2", "13.3", "13.4", "17", "18"]:
            quantity = manual_quantity
        elif code == "9":
            quantity = lobby
        elif code == "13.5":
            quantity = rooms * factor
        elif code == "14":
            quantity = rooms
        elif code == "15.1":
            quantity = wooden_door
        elif code == "15.2":
            quantity = steel_door
        elif code in ["16.1", "16.2", "16.3"]:
            quantity = gfa * (factor / 100.0) * overlap * waste
        else:
            quantity = manual_quantity

        amount = quantity * unit_price

        cleaned_rows.append({
            "code": code,
            "description": default_row["description"],
            "unit": default_row["unit"],
            "factor": factor,
            "overlap": overlap,
            "waste": waste,
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_architectural_detail(rows, base_values):
    base_values = base_values if isinstance(base_values, dict) else {}
    cleaned_rows = clean_architectural_detail_rows(rows, base_values)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gfa = _safe_float(base_values.get("gfa", 0.0))
    derived_unit_price = detail_total / gfa if gfa > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_default_structural_detail_rows():
    return [
        {"code": "1", "description": "Sub/Superstructure", "unit": "m3", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "2", "description": "Bekisting", "unit": "m2", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "3", "description": "Pembesian", "unit": "kg", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "4", "description": "Readymix Concrete", "unit": "m3", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "5", "description": "Rebar", "unit": "kg", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "6", "description": "Prestress Works", "unit": "ls", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "7", "description": "Steelworks", "unit": "kg", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
        {"code": "8", "description": "Others", "unit": "m3", "ratio": 0.0, "waste_factor": 0.0, "quantity": 0.0, "unit_price": 0.0, "amount": 0.0},
    ]

def clean_structural_detail_rows(rows, gba):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")

    if not isinstance(rows, list) or len(rows) == 0:
        rows = get_default_structural_detail_rows()

    gba = _safe_float(gba)
    default_rows = get_default_structural_detail_rows()
    cleaned_rows = []
    sub_super_qty = 0.0

    for idx, default_row in enumerate(default_rows):
        row = rows[idx] if idx < len(rows) else {}
        row = row if isinstance(row, dict) else {}

        ratio = _safe_float(row.get("ratio", default_row["ratio"]))
        waste_factor = _safe_float(row.get("waste_factor", default_row["waste_factor"]))
        unit_price = _safe_float(row.get("unit_price", 0.0))
        description = default_row["description"]

        if idx == 0:
            quantity = gba * ratio
            sub_super_qty = quantity
        elif idx in [1, 2, 3]:
            quantity = sub_super_qty * ratio
        elif idx == 4:
            quantity = sub_super_qty * ratio * waste_factor
        elif idx in [5, 6]:
            quantity = _safe_float(row.get("quantity", 0.0))
        else:
            quantity = sub_super_qty

        amount = quantity * unit_price

        cleaned_rows.append({
            "code": default_row["code"],
            "description": description,
            "unit": default_row["unit"],
            "ratio": ratio,
            "waste_factor": waste_factor,
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    return cleaned_rows

def calculate_structural_detail(rows, gba):
    cleaned_rows = clean_structural_detail_rows(rows, gba)
    detail_total = sum(_safe_float(row.get("amount", 0.0)) for row in cleaned_rows)
    gba = _safe_float(gba)
    derived_unit_price = detail_total / gba if gba > 0 else 0.0

    return cleaned_rows, detail_total, derived_unit_price

def get_earthwork_price_difference_status(curr_proj, gba):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("earthwork_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_earth", 0.0))
    detail_total = _safe_float(data.get("earthwork_detail_total", 0.0))
    gba = _safe_float(gba)

    current_total = gba * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = current_total - detail_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_foundation_price_difference_status(curr_proj, gba):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("foundation_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_found", 0.0))
    detail_total = _safe_float(data.get("foundation_detail_total", 0.0))
    gba = _safe_float(gba)

    current_total = gba * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_ffe_price_difference_status(curr_proj, rooms):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("ffe_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_ffe", 0.0))
    detail_total = _safe_float(data.get("ffe_detail_total", 0.0))
    rooms = _safe_float(rooms)

    current_total = rooms * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_mep_price_difference_status(curr_proj, gba):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("mep_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_mep", 0.0))
    detail_total = _safe_float(data.get("mep_detail_total", 0.0))
    gba = _safe_float(gba)

    current_total = gba * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_utility_price_difference_status(curr_proj, gba):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("utility_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_util", 0.0))
    detail_total = _safe_float(data.get("utility_detail_total", 0.0))
    gba = _safe_float(gba)

    current_total = gba * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_consultancy_price_difference_status(curr_proj, gfa):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("consultancy_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("sc_cons", 0.0))
    detail_total = _safe_float(data.get("consultant_subtotal_excl_qs_pm", 0.0))
    if detail_total <= 0:
        detail_total = _safe_float(data.get("consultancy_detail_total", 0.0))
    gfa = _safe_float(gfa)

    current_total = gfa * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_architectural_price_difference_status(curr_proj, gfa):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("architectural_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_arch", 0.0))
    detail_total = _safe_float(data.get("architectural_detail_total", 0.0))
    gfa = _safe_float(gfa)

    current_total = gfa * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def get_structural_price_difference_status(curr_proj, gba):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    derived_rate = _safe_float(data.get("structural_derived_unit_price", 0.0))
    current_rate = _safe_float(data.get("u_struc", 0.0))
    detail_total = _safe_float(data.get("structural_detail_total", 0.0))
    gba = _safe_float(gba)

    current_total = gba * current_rate
    rate_difference = derived_rate - current_rate
    total_difference = detail_total - current_total

    has_difference = (
        detail_total > 0
        and derived_rate > 0
        and abs(rate_difference) > 1
    )

    return {
        "has_difference": has_difference,
        "derived_rate": derived_rate,
        "current_rate": current_rate,
        "rate_difference": rate_difference,
        "detail_total": detail_total,
        "current_total": current_total,
        "total_difference": total_difference,
    }

def build_detail_rate_review_rows(curr_proj, gba, gfa=None, rooms=None):
    data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
    gba = _safe_float(gba)
    gfa = _safe_float(gfa if gfa is not None else gba)
    rooms = _safe_float(rooms if rooms is not None else data.get("m_rooms", 0.0))

    review_specs = [
        {
            "Section": "Earthworks",
            "Cost Key": "u_earth",
            "Basis": "GBA",
            "basis_value": gba,
            "current_rate": _safe_float(data.get("u_earth", 0.0)),
            "derived_rate": _safe_float(data.get("earthwork_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("earthwork_detail_total", 0.0)),
        },
        {
            "Section": "Foundation",
            "Cost Key": "u_found",
            "Basis": "GBA",
            "basis_value": gba,
            "current_rate": _safe_float(data.get("u_found", 0.0)),
            "derived_rate": _safe_float(data.get("foundation_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("foundation_detail_total", 0.0)),
        },
        {
            "Section": "Structural",
            "Cost Key": "u_struc",
            "Basis": "GBA",
            "basis_value": gba,
            "current_rate": _safe_float(data.get("u_struc", 0.0)),
            "derived_rate": _safe_float(data.get("structural_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("structural_detail_total", 0.0)),
        },
        {
            "Section": "Architectural",
            "Cost Key": "u_arch",
            "Basis": "GFA",
            "basis_value": gfa,
            "current_rate": _safe_float(data.get("u_arch", 0.0)),
            "derived_rate": _safe_float(data.get("architectural_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("architectural_detail_total", 0.0)),
        },
        {
            "Section": "Consultancy",
            "Cost Key": "sc_cons",
            "Basis": "GFA",
            "basis_value": gfa,
            "current_rate": _safe_float(data.get("sc_cons", 0.0)),
            "derived_rate": _safe_float(data.get("consultancy_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("consultant_subtotal_excl_qs_pm", data.get("consultancy_detail_total", 0.0))),
        },
        {
            "Section": "FF&E",
            "Cost Key": "u_ffe",
            "Basis": "Rooms",
            "basis_value": rooms,
            "current_rate": _safe_float(data.get("u_ffe", 0.0)),
            "derived_rate": _safe_float(data.get("ffe_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("ffe_detail_total", 0.0)),
        },
        {
            "Section": "MEP",
            "Cost Key": "u_mep",
            "Basis": "GBA",
            "basis_value": gba,
            "current_rate": _safe_float(data.get("u_mep", 0.0)),
            "derived_rate": _safe_float(data.get("mep_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("mep_detail_total", 0.0)),
        },
        {
            "Section": "Utility",
            "Cost Key": "u_util",
            "Basis": "GBA",
            "basis_value": gba,
            "current_rate": _safe_float(data.get("u_util", 0.0)),
            "derived_rate": _safe_float(data.get("utility_derived_unit_price", 0.0)),
            "detail_total": _safe_float(data.get("utility_detail_total", 0.0)),
        },
    ]

    rows = []
    for spec in review_specs:
        current_rate = spec["current_rate"]
        derived_rate = spec["derived_rate"]
        detail_total = spec["detail_total"]
        basis_value = _safe_float(spec.get("basis_value", gba))
        current_total = basis_value * current_rate
        rate_difference = derived_rate - current_rate
        total_difference = detail_total - current_total

        if detail_total <= 0:
            status = "No detail / zero detail"
        elif abs(rate_difference) <= 1:
            status = "Matching"
        else:
            status = "Different"

        rows.append({
            "Section": spec["Section"],
            "Cost Key": spec["Cost Key"],
            "Basis": spec["Basis"],
            "Current Cost Rate": current_rate,
            "Detail-Derived Rate": derived_rate,
            "Difference / Unit": rate_difference,
            "Detail Total": detail_total,
            "Current Total": current_total,
            "Difference Total": total_difference,
            "Status": status,
        })

    return rows

from project_database import PROJECT_DATABASE

#region --- DO NOT CHANGE#2 (OR I WILL KICK YOUR FACE)---
def mi(name):
    return f":material/{name}:"

def icon_safe(name):
    return mi(name) if "mi" in globals() else None    

def cb_switch_project():
    # Use .get() to avoid the AttributeError if the key is missing
    selected_label = st.session_state.get("project_selector")
    
    if not selected_label:
        return

    proj_ids = list(st.session_state.projects.keys())
    proj_labels = [f"{st.session_state.projects[pid]['name']} ({st.session_state.projects[pid]['type']})" for pid in proj_ids]
    
    if selected_label in proj_labels:
        selected_idx = proj_labels.index(selected_label)
        st.session_state.current_proj_id = proj_ids[selected_idx]
        save_data()     

def create_new_feasibility_study(study_name, project_type="Hotel"):
    clean_name = str(study_name).strip()

    if clean_name == "":
        st.error("Please enter feasibility study name.")
        return False

    st.session_state.current_study_name = clean_name

    st.session_state.projects = {
        "proj_1": {
            "name": clean_name,
            "type": project_type,
            "data": {}
        }
    }

    st.session_state.current_proj_id = "proj_1"
    st.session_state.proj_counter = 1

    st.session_state.loaded_snapshot_id = None
    st.session_state.loaded_snapshot_name = None

    st.session_state.report_config = copy.deepcopy(DEFAULT_REPORT_CONFIG)

    safe_keys = {"projects", "current_proj_id", "proj_counter", "current_study_name"}
    keys_to_clear = [
        k for k in st.session_state.keys()
        if "base_table_" in k
        or "area_editor_" in k
        or "rename_input_" in k
        or "renaming_" in k
        or "deleting_" in k
        or "fs_load_page" in k
        or "proj_" in k
    ]

    for k in keys_to_clear:
        if k not in safe_keys:
            del st.session_state[k]

    for stale_key in ["summary_calculations", "recap_math_engine", "project_selector"]:
        if stale_key in st.session_state:
            del st.session_state[stale_key]

    clear_project_ui_cache()

    return True

def generate_excel_template():
    """Generates an Excel template in memory with formatting and data validation."""
    
    # 1. Define the required columns based on your app's structure
    columns = [
        "FL", 
        "Space Type", 
        "Floor to Floor Height (m)", 
        "Typical Unit", 
        "Unit Area (m2)",
        "Parkir", 
        "Roof/Deck", 
        "MEP Outdoor", 
        "Office", 
        "Koridor/Lobby"
    ]
    
    # Create an empty DataFrame with these columns
    df = pd.DataFrame(columns=columns)
    
    # 2. Use BytesIO to keep the file in memory (no need to save to disk)
    output = BytesIO()
    
    # 3. Use XlsxWriter as the engine to enable Excel formatting
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='GBA_Stacking_Plan')
        
        # Access the underlying workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['GBA_Stacking_Plan']
        
        # Create a format object for headers (Bold, Green Background, Borders)
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Apply header format and set standard column widths
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 18) # Set column width to 18 pixels
            
        # 4. Add Data Validation (Dropdowns) to the Excel file
        # "Space Type" is the 2nd column (Column B)
        space_types = ["Roof", "Unit", "Lobby", "Ramp", "Carpark", "Facility", "Office"]
        
        # Apply the dropdown to rows 2 through 100 in Column B
        worksheet.data_validation('B2:B100', {
            'validate': 'list',
            'source': space_types,
            'input_title': 'Select Space Type',
            'input_message': 'Choose from the dropdown list.',
            'error_title': 'Invalid Input',
            'error_message': 'Please select a valid Space Type from the list.'
        })

    # Return the binary data
    return output.getvalue()
#endregion

def render_feasibility_study_landing(): #start page
    st.title("Feasibility Study")

    st.divider()

    active_file_id = st.session_state.get("loaded_snapshot_id")
    active_file_name = st.session_state.get("loaded_snapshot_name")

    # ==================================================
    # LANDING MODE STATE
    # ==================================================
    if "fs_landing_mode" not in st.session_state:
        st.session_state.fs_landing_mode = None

    # ==================================================
    # AFTER FILE IS CREATED / LOADED
    # ==================================================

    if active_file_id and st.session_state.fs_landing_mode is None:
        
        c1, c2 = st.columns([1, 1])
        
        c1.success(f"**{active_file_name}** is currently loaded (You can start calculating your project)", icon=":material/check:")
        c2.info("Use **Quick Save** button on the sidebar to save calculation", icon=":material/help:")

        col_msg, col_back = st.columns([1, 5])
    
        with col_msg:
            if st.button("Previous Page", icon=":material/arrow_back:", key="go_back_to_load_list", width="stretch"):
                st.session_state.fs_landing_mode = "home"
                st.rerun()

        return

    # ==================================================
    # FIRST-SIGHT QUESTION
    # ==================================================
    if st.session_state.fs_landing_mode is None or st.session_state.fs_landing_mode == "home":

        st.info("""

        **Welcome to Project Feasibility Study - Agung Sedayu Group**

        To start calculating, first create a new project by clicking the button below.""", icon=":material/waving_hand:")

        st.space(size="small")

        col_create_btn, col_load_btn, col_empty = st.columns([1, 1, 2], gap="small", vertical_alignment="center")

        with col_create_btn:
            if st.button(
                "Create New Feasibility Study",
                key="landing_choose_create_study",
                type="primary",
                width="stretch",
                icon=":material/create_new_folder:"
            ):
                st.session_state.fs_landing_mode = "create"
                st.rerun()

        with col_load_btn:
            st.info("**or load saved FS below:**")


    # ==================================================
    # CREATE NEW STUDY MODE
    # ==================================================
    elif st.session_state.fs_landing_mode == "create":
        col_title, col_back = st.columns([5, 1])

        with col_title:
            st.subheader("Create New Feasibility Study")
            st.caption("Enter a study name. A new default project will be created automatically.")

        with col_back:
            if st.button("Back", key="landing_create_back", width="stretch"):
                st.session_state.fs_landing_mode = None
                st.rerun()


        col_title, col_back = st.columns([5, 1])

        with col_title.form("create_new_feasibility_study_form", clear_on_submit=False):
            study_name = st.text_input(
                "Feasibility Study Name",
                placeholder="e.g. Project X - Option 1 - Rev 0"
            )

            create_clicked = st.form_submit_button(
                "Create New Feasibility Study",
                type="primary",
                width="stretch"
            )

            if create_clicked:
                if study_name.strip() == "":
                    st.warning("Please enter feasibility study name.")
                else:
                    # Uses default project type from create_new_feasibility_study()
                    # Usually default = Hotel
                    if create_new_feasibility_study(study_name):
                        st.session_state.fs_landing_mode = None
                        st.success(f"Started local study **{study_name.strip()}**. Use Archive > Online Backup > Save to archive it.")
                        st.rerun()
    
    snapshots = load_snapshots()

    if not snapshots:
        st.info("No saved feasibility studies yet.")

    else:
        # ==================================================
        # PAGINATION SETUP
        # ==================================================
        PAGE_SIZE = 10

        if "fs_load_page" not in st.session_state:
            st.session_state.fs_load_page = 0

        total_items = len(snapshots)
        total_pages = max(1, (total_items + PAGE_SIZE - 1) // PAGE_SIZE)

        # Safety clamp if files were deleted / changed
        if st.session_state.fs_load_page >= total_pages:
            st.session_state.fs_load_page = total_pages - 1

        if st.session_state.fs_load_page < 0:
            st.session_state.fs_load_page = 0

        start_idx = st.session_state.fs_load_page * PAGE_SIZE
        end_idx = start_idx + PAGE_SIZE
        page_snapshots = snapshots[start_idx:end_idx]

        st.divider()

        # ==================================================
        # SAVED FILE LIST
        # ==================================================
        for snap in page_snapshots:
            snap_id = snap["id"]
            snap_name = snap.get("snapshot_name", "Untitled File")

            is_active = st.session_state.get("loaded_snapshot_id") == snap_id
            active_label = " - ACTIVE" if is_active else ""

            saved_time = ""
            if "format_snapshot_time" in globals():
                saved_time = format_snapshot_time(snap.get("created_at"))

            col_file, col_date, col_action= st.columns([4, 2, 1])

            with col_file:
                st.markdown(f"**{snap_name}**")
                if is_active:
                    st.badge("Currently loaded file", icon=":material/check:", color="green")

                if saved_time:
                    st.caption(f"Saved: {saved_time}{active_label}")
                else:
                    st.caption(f"Saved file{active_label}")

            with col_action:
                if st.button(
                    "Load",
                    key=f"landing_load_file_{snap_id}",
                    type="primary",
                    width="stretch"
                ):
                    data = load_snapshot_data(snap_id)

                    if data:
                        restore_app_payload(data)

                        st.session_state.loaded_snapshot_id = snap_id
                        st.session_state.loaded_snapshot_name = snap_name
                        st.session_state.current_study_name = snap_name
                        st.session_state.fs_landing_mode = None

                        save_data()
                        st.success(f"Loaded **{snap_name}**.")
                        st.rerun()

            st.divider()

        # ==================================================
        # PAGINATION CONTROLS
        # ==================================================
        c1, col_prev, col_page, col_next, c2 = st.columns([5, 1, 2, 1, 5])

        with col_prev:
            if st.button(
                "Previous",
                key="fs_load_prev_page",
                width="stretch",
                disabled=st.session_state.fs_load_page <= 0
            ):
                st.session_state.fs_load_page -= 1
                st.rerun()

        with col_page:
            st.markdown(
                f"<div style='text-align:center; padding-top: 0.45rem;'>"
                f"Page {st.session_state.fs_load_page + 1} of {total_pages}"
                f"</div>",
                unsafe_allow_html=True
            )

        with col_next:
            if st.button(
                "Next",
                key="fs_load_next_page",
                width="stretch",
                disabled=st.session_state.fs_load_page >= total_pages - 1
            ):
                st.session_state.fs_load_page += 1
                st.rerun()

    return                
    st.divider()
    st.caption("For rename, delete, import, export, or full archive management, use the Feasibility Study Archive page.")

def show_project_database():  # database page
    st.title("Project Database")

    # ==================================================
    # CONTEXT MAP
    # Converts internal database keys into readable QS labels
    # ==================================================
    FIELD_CONTEXT = {
        # Structure
        "struc_earth": {
            "Group": "Structure",
            "Item": "Earthwork",
            "Basis": "Rp / m2 GBA",
            "Type": "currency",
            "Note": "Applied to gross building area."
        },
        "struc_found": {
            "Group": "Structure",
            "Item": "Foundation Work",
            "Basis": "Rp / m2 GBA",
            "Type": "currency",
            "Note": "Applied to gross building area."
        },
        "struc_work": {
            "Group": "Structure",
            "Item": "Main Structure Work",
            "Basis": "Rp / m2 GBA",
            "Type": "currency",
            "Note": "Applied to gross building area."
        },

        # Architecture Base
        "arch_base": {
            "Group": "Architecture",
            "Item": "Base Architectural Work",
            "Basis": "Rp / m2 GFA",
            "Type": "currency",
            "Note": "General architectural finishing rate."
        },
        "lobby": {
            "Group": "Architecture",
            "Item": "Lobby Finishing Premium",
            "Basis": "Rp / m2",
            "Type": "currency",
            "Note": "Additional lobby finishing allowance."
        },

        # Facade
        "facade_precast_rate": {
            "Group": "Facade",
            "Item": "Precast Facade Rate",
            "Basis": "Rp / m2 facade",
            "Type": "currency",
            "Note": "Applied according to precast facade ratio."
        },
        "facade_window_rate": {
            "Group": "Facade",
            "Item": "Window / Glass Facade Rate",
            "Basis": "Rp / m2 facade",
            "Type": "currency",
            "Note": "Applied according to window facade ratio."
        },
        "facade_double_rate": {
            "Group": "Facade",
            "Item": "Double Facade Rate",
            "Basis": "Rp / m2 facade",
            "Type": "currency",
            "Note": "Applied according to double facade ratio."
        },
        "facade_precast_pct": {
            "Group": "Facade",
            "Item": "Precast Facade Ratio",
            "Basis": "% of facade area",
            "Type": "percent",
            "Note": "Facade composition assumption."
        },
        "facade_window_pct": {
            "Group": "Facade",
            "Item": "Window Facade Ratio",
            "Basis": "% of facade area",
            "Type": "percent",
            "Note": "Facade composition assumption."
        },
        "facade_double_pct": {
            "Group": "Facade",
            "Item": "Double Facade Ratio",
            "Basis": "% of facade area",
            "Type": "percent",
            "Note": "Facade composition assumption."
        },

        # Doors & Hardware
        "door_wood": {
            "Group": "Doors & Hardware",
            "Item": "Wooden Door",
            "Basis": "Rp / unit",
            "Type": "currency",
            "Note": "Door supply and installation allowance."
        },
        "door_steel": {
            "Group": "Doors & Hardware",
            "Item": "Steel Door",
            "Basis": "Rp / unit",
            "Type": "currency",
            "Note": "Door supply and installation allowance."
        },
        "door_glass": {
            "Group": "Doors & Hardware",
            "Item": "Glass Door",
            "Basis": "Rp / unit",
            "Type": "currency",
            "Note": "Door supply and installation allowance."
        },
        "hw_wood": {
            "Group": "Doors & Hardware",
            "Item": "Wooden Door Hardware",
            "Basis": "Rp / set",
            "Type": "currency",
            "Note": "Hardware set allowance."
        },
        "hw_steel": {
            "Group": "Doors & Hardware",
            "Item": "Steel Door Hardware",
            "Basis": "Rp / set",
            "Type": "currency",
            "Note": "Hardware set allowance."
        },

        # Flooring
        "fl_waste": {
            "Group": "Flooring",
            "Item": "Flooring Wastage",
            "Basis": "%",
            "Type": "percent",
            "Note": "Material waste allowance."
        },
        "fl_ht_pct": {
            "Group": "Flooring",
            "Item": "Homogeneous Tile Ratio",
            "Basis": "% of floor area",
            "Type": "percent",
            "Note": "Floor finish composition."
        },
        "fl_vinyl_pct": {
            "Group": "Flooring",
            "Item": "Vinyl Floor Ratio",
            "Basis": "% of floor area",
            "Type": "percent",
            "Note": "Floor finish composition."
        },
        "fl_marmer_pct": {
            "Group": "Flooring",
            "Item": "Marble Floor Ratio",
            "Basis": "% of floor area",
            "Type": "percent",
            "Note": "Floor finish composition."
        },

        # Interior / Specialist
        "gondola": {
            "Group": "Specialist Works",
            "Item": "Gondola System",
            "Basis": "Rp / project",
            "Type": "currency",
            "Note": "Facade maintenance equipment allowance."
        },
        "carpet": {
            "Group": "Interior",
            "Item": "Carpet Finish",
            "Basis": "Rp / m2",
            "Type": "currency",
            "Note": "Carpet finishing rate."
        },
        "glass": {
            "Group": "Interior",
            "Item": "Interior Glass / Mirror",
            "Basis": "Rp / m2",
            "Type": "currency",
            "Note": "Interior glass allowance."
        },
        "ffe": {
            "Group": "Interior",
            "Item": "FF&E",
            "Basis": "Rp / room or unit",
            "Type": "currency",
            "Note": "Furniture, fixtures, and equipment allowance."
        },
        "misc": {
            "Group": "Interior",
            "Item": "Miscellaneous Interior Allowance",
            "Basis": "Rp allowance",
            "Type": "currency",
            "Note": "Project-specific miscellaneous allowance."
        },
        "kitchen": {
            "Group": "Interior",
            "Item": "Kitchen Equipment",
            "Basis": "Rp allowance",
            "Type": "currency",
            "Note": "Kitchen equipment allowance."
        },

        # Sanitary
        "san_room_rate": {
            "Group": "Sanitary",
            "Item": "Typical Room Sanitary",
            "Basis": "Rp / room",
            "Type": "currency",
            "Note": "Sanitary allowance for typical room or unit."
        },
        "san_pub_f": {
            "Group": "Sanitary",
            "Item": "Public Female Toilet",
            "Basis": "Rp / toilet set",
            "Type": "currency",
            "Note": "Public toilet sanitary allowance."
        },
        "san_pub_m": {
            "Group": "Sanitary",
            "Item": "Public Male Toilet",
            "Basis": "Rp / toilet set",
            "Type": "currency",
            "Note": "Public toilet sanitary allowance."
        },
        "san_dis": {
            "Group": "Sanitary",
            "Item": "Accessible Toilet",
            "Basis": "Rp / toilet set",
            "Type": "currency",
            "Note": "Disabled toilet sanitary allowance."
        },
        "san_mushola": {
            "Group": "Sanitary",
            "Item": "Mushola Ablution Area",
            "Basis": "Rp / area",
            "Type": "currency",
            "Note": "Wudhu / mushola sanitary allowance."
        },
        "san_room_qty": {
            "Group": "Sanitary",
            "Item": "Sanitary Quantity per Room",
            "Basis": "Qty / room",
            "Type": "number",
            "Note": "Typical sanitary quantity assumption."
        },

        # MEP & Utility
        "mep": {
            "Group": "MEP",
            "Item": "MEP Works",
            "Basis": "Rp / m2 GBA",
            "Type": "currency",
            "Note": "Mechanical, electrical, and plumbing rate."
        },
        "utility": {
            "Group": "Utility",
            "Item": "Infrastructure / Utility Works",
            "Basis": "Rp / m2 GBA",
            "Type": "currency",
            "Note": "External or supporting utility allowance."
        },

        # External Works
        "ext_land": {
            "Group": "External Works",
            "Item": "Landscape / External Works",
            "Basis": "Rp / m2",
            "Type": "currency",
            "Note": "External area and landscape allowance."
        },
        "railing_rate": {
            "Group": "External Works",
            "Item": "Railing",
            "Basis": "Rp / m",
            "Type": "currency",
            "Note": "Railing work allowance."
        },
        "railing_qty": {
            "Group": "External Works",
            "Item": "Railing Quantity Ratio",
            "Basis": "Qty",
            "Type": "number",
            "Note": "Default railing quantity assumption."
        },
        "skylight_rate": {
            "Group": "External Works",
            "Item": "Skylight",
            "Basis": "Rp / m2",
            "Type": "currency",
            "Note": "Skylight work allowance."
        },

        # Facilities
        "fac_pub": {
            "Group": "Facilities",
            "Item": "Public Facility Allowance",
            "Basis": "Rp / room or unit",
            "Type": "currency",
            "Note": "Public facility cost allowance."
        },
        "fac_res": {
            "Group": "Facilities",
            "Item": "Residential Facility Allowance",
            "Basis": "Rp / room or unit",
            "Type": "currency",
            "Note": "Residential facility cost allowance."
        },
        "fac_proj": {
            "Group": "Facilities",
            "Item": "Project Facility Allowance",
            "Basis": "Rp / project",
            "Type": "currency",
            "Note": "Lump-sum project facility allowance."
        },

        # Soft Cost
        "cons": {
            "Group": "Soft Cost",
            "Item": "Consultancy Cost",
            "Basis": "Rp / m2 GFA",
            "Type": "currency",
            "Note": "Consultant / professional fee allowance."
        },
    }

    GROUP_ORDER = [
        "Structure",
        "Architecture",
        "Facade",
        "Doors & Hardware",
        "Flooring",
        "Interior",
        "Sanitary",
        "MEP",
        "Utility",
        "External Works",
        "Facilities",
        "Specialist Works",
        "Soft Cost",
        "Other"
    ]

    # ==================================================
    # HELPER FUNCTIONS
    # ==================================================
    def prettify_key(raw_key):
        text = raw_key.replace("_", " ").replace(".", " - ")
        return text.title()

    def get_field_context(raw_key):
        # Exact match
        if raw_key in FIELD_CONTEXT:
            return FIELD_CONTEXT[raw_key]

        # Nested flooring rates, e.g. fl_ht_rate.Type1
        if raw_key.startswith("fl_ht_rate."):
            subtype = raw_key.split(".", 1)[1]
            return {
                "Group": "Flooring",
                "Item": f"Homogeneous Tile Rate - {subtype}",
                "Basis": "Rp / m2",
                "Type": "currency",
                "Note": "Floor finish unit rate."
            }

        if raw_key.startswith("fl_vinyl_rate."):
            subtype = raw_key.split(".", 1)[1]
            return {
                "Group": "Flooring",
                "Item": f"Vinyl Floor Rate - {subtype}",
                "Basis": "Rp / m2",
                "Type": "currency",
                "Note": "Floor finish unit rate."
            }

        if raw_key.startswith("fl_marmer_rate."):
            subtype = raw_key.split(".", 1)[1]
            return {
                "Group": "Flooring",
                "Item": f"Marble Floor Rate - {subtype}",
                "Basis": "Rp / m2",
                "Type": "currency",
                "Note": "Floor finish unit rate."
            }

        return {
            "Group": "Other",
            "Item": prettify_key(raw_key),
            "Basis": "-",
            "Type": "number",
            "Note": "Unmapped database item."
        }

    def format_value(value, value_type):
        try:
            num = _safe_float(value)
        except Exception:
            return str(value)

        if value_type == "currency":
            return f"Rp {num:,.0f}"
        elif value_type == "percent":
            return f"{num:,.1f}%"
        elif value_type == "number":
            return f"{num:,.2f}".rstrip("0").rstrip(".")
        else:
            return str(value)

    def flatten_project_database():
        rows = []

        for project_type, metrics in PROJECT_DATABASE.items():
            for key_name, value in metrics.items():

                if isinstance(value, dict):
                    for sub_key, sub_value in value.items():
                        raw_key = f"{key_name}.{sub_key}"
                        ctx = get_field_context(raw_key)

                        rows.append({
                            "Project Type": project_type,
                            "Group": ctx["Group"],
                            "Cost Item": ctx["Item"],
                            "Basis": ctx["Basis"],
                            "Value": sub_value,
                            "Formatted Value": format_value(sub_value, ctx["Type"]),
                            "Note": ctx["Note"],
                            "Internal Key": raw_key
                        })

                else:
                    raw_key = key_name
                    ctx = get_field_context(raw_key)

                    rows.append({
                        "Project Type": project_type,
                        "Group": ctx["Group"],
                        "Cost Item": ctx["Item"],
                        "Basis": ctx["Basis"],
                        "Value": value,
                        "Formatted Value": format_value(value, ctx["Type"]),
                        "Note": ctx["Note"],
                        "Internal Key": raw_key
                    })

        df = pd.DataFrame(rows)

        df["Group Sort"] = df["Group"].apply(
            lambda x: GROUP_ORDER.index(x) if x in GROUP_ORDER else 999
        )

        df = df.sort_values(
            by=["Group Sort", "Cost Item", "Project Type"]
        ).drop(columns=["Group Sort"])

        return df

    df_long = flatten_project_database()

    project_types = list(PROJECT_DATABASE.keys())

    # ==================================================
    # PAGE STYLE
    # ==================================================
    st.markdown("""
    <style>
        .db-card {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 14px;
            padding: 14px 16px;
            box-shadow: 0 1px 3px rgba(16,24,40,0.04);
        }

        .db-label {
            font-size: 11px;
            color: #6B7280;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            font-weight: 700;
            margin-bottom: 5px;
        }

        .db-value {
            font-size: 20px;
            color: #111827;
            font-weight: 750;
            line-height: 1.2;
        }

        .db-sub {
            font-size: 12px;
            color: #6B7280;
            margin-top: 4px;
        }

        .db-note {
            background: #F9FAFB;
            border: 1px solid #E5E7EB;
            border-left: 4px solid #3E4095;
            border-radius: 12px;
            padding: 12px 14px;
            margin-bottom: 14px;
            font-size: 13px;
            color: #4B5563;
            line-height: 1.55;
        }
    </style>
    """, unsafe_allow_html=True)

    # ==================================================
    # SUMMARY CARDS
    # ==================================================
    total_project_types = len(project_types)
    total_items = df_long["Cost Item"].nunique()
    total_groups = df_long["Group"].nunique()

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            f"""
            <div class="db-card">
                <div class="db-label">Project Types</div>
                <div class="db-value">{total_project_types}</div>
                <div class="db-sub">Available benchmark templates</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            f"""
            <div class="db-card">
                <div class="db-label">Cost Items</div>
                <div class="db-value">{total_items}</div>
                <div class="db-sub">Contextual database fields</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            f"""
            <div class="db-card">
                <div class="db-label">Cost Groups</div>
                <div class="db-value">{total_groups}</div>
                <div class="db-sub">Grouped by QS discipline</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()
    # ==================================================
    # MAIN TABS
    # ==================================================
    tab_search, tab_project, tab_matrix, tab_raw = st.tabs([
        "Search Database",
        "Project Type View",
        "Comparison Matrix",
        "Raw Audit View"
    ])

    # ==================================================
    # TAB 1: PROJECT TYPE VIEW
    # ==================================================
    with tab_project:
        selected_project_type = st.selectbox(
            "Select Project Type",
            options=project_types,
            index=0
        )

        selected_groups = st.multiselect(
            "Filter Cost Group",
            options=GROUP_ORDER,
            default=[
                "Structure",
                "Architecture",
                "Facade",
                "Flooring",
                "MEP",
                "Utility",
                "Soft Cost"
            ]
        )

        df_project = df_long[
            (df_long["Project Type"] == selected_project_type)
            & (df_long["Group"].isin(selected_groups))
        ].copy()

        display_df = df_project[[
            "Group",
            "Cost Item",
            "Basis",
            "Formatted Value",
            "Note"
        ]].rename(columns={
            "Formatted Value": "Rate / Assumption"
        })

        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            column_config={
                "Group": st.column_config.TextColumn("Cost Group", width="medium"),
                "Cost Item": st.column_config.TextColumn("Cost Item", width="large"),
                "Basis": st.column_config.TextColumn("Basis", width="medium"),
                "Rate / Assumption": st.column_config.TextColumn("Rate / Assumption", width="medium"),
                "Note": st.column_config.TextColumn("Context", width="large"),
            }
        )

    # ==================================================
    # TAB 2: COMPARISON MATRIX
    # ==================================================
    with tab_matrix:
        st.caption("Compare contextual cost items across project types.")

        matrix_groups = st.multiselect(
            "Select Groups to Compare",
            options=GROUP_ORDER,
            default=["Structure", "Architecture", "Facade", "MEP", "Utility", "Soft Cost"],
            key="db_matrix_groups"
        )

        compare_project_types = st.multiselect(
            "Select Project Types",
            options=project_types,
            default=project_types[:5],
            key="db_compare_project_types"
        )

        df_matrix_source = df_long[
            (df_long["Group"].isin(matrix_groups))
            & (df_long["Project Type"].isin(compare_project_types))
        ].copy()

        matrix_df = df_matrix_source.pivot_table(
            index=["Group", "Cost Item", "Basis"],
            columns="Project Type",
            values="Formatted Value",
            aggfunc="first"
        ).reset_index()

        st.dataframe(
            matrix_df,
            width="stretch",
            hide_index=True
        )

    # ==================================================
    # TAB 3: RAW AUDIT VIEW
    # ==================================================
    with tab_raw:
        st.caption("Audit view showing internal keys beside contextual labels.")

        raw_groups = st.multiselect(
            "Filter Raw View by Group",
            options=GROUP_ORDER,
            default=GROUP_ORDER,
            key="db_raw_groups"
        )

        raw_df = df_long[df_long["Group"].isin(raw_groups)].copy()

        st.dataframe(
            raw_df[[
                "Project Type",
                "Group",
                "Cost Item",
                "Basis",
                "Formatted Value",
                "Internal Key",
                "Note"
            ]].rename(columns={
                "Formatted Value": "Rate / Assumption"
            }),
            width="stretch",
            hide_index=True
        )

    # ==================================================
    # TAB 4: SEARCH DATABASE
    # ==================================================
    with tab_search:

        c1, c2, c3 = st.columns(3)

        with c1:
            search_item = st.text_input(
                "Item Name",
                placeholder="e.g. Foundation, Utility, MEP",
                key="db_search_item"
            )

        with c2:
            search_group = st.text_input(
                "Cost Group",
                placeholder="e.g. Structure, Architecture, Flooring",
                key="db_search_group"
            )

        with c3:
            search_project_type = st.text_input(
                "Project Type",
                placeholder="e.g. Hotel, Apartment, Retail",
                key="db_search_project_type"
            )

        has_search_input = (
            search_item.strip() != ""
            or search_group.strip() != ""
            or search_project_type.strip() != ""
        )

        if not has_search_input:
            st.info("Enter item name, cost group, or project type to search the database.")
        else:
            df_search = df_long.copy()

            if search_item.strip():
                item_query = search_item.strip().lower()
                df_search = df_search[
                    df_search["Cost Item"].str.lower().str.contains(item_query, na=False)
                    | df_search["Internal Key"].str.lower().str.contains(item_query, na=False)
                    | df_search["Note"].str.lower().str.contains(item_query, na=False)
                ]

            if search_group.strip():
                group_query = search_group.strip().lower()
                df_search = df_search[
                    df_search["Group"].str.lower().str.contains(group_query, na=False)
                ]

            if search_project_type.strip():
                project_query = search_project_type.strip().lower()
                df_search = df_search[
                    df_search["Project Type"].str.lower().str.contains(project_query, na=False)
                ]

            st.divider()

            st.caption(f"Search result: {len(df_search):,} rows")

            if df_search.empty:
                st.info("No matching database item found.")
            else:
                st.dataframe(
                    df_search[[
                        "Group",
                        "Cost Item",
                        "Project Type",
                        "Basis",
                        "Formatted Value",
                        "Internal Key",
                        "Note"
                    ]].rename(columns={
                        "Group": "Cost Group",
                        "Formatted Value": "Rate / Assumption"
                    }),
                    width="stretch",
                    hide_index=True,
                    column_config={
                        "Cost Group": st.column_config.TextColumn("Cost Group", width="medium"),
                        "Cost Item": st.column_config.TextColumn("Cost Item", width="large"),
                        "Project Type": st.column_config.TextColumn("Project Type", width="medium"),
                        "Basis": st.column_config.TextColumn("Basis", width="medium"),
                        "Rate / Assumption": st.column_config.TextColumn("Rate / Assumption", width="medium"),
                        "Internal Key": st.column_config.TextColumn("Internal Key", width="medium"),
                        "Note": st.column_config.TextColumn("Context", width="large"),
                    }
                )

def sync_door_rows_from_area(area_df, old_door_records):
    F2F_COL = "Floor to Floor Height (m)"
    TYPICAL_UNIT_COL = "Typical Unit"

    DOOR_WOOD_COL = "Pintu Kayu"
    DOOR_STEEL_COL = "Pintu Besi"
    DOOR_GLASS_COL = "Pintu Kaca"

    if not isinstance(old_door_records, list):
        old_door_records = []

    old_by_fl = {
        str(row.get("FL", "")).strip(): row
        for row in old_door_records
        if isinstance(row, dict)
    }

    synced = []

    for _, row in area_df.iterrows():
        fl = str(row.get("FL", "")).strip()
        old = old_by_fl.get(fl, {})

        synced.append(
            {
                "FL": fl,
                "Space Type": str(row.get("Space Type", "")).strip(),
                F2F_COL: _safe_float(row.get(F2F_COL, 0.0)),
                TYPICAL_UNIT_COL: int(_safe_float(row.get(TYPICAL_UNIT_COL, 0))),
                DOOR_WOOD_COL: int(_safe_float(old.get(DOOR_WOOD_COL, 0))),
                DOOR_STEEL_COL: int(_safe_float(old.get(DOOR_STEEL_COL, 0))),
                DOOR_GLASS_COL: int(_safe_float(old.get(DOOR_GLASS_COL, 0))),
            }
        )

    return synced

def save_door_table_to_cloud(
    edited_door_df,
    area_df,
    door_draft_key,
    door_committed_key,
    door_editor_key,
):
    """
    Door/Pintu has a special save path because rows are generated from Area/GBA
    but user-entered door quantities must be preserved.
    """
    DOOR_WOOD_COL = "Pintu Kayu"
    DOOR_STEEL_COL = "Pintu Besi"
    DOOR_GLASS_COL = "Pintu Kaca"

    current_door_records = edited_door_df.to_dict("records")

    synced_door_records = sync_door_rows_from_area(
        area_df,
        current_door_records,
    )

    cleaned_door = clean_door_records(synced_door_records)

    st.session_state[door_draft_key] = copy.deepcopy(cleaned_door)
    st.session_state[door_committed_key] = copy.deepcopy(cleaned_door)

    set_data("door_table", copy.deepcopy(cleaned_door))

    saved_door_df = pd.DataFrame(cleaned_door)

    total_door_wood = (
        int(saved_door_df[DOOR_WOOD_COL].sum())
        if DOOR_WOOD_COL in saved_door_df.columns
        else 0
    )
    total_door_steel = (
        int(saved_door_df[DOOR_STEEL_COL].sum())
        if DOOR_STEEL_COL in saved_door_df.columns
        else 0
    )
    total_door_glass = (
        int(saved_door_df[DOOR_GLASS_COL].sum())
        if DOOR_GLASS_COL in saved_door_df.columns
        else 0
    )

    update_data({
        "door_wood_calc": total_door_wood,
        "door_steel_calc": total_door_steel,
        "door_glass_calc": total_door_glass,
    })

    if door_editor_key in st.session_state:
        del st.session_state[door_editor_key]

    save_ok = save_data_force()

    if save_ok:
        st.success("Door table saved to cloud.")
        return True

    st.error("Cloud save failed. Door data changed locally, but was not saved. Do not log out yet.")
    return False


def save_after_user_action(success_message="Saved to cloud.", fail_message="Changed locally, but cloud save failed. Do not log out yet."):
    """
    Use this only after explicit structural user actions:
    - create component
    - edit component
    - delete component

    Do not use this during app startup/load/repair.
    Do not use this after every table edit.
    """
    save_ok = save_data_force()

    if save_ok:
        st.success(success_message)
        return True

    st.error(fail_message)
    return False
#endregion

def show_area_calculator():
    st.title("Area Analysis")

    curr_id, curr_proj = get_current_project()

    if "data" not in curr_proj or not isinstance(curr_proj["data"], dict):
        curr_proj["data"] = {}

    def get_area_val(key, default=0.0):
        return curr_proj["data"].get(key, default)

    def safe_sum(df, col):
        return _safe_float(df[col].sum()) if col in df.columns else 0.0

    F2F_COL = "Floor to Floor Height (m)"
    UNIT_AREA_COL = "Unit Area"
    TYPICAL_UNIT_COL = "Typical Unit"

    breakdown_cols = [
        "Parkir",
        "Roof/Deck",
        "MEP Outdoor",
        "Koridor/Lobby",
        "Stair, MEP, Etc",
        UNIT_AREA_COL,
        "Office",
    ]

    editable_cols = ["FL", "Space Type", F2F_COL, TYPICAL_UNIT_COL] + breakdown_cols

    # Stable state keys
    area_committed_key = f"area_committed_{curr_id}"
    area_draft_key = f"area_draft_{curr_id}"
    area_editor_key = f"area_editor_{curr_id}"

    def clear_area_editor_state():
        for k in [
            f"{area_editor_key}_Detailed Category",
            f"{area_editor_key}_Consultant Summary",
        ]:
            if k in st.session_state:
                del st.session_state[k]

    door_committed_key = f"door_committed_{curr_id}"
    door_draft_key = f"door_draft_{curr_id}"
    door_editor_key = f"door_editor_{curr_id}"

    # ==================================================
    # INITIALIZE AREA COMMITTED + DRAFT
    # ==================================================
    if area_committed_key not in st.session_state:
        saved_area = curr_proj["data"].get("area_table_data", [])

        if isinstance(saved_area, list) and len(saved_area) > 0:
            st.session_state[area_committed_key] = clean_area_records(saved_area)
        else:
            default_area = generate_area_rows(
                int(get_area_val("up_in", 5)),
                int(get_area_val("base_in", 1)),
            )
            st.session_state[area_committed_key] = clean_area_records(default_area)
            set_data("area_table", st.session_state[area_committed_key])

    if area_draft_key not in st.session_state:
        st.session_state[area_draft_key] = copy.deepcopy(
            st.session_state[area_committed_key]
        )

    # Always calculate from committed data only
    calc_df = calculate_area_dataframe(st.session_state[area_committed_key])
    edited_df = calc_df  # compatibility with your older downstream naming

    # ==================================================
    # TOP CONTROLS
    # ==================================================
    tower_name = curr_proj.get("name", "Unnamed Component")
    curr_proj["data"]["tname"] = tower_name

    with st.container(border=True):
        c_name, c_h, c_b, c_u = st.columns(4)

        c_name.text_input(
            "Project Name",
            value=tower_name,
            key=f"wid_tname_{curr_id}",
            disabled=True,
            help="This follows the active component name from the sidebar.",
        )

        base_key = f"wid_base_{curr_id}"
        up_key = f"wid_up_{curr_id}"
        land_key = f"wid_m_land_{curr_id}"

        c_h.number_input(
            "Basements / LG",
            min_value=0,
            value=int(get_area_val("base_in", 1)),
            step=1,
            key=base_key,
            help="1 = LG, 2 = B1, 3 = B2, etc",
        )

        c_b.number_input(
            "Floors",
            min_value=1,
            value=int(get_area_val("up_in", 5)),
            step=1,
            key=up_key,
        )

        c_u.number_input(
            "Luas Tanah (m2)",
            min_value=0.0,
            value=_safe_float(get_area_val("m_land", 0.0)),
            step=100.0,
            key=land_key,
        )

        curr_proj["data"]["base_in"] = int(st.session_state[base_key])
        curr_proj["data"]["up_in"] = int(st.session_state[up_key])
        curr_proj["data"]["m_land"] = _safe_float(st.session_state[land_key])

    # ==================================================
    # FACADE INPUTS - NORMAL WIDGETS USE SESSION KEYS
    # ==================================================
    with st.container(border=True):
        st.markdown("##### Facade Area Inputs")

        f1, f2, f3, f4 = st.columns(4)

        keliling_key = f"area_keliling_facade_{curr_id}"
        railing_len_key = f"area_panjang_railing_{curr_id}"
        railing_h_key = f"area_tinggi_railing_{curr_id}"
        tol_key = f"area_facade_tolerance_pct_{curr_id}"

        f1.number_input(
            "Keliling Facade (m)",
            min_value=0.0,
            value=_safe_float(get_area_val("area_keliling_facade", 0.0)),
            step=1.0,
            key=keliling_key,
            help="Building facade perimeter / keliling bangunan.",
        )

        f2.number_input(
            "Panjang Railing(m)",
            min_value=0.0,
            value=_safe_float(get_area_val("area_panjang_railing", 0.0)),
            step=0.1,
            key=railing_len_key,
            help="Average railing length per typical unit.",
        )

        f3.number_input(
            "Tinggi Railing(m)",
            min_value=0.0,
            value=_safe_float(get_area_val("area_tinggi_railing", 0.0)),
            step=0.1,
            key=railing_h_key,
            help="Average railing height.",
        )

        f4.number_input(
            "Tolerance (%)",
            min_value=0.0,
            value=_safe_float(get_area_val("area_facade_tolerance_pct", 15.0)),
            step=1.0,
            key=tol_key,
            help="Tekukan dan Overlap",
        )

        keliling_facade = _safe_float(st.session_state[keliling_key])
        panjang_railing = _safe_float(st.session_state[railing_len_key])
        tinggi_railing = _safe_float(st.session_state[railing_h_key])
        facade_tolerance_pct = _safe_float(st.session_state[tol_key])

        curr_proj["data"]["area_keliling_facade"] = keliling_facade
        curr_proj["data"]["area_panjang_railing"] = panjang_railing
        curr_proj["data"]["area_tinggi_railing"] = tinggi_railing
        curr_proj["data"]["area_facade_tolerance_pct"] = facade_tolerance_pct
        curr_proj["data"]["area_railing_length_per_room_calc"] = panjang_railing

    # ==================================================
    # GLOBAL CALCULATIONS FROM COMMITTED TABLE ONLY
    # ==================================================
    total_floor_height = safe_sum(edited_df, F2F_COL)

    total_typical_units = (
        int(pd.to_numeric(edited_df[TYPICAL_UNIT_COL], errors="coerce").fillna(0).sum())
        if TYPICAL_UNIT_COL in edited_df.columns
        else 0
    )

    total_lobby_interior = safe_sum(edited_df, "Koridor/Lobby")

    facade_wall_area = total_floor_height * keliling_facade
    facade_railing_area = total_typical_units * panjang_railing * tinggi_railing
    facade_subtotal = facade_wall_area + facade_railing_area
    facade_tolerance_area = facade_subtotal * facade_tolerance_pct / 100
    total_facade_area = facade_subtotal + facade_tolerance_area

    curr_proj["data"]["area_lobby_interior_calc"] = total_lobby_interior
    curr_proj["data"]["area_rooms_calc"] = total_typical_units
    curr_proj["data"]["area_typical_units_total_calc"] = total_typical_units
    curr_proj["data"]["area_facade_wall_calc"] = facade_wall_area
    curr_proj["data"]["area_facade_railing_calc"] = facade_railing_area
    curr_proj["data"]["area_facade_subtotal_calc"] = facade_subtotal
    curr_proj["data"]["area_facade_tolerance_area_calc"] = facade_tolerance_area
    curr_proj["data"]["area_facade_calc"] = total_facade_area

    # ==================================================
    # PAGE SELECTOR
    # ==================================================
    area_page_options = [
            "Excel",
            "GBA Input",
            "Pintu",
            "Eksternal",
            "Residential",
            "Earthworks",
            "Structural",
            "Foundation",
            "Architectural",
            "FF&E",
            "MEP",
            "Utility",
            "Consultancy",
    ]

    area_page_key = f"area_page_{curr_id}"
    if st.session_state.get(area_page_key) not in [None, *area_page_options]:
        st.session_state[area_page_key] = "Excel"

    if callable(getattr(st, "segmented_control", None)):
        area_page = st.segmented_control(
            "Area Page",
            options=area_page_options,
            default="Excel",
            key=area_page_key,
        )
    else:
        area_page = st.radio(
            "Area Page",
            options=area_page_options,
            horizontal=True,
            key=area_page_key,
        )

    # ==================================================
    # TAB 1 - EXCEL FORM
    # ==================================================
    if area_page == "Excel":
        st.subheader("Area Analysis (Excel)")
        st.caption(
            "Download or upload the Area Calculator Excel Form, then review the current values below."
        )

        summary_area_totals = calculate_area_totals_from_table(
            curr_proj["data"].get("area_table_data", [])
        )
        summary_residential_total = _safe_float(
            curr_proj["data"].get("area_res_fac_amount_calc", 0.0)
        )
        summary_external_total = _safe_float(
            curr_proj["data"].get("area_external_amount_calc", 0.0)
        )
        summary_external_rate = _safe_float(curr_proj["data"].get("u_ext", 0.0))
        summary_residential_rate = _safe_float(curr_proj["data"].get("u_fac_res", 0.0))
        summary_landscape = (
            summary_external_total / summary_external_rate
            if summary_external_total > 0 and summary_external_rate > 0
            else _safe_float(curr_proj["data"].get("area_landscape_qty_calc", 0.0))
        )
        summary_residential_facility = (
            summary_residential_total / summary_residential_rate
            if summary_residential_total > 0 and summary_residential_rate > 0
            else 0.0
        )

        core_summary_df = pd.DataFrame(
            [
                {"Item": "Luas Tanah", "Value": _safe_float(curr_proj["data"].get("m_land", 0.0)), "Unit": "m2"},
                {"Item": "GBA", "Value": summary_area_totals["gba"], "Unit": "m2"},
                {"Item": "GFA", "Value": summary_area_totals["gfa"], "Unit": "m2"},
                {"Item": "SGFA", "Value": summary_area_totals["sgfa"], "Unit": "m2"},
                {"Item": "NFA", "Value": summary_area_totals["nfa"], "Unit": "m2"},
                {"Item": "Lobby/Koridor", "Value": _safe_float(curr_proj["data"].get("area_lobby_interior_calc", 0.0)), "Unit": "m2"},
                {"Item": "Rooms/Units", "Value": _safe_float(curr_proj["data"].get("area_rooms_calc", 0.0)), "Unit": "unit"},
            ]
        )
        opening_summary_df = pd.DataFrame(
            [
                {"Item": "Facade", "Value": _safe_float(curr_proj["data"].get("area_facade_calc", 0.0)), "Unit": "m2"},
                {"Item": "Landscape", "Value": summary_landscape, "Unit": "m2"},
                {"Item": "Residential Facility", "Value": summary_residential_facility, "Unit": "m2"},
                {"Item": "Railing", "Value": _safe_float(curr_proj["data"].get("area_railing_length_per_room_calc", curr_proj["data"].get("area_panjang_railing", 0.0))), "Unit": "m'/room"},
                {"Item": "Wooden Door", "Value": _safe_float(curr_proj["data"].get("area_door_wood_calc", 0.0)), "Unit": "unit"},
                {"Item": "Steel Door", "Value": _safe_float(curr_proj["data"].get("area_door_steel_calc", 0.0)), "Unit": "unit"},
                {"Item": "Glass Door", "Value": _safe_float(curr_proj["data"].get("area_door_glass_calc", 0.0)), "Unit": "unit"},
            ]
        )

        s1, s2 = st.columns(2)
        with s1:
            st.markdown("##### Core Area")
            st.dataframe(
                core_summary_df,
                width="stretch",
                hide_index=True,
                column_config={
                    "Value": st.column_config.NumberColumn("Value", format="%.2f"),
                },
            )
        with s2:
            st.markdown("##### Opening / External")
            st.dataframe(
                opening_summary_df,
                width="stretch",
                hide_index=True,
                column_config={
                    "Value": st.column_config.NumberColumn("Value", format="%.2f"),
                },
            )

        st.markdown("##### Excel Form")

        with st.container():
            fcol1, fcol2, fcol3, fcol4 = st.columns(4)

            include_roof_machine = fcol1.checkbox(
                "Include Roof Machine",
                value=True,
                key=f"excel_form_roof_machine_{curr_id}",
            )

            include_roof = fcol2.checkbox(
                "Include Roof",
                value=True,
                key=f"excel_form_roof_{curr_id}",
            )

            excel_upper_floors = fcol3.number_input(
                "Floors",
                min_value=1,
                value=int(st.session_state[up_key]),
                step=1,
                key=f"excel_form_floors_{curr_id}",
            )

            excel_basements = fcol4.number_input(
                "Basements / LG",
                min_value=0,
                value=int(st.session_state[base_key]),
                step=1,
                key=f"excel_form_basements_{curr_id}",
                help="1 = LG, 2 = LG + B1, 3 = LG + B1 + B2",
            )

            excel_form_bytes = create_area_excel_form_bytes(
                project_name=tower_name,
                upper_floors=excel_upper_floors,
                basements=excel_basements,
                include_roof_machine=include_roof_machine,
                include_roof=include_roof,
                earthwork_detail_rows=curr_proj["data"].get(
                    "earthwork_detail_rows",
                    get_default_earthwork_detail_rows(),
                ),
                consultancy_detail_rows=curr_proj["data"].get(
                    "consultancy_detail_rows",
                    get_default_consultancy_detail_rows(),
                ),
                ffe_detail_rows=curr_proj["data"].get(
                    "ffe_detail_rows",
                    get_default_ffe_detail_rows(),
                ),
                mep_detail_rows=curr_proj["data"].get(
                    "mep_detail_rows",
                    get_default_mep_detail_rows(),
                ),
                utility_detail_rows=curr_proj["data"].get(
                    "utility_detail_rows",
                    get_default_utility_detail_rows(),
                ),
                foundation_detail_rows=curr_proj["data"].get(
                    "foundation_detail_rows",
                    get_default_foundation_detail_rows(),
                ),
                structural_detail_rows=curr_proj["data"].get(
                    "structural_detail_rows",
                    get_default_structural_detail_rows(),
                ),
                architectural_detail_rows=curr_proj["data"].get(
                    "architectural_detail_rows",
                    get_default_architectural_detail_rows(),
                ),
                earthwork_gba=_safe_float(curr_proj["data"].get("m_gba", 0.0)) or safe_sum(edited_df, "GBA"),
                foundation_gba=_safe_float(curr_proj["data"].get("m_gba", 0.0)) or safe_sum(edited_df, "GBA"),
                structural_gba=_safe_float(curr_proj["data"].get("m_gba", 0.0)) or safe_sum(edited_df, "GBA"),
                architectural_base_values=get_architectural_detail_base_values(curr_proj["data"], edited_df),
                consultancy_base_values=get_consultancy_detail_base_values(curr_proj["data"], edited_df),
                ffe_rooms=_safe_float(curr_proj["data"].get("m_rooms", 0.0)) or _safe_float(curr_proj["data"].get("area_rooms_calc", 0.0)),
                mep_gba=_safe_float(curr_proj["data"].get("m_gba", 0.0)) or safe_sum(edited_df, "GBA"),
                utility_gba=_safe_float(curr_proj["data"].get("m_gba", 0.0)) or safe_sum(edited_df, "GBA"),
            )

            def safe_filename_part(value, fallback="Unnamed"):
                text = str(value or fallback).strip()

                if text == "":
                    text = fallback

                for bad_char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
                    text = text.replace(bad_char, "-")

                # Clean double spaces / double dashes lightly
                text = " ".join(text.split())

                return text


            project_name_for_file = safe_filename_part(
                curr_proj.get("name", "Unnamed Component"),
                "Unnamed Component",
            )

            project_type_for_file = safe_filename_part(
                curr_proj.get("type", "Unknown Type"),
                "Unknown Type",
            )

            active_file_name_for_file = safe_filename_part(
                st.session_state.get("loaded_snapshot_name", "Unsaved File"),
                "Unsaved File",
            )

            excel_download_filename = (
                f"{project_name_for_file} - "
                f"{project_type_for_file} - "
                f"{active_file_name_for_file}.xlsx"
            )

            st.download_button(
                label="Download Area Excel Form",
                data=excel_form_bytes,
                file_name=excel_download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

            uploaded_excel = st.file_uploader(
                "Upload Area Calculator Excel Form",
                type=["xlsx"],
                key=f"area_excel_upload_{curr_id}",
                help="Upload the Excel form generated from this app.",
            )

            if uploaded_excel is not None:
                import_clicked = st.button(
                    "Import Excel to Area Calculator",
                    type="primary",
                    key=f"import_area_excel_{curr_id}",
                    width="stretch",
                )

                if import_clicked:
                    excel_bytes = uploaded_excel.getvalue()

                    try:
                        imported_area_records = read_area_input_sheet(excel_bytes)
                        st.write("Imported area rows:", len(imported_area_records))
                        st.dataframe(pd.DataFrame(imported_area_records).head(10), width="stretch")

                        if not imported_area_records:
                            st.error("Excel import failed: no area rows found in Area Input.")
                            st.stop()

                        imported_area_df = calculate_area_dataframe(imported_area_records)

                        imported_door_records = read_pintu_sheet(
                            excel_bytes,
                            area_df=imported_area_df,
                        )

                        imported_external_records, imported_landscape_data = read_external_sheet(
                            excel_bytes
                        )

                        imported_res_fac_records = read_residential_area_sheet(excel_bytes)

                        imported_earthwork_rows = None
                        try:
                            imported_earthwork_rows = read_earthworks_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_architectural_rows = None
                        try:
                            imported_architectural_rows = read_architectural_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_architectural_facade_inputs = read_architectural_facade_inputs(excel_bytes)

                        imported_consultancy_rows = None
                        try:
                            imported_consultancy_rows = read_consultancy_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_foundation_rows = None
                        try:
                            imported_foundation_rows = read_foundation_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_ffe_rows = None
                        try:
                            imported_ffe_rows = read_ffe_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_mep_rows = None
                        try:
                            imported_mep_rows = read_mep_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_utility_rows = None
                        try:
                            imported_utility_rows = read_utility_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        imported_structural_rows = None
                        try:
                            imported_structural_rows = read_structural_sheet(excel_bytes)
                        except ExcelImportError as e:
                            st.warning(str(e))

                        st.session_state[area_draft_key] = copy.deepcopy(imported_area_records)
                        st.session_state[area_committed_key] = copy.deepcopy(imported_area_records)
                        set_data("area_table", copy.deepcopy(imported_area_records))

                        if imported_door_records:
                            st.session_state[door_draft_key] = copy.deepcopy(imported_door_records)
                            st.session_state[door_committed_key] = copy.deepcopy(imported_door_records)
                            curr_proj["data"]["area_door_table_data"] = copy.deepcopy(imported_door_records)

                            door_df_imported = pd.DataFrame(imported_door_records)

                            curr_proj["data"]["area_door_wood_calc"] = (
                                int(door_df_imported["Pintu Kayu"].sum())
                                if "Pintu Kayu" in door_df_imported.columns
                                else 0
                            )

                            curr_proj["data"]["area_door_steel_calc"] = (
                                int(door_df_imported["Pintu Besi"].sum())
                                if "Pintu Besi" in door_df_imported.columns
                                else 0
                            )

                            curr_proj["data"]["area_door_glass_calc"] = (
                                int(door_df_imported["Pintu Kaca"].sum())
                                if "Pintu Kaca" in door_df_imported.columns
                                else 0
                            )

                        if imported_external_records:
                            external_key = f"external_table_{curr_id}"
                            st.session_state[external_key] = copy.deepcopy(imported_external_records)
                            curr_proj["data"]["area_external_table_data"] = copy.deepcopy(
                                imported_external_records
                            )

                        for k, v in imported_landscape_data.items():
                            curr_proj["data"][k] = v

                        if imported_architectural_facade_inputs:
                            for k, v in imported_architectural_facade_inputs.items():
                                curr_proj["data"][k] = v

                            imported_total_floor_height = safe_sum(imported_area_df, F2F_COL)
                            imported_total_typical_units = (
                                int(pd.to_numeric(imported_area_df[TYPICAL_UNIT_COL], errors="coerce").fillna(0).sum())
                                if TYPICAL_UNIT_COL in imported_area_df.columns
                                else 0
                            )
                            imported_keliling_facade = _safe_float(
                                imported_architectural_facade_inputs.get("area_keliling_facade", 0.0)
                            )
                            imported_panjang_railing = _safe_float(
                                imported_architectural_facade_inputs.get("area_panjang_railing", 0.0)
                            )
                            imported_tinggi_railing = _safe_float(
                                imported_architectural_facade_inputs.get("area_tinggi_railing", 0.0)
                            )
                            imported_facade_tolerance_pct = _safe_float(
                                imported_architectural_facade_inputs.get("area_facade_tolerance_pct", 0.0)
                            )
                            imported_facade_wall_area = imported_total_floor_height * imported_keliling_facade
                            imported_facade_railing_area = (
                                imported_total_typical_units
                                * imported_panjang_railing
                                * imported_tinggi_railing
                            )
                            imported_facade_subtotal = imported_facade_wall_area + imported_facade_railing_area
                            imported_facade_tolerance_area = (
                                imported_facade_subtotal
                                * imported_facade_tolerance_pct
                                / 100
                            )
                            imported_total_facade_area = (
                                imported_facade_subtotal
                                + imported_facade_tolerance_area
                            )

                            curr_proj["data"]["area_railing_length_per_room_calc"] = imported_panjang_railing
                            curr_proj["data"]["area_facade_wall_calc"] = imported_facade_wall_area
                            curr_proj["data"]["area_facade_railing_calc"] = imported_facade_railing_area
                            curr_proj["data"]["area_facade_subtotal_calc"] = imported_facade_subtotal
                            curr_proj["data"]["area_facade_tolerance_area_calc"] = imported_facade_tolerance_area
                            curr_proj["data"]["area_facade_calc"] = imported_total_facade_area

                        if imported_res_fac_records:
                            res_fac_key = f"res_fac_table_{curr_id}"
                            st.session_state[res_fac_key] = copy.deepcopy(imported_res_fac_records)
                            curr_proj["data"]["area_res_fac_table_data"] = copy.deepcopy(
                                imported_res_fac_records
                            )

                        if imported_earthwork_rows is not None:
                            earthwork_import_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
                            if earthwork_import_gba <= 0:
                                earthwork_import_gba = safe_sum(imported_area_df, "GBA")

                            (
                                imported_earthwork_rows,
                                imported_earthwork_total,
                                imported_earthwork_derived_unit_price,
                            ) = calculate_earthwork_detail(imported_earthwork_rows, earthwork_import_gba)

                            curr_proj["data"]["earthwork_detail_enabled"] = True
                            curr_proj["data"]["earthwork_detail_rows"] = imported_earthwork_rows
                            curr_proj["data"]["earthwork_detail_total"] = imported_earthwork_total
                            curr_proj["data"]["earthwork_derived_unit_price"] = imported_earthwork_derived_unit_price

                            earthwork_diff_status = get_earthwork_price_difference_status(
                                curr_proj,
                                earthwork_import_gba,
                            )

                            if earthwork_diff_status["has_difference"]:
                                st.session_state["earthwork_import_warning"] = {
                                    "derived_rate": earthwork_diff_status["derived_rate"],
                                    "current_rate": earthwork_diff_status["current_rate"],
                                    "rate_difference": earthwork_diff_status["rate_difference"],
                                    "detail_total": earthwork_diff_status["detail_total"],
                                    "current_total": earthwork_diff_status["current_total"],
                                    "total_difference": earthwork_diff_status["total_difference"],
                                }
                            else:
                                st.session_state.pop("earthwork_import_warning", None)

                        if imported_architectural_rows is not None:
                            architectural_base_values = get_architectural_detail_base_values(
                                curr_proj["data"],
                                imported_area_df,
                            )

                            (
                                imported_architectural_rows,
                                imported_architectural_total,
                                imported_architectural_derived_unit_price,
                            ) = calculate_architectural_detail(imported_architectural_rows, architectural_base_values)

                            curr_proj["data"]["architectural_detail_rows"] = imported_architectural_rows
                            curr_proj["data"]["architectural_detail_total"] = imported_architectural_total
                            curr_proj["data"]["architectural_derived_unit_price"] = imported_architectural_derived_unit_price

                        if imported_consultancy_rows is not None:
                            consultancy_base_values = get_consultancy_detail_base_values(
                                curr_proj["data"],
                                imported_area_df,
                            )

                            imported_consultancy_outputs = get_consultancy_detail_outputs(
                                imported_consultancy_rows,
                                consultancy_base_values,
                            )
                            store_consultancy_detail_outputs(
                                curr_proj["data"],
                                imported_consultancy_outputs,
                            )

                        if imported_foundation_rows is not None:
                            foundation_import_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
                            if foundation_import_gba <= 0:
                                foundation_import_gba = safe_sum(imported_area_df, "GBA")

                            (
                                imported_foundation_rows,
                                imported_foundation_total,
                                imported_foundation_derived_unit_price,
                            ) = calculate_foundation_detail(imported_foundation_rows, foundation_import_gba)

                            curr_proj["data"]["foundation_detail_rows"] = imported_foundation_rows
                            curr_proj["data"]["foundation_detail_total"] = imported_foundation_total
                            curr_proj["data"]["foundation_derived_unit_price"] = imported_foundation_derived_unit_price

                        if imported_ffe_rows is not None:
                            ffe_import_rooms = _safe_float(curr_proj["data"].get("m_rooms", 0.0))
                            if ffe_import_rooms <= 0:
                                ffe_import_rooms = _safe_float(curr_proj["data"].get("area_rooms_calc", 0.0))

                            (
                                imported_ffe_rows,
                                imported_ffe_total,
                                imported_ffe_derived_unit_price,
                            ) = calculate_ffe_detail(imported_ffe_rows, ffe_import_rooms)

                            curr_proj["data"]["ffe_detail_rows"] = imported_ffe_rows
                            curr_proj["data"]["ffe_detail_total"] = imported_ffe_total
                            curr_proj["data"]["ffe_derived_unit_price"] = imported_ffe_derived_unit_price

                        if imported_mep_rows is not None:
                            mep_import_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
                            if mep_import_gba <= 0:
                                mep_import_gba = safe_sum(imported_area_df, "GBA")

                            (
                                imported_mep_rows,
                                imported_mep_total,
                                imported_mep_derived_unit_price,
                            ) = calculate_mep_detail(imported_mep_rows, mep_import_gba)

                            curr_proj["data"]["mep_detail_rows"] = imported_mep_rows
                            curr_proj["data"]["mep_detail_total"] = imported_mep_total
                            curr_proj["data"]["mep_derived_unit_price"] = imported_mep_derived_unit_price

                        if imported_utility_rows is not None:
                            utility_import_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
                            if utility_import_gba <= 0:
                                utility_import_gba = safe_sum(imported_area_df, "GBA")

                            (
                                imported_utility_rows,
                                imported_utility_total,
                                imported_utility_derived_unit_price,
                            ) = calculate_utility_detail(imported_utility_rows, utility_import_gba)

                            curr_proj["data"]["utility_detail_rows"] = imported_utility_rows
                            curr_proj["data"]["utility_detail_total"] = imported_utility_total
                            curr_proj["data"]["utility_derived_unit_price"] = imported_utility_derived_unit_price

                        if imported_structural_rows is not None:
                            structural_import_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
                            if structural_import_gba <= 0:
                                structural_import_gba = safe_sum(imported_area_df, "GBA")

                            (
                                imported_structural_rows,
                                imported_structural_total,
                                imported_structural_derived_unit_price,
                            ) = calculate_structural_detail(imported_structural_rows, structural_import_gba)

                            curr_proj["data"]["structural_detail_rows"] = imported_structural_rows
                            curr_proj["data"]["structural_detail_total"] = imported_structural_total
                            curr_proj["data"]["structural_derived_unit_price"] = imported_structural_derived_unit_price

                        clear_area_editor_state()

                        for stale_key in [
                            door_editor_key,
                            f"other_external_editor_{curr_id}",
                            f"res_fac_editor_{curr_id}",
                            f"architectural_detail_editor_{curr_id}",
                            f"consultancy_detail_editor_{curr_id}",
                            f"earthwork_detail_editor_{curr_id}",
                            f"ffe_detail_editor_{curr_id}",
                            f"foundation_detail_editor_{curr_id}",
                            f"mep_detail_editor_{curr_id}",
                            f"structural_detail_editor_{curr_id}",
                            f"utility_detail_editor_{curr_id}",
                            f"area_keliling_facade_{curr_id}",
                            f"area_panjang_railing_{curr_id}",
                            f"area_tinggi_railing_{curr_id}",
                            f"area_facade_tolerance_pct_{curr_id}",
                        ]:
                            if stale_key in st.session_state:
                                del st.session_state[stale_key]

                        save_ok = save_after_user_action("Area Excel Import")

                        if save_ok:
                            st.success(
                                f"Excel imported successfully. "
                                f"{len(imported_area_records)} area rows loaded."
                            )
                            st.rerun()
                        else:
                            st.error("Cloud save failed. Imported area data changed locally, but was not saved. Do not log out yet.")

                    except ExcelImportError as e:
                        st.error(f"Excel import failed: {e}")
                    except Exception as e:
                        st.error("Excel import failed. Please upload a valid .xlsx file generated from this app.")
                        with st.expander("Technical details"):
                            st.code(str(e))

            earthwork_import_warning = st.session_state.get("earthwork_import_warning")
            if isinstance(earthwork_import_warning, dict):
                warning_derived_rate = _safe_float(
                    earthwork_import_warning.get(
                        "derived_rate",
                        curr_proj.get("data", {}).get("earthwork_derived_unit_price", 0.0),
                    )
                )

                project_type = curr_proj.get("type", "")
                database_price = _safe_float(
                    PROJECT_DATABASE.get(project_type, {}).get("struc_earth", 0.0)
                )

                warn_col, dismiss_col = st.columns([5, 1])

                with warn_col:
                    st.warning(
                        f"Default Earthwork rate from database for {project_type} is "
                        f"Rp {database_price:,.0f}/m2. "
                        f"Your Earthworks detail calculation is "
                        f"Rp {warning_derived_rate:,.0f}/m2. "
                        "Review before finalizing Cost Analysis."
                    )

                with dismiss_col:
                    if st.button(
                        "Dismiss",
                        key=f"dismiss_earthwork_import_warning_{curr_id}",
                        use_container_width=True,
                    ):
                        st.session_state.pop("earthwork_import_warning", None)
                        st.rerun()

    # ==================================================
    # TAB 2 - GBA INPUT DRAFT ONLY
    # ==================================================
    elif area_page == "GBA Input":
        st.subheader("GBA Input Draft")
        st.caption(
            "Edit here first. Calculations and cost sync will not update until you click Save Change."
        )

        c_gen, c_reset = st.columns(2)

        with c_gen:
            generate_clicked = st.button(
                "Generate Table",
                type="secondary",
                key=f"generate_area_table_{curr_id}",
                width="stretch",
            )

        with c_reset:
            reset_clicked = st.button(
                "Reset Draft",
                key=f"reset_area_draft_{curr_id}",
                width="stretch",
            )

        if generate_clicked:
            new_data = generate_area_rows(
                st.session_state[up_key],
                st.session_state[base_key],
            )

            st.session_state[area_draft_key] = clean_area_records(new_data)
            st.session_state[area_committed_key] = clean_area_records(new_data)
            set_data("area_table", copy.deepcopy(st.session_state[area_committed_key]))

            clear_area_editor_state()

            save_ok = save_after_user_action("Generate Area Table")

            if save_ok:
                st.success("New area table generated and committed.")
                st.rerun()
            else:
                st.error("Cloud save failed. Generated area table changed locally, but was not saved. Do not log out yet.")

        if reset_clicked:
            st.session_state[area_draft_key] = copy.deepcopy(
                st.session_state[area_committed_key]
            )

            clear_area_editor_state()

            st.warning("Draft reset to last applied table.")
            st.rerun()

        # Clean draft before showing it in editor, so deleted cells do not display as None
        st.session_state[area_draft_key] = clean_area_records(
            st.session_state[area_draft_key]
        )

        area_input_mode_key = f"area_input_mode_{curr_id}"

        area_input_mode = st.radio(
            "Area Input Mode",
            options=[
                "Detailed Category",
                "Consultant Summary",
            ],
            horizontal=True,
            key=area_input_mode_key,
            help=(
                "Detailed Category = input each physical category. "
                "Consultant Summary = input cumulative NFA, SGFA, and GFA per floor."
            ),
        )

        draft_df = area_records_to_input_view(
            st.session_state[area_draft_key],
            area_input_mode,
        )

        edited_draft_df = st.data_editor(
            draft_df,
            num_rows="dynamic",
            width="stretch",
            key=f"{area_editor_key}_{area_input_mode}",
            hide_index=True,
            column_order=list(draft_df.columns),
            column_config={
                "Space Type": st.column_config.SelectboxColumn(
                    "Space Type",
                    options=[
                        "Roof",
                        "Unit",
                        "Lobby",
                        "Ramp",
                        "Carpark",
                        "Facility",
                        "Office",
                    ],
                    required=True,
                ),
                F2F_COL: st.column_config.NumberColumn(
                    "Height (m)",
                    min_value=0.0,
                    step=0.1,
                    format="%.2f m",
                    help="Vertical height from one finished floor level to the next.",
                ),
                TYPICAL_UNIT_COL: st.column_config.NumberColumn(
                    "Typical Unit",
                    min_value=0,
                    step=1,
                    format="%d",
                    help="Number of typical units on this floor. This is a count, not an area.",
                ),
                UNIT_AREA_COL: st.column_config.NumberColumn(
                    "Unit Area",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                    help="Detailed mode: direct Unit/NFA area.",
                ),
                "Office": st.column_config.NumberColumn(
                    "Office",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                ),
                "Parkir": st.column_config.NumberColumn(
                    "Parkir",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                ),
                "Roof/Deck": st.column_config.NumberColumn(
                    "Roof/Deck",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                ),
                "MEP Outdoor": st.column_config.NumberColumn(
                    "MEP Outdoor",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                ),
                "Koridor/Lobby": st.column_config.NumberColumn(
                    "Koridor/Lobby",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                ),
                "Stair, MEP, Etc": st.column_config.NumberColumn(
                    "Stair, MEP, Etc",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                ),
                "GFA Input": st.column_config.NumberColumn(
                    "GFA",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                    help="Consultant mode: cumulative GFA. Converted to Stair, MEP, Etc = GFA - SGFA.",
                ),
                "SGFA Input": st.column_config.NumberColumn(
                    "SGFA",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                    help="Consultant mode: cumulative SGFA. Converted to Koridor/Lobby = SGFA - NFA.",
                ),
                "NFA Input": st.column_config.NumberColumn(
                    "NFA",
                    min_value=0.0,
                    step=10.0,
                    format="%.2f m2",
                    help="Consultant mode: NFA. Converted to Unit Area.",
                ),
            },
        )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_area_clicked = st.button(
                "Save Change",
                key=f"save_area_table_{curr_id}",
                type="primary",
                width="stretch"
            )

        if save_area_clicked:
            warnings = validate_consultant_summary_input(edited_draft_df)

            if warnings:
                for msg in warnings:
                    st.warning(msg)

            cleaned_records = input_view_to_area_records(
                edited_draft_df,
                area_input_mode,
            )
            cleaned_records = clean_area_records(cleaned_records)

            st.session_state[area_draft_key] = copy.deepcopy(cleaned_records)
            st.session_state[area_committed_key] = copy.deepcopy(cleaned_records)
            set_data("area_table", copy.deepcopy(cleaned_records))

            clear_area_editor_state()

            save_ok = save_data_force()

            if save_ok:
                st.success("Saved to cloud.")
                st.rerun()
            else:
                st.error("Cloud save failed. Do not log out yet.")


    # ==================================================
    # TAB 2 - READ-ONLY GBA SUMMARY
    # ==================================================
    elif area_page == "GBA Summary":
        st.subheader("GBA Summary")
        st.caption("Read-only calculation from the last applied table.")

        total_gba = safe_sum(edited_df, "GBA")
        total_gfa = safe_sum(edited_df, "GFA")
        total_sgfa = safe_sum(edited_df, "SGFA")
        total_nfa = safe_sum(edited_df, "NFA")

        efficiency = (total_nfa / total_gfa * 100) if total_gfa > 0 else 0
        gfa_ratio = (total_gfa / total_gba * 100) if total_gba > 0 else 0
        sgfa_ratio = (total_sgfa / total_gba * 100) if total_gba > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("GBA", f"{total_gba:,.2f} m2")
        k2.metric("GFA", f"{total_gfa:,.2f} m2", f"{gfa_ratio:.1f}% of GBA")
        k3.metric("SGFA", f"{total_sgfa:,.2f} m2", f"{sgfa_ratio:.1f}% of GBA")
        k4.metric("NFA", f"{total_nfa:,.2f} m2", f"Efficiency {efficiency:.1f}%")

        st.markdown("##### Applied Area Table")

        st.dataframe(
            edited_df,
            width="stretch",
            hide_index=True,
            column_config={
                col: st.column_config.NumberColumn(col, format="%.2f")
                for col in breakdown_cols + ["TOTAL", "GBA", "GFA", "SGFA", "NFA"]
            },
        )

        st.markdown("##### Category Total")

        category_summary_cols = breakdown_cols + ["GBA", "GFA", "SGFA", "NFA"]

        total_summary = pd.DataFrame(
            [
                {
                    "Summary": "TOTAL",
                    **{
                        col: (
                            int(edited_df[col].sum())
                            if col == TYPICAL_UNIT_COL
                            else _safe_float(edited_df[col].sum())
                        )
                        for col in category_summary_cols
                        if col in edited_df.columns
                    },
                }
            ]
        )

        st.dataframe(
            total_summary,
            width="stretch",
            hide_index=True,
        )

        # Simple stable Plotly section view
        st.markdown("##### Building Area Section")

        area_cols = [
            "Office",
            UNIT_AREA_COL,
            "Stair, MEP, Etc",
            "Koridor/Lobby",
            "Parkir",
            "MEP Outdoor",
            "Roof/Deck",
        ]

        draw_df = edited_df.iloc[::-1].reset_index(drop=True).copy()

        for col in area_cols + ["GBA"]:
            if col not in draw_df.columns:
                draw_df[col] = 0.0
            draw_df[col] = pd.to_numeric(draw_df[col], errors="coerce").fillna(0.0)

        fig_mass = go.Figure()

        y_positions = list(range(len(draw_df)))
        floor_labels = draw_df["FL"].astype(str).tolist()

        PRO_COLORS = {
            "Unit Area": "#9FBBD6",
            "Office": "#B5CEE5",
            "Koridor/Lobby": "#B6D0AA",
            "Stair, MEP, Etc": "#CCC7BE",
            "Parkir": "#AEB3BA",
            "Roof/Deck": "#D1D8E2",
            "MEP Outdoor": "#BDC5CF",
        }

        hierarchy_groups = [
            {
                "label": "NFA",
                "cols": ["Office", UNIT_AREA_COL],
                "color": "rgba(159,187,214,0.25)",
            },
            {
                "label": "GFA",
                "cols": ["Stair, MEP, Etc"],
                "color": "rgba(204,199,190,0.28)",
            },
            {
                "label": "SGFA",
                "cols": ["Koridor/Lobby"],
                "color": "rgba(182,208,170,0.28)",
            },
            {
                "label": "GBA Only",
                "cols": ["Parkir", "MEP Outdoor", "Roof/Deck"],
                "color": "rgba(174,179,186,0.25)",
            },
        ]

        # Calculate centered base for every segment
        bases = {col: [] for col in area_cols}
        widths = {col: [] for col in area_cols}

        for _, row in draw_df.iterrows():
            gba = _safe_float(row.get("GBA", 0.0))
            curr_x = -gba / 2 if gba > 0 else 0

            for col in area_cols:
                val = _safe_float(row.get(col, 0.0))

                bases[col].append(curr_x)
                widths[col].append(val)

                curr_x += val

        # Minimum segment share before showing inside label
        # Adjust this if labels are too crowded.
        MIN_LABEL_SHARE = 0.11

        text_labels = {col: [] for col in area_cols}
        hover_texts = {col: [] for col in area_cols}

        for _, row in draw_df.iterrows():
            gba = _safe_float(row.get("GBA", 0.0))

            for col in area_cols:
                val = _safe_float(row.get(col, 0.0))
                share = (val / gba) if gba > 0 else 0

                if val > 0 and share >= MIN_LABEL_SHARE:
                    if col == "Koridor/Lobby":
                        label_name = "Corridor"
                    elif col == "Stair, MEP, Etc":
                        label_name = "Service"
                    elif col == "MEP Outdoor":
                        label_name = "MEP Out."
                    elif col == "Roof/Deck":
                        label_name = "Roof"
                    elif col == UNIT_AREA_COL:
                        label_name = "Unit"
                    else:
                        label_name = col

                    text_labels[col].append(f"{label_name}<br>{val:,.0f} sqm")
                else:
                    text_labels[col].append("")

                hover_texts[col].append(
                    f"<b>{col}</b>"
                    f"<br>Area: {val:,.0f} sqm"
                    f"<br>Floor GBA: {gba:,.0f} sqm"
                    f"<br>Share: {share * 100:.1f}%"
                )

        for col in area_cols:
            fig_mass.add_trace(
                go.Bar(
                    y=y_positions,
                    x=widths[col],
                    base=bases[col],
                    name=col,
                    orientation="h",
                    marker=dict(
                        color=PRO_COLORS.get(col, "#DDDDDD"),
                        line=dict(color="#111827", width=0.6),
                    ),
                    text=text_labels[col],
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(
                        color="#111827",
                        size=10,
                        family="Arial",
                    ),
                    hoverinfo="text",
                    hovertext=hover_texts[col],
                    cliponaxis=False,
                )
            )

        max_gba = _safe_float(draw_df["GBA"].max()) if len(draw_df) else 0

        # Optional: outer outline per floor
        floor_bar_height = 0.84

        for i, row in draw_df.iterrows():
            gba = _safe_float(row.get("GBA", 0.0))

            if gba > 0:
                fig_mass.add_shape(
                    type="rect",
                    x0=-gba / 2,
                    x1=gba / 2,
                    y0=i - floor_bar_height / 2,
                    y1=i + floor_bar_height / 2,
                    line=dict(color="#111827", width=1.2),
                    fillcolor="rgba(0,0,0,0)",
                    layer="above",
                )

        fig_mass.update_layout(
            barmode="stack",
            height=max(420, len(draw_df) * 38),
            margin=dict(l=8, r=8, t=30, b=12),
            plot_bgcolor="#FFFFFF",
            paper_bgcolor="#FFFFFF",
            showlegend=True,
            legend=dict(orientation="h", y=1.05),
            uniformtext=dict(
                minsize=8,
                mode="hide",
            ),
            xaxis=dict(
                title=None,
                range=[-max_gba * 0.60, max_gba * 0.60] if max_gba > 0 else None,
                showgrid=False,
                zeroline=True,
                zerolinecolor="#111827",
                zerolinewidth=1,
                showticklabels=False,
                fixedrange=True,
            ),
            yaxis=dict(
                tickmode="array",
                tickvals=y_positions,
                ticktext=floor_labels,
                fixedrange=True,
            ),
        )

        st.plotly_chart(
            fig_mass,
            width="stretch",
            config={"displayModeBar": False},
        )

    # ==================================================
    # TAB 3 - DOORS, SEPARATE DRAFT/COMMIT
    # ==================================================
    elif area_page == "Pintu":
        st.subheader("Area Analysis (Doors)")

        DOOR_WOOD_COL = "Pintu Kayu"
        DOOR_STEEL_COL = "Pintu Besi"
        DOOR_GLASS_COL = "Pintu Kaca"

        if door_committed_key not in st.session_state:
            saved_door = curr_proj["data"].get("area_door_table_data", [])

            if isinstance(saved_door, list) and len(saved_door) > 0:
                st.session_state[door_committed_key] = clean_door_records(saved_door)
            else:
                st.session_state[door_committed_key] = clean_door_records(
                    sync_door_rows_from_area(
                        edited_df,
                        [],
                    )
                )

        if door_draft_key not in st.session_state:
            st.session_state[door_draft_key] = copy.deepcopy(
                st.session_state[door_committed_key]
            )

        # Clean draft before showing it in editor, so deleted cells do not display as None
        st.session_state[door_draft_key] = clean_door_records(
            st.session_state[door_draft_key]
        )

        door_df = pd.DataFrame(st.session_state[door_draft_key])

        edited_door_df = st.data_editor(
            door_df,
            num_rows="fixed",
            width="stretch",
            key=door_editor_key,
            hide_index=True,
            column_order=[
                "FL",
                "Space Type",
                F2F_COL,
                TYPICAL_UNIT_COL,
                DOOR_WOOD_COL,
                DOOR_STEEL_COL,
                DOOR_GLASS_COL,
            ],
            disabled=[
                "FL",
                "Space Type",
                F2F_COL,
                TYPICAL_UNIT_COL,
            ],
            column_config={
                "FL": st.column_config.TextColumn("FL", disabled=True),
                "Space Type": st.column_config.TextColumn("Space Type", disabled=True),
                F2F_COL: st.column_config.NumberColumn(
                    "Height (m)",
                    disabled=True,
                    format="%.2f m",
                ),
                TYPICAL_UNIT_COL: st.column_config.NumberColumn(
                    "Typical Unit",
                    disabled=True,
                    format="%d",
                ),
                DOOR_WOOD_COL: st.column_config.NumberColumn(
                    "Pintu Kayu",
                    min_value=0,
                    step=1,
                    format="%d",
                ),
                DOOR_STEEL_COL: st.column_config.NumberColumn(
                    "Pintu Besi",
                    min_value=0,
                    step=1,
                    format="%d",
                ),
                DOOR_GLASS_COL: st.column_config.NumberColumn(
                    "Pintu Kaca",
                    min_value=0,
                    step=1,
                    format="%d",
                ),
            },
        )

        st.session_state[door_draft_key] = edited_door_df.to_dict("records")

        preview_door_df = pd.DataFrame(
            clean_door_records(edited_door_df.to_dict("records"))
        )

        preview_door_wood = (
            int(preview_door_df[DOOR_WOOD_COL].sum())
            if DOOR_WOOD_COL in preview_door_df.columns
            else 0
        )
        preview_door_steel = (
            int(preview_door_df[DOOR_STEEL_COL].sum())
            if DOOR_STEEL_COL in preview_door_df.columns
            else 0
        )
        preview_door_glass = (
            int(preview_door_df[DOOR_GLASS_COL].sum())
            if DOOR_GLASS_COL in preview_door_df.columns
            else 0
        )

        st.caption(
            f"Current Edited Total Pintu: "
            f"Kayu {preview_door_wood:,} unit | "
            f"Besi {preview_door_steel:,} unit | "
            f"Kaca {preview_door_glass:,} unit"
        )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_door_cloud_clicked = st.button(
                "Save Change",
                key=f"save_door_table_to_cloud_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_door_cloud_clicked:
            if save_door_table_to_cloud(
                edited_door_df=edited_door_df,
                area_df=edited_df,
                door_draft_key=door_draft_key,
                door_committed_key=door_committed_key,
                door_editor_key=door_editor_key,
            ):
                st.rerun()

    # ==================================================
    # TAB 4 - EXTERNAL WORKS
    # kept stable enough; still using editor but isolated from GBA
    # ==================================================
    elif area_page == "Eksternal":
        st.subheader("Area Analysis (External Works)")

        external_key = f"external_table_{curr_id}"
        other_external_editor_key = f"other_external_editor_{curr_id}"

        st.markdown("##### Landscape Works")

        with st.container(border=True):
            l1, l2, l3 = st.columns(3)

            landscape_qty_key = f"landscape_qty_{curr_id}"
            hardscape_pct_key = f"hardscape_pct_{curr_id}"
            softscape_pct_key = f"softscape_pct_{curr_id}"

            l1.number_input(
                "Landscape Area (m2)",
                min_value=0.0,
                step=1.0,
                value=_safe_float(curr_proj["data"].get("area_landscape_qty_calc", 0.0)),
                format="%.2f",
                key=landscape_qty_key,
            )

            l2.number_input(
                "Hardscape %",
                min_value=0.0,
                max_value=100.0,
                step=1.0,
                value=_safe_float(curr_proj["data"].get("area_hardscape_pct_calc", 0.0)),
                format="%.2f",
                key=hardscape_pct_key,
            )

            l3.number_input(
                "Softscape %",
                min_value=0.0,
                max_value=100.0,
                step=1.0,
                value=_safe_float(curr_proj["data"].get("area_softscape_pct_calc", 0.0)),
                format="%.2f",
                key=softscape_pct_key,
            )

            r1, r2 = st.columns(2)

            hardscape_rate_key = f"hardscape_rate_{curr_id}"
            softscape_rate_key = f"softscape_rate_{curr_id}"

            r1.number_input(
                "Hardscape Rate (Rp/m2)",
                min_value=0.0,
                step=50000.0,
                value=_safe_float(curr_proj["data"].get("area_hardscape_rate_calc", 0.0)),
                format="%.0f",
                key=hardscape_rate_key,
            )

            r2.number_input(
                "Softscape Rate (Rp/m2)",
                min_value=0.0,
                step=50000.0,
                value=_safe_float(curr_proj["data"].get("area_softscape_rate_calc", 0.0)),
                format="%.0f",
                key=softscape_rate_key,
            )

            landscape_qty = _safe_float(st.session_state[landscape_qty_key])
            hardscape_pct = _safe_float(st.session_state[hardscape_pct_key])
            softscape_pct = _safe_float(st.session_state[softscape_pct_key])
            hardscape_rate = _safe_float(st.session_state[hardscape_rate_key])
            softscape_rate = _safe_float(st.session_state[softscape_rate_key])

            total_landscape_pct = hardscape_pct + softscape_pct

            if total_landscape_pct != 100 and total_landscape_pct > 0:
                st.warning(
                    f"Hardscape + Softscape = {total_landscape_pct:.2f}%. Check percentage split."
                )

            hardscape_area = landscape_qty * hardscape_pct / 100
            softscape_area = landscape_qty * softscape_pct / 100

            hardscape_amount = hardscape_area * hardscape_rate
            softscape_amount = softscape_area * softscape_rate
            landscape_amount = hardscape_amount + softscape_amount
            landscape_rate = landscape_amount / landscape_qty if landscape_qty > 0 else 0.0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Hardscape Area", f"{hardscape_area:,.2f} m2")
            m2.metric("Softscape Area", f"{softscape_area:,.2f} m2")
            m3.metric("Landscape Rate", f"Rp {landscape_rate:,.0f}/m2")
            m4.metric("Landscape Total", f"Rp {landscape_amount:,.0f}")

        st.markdown("##### Other External Works")

        default_other_records = [
            {"No": "2", "Item": "SBO : PJU", "Unit": "tk", "Qty": 0.0, "Rate": 0.0},
            {"No": "3", "Item": "Drainage System", "Unit": "m2", "Qty": 0.0, "Rate": 0.0},
            {"No": "4", "Item": "Boundary Wall & Gates", "Unit": "m1", "Qty": 0.0, "Rate": 0.0},
            {"No": "5", "Item": "Infrastructure-Access road", "Unit": "m2", "Qty": 0.0, "Rate": 0.0},
            {"No": "6", "Item": "Others", "Unit": "ls", "Qty": 0.0, "Rate": 0.0},
        ]

        if external_key not in st.session_state:
            saved_other = curr_proj["data"].get("area_external_table_data", [])
            st.session_state[external_key] = (
                saved_other if isinstance(saved_other, list) and len(saved_other) > 0 else default_other_records
            )

        other_df = pd.DataFrame(st.session_state[external_key])

        for col in ["No", "Item", "Unit"]:
            if col not in other_df.columns:
                other_df[col] = ""
            other_df[col] = other_df[col].astype(str)

        for col in ["Qty", "Rate"]:
            if col not in other_df.columns:
                other_df[col] = 0.0
            other_df[col] = pd.to_numeric(other_df[col], errors="coerce").fillna(0.0)

        other_df["Amount"] = other_df["Qty"] * other_df["Rate"]

        edited_other_df = st.data_editor(
            other_df,
            num_rows="dynamic",
            width="stretch",
            key=other_external_editor_key,
            hide_index=True,
            column_order=["No", "Item", "Unit", "Qty", "Rate", "Amount"],
            disabled=["Amount"],
            column_config={
                "No": st.column_config.TextColumn("No"),
                "Item": st.column_config.TextColumn("Item"),
                "Unit": st.column_config.SelectboxColumn(
                    "Unit",
                    options=["m2", "m1", "tk", "ls"],
                    required=True,
                ),
                "Qty": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "Rate": st.column_config.NumberColumn("Rate", min_value=0.0, step=50000.0, format="%.0f"),
                "Amount": st.column_config.NumberColumn("Amount", disabled=True, format="%.0f"),
            },
        )

        for col in ["Qty", "Rate"]:
            edited_other_df[col] = pd.to_numeric(edited_other_df[col], errors="coerce").fillna(0.0)

        edited_other_df["Amount"] = edited_other_df["Qty"] * edited_other_df["Rate"]

        saved_other_records = edited_other_df[["No", "Item", "Unit", "Qty", "Rate"]].to_dict("records")
        st.session_state[external_key] = saved_other_records

        total_other_amount = _safe_float(edited_other_df["Amount"].sum())
        total_external_amount = landscape_amount + total_other_amount

        g1, g2, g3 = st.columns(3)
        g1.metric("Landscape Works", f"Rp {landscape_amount:,.0f}")
        g2.metric("Other External Works", f"Rp {total_other_amount:,.0f}")
        g3.metric("External Works Grand Total", f"Rp {total_external_amount:,.0f}")

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_external_clicked = st.button(
                "Save Change",
                key=f"save_external_{curr_id}",
                type="primary",
                width="stretch"
            )

        if save_external_clicked:
            curr_proj["data"]["area_external_table_data"] = saved_other_records
            curr_proj["data"]["area_external_amount_calc"] = total_external_amount
            curr_proj["data"]["area_landscape_qty_calc"] = landscape_qty
            curr_proj["data"]["area_landscape_rate_calc"] = landscape_rate
            curr_proj["data"]["area_landscape_amount_calc"] = landscape_amount
            curr_proj["data"]["area_hardscape_pct_calc"] = hardscape_pct
            curr_proj["data"]["area_softscape_pct_calc"] = softscape_pct
            curr_proj["data"]["area_hardscape_area_calc"] = hardscape_area
            curr_proj["data"]["area_softscape_area_calc"] = softscape_area
            curr_proj["data"]["area_hardscape_rate_calc"] = hardscape_rate
            curr_proj["data"]["area_softscape_rate_calc"] = softscape_rate

            save_ok = save_data_force()

            if save_ok:
                st.success("Saved to cloud.")
                st.rerun()
            else:
                st.error("Cloud save failed. Do not log out yet.")

    # ==================================================
    # TAB 5 - RESIDENTIAL FACILITY
    # ==================================================
    elif area_page == "Residential":
        st.subheader("Area Analysis (Residential Facility Works)")

        res_fac_key = f"res_fac_table_{curr_id}"
        res_fac_editor_key = f"res_fac_editor_{curr_id}"

        default_res_fac_records = [
            {"No": "1", "Item": "Swimming Pool", "Unit": "m2", "Qty": 0.0, "Rate": 0.0},
            {"No": "2", "Item": "Club House / Fitness Centre", "Unit": "ls", "Qty": 0.0, "Rate": 0.0},
            {"No": "3", "Item": "Pool Deck", "Unit": "m2", "Qty": 0.0, "Rate": 0.0},
        ]

        if res_fac_key not in st.session_state:
            saved_res_fac = curr_proj["data"].get("area_res_fac_table_data", [])
            st.session_state[res_fac_key] = (
                saved_res_fac
                if isinstance(saved_res_fac, list) and len(saved_res_fac) > 0
                else default_res_fac_records
            )

        res_fac_df = pd.DataFrame(st.session_state[res_fac_key])

        for col in ["No", "Item", "Unit"]:
            if col not in res_fac_df.columns:
                res_fac_df[col] = ""
            res_fac_df[col] = res_fac_df[col].astype(str)

        for col in ["Qty", "Rate"]:
            if col not in res_fac_df.columns:
                res_fac_df[col] = 0.0
            res_fac_df[col] = pd.to_numeric(res_fac_df[col], errors="coerce").fillna(0.0)

        res_fac_df["Amount"] = res_fac_df["Qty"] * res_fac_df["Rate"]

        edited_res_fac_df = st.data_editor(
            res_fac_df,
            num_rows="dynamic",
            width="stretch",
            key=res_fac_editor_key,
            hide_index=True,
            column_order=["No", "Item", "Unit", "Qty", "Rate", "Amount"],
            disabled=["Amount"],
            column_config={
                "No": st.column_config.TextColumn("No"),
                "Item": st.column_config.TextColumn("Item"),
                "Unit": st.column_config.SelectboxColumn(
                    "Unit",
                    options=["m2", "m1", "tk", "ls", "unit"],
                    required=True,
                ),
                "Qty": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "Rate": st.column_config.NumberColumn("Rate", min_value=0.0, step=50000.0, format="%.0f"),
                "Amount": st.column_config.NumberColumn("Amount", disabled=True, format="%.0f"),
            },
        )

        for col in ["Qty", "Rate"]:
            edited_res_fac_df[col] = pd.to_numeric(edited_res_fac_df[col], errors="coerce").fillna(0.0)

        edited_res_fac_df["Amount"] = edited_res_fac_df["Qty"] * edited_res_fac_df["Rate"]

        total_res_fac_amount = _safe_float(edited_res_fac_df["Amount"].sum())

        saved_res_fac_records = edited_res_fac_df[["No", "Item", "Unit", "Qty", "Rate"]].to_dict("records")

        st.session_state[res_fac_key] = saved_res_fac_records

        st.metric(
            "Residential Facility Total Sync Source",
            f"Rp {total_res_fac_amount:,.0f}",
        )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_res_fac_clicked = st.button(
                "Save Change",
                key=f"save_res_fac_{curr_id}",
                type="primary",
                width="stretch"
            )

        if save_res_fac_clicked:
            curr_proj["data"]["area_res_fac_table_data"] = saved_res_fac_records
            curr_proj["data"]["area_res_fac_amount_calc"] = total_res_fac_amount

            save_ok = save_data_force()

            if save_ok:
                st.success("Saved to cloud.")
                st.rerun()
            else:
                st.error("Cloud save failed. Do not log out yet.")

    # ==================================================
    # TAB 6 - EARTHWORKS
    # ==================================================
    elif area_page == "Earthworks":
        st.subheader("Area Analysis (Earthworks)")
        st.caption("Preview only: this Earthworks breakdown does not control Cost Analysis yet.")

        earthwork_detail_enabled = st.checkbox(
            "Use detailed Earthworks breakdown",
            value=bool(curr_proj["data"].get("earthwork_detail_enabled", False)),
            key=f"earthwork_detail_enabled_{curr_id}",
        )

        saved_earthwork_detail_rows = curr_proj["data"].get(
            "earthwork_detail_rows",
            get_default_earthwork_detail_rows(),
        )
        earthwork_detail_rows, earthwork_detail_total, earthwork_derived_unit_price = calculate_earthwork_detail(
            saved_earthwork_detail_rows,
            _safe_float(curr_proj["data"].get("m_gba", 0.0)) or safe_sum(edited_df, "GBA"),
        )

        edited_earthwork_detail = st.data_editor(
            pd.DataFrame(earthwork_detail_rows),
            key=f"earthwork_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["amount"],
            column_config={
                "code": st.column_config.TextColumn("Item code"),
                "description": st.column_config.TextColumn("Description"),
                "unit": st.column_config.SelectboxColumn("Unit", options=["m2", "ls"]),
                "quantity": st.column_config.NumberColumn("Quantity", min_value=0.0, step=1.0),
                "unit_price": st.column_config.NumberColumn("Unit Price", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        earthwork_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
        if earthwork_gba <= 0:
            earthwork_gba = safe_sum(edited_df, "GBA")

        earthwork_detail_rows, earthwork_detail_total, earthwork_derived_unit_price = calculate_earthwork_detail(
            edited_earthwork_detail,
            earthwork_gba,
        )

        manual_earthwork_price = _safe_float(
            curr_proj["data"].get(
                "u_earth",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("struc_earth", 0.0),
            )
        )
        current_earthworks_total = earthwork_gba * manual_earthwork_price
        earthwork_detail_difference = current_earthworks_total - earthwork_detail_total

        if earthwork_detail_enabled and earthwork_gba <= 0:
            st.warning("Detailed Earthworks preview needs GBA greater than 0 to derive an Earthwork Price.")

        ew_c1, ew_c2, ew_c3 = st.columns(3)
        ew_c1.metric("GBA", f"{earthwork_gba:,.0f} m2")
        ew_c2.metric("Earthworks Detail Total", f"Rp {earthwork_detail_total:,.0f}")
        ew_c3.metric("Derived Earthwork Price", f"Rp {earthwork_derived_unit_price:,.0f}/m2")

        ew_c4, ew_c5, ew_c6 = st.columns(3)
        ew_c4.metric("Current Manual Earthwork Price", f"Rp {manual_earthwork_price:,.0f}/m2")
        ew_c5.metric("Current Earthworks Total", f"Rp {current_earthworks_total:,.0f}")
        ew_c6.metric("Difference", f"Rp {earthwork_detail_difference:,.0f}")

        curr_proj["data"]["earthwork_detail_total"] = earthwork_detail_total
        curr_proj["data"]["earthwork_derived_unit_price"] = earthwork_derived_unit_price
        earthwork_diff_status = get_earthwork_price_difference_status(curr_proj, earthwork_gba)

        if earthwork_diff_status["has_difference"]:
            st.warning(
                "Earthworks detail has changed. "
                f"Derived Earthwork Price is Rp {earthwork_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {earthwork_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )
            diff_c1, diff_c2, diff_c3 = st.columns(3)
            diff_c1.metric("Difference per m2", f"Rp {earthwork_diff_status['rate_difference']:,.0f}/m2")
            diff_c2.metric("Earthworks Detail Total", f"Rp {earthwork_diff_status['detail_total']:,.0f}")
            diff_c3.metric("Current Total Difference", f"Rp {earthwork_diff_status['total_difference']:,.0f}")

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_earthwork_detail_clicked = st.button(
                "Save Earthworks Detail",
                key=f"earthwork_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_earthwork_detail_clicked:
            curr_proj["data"]["earthwork_detail_enabled"] = earthwork_detail_enabled
            curr_proj["data"]["earthwork_detail_rows"] = earthwork_detail_rows
            curr_proj["data"]["earthwork_detail_total"] = earthwork_detail_total
            curr_proj["data"]["earthwork_derived_unit_price"] = earthwork_derived_unit_price

            save_ok = save_after_user_action("Save Earthworks Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 7 - FOUNDATION
    # ==================================================
    elif area_page == "Foundation":
        st.subheader("Area Analysis (Foundation)")
        st.caption("Foundation Detail calculates a suggested Foundation Rate only. Cost Analysis is updated only from the explicit Apply button.")

        foundation_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
        if foundation_gba <= 0:
            foundation_gba = safe_sum(edited_df, "GBA")

        saved_foundation_detail_rows = curr_proj["data"].get(
            "foundation_detail_rows",
            get_default_foundation_detail_rows(),
        )
        foundation_detail_rows, foundation_detail_total, foundation_derived_unit_price = calculate_foundation_detail(
            saved_foundation_detail_rows,
            foundation_gba,
        )

        current_foundation_rate = _safe_float(
            curr_proj["data"].get(
                "u_found",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("struc_found", 0.0),
            )
        )

        edited_foundation_detail = st.data_editor(
            pd.DataFrame(foundation_detail_rows),
            key=f"foundation_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "quantity": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        foundation_detail_rows, foundation_detail_total, foundation_derived_unit_price = calculate_foundation_detail(
            edited_foundation_detail,
            foundation_gba,
        )

        current_foundation_total = foundation_gba * current_foundation_rate
        foundation_rate_difference = foundation_derived_unit_price - current_foundation_rate
        foundation_total_difference = foundation_detail_total - current_foundation_total

        if foundation_gba <= 0:
            st.warning("Foundation Detail needs GBA greater than 0 to derive a Foundation Rate.")

        st.info(
            "Foundation Detail Review\n\n"
            f"GBA: {foundation_gba:,.0f} m2\n\n"
            f"Foundation Detail Total: Rp {foundation_detail_total:,.0f}\n\n"
            f"Derived Foundation Rate: Rp {foundation_derived_unit_price:,.0f}/m2\n\n"
            f"Current Foundation Rate from Cost Analysis: Rp {current_foundation_rate:,.0f}/m2\n\n"
            f"Difference: Rp {foundation_rate_difference:,.0f}/m2 "
            f"(Rp {foundation_total_difference:,.0f} total)"
        )

        fd_c1, fd_c2, fd_c3 = st.columns(3)
        fd_c1.metric("GBA", f"{foundation_gba:,.0f} m2")
        fd_c2.metric("Foundation Detail Total", f"Rp {foundation_detail_total:,.0f}")
        fd_c3.metric("Derived Foundation Rate", f"Rp {foundation_derived_unit_price:,.0f}/m2")

        fd_c4, fd_c5, fd_c6 = st.columns(3)
        fd_c4.metric("Current Foundation Rate", f"Rp {current_foundation_rate:,.0f}/m2")
        fd_c5.metric("Current Foundation Total", f"Rp {current_foundation_total:,.0f}")
        fd_c6.metric("Difference", f"Rp {foundation_rate_difference:,.0f}/m2")

        curr_proj["data"]["foundation_detail_total"] = foundation_detail_total
        curr_proj["data"]["foundation_derived_unit_price"] = foundation_derived_unit_price
        foundation_diff_status = get_foundation_price_difference_status(curr_proj, foundation_gba)

        if foundation_diff_status["has_difference"]:
            st.warning(
                "Foundation detail has changed. "
                f"Derived Foundation Rate is Rp {foundation_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {foundation_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_foundation_detail_clicked = st.button(
                "Save Foundation Detail",
                key=f"foundation_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_foundation_detail_clicked:
            curr_proj["data"]["foundation_detail_rows"] = foundation_detail_rows
            curr_proj["data"]["foundation_detail_total"] = foundation_detail_total
            curr_proj["data"]["foundation_derived_unit_price"] = foundation_derived_unit_price

            save_ok = save_after_user_action("Save Foundation Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 8 - STRUCTURAL
    # ==================================================
    elif area_page == "Structural":
        st.subheader("Area Analysis (Structural)")
        st.caption("Structural Detail calculates a suggested Structural Rate only. Cost Analysis is updated only from the explicit Apply button.")

        structural_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
        if structural_gba <= 0:
            structural_gba = safe_sum(edited_df, "GBA")

        saved_structural_detail_rows = curr_proj["data"].get(
            "structural_detail_rows",
            get_default_structural_detail_rows(),
        )
        structural_detail_rows, structural_detail_total, structural_derived_unit_price = calculate_structural_detail(
            saved_structural_detail_rows,
            structural_gba,
        )

        current_structural_rate = _safe_float(
            curr_proj["data"].get(
                "u_struc",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("struc_work", 0.0),
            )
        )
        current_structural_total = structural_gba * current_structural_rate
        structural_rate_difference = structural_derived_unit_price - current_structural_rate
        structural_total_difference = structural_detail_total - current_structural_total

        edited_structural_detail = st.data_editor(
            pd.DataFrame(structural_detail_rows),
            key=f"structural_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "ratio": st.column_config.NumberColumn("Ratio", min_value=0.0, step=0.01, format="%.4f"),
                "waste_factor": st.column_config.NumberColumn("Waste Factor", min_value=0.0, step=0.01, format="%.4f"),
                "quantity": st.column_config.NumberColumn("Qty", format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        structural_detail_rows, structural_detail_total, structural_derived_unit_price = calculate_structural_detail(
            edited_structural_detail,
            structural_gba,
        )

        current_structural_total = structural_gba * current_structural_rate
        structural_rate_difference = structural_derived_unit_price - current_structural_rate
        structural_total_difference = structural_detail_total - current_structural_total

        if structural_gba <= 0:
            st.warning("Structural Detail needs GBA greater than 0 to derive a Structural Rate.")

        st.info(
            "Structural Detail Review\n\n"
            f"GBA: {structural_gba:,.0f} m2\n\n"
            f"Structural Detail Total: Rp {structural_detail_total:,.0f}\n\n"
            f"Derived Structural Rate: Rp {structural_derived_unit_price:,.0f}/m2\n\n"
            f"Current Structural Rate from Cost Analysis: Rp {current_structural_rate:,.0f}/m2\n\n"
            f"Difference: Rp {structural_rate_difference:,.0f}/m2 "
            f"(Rp {structural_total_difference:,.0f} total)"
        )

        st.markdown("##### Structural Detail Review")
        sd_c1, sd_c2, sd_c3 = st.columns(3)
        sd_c1.metric("GBA", f"{structural_gba:,.0f} m2")
        sd_c2.metric("Structural Detail Total", f"Rp {structural_detail_total:,.0f}")
        sd_c3.metric("Derived Structural Rate", f"Rp {structural_derived_unit_price:,.0f}/m2")

        sd_c4, sd_c5, sd_c6 = st.columns(3)
        sd_c4.metric("Current Structural Rate", f"Rp {current_structural_rate:,.0f}/m2")
        sd_c5.metric("Current Structural Total", f"Rp {current_structural_total:,.0f}")
        sd_c6.metric("Difference", f"Rp {structural_rate_difference:,.0f}/m2")

        curr_proj["data"]["structural_detail_total"] = structural_detail_total
        curr_proj["data"]["structural_derived_unit_price"] = structural_derived_unit_price
        structural_diff_status = get_structural_price_difference_status(curr_proj, structural_gba)

        if structural_diff_status["has_difference"]:
            st.warning(
                "Structural detail has changed. "
                f"Derived Structural Rate is Rp {structural_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {structural_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_structural_detail_clicked = st.button(
                "Save Structural Detail",
                key=f"structural_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_structural_detail_clicked:
            curr_proj["data"]["structural_detail_rows"] = structural_detail_rows
            curr_proj["data"]["structural_detail_total"] = structural_detail_total
            curr_proj["data"]["structural_derived_unit_price"] = structural_derived_unit_price

            save_ok = save_after_user_action("Save Structural Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 9 - ARCHITECTURAL
    # ==================================================
    elif area_page == "Architectural":
        st.subheader("Area Analysis (Architectural)")
        st.caption("Architectural Detail calculates a suggested Architectural Rate only. Cost Analysis is updated only from the explicit Apply button.")

        architectural_base_values = get_architectural_detail_base_values(curr_proj["data"], edited_df)
        architectural_gfa = _safe_float(architectural_base_values.get("gfa", 0.0))

        saved_architectural_detail_rows = curr_proj["data"].get(
            "architectural_detail_rows",
            get_default_architectural_detail_rows(),
        )
        architectural_detail_rows, architectural_detail_total, architectural_derived_unit_price = calculate_architectural_detail(
            saved_architectural_detail_rows,
            architectural_base_values,
        )

        current_architectural_rate = _safe_float(
            curr_proj["data"].get(
                "u_arch",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("arch_base", 0.0),
            )
        )

        base_c1, base_c2, base_c3, base_c4 = st.columns(4)
        base_c1.metric("GFA", f"{architectural_gfa:,.0f} m2")
        base_c2.metric("Facade", f"{_safe_float(architectural_base_values.get('facade', 0.0)):,.0f} m2")
        base_c3.metric("Rooms", f"{_safe_float(architectural_base_values.get('rooms', 0.0)):,.0f}")
        base_c4.metric("Lobby Interior", f"{_safe_float(architectural_base_values.get('lobby', 0.0)):,.0f} m2")

        edited_architectural_detail = st.data_editor(
            pd.DataFrame(architectural_detail_rows),
            key=f"architectural_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "factor": st.column_config.NumberColumn("Factor / %", min_value=0.0, step=1.0, format="%.4f"),
                "overlap": st.column_config.NumberColumn("Overlap", min_value=0.0, step=0.1, format="%.4f"),
                "waste": st.column_config.NumberColumn("Waste", min_value=0.0, step=0.1, format="%.4f"),
                "quantity": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        architectural_detail_rows, architectural_detail_total, architectural_derived_unit_price = calculate_architectural_detail(
            edited_architectural_detail,
            architectural_base_values,
        )

        current_architectural_total = architectural_gfa * current_architectural_rate
        architectural_rate_difference = architectural_derived_unit_price - current_architectural_rate
        architectural_total_difference = architectural_detail_total - current_architectural_total

        if architectural_gfa <= 0:
            st.warning("Architectural Detail needs GFA greater than 0 to derive an Architectural Rate.")

        st.info(
            "Architectural Detail Review\n\n"
            f"GFA: {architectural_gfa:,.0f} m2\n\n"
            f"Architectural Detail Total: Rp {architectural_detail_total:,.0f}\n\n"
            f"Derived Architectural Rate: Rp {architectural_derived_unit_price:,.0f}/m2\n\n"
            f"Current Architectural Rate from Cost Analysis: Rp {current_architectural_rate:,.0f}/m2\n\n"
            f"Difference: Rp {architectural_rate_difference:,.0f}/m2 "
            f"(Rp {architectural_total_difference:,.0f} total)"
        )

        ar_c1, ar_c2, ar_c3 = st.columns(3)
        ar_c1.metric("GFA", f"{architectural_gfa:,.0f} m2")
        ar_c2.metric("Architectural Detail Total", f"Rp {architectural_detail_total:,.0f}")
        ar_c3.metric("Derived Architectural Rate", f"Rp {architectural_derived_unit_price:,.0f}/m2")

        ar_c4, ar_c5, ar_c6 = st.columns(3)
        ar_c4.metric("Current Architectural Rate", f"Rp {current_architectural_rate:,.0f}/m2")
        ar_c5.metric("Current Architectural Total", f"Rp {current_architectural_total:,.0f}")
        ar_c6.metric("Difference", f"Rp {architectural_rate_difference:,.0f}/m2")

        curr_proj["data"]["architectural_detail_total"] = architectural_detail_total
        curr_proj["data"]["architectural_derived_unit_price"] = architectural_derived_unit_price
        architectural_diff_status = get_architectural_price_difference_status(curr_proj, architectural_gfa)

        if architectural_diff_status["has_difference"]:
            st.warning(
                "Architectural detail has changed. "
                f"Derived Architectural Rate is Rp {architectural_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {architectural_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_architectural_detail_clicked = st.button(
                "Save Architectural Detail",
                key=f"architectural_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_architectural_detail_clicked:
            curr_proj["data"]["architectural_detail_rows"] = architectural_detail_rows
            curr_proj["data"]["architectural_detail_total"] = architectural_detail_total
            curr_proj["data"]["architectural_derived_unit_price"] = architectural_derived_unit_price

            save_ok = save_after_user_action("Save Architectural Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 10 - CONSULTANCY
    # ==================================================
    elif area_page == "Consultancy":
        st.subheader("Area Analysis (Consultancy)")
        st.caption("Consultancy Detail calculates a suggested Consultancy Rate only. Cost Analysis is updated only from the explicit Apply button.")

        consultancy_base_values = get_consultancy_detail_base_values(curr_proj["data"], edited_df)
        consultancy_gfa = _safe_float(consultancy_base_values.get("gfa", 0.0))
        consultancy_koridor_lobby = _safe_float(consultancy_base_values.get("koridor_lobby", 0.0))
        consultancy_landscape_qty = _safe_float(consultancy_base_values.get("landscape_qty", 0.0))

        saved_consultancy_detail_rows = curr_proj["data"].get(
            "consultancy_detail_rows",
            get_default_consultancy_detail_rows(),
        )
        consultancy_outputs = get_consultancy_detail_outputs(
            saved_consultancy_detail_rows,
            consultancy_base_values,
        )
        consultancy_detail_rows = consultancy_outputs["consultancy_detail_rows"]
        consultancy_detail_total = consultancy_outputs["consultancy_detail_total"]
        consultancy_derived_unit_price = consultancy_outputs["consultancy_derived_unit_price"]

        current_consultancy_rate = _safe_float(
            curr_proj["data"].get(
                "sc_cons",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("cons", 0.0),
            )
        )

        cons_base_c1, cons_base_c2, cons_base_c3 = st.columns(3)
        cons_base_c1.metric("Current Project GFA", f"{consultancy_gfa:,.0f} m2")
        cons_base_c2.metric("Current Koridor/Lobby", f"{consultancy_koridor_lobby:,.0f} m2")
        cons_base_c3.metric("Current Landscape Qty", f"{consultancy_landscape_qty:,.0f} m2")

        reminder_df = pd.DataFrame([
            {"Source": "Current Project GFA", "Value": consultancy_gfa, "Unit": "m2", "Import Rule": "Reminder only - not imported as a detail row"},
            {"Source": "Current Koridor/Lobby area", "Value": consultancy_koridor_lobby, "Unit": "m2", "Import Rule": "Source for Interior Designer quantity"},
            {"Source": "Current Landscape qty", "Value": consultancy_landscape_qty, "Unit": "m2", "Import Rule": "Source for Landscaping Consultant quantity"},
        ])
        st.dataframe(reminder_df, hide_index=True, width="stretch")

        edited_consultancy_detail = st.data_editor(
            pd.DataFrame(consultancy_detail_rows),
            key=f"consultancy_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "quantity", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "manual_quantity": st.column_config.NumberColumn("Manual Month Qty", min_value=0.0, step=1.0, format="%.2f", help="Used only for Quantity Surveyor and Project Management Fee."),
                "quantity": st.column_config.NumberColumn("Calculated Qty", min_value=0.0, step=1.0, format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        consultancy_outputs = get_consultancy_detail_outputs(
            edited_consultancy_detail,
            consultancy_base_values,
        )
        consultancy_detail_rows = consultancy_outputs["consultancy_detail_rows"]
        consultancy_detail_total = consultancy_outputs["consultancy_detail_total"]
        consultancy_derived_unit_price = consultancy_outputs["consultancy_derived_unit_price"]

        current_consultancy_total = consultancy_gfa * current_consultancy_rate
        consultancy_rate_difference = consultancy_derived_unit_price - current_consultancy_rate
        consultancy_total_difference = consultancy_outputs["consultant_subtotal_excl_qs_pm"] - current_consultancy_total

        if consultancy_gfa <= 0:
            st.warning("Consultancy Detail needs GFA greater than 0 to derive a Consultancy Rate.")

        st.info(
            "Consultancy Detail Review\n\n"
            f"GFA: {consultancy_gfa:,.0f} m2\n\n"
            f"Koridor/Lobby: {consultancy_koridor_lobby:,.0f} m2\n\n"
            f"Landscape Qty: {consultancy_landscape_qty:,.0f} m2\n\n"
            f"Consultancy Detail Total: Rp {consultancy_detail_total:,.0f}\n\n"
            f"Consultant Subtotal excl. QS/PM: Rp {consultancy_outputs['consultant_subtotal_excl_qs_pm']:,.0f}\n\n"
            f"Derived Consultancy Rate excl. QS/PM: Rp {consultancy_derived_unit_price:,.0f}/m2\n\n"
            f"QS from Detail: {consultancy_outputs['qs_duration_from_consultancy']:,.2f} month x Rp {consultancy_outputs['qs_rate_from_consultancy']:,.0f}/month\n\n"
            f"PM from Detail: {consultancy_outputs['pm_duration_from_consultancy']:,.2f} month x Rp {consultancy_outputs['pm_rate_from_consultancy']:,.0f}/month\n\n"
            f"Current Consultancy Rate from Cost Analysis: Rp {current_consultancy_rate:,.0f}/m2\n\n"
            f"Difference: Rp {consultancy_rate_difference:,.0f}/m2 "
            f"(Rp {consultancy_total_difference:,.0f} total)"
        )

        cons_c1, cons_c2, cons_c3 = st.columns(3)
        cons_c1.metric("GFA", f"{consultancy_gfa:,.0f} m2")
        cons_c2.metric("Consultancy Detail Total", f"Rp {consultancy_detail_total:,.0f}")
        cons_c3.metric("Consultant Rate excl. QS/PM", f"Rp {consultancy_derived_unit_price:,.0f}/m2")

        cons_c4, cons_c5, cons_c6 = st.columns(3)
        cons_c4.metric("Current Consultancy Rate", f"Rp {current_consultancy_rate:,.0f}/m2")
        cons_c5.metric("Consultant Subtotal excl. QS/PM", f"Rp {consultancy_outputs['consultant_subtotal_excl_qs_pm']:,.0f}")
        cons_c6.metric("Difference", f"Rp {consultancy_rate_difference:,.0f}/m2")

        store_consultancy_detail_outputs(curr_proj["data"], consultancy_outputs)
        consultancy_diff_status = get_consultancy_price_difference_status(curr_proj, consultancy_gfa)

        if consultancy_diff_status["has_difference"]:
            st.warning(
                "Consultancy detail has changed. "
                f"Derived Consultancy Rate is Rp {consultancy_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {consultancy_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_consultancy_detail_clicked = st.button(
                "Save Consultancy Detail",
                key=f"consultancy_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_consultancy_detail_clicked:
            store_consultancy_detail_outputs(curr_proj["data"], consultancy_outputs)

            save_ok = save_after_user_action("Save Consultancy Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 11 - FF&E
    # ==================================================
    elif area_page == "FF&E":
        st.subheader("Area Analysis (FF&E)")
        st.caption("FF&E Detail calculates a suggested FF&E Rate only. Cost Analysis is updated only from the explicit Apply button.")

        ffe_rooms = _safe_float(curr_proj["data"].get("m_rooms", 0.0))
        if ffe_rooms <= 0:
            ffe_rooms = _safe_float(curr_proj["data"].get("area_rooms_calc", 0.0))
        if ffe_rooms <= 0:
            ffe_rooms = _safe_float(curr_proj["data"].get("area_typical_units_total_calc", 0.0))

        saved_ffe_detail_rows = curr_proj["data"].get(
            "ffe_detail_rows",
            get_default_ffe_detail_rows(),
        )
        ffe_detail_rows, ffe_detail_total, ffe_derived_unit_price = calculate_ffe_detail(
            saved_ffe_detail_rows,
            ffe_rooms,
        )

        current_ffe_rate = _safe_float(
            curr_proj["data"].get(
                "u_ffe",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("ffe", 0.0),
            )
        )

        st.metric("Current Project Rooms", f"{ffe_rooms:,.0f}")
        reminder_df = pd.DataFrame([{
            "Reminder": "Current Project Rooms",
            "Rooms": ffe_rooms,
            "Import Rule": "Reminder only - not imported as a detail row",
        }])
        st.dataframe(reminder_df, hide_index=True, width="stretch")

        edited_ffe_detail = st.data_editor(
            pd.DataFrame(ffe_detail_rows),
            key=f"ffe_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "quantity": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        ffe_detail_rows, ffe_detail_total, ffe_derived_unit_price = calculate_ffe_detail(
            edited_ffe_detail,
            ffe_rooms,
        )

        current_ffe_total = ffe_rooms * current_ffe_rate
        ffe_rate_difference = ffe_derived_unit_price - current_ffe_rate
        ffe_total_difference = ffe_detail_total - current_ffe_total

        if ffe_rooms <= 0:
            st.warning("FF&E Detail needs Rooms greater than 0 to derive an FF&E Rate.")

        st.info(
            "FF&E Detail Review\n\n"
            f"Rooms: {ffe_rooms:,.0f}\n\n"
            f"FF&E Detail Total: Rp {ffe_detail_total:,.0f}\n\n"
            f"Derived FF&E Rate: Rp {ffe_derived_unit_price:,.0f}/room\n\n"
            f"Current FF&E Rate from Cost Analysis: Rp {current_ffe_rate:,.0f}/room\n\n"
            f"Difference: Rp {ffe_rate_difference:,.0f}/room "
            f"(Rp {ffe_total_difference:,.0f} total)"
        )

        fe_c1, fe_c2, fe_c3 = st.columns(3)
        fe_c1.metric("Rooms", f"{ffe_rooms:,.0f}")
        fe_c2.metric("FF&E Detail Total", f"Rp {ffe_detail_total:,.0f}")
        fe_c3.metric("Derived FF&E Rate", f"Rp {ffe_derived_unit_price:,.0f}/room")

        fe_c4, fe_c5, fe_c6 = st.columns(3)
        fe_c4.metric("Current FF&E Rate", f"Rp {current_ffe_rate:,.0f}/room")
        fe_c5.metric("Current FF&E Total", f"Rp {current_ffe_total:,.0f}")
        fe_c6.metric("Difference", f"Rp {ffe_rate_difference:,.0f}/room")

        curr_proj["data"]["ffe_detail_total"] = ffe_detail_total
        curr_proj["data"]["ffe_derived_unit_price"] = ffe_derived_unit_price
        ffe_diff_status = get_ffe_price_difference_status(curr_proj, ffe_rooms)

        if ffe_diff_status["has_difference"]:
            st.warning(
                "FF&E detail has changed. "
                f"Derived FF&E Rate is Rp {ffe_diff_status['derived_rate']:,.0f}/room, "
                f"while Cost Analysis currently uses Rp {ffe_diff_status['current_rate']:,.0f}/room. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_ffe_detail_clicked = st.button(
                "Save FF&E Detail",
                key=f"ffe_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_ffe_detail_clicked:
            curr_proj["data"]["ffe_detail_rows"] = ffe_detail_rows
            curr_proj["data"]["ffe_detail_total"] = ffe_detail_total
            curr_proj["data"]["ffe_derived_unit_price"] = ffe_derived_unit_price

            save_ok = save_after_user_action("Save FF&E Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 11 - MEP
    # ==================================================
    elif area_page == "MEP":
        st.subheader("Area Analysis (MEP)")
        st.caption("MEP Detail calculates a suggested MEP Rate only. Cost Analysis is updated only from the explicit Apply button.")

        mep_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
        if mep_gba <= 0:
            mep_gba = safe_sum(edited_df, "GBA")

        saved_mep_detail_rows = curr_proj["data"].get(
            "mep_detail_rows",
            get_default_mep_detail_rows(),
        )
        mep_detail_rows, mep_detail_total, mep_derived_unit_price = calculate_mep_detail(
            saved_mep_detail_rows,
            mep_gba,
        )

        current_mep_rate = _safe_float(
            curr_proj["data"].get(
                "u_mep",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("mep", 0.0),
            )
        )

        st.metric("Current Project GBA", f"{mep_gba:,.0f} m2")
        reminder_df = pd.DataFrame([{
            "Reminder": "Current Project GBA",
            "GBA": mep_gba,
            "Import Rule": "Reminder only - not imported as a detail row",
        }])
        st.dataframe(reminder_df, hide_index=True, width="stretch")

        edited_mep_detail = st.data_editor(
            pd.DataFrame(mep_detail_rows),
            key=f"mep_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "quantity": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        mep_detail_rows, mep_detail_total, mep_derived_unit_price = calculate_mep_detail(
            edited_mep_detail,
            mep_gba,
        )

        current_mep_total = mep_gba * current_mep_rate
        mep_rate_difference = mep_derived_unit_price - current_mep_rate
        mep_total_difference = mep_detail_total - current_mep_total

        if mep_gba <= 0:
            st.warning("MEP Detail needs GBA greater than 0 to derive an MEP Rate.")

        st.info(
            "MEP Detail Review\n\n"
            f"GBA: {mep_gba:,.0f} m2\n\n"
            f"MEP Detail Total: Rp {mep_detail_total:,.0f}\n\n"
            f"Derived MEP Rate: Rp {mep_derived_unit_price:,.0f}/m2\n\n"
            f"Current MEP Rate from Cost Analysis: Rp {current_mep_rate:,.0f}/m2\n\n"
            f"Difference: Rp {mep_rate_difference:,.0f}/m2 "
            f"(Rp {mep_total_difference:,.0f} total)"
        )

        mep_c1, mep_c2, mep_c3 = st.columns(3)
        mep_c1.metric("GBA", f"{mep_gba:,.0f} m2")
        mep_c2.metric("MEP Detail Total", f"Rp {mep_detail_total:,.0f}")
        mep_c3.metric("Derived MEP Rate", f"Rp {mep_derived_unit_price:,.0f}/m2")

        mep_c4, mep_c5, mep_c6 = st.columns(3)
        mep_c4.metric("Current MEP Rate", f"Rp {current_mep_rate:,.0f}/m2")
        mep_c5.metric("Current MEP Total", f"Rp {current_mep_total:,.0f}")
        mep_c6.metric("Difference", f"Rp {mep_rate_difference:,.0f}/m2")

        curr_proj["data"]["mep_detail_total"] = mep_detail_total
        curr_proj["data"]["mep_derived_unit_price"] = mep_derived_unit_price
        mep_diff_status = get_mep_price_difference_status(curr_proj, mep_gba)

        if mep_diff_status["has_difference"]:
            st.warning(
                "MEP detail has changed. "
                f"Derived MEP Rate is Rp {mep_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {mep_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_mep_detail_clicked = st.button(
                "Save MEP Detail",
                key=f"mep_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_mep_detail_clicked:
            curr_proj["data"]["mep_detail_rows"] = mep_detail_rows
            curr_proj["data"]["mep_detail_total"] = mep_detail_total
            curr_proj["data"]["mep_derived_unit_price"] = mep_derived_unit_price

            save_ok = save_after_user_action("Save MEP Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 12 - UTILITY
    # ==================================================
    elif area_page == "Utility":
        st.subheader("Area Analysis (Utility)")
        st.caption("Utility Detail calculates a suggested Utility Rate only. Cost Analysis is updated only from the explicit Apply button.")

        utility_gba = _safe_float(curr_proj["data"].get("m_gba", 0.0))
        if utility_gba <= 0:
            utility_gba = safe_sum(edited_df, "GBA")

        saved_utility_detail_rows = curr_proj["data"].get(
            "utility_detail_rows",
            get_default_utility_detail_rows(),
        )
        utility_detail_rows, utility_detail_total, utility_derived_unit_price = calculate_utility_detail(
            saved_utility_detail_rows,
            utility_gba,
        )

        current_utility_rate = _safe_float(
            curr_proj["data"].get(
                "u_util",
                PROJECT_DATABASE.get(curr_proj.get("type", "Hotel"), {}).get("utility", 0.0),
            )
        )

        st.metric("Current Project GBA", f"{utility_gba:,.0f} m2")
        reminder_df = pd.DataFrame([{
            "Reminder": "Current Project GBA",
            "GBA": utility_gba,
            "Import Rule": "Reminder only - not imported as a detail row",
        }])
        st.dataframe(reminder_df, hide_index=True, width="stretch")

        edited_utility_detail = st.data_editor(
            pd.DataFrame(utility_detail_rows),
            key=f"utility_detail_editor_{curr_id}",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=["code", "description", "unit", "amount"],
            column_config={
                "code": st.column_config.TextColumn("No"),
                "description": st.column_config.TextColumn("Description", width="large"),
                "unit": st.column_config.TextColumn("Unit"),
                "quantity": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Rate", min_value=0.0, step=100000.0, format="Rp %.0f"),
                "amount": st.column_config.NumberColumn("Amount", format="Rp %.0f"),
            },
        )

        utility_detail_rows, utility_detail_total, utility_derived_unit_price = calculate_utility_detail(
            edited_utility_detail,
            utility_gba,
        )

        current_utility_total = utility_gba * current_utility_rate
        utility_rate_difference = utility_derived_unit_price - current_utility_rate
        utility_total_difference = utility_detail_total - current_utility_total

        if utility_gba <= 0:
            st.warning("Utility Detail needs GBA greater than 0 to derive a Utility Rate.")

        st.info(
            "Utility Detail Review\n\n"
            f"GBA: {utility_gba:,.0f} m2\n\n"
            f"Utility Detail Total: Rp {utility_detail_total:,.0f}\n\n"
            f"Derived Utility Rate: Rp {utility_derived_unit_price:,.0f}/m2\n\n"
            f"Current Utility Rate from Cost Analysis: Rp {current_utility_rate:,.0f}/m2\n\n"
            f"Difference: Rp {utility_rate_difference:,.0f}/m2 "
            f"(Rp {utility_total_difference:,.0f} total)"
        )

        utility_c1, utility_c2, utility_c3 = st.columns(3)
        utility_c1.metric("GBA", f"{utility_gba:,.0f} m2")
        utility_c2.metric("Utility Detail Total", f"Rp {utility_detail_total:,.0f}")
        utility_c3.metric("Derived Utility Rate", f"Rp {utility_derived_unit_price:,.0f}/m2")

        utility_c4, utility_c5, utility_c6 = st.columns(3)
        utility_c4.metric("Current Utility Rate", f"Rp {current_utility_rate:,.0f}/m2")
        utility_c5.metric("Current Utility Total", f"Rp {current_utility_total:,.0f}")
        utility_c6.metric("Difference", f"Rp {utility_rate_difference:,.0f}/m2")

        curr_proj["data"]["utility_detail_total"] = utility_detail_total
        curr_proj["data"]["utility_derived_unit_price"] = utility_derived_unit_price
        utility_diff_status = get_utility_price_difference_status(curr_proj, utility_gba)

        if utility_diff_status["has_difference"]:
            st.warning(
                "Utility detail has changed. "
                f"Derived Utility Rate is Rp {utility_diff_status['derived_rate']:,.0f}/m2, "
                f"while Cost Analysis currently uses Rp {utility_diff_status['current_rate']:,.0f}/m2. "
                "Cost Analysis has not been updated automatically."
            )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_utility_detail_clicked = st.button(
                "Save Utility Detail",
                key=f"utility_detail_save_{curr_id}",
                type="primary",
                width="stretch",
            )

        if save_utility_detail_clicked:
            curr_proj["data"]["utility_detail_rows"] = utility_detail_rows
            curr_proj["data"]["utility_detail_total"] = utility_detail_total
            curr_proj["data"]["utility_derived_unit_price"] = utility_derived_unit_price

            save_ok = save_after_user_action("Save Utility Detail")

            if save_ok:
                st.rerun()

    # ==================================================
    # TAB 13 - DETAILS
    # ==================================================
    elif area_page == "Details":
        st.subheader("Details")
        st.caption("Calculated detail areas from the last applied GBA table.")

        summary_cols = [
            TYPICAL_UNIT_COL,
            "Parkir",
            "Roof/Deck",
            "MEP Outdoor",
            "Koridor/Lobby",
            "Stair, MEP, Etc",
            UNIT_AREA_COL,
            "Office",
            "GBA",
            "GFA",
            "SGFA",
            "NFA",
        ]

        summary_cols = [c for c in summary_cols if c in edited_df.columns]

        total_summary = pd.DataFrame(
            [
                {
                    "Summary": "TOTAL",
                    **{
                        col: (
                            int(edited_df[col].sum())
                            if col == TYPICAL_UNIT_COL
                            else _safe_float(edited_df[col].sum())
                        )
                        for col in summary_cols
                    },
                }
            ]
        )

        st.markdown("##### Total Area Summary From Applied GBA Table")

        st.dataframe(
            total_summary,
            width="stretch",
            hide_index=True,
        )

        st.markdown("##### Area Summary by Space Type")

        group_cols = [
            TYPICAL_UNIT_COL,
            "Parkir",
            "Roof/Deck",
            "MEP Outdoor",
            "Koridor/Lobby",
            "Stair, MEP, Etc",
            UNIT_AREA_COL,
            "Office",
            "GBA",
            "GFA",
            "SGFA",
            "NFA",
        ]

        group_cols = [c for c in group_cols if c in edited_df.columns]

        area_by_space_type = (
            edited_df.groupby("Space Type", dropna=False)[group_cols]
            .sum()
            .reset_index()
        )

        if TYPICAL_UNIT_COL in area_by_space_type.columns:
            area_by_space_type[TYPICAL_UNIT_COL] = area_by_space_type[
                TYPICAL_UNIT_COL
            ].astype(int)

        st.dataframe(
            area_by_space_type,
            width="stretch",
            hide_index=True,
        )

        st.markdown("##### Facade Area")

        facade_formula_df = pd.DataFrame(
            [
                {
                    "Item": "Facade Wall Area",
                    "Formula": "Total Floor Height x Keliling Facade",
                    "Input 1": f"{total_floor_height:,.2f} m",
                    "Input 2": f"{keliling_facade:,.2f} m",
                    "Result": facade_wall_area,
                },
                {
                    "Item": "Railing Area",
                    "Formula": "Total Typical Unit x Panjang Railing x Tinggi Railing",
                    "Input 1": f"{total_typical_units:,} unit",
                    "Input 2": f"{panjang_railing:,.2f} m x {tinggi_railing:,.2f} m",
                    "Result": facade_railing_area,
                },
                {
                    "Item": "Subtotal",
                    "Formula": "Facade Wall Area + Railing Area",
                    "Input 1": "-",
                    "Input 2": "-",
                    "Result": facade_subtotal,
                },
                {
                    "Item": "Tekukan dan Overlap",
                    "Formula": f"{facade_tolerance_pct:.2f}%",
                    "Input 1": "-",
                    "Input 2": "-",
                    "Result": facade_tolerance_area,
                },
                {
                    "Item": "Total Facade Area",
                    "Formula": "Subtotal + Tekukan dan Overlap",
                    "Input 1": "-",
                    "Input 2": "-",
                    "Result": total_facade_area,
                },
            ]
        )

        st.dataframe(
            facade_formula_df,
            width="stretch",
            hide_index=True,
            column_config={
                "Result": st.column_config.NumberColumn(
                    "Result (m2)",
                    format="%.2f",
                )
            },
        )

        fc1, fc2, fc3, fc4 = st.columns(4)

        fc1.metric("Total Floor Height", f"{total_floor_height:,.2f} m")
        fc2.metric("Facade Wall Area", f"{facade_wall_area:,.2f} m2")
        fc3.metric("Railing Area", f"{facade_railing_area:,.2f} m2")
        fc4.metric(
            "Facade Sync Source",
            f"{total_facade_area:,.2f} m2",
            help="This value syncs into cost estimator facade quantity.",
        )

        st.markdown("##### Cost Analysis Sync Sources")

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Lobby Interior Sync Source",
            f"{total_lobby_interior:,.0f} m2",
            help="Calculated from total Koridor/Lobby in applied GBA table.",
        )

        c2.metric(
            "Ruang / Units Sync Source",
            f"{total_typical_units:,} unit",
            help="Calculated from total Typical Unit in applied GBA table.",
        )

        c3.metric(
            "Facade Area Sync Source",
            f"{total_facade_area:,.0f} m2",
        )
        
def update_price(metric_key, db_key): #this function pulls~
    """Update flooring price based on spec radio selection."""
    c_id = st.session_state.current_proj_id
    
    # 1. Safety check: Ensure the project actually exists
    if c_id not in st.session_state.projects:
        return 

    p_type = st.session_state.projects[c_id]["type"]
    c_type_key = f"{c_id}_{p_type}"
    widget_key = f"temp_spec_{metric_key}_{c_type_key}"
    
    # 2. Use .get() to securely fetch the value without throwing a KeyError
    selected_spec = st.session_state.get(widget_key)
    
    # 3. If the widget was destroyed or doesn't exist, quietly exit
    if selected_spec is None:
        return

    # 4. Proceed with normal update
    st.session_state.projects[c_id]["data"][f"{metric_key}_spec_type"] = selected_spec
    db_val = PROJECT_DATABASE.get(p_type, {}).get(db_key, {})
    
    if isinstance(db_val, dict):
        new_val = db_val.get(selected_spec, 0.0)
        st.session_state[f"u_fl_{metric_key}_{c_type_key}"] = _safe_float(new_val)

def show_cost_estimator(): #cost calculator page
    st.title("Cost Analysis")

    st.markdown("""
        <style>
            .metric-container {
                position: relative;
                display: inline-block;
                width: 100%;
            }
            .custom-tooltip {
                visibility: hidden;
                width: 240px;
                background-color: #FFFFFF;
                color: #262730;
                text-align: left;
                border-radius: 8px;
                padding: 12px;
                position: absolute;
                z-index: 1000;
                bottom: 110%; 
                left: 50%;
                transform: translateX(-50%);
                opacity: 0;
                transition: opacity 0.3s;
                font-size: 12px;
                border: 1px solid #CDDC39;
                box-shadow: 0px 4px 15px rgba(0,0,0,0.4);
                line-height: 1.5;
            }
            .metric-container:hover .custom-tooltip {
                visibility: visible;
                opacity: 1;
            }
            /* Triangle arrow for tooltip */
            .custom-tooltip::after {
                content: "";
                position: absolute;
                top: 100%;
                left: 50%;
                margin-left: -5px;
                border-width: 5px;
                border-style: solid;
                border-color: #262730 transparent transparent transparent;
            }
        </style>
        """, unsafe_allow_html=True)

    curr_id, curr_proj = get_current_project()

    if "data" not in curr_proj:
        st.session_state.projects[curr_id]["data"] = {}

    def get_val(key, default=0.0):
        data_dict = st.session_state.projects[curr_id]["data"]
        val = data_dict.get(key, default)
        
        # If the value is a list (for Custom Items), return it immediately
        if isinstance(val, list):
            return val
            
        # For everything else (numbers/strings), try to force to _safe_float
        try:
            return _safe_float(val)
        except (ValueError, TypeError):
            # If it's not a number (like "Type1"), return it as is
            return val

    # --- PROJECT SETUP ---
    curr_type = curr_proj["type"]
    pt_data = PROJECT_DATABASE[curr_type]
    curr_type_key = f"{curr_id}_{curr_type}"

    # --- TABS ---
    tab2, tab3, tab5, tab4, tab6, tab7, tab8 = st.tabs([
        "1. Ukuran", "2. Rasio",
        "3. Soft Costs", "4. Harga",
        "5. Item Tambahan", "6. Hasil",
        "7. Pembuktian",
    ])

    with tab2:
        save_cost_placeholder = st.empty()

        with save_cost_placeholder.container():
            save_cost_clicked = st.button(
                "Save Cost Analysis",
                key=f"save_cost_analysis_{curr_id}",
                type="primary",
                use_container_width=True
            )

        if save_cost_clicked:
            save_ok = save_after_user_action("Save Cost Analysis")

            if save_ok:
                st.success("Cost Analysis saved to cloud.")
                st.rerun()
            else:
                st.error("Cost Analysis changed locally, but cloud save failed. Do not log out yet.")

        # ==================================================
        # IMPORT AREA ANALYSIS TOTALS INTO COST ANALYSIS
        # ==================================================
        area_table_data = curr_proj.get("data", {}).get("area_table_data", [])
        area_totals = calculate_area_totals_from_table(area_table_data)
        area_land = _safe_float(curr_proj.get("data", {}).get("m_land", 0.0))
        area_lobby_interior = _safe_float(curr_proj.get("data", {}).get("area_lobby_interior_calc", 0.0))
        area_rooms = int(_safe_float(curr_proj.get("data", {}).get("area_rooms_calc", 0)))
        suggested_facade = _safe_float(curr_proj.get("data", {}).get("area_facade_calc", 0.0))
        suggested_door_wood = int(_safe_float(curr_proj["data"].get("area_door_wood_calc", 0)))
        suggested_door_steel = int(_safe_float(curr_proj["data"].get("area_door_steel_calc", 0)))
        suggested_door_glass = int(_safe_float(curr_proj["data"].get("area_door_glass_calc", 0)))
        suggested_railing_qty = _safe_float(
            curr_proj["data"].get(
                "area_railing_length_per_room_calc",
                curr_proj["data"].get("area_panjang_railing", 0.0)
            )
        )
        suggested_external_total = _safe_float(
            curr_proj.get("data", {}).get("area_external_amount_calc", 0.0)
        )

        current_external_rate = _safe_float(
            curr_proj.get("data", {}).get("u_ext", pt_data["ext_land"])
        )

        if suggested_external_total > 0 and current_external_rate > 0:
            suggested_landscape_area = suggested_external_total / current_external_rate
        else:
            suggested_landscape_area = 0.0

        suggested_res_fac_total = _safe_float(
            curr_proj.get("data", {}).get("area_res_fac_amount_calc", 0.0)
        )

        current_res_fac_rate = _safe_float(
            curr_proj.get("data", {}).get("u_fac_res", pt_data["fac_res"])
        )

        if suggested_res_fac_total > 0 and current_res_fac_rate > 0:
            suggested_res_fac_area = suggested_res_fac_total / current_res_fac_rate
        else:
            suggested_res_fac_area = 0.0

        if isinstance(area_table_data, list) and len(area_table_data) > 0:
            area_sync_df = pd.DataFrame(area_table_data)
            if "Koridor/Lobby" in area_sync_df.columns:
                area_lobby_interior = _safe_float(
                    pd.to_numeric(
                        area_sync_df["Koridor/Lobby"],
                        errors="coerce"
                    ).fillna(0.0).sum()
                )
            if "Typical Unit" in area_sync_df.columns:
                area_rooms = int(
                    pd.to_numeric(
                        area_sync_df["Typical Unit"],
                        errors="coerce"
                    ).fillna(0).sum()
                )

        sync_groups = [
            {
                "title": "Core Area",
                "items": [
                    {"label": "Luas Tanah", "data_key": "m_land", "widget_key": "m_land", "value": area_land, "unit": "m2"},
                    {"label": "GBA", "data_key": "m_gba", "widget_key": "m_gba", "value": area_totals["gba"], "unit": "m2"},
                    {"label": "GFA", "data_key": "m_gfa", "widget_key": "m_gfa", "value": area_totals["gfa"], "unit": "m2"},
                    {"label": "SGFA", "data_key": "m_sgfa", "widget_key": "m_sgfa", "value": area_totals["sgfa"], "unit": "m2"},
                    {"label": "NFA", "data_key": "m_nfa", "widget_key": "m_nfa", "value": area_totals["nfa"], "unit": "m2"},
                    {"label": "Lobby/Koridor", "data_key": "m_lobby", "widget_key": "m_lobby", "value": area_lobby_interior, "unit": "m2"},
                    {"label": "Rooms/Units", "data_key": "m_rooms", "widget_key": "m_rooms", "value": area_rooms, "unit": "unit"},
                ],
            },
            {
                "title": "Opening / External",
                "items": [
                    {"label": "Facade", "data_key": "m_facade", "widget_key": "m_facade", "value": suggested_facade, "unit": "m2"},
                    {"label": "Landscape", "data_key": "m_land_m2", "widget_key": "m_land_m2", "value": suggested_landscape_area, "unit": "m2"},
                    {"label": "Residential Facility", "data_key": "m_fac_res", "widget_key": "m_fac_res", "value": suggested_res_fac_area, "unit": "m2"},
                    {"label": "Railing", "data_key": "r_rail_qty", "widget_key": "r_rail_qty", "value": suggested_railing_qty, "unit": "m'/room"},
                    {"label": "Wooden Door", "data_key": "m_door_w", "widget_key": "m_door_w", "value": suggested_door_wood, "unit": "unit"},
                    {"label": "Steel Door", "data_key": "m_door_s", "widget_key": "m_door_s", "value": suggested_door_steel, "unit": "unit"},
                    {"label": "Glass Door", "data_key": "m_door_g", "widget_key": "m_door_g", "value": suggested_door_glass, "unit": "unit"},
                ],
            },
            {
                "title": "Detail-Derived Rates",
                "items": [
                    {"label": "Earthworks Rate", "data_key": "u_earth", "widget_key": "u_earth", "session_key": f"u_earth_{curr_type_key}", "value": curr_proj["data"].get("earthwork_derived_unit_price", 0.0), "unit": "Rp/m2"},
                    {"label": "Foundation Rate", "data_key": "u_found", "widget_key": "u_found", "session_key": f"u_found_{curr_type_key}", "value": curr_proj["data"].get("foundation_derived_unit_price", 0.0), "unit": "Rp/m2"},
                    {"label": "Structural Rate", "data_key": "u_struc", "widget_key": "u_struc", "session_key": f"u_struc_{curr_type_key}", "value": curr_proj["data"].get("structural_derived_unit_price", 0.0), "unit": "Rp/m2"},
                    {"label": "Architectural Rate", "data_key": "u_arch", "widget_key": "u_arch", "session_key": f"u_arch_{curr_type_key}", "value": curr_proj["data"].get("architectural_derived_unit_price", 0.0), "unit": "Rp/m2"},
                    {"label": "FF&E Rate", "data_key": "u_ffe", "widget_key": "u_ffe", "session_key": f"u_ffe_{curr_type_key}", "value": curr_proj["data"].get("ffe_derived_unit_price", 0.0), "unit": "Rp/room"},
                    {"label": "MEP Rate", "data_key": "u_mep", "widget_key": "u_mep", "session_key": f"u_mep_{curr_type_key}", "value": curr_proj["data"].get("mep_derived_unit_price", 0.0), "unit": "Rp/m2"},
                    {"label": "Utility Rate", "data_key": "u_util", "widget_key": "u_util", "session_key": f"u_util_{curr_type_key}", "value": curr_proj["data"].get("utility_derived_unit_price", 0.0), "unit": "Rp/m2"},
                    {"label": "QS Duration", "data_key": "sc_qs_m", "widget_key": "sc_qs_m", "value": curr_proj["data"].get("qs_duration_from_consultancy", 0.0), "unit": "month"},
                    {"label": "QS Monthly Rate", "data_key": "sc_qs_r", "widget_key": "sc_qs_r", "value": curr_proj["data"].get("qs_rate_from_consultancy", 0.0), "unit": "Rp/month"},
                    {"label": "PM Duration", "data_key": "sc_pm_m", "widget_key": "sc_pm_m", "value": curr_proj["data"].get("pm_duration_from_consultancy", 0.0), "unit": "month"},
                    {"label": "PM Monthly Rate", "data_key": "sc_pm_r", "widget_key": "sc_pm_r", "value": curr_proj["data"].get("pm_rate_from_consultancy", 0.0), "unit": "Rp/month"},
                    {"label": "Consultant Rate excl. QS/PM", "data_key": "sc_cons", "widget_key": "sc_cons", "session_key": f"sc_cons_{curr_type_key}", "value": curr_proj["data"].get("consultant_rate_excl_qs_pm", curr_proj["data"].get("consultancy_derived_unit_price", 0.0)), "unit": "Rp/m2"},
                ],
            },
        ]

        syncable_items = [
            item
            for group in sync_groups
            for item in group["items"]
            if _safe_float(item.get("value", 0.0)) > 0
        ]

        sync_col1, sync_col2 = st.columns([1, 3], vertical_alignment="center")
        confirm_sync_key = f"confirm_use_area_analysis_{curr_id}"
        show_sync_details_key = f"show_area_sync_details_{curr_id}"

        def sync_area_value(data_key, widget_key, value, session_key=None):
            """
            Sync Area Analysis value into Cost Analysis only if value > 0.
            This prevents empty / unfilled Area Analysis values from overwriting existing Cost Analysis inputs.
            """
            value = _safe_float(value)

            if value > 0:
                st.session_state.projects[curr_id]["data"][data_key] = value
                st.session_state[session_key or f"{widget_key}_{curr_id}"] = value

        def perform_area_analysis_sync():
            for group in sync_groups:
                for item in group["items"]:
                    sync_area_value(
                        item["data_key"],
                        item["widget_key"],
                        item["value"],
                        item.get("session_key"),
                    )

            return save_after_user_action("Use Area Analysis in Cost Analysis")

        with sync_col1:
            if st.button(
                "Use Area Analysis",
                key=f"use_area_analysis_{curr_id}",
                type="primary",
                width="stretch",
                icon=mi("sync") if "mi" in globals() else None,
            ):
                st.session_state[confirm_sync_key] = True
                st.session_state[show_sync_details_key] = True

            on = st.toggle(
                "Show details of synced values",
                key=show_sync_details_key,
            )

        if st.session_state.get(confirm_sync_key, False):
            st.warning(
                "Review the values below. This will update Cost Analysis inputs using nonzero values from Area Analysis and detail-derived rates. Existing Cost Analysis values will not be overwritten by zero values."
            )

            confirm_col, cancel_col, _ = st.columns([1, 1, 4])

            with confirm_col:
                confirm_sync_clicked = st.button(
                    "Confirm Sync",
                    key=f"confirm_area_analysis_sync_{curr_id}",
                    type="primary",
                    width="stretch",
                )

            with cancel_col:
                cancel_sync_clicked = st.button(
                    "Cancel",
                    key=f"cancel_area_analysis_sync_{curr_id}",
                    width="stretch",
                )

            if confirm_sync_clicked:
                save_ok = perform_area_analysis_sync()

                if save_ok:
                    st.session_state.pop(confirm_sync_key, None)
                    st.success("Area metrics, opening/external quantities, and detail-derived rates applied to Cost Analysis.")
                    st.rerun()
                else:
                    st.error("Area Analysis values were applied locally, but cloud save failed. Do not log out yet.")

            if cancel_sync_clicked:
                st.session_state.pop(confirm_sync_key, None)
                st.rerun()

        if on:
            with sync_col2:
                if syncable_items:
                    st.info("Source values available from Area Analysis. Values of 0 will not overwrite existing Cost Analysis inputs. Consultancy Rate excludes QS and PM to avoid double counting; QS and PM sync to their monthly duration/rate fields.")

                    for group in sync_groups:
                        preview_rows = []
                        for item in group["items"]:
                            value = _safe_float(item.get("value", 0.0))
                            preview_rows.append({
                                "Value": item["label"],
                                "Source Value": value,
                                "Unit": item.get("unit", ""),
                                "Will Sync": "Yes" if value > 0 else "No - zero skipped",
                            })

                        st.markdown(f"**{group['title']}**")
                        st.dataframe(
                            pd.DataFrame(preview_rows),
                            hide_index=True,
                            width="stretch",
                            column_config={
                                "Source Value": st.column_config.NumberColumn(
                                    "Source Value",
                                    format="%.2f",
                                ),
                            },
                        )
                else:
                    st.info("No data found in Area Analysis to sync.")
            
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)

        with col_m1:
            with st.expander("Ukuran Proyek", expanded=True):
                st.subheader("Ukuran")
                c_land_input, c_land_info = st.columns([5, 1])
                with c_land_input:
                    land_area = st.number_input("Luas Tanah (m2)", value=get_val("m_land", 0.0), step=100.0, key=f"m_land_{curr_id}")
                with c_land_info:
                    st.write("")
                    with st.popover(" "):
                        st.info(
                            f"""
**Luas Tanah**
Used as project land reference.
Current Luas Tanah: {_safe_float(land_area):,.0f} m2
                            """
                        )

                c_gba_input, c_gba_info = st.columns([5, 1])
                with c_gba_input:
                    gba = st.number_input("GBA (m2)", value=get_val("m_gba", 0.0), step=100.0, key=f"m_gba_{curr_id}")
                with c_gba_info:
                    st.write("")
                    with st.popover(" "):
                        gba_qty = _safe_float(gba)
                        gba_affected = [
                            ("Earthwork", _safe_float(get_val("u_earth", pt_data.get("struc_earth", 0.0)))),
                            ("Foundation", _safe_float(get_val("u_found", pt_data.get("struc_found", 0.0)))),
                            ("Structure", _safe_float(get_val("u_struc", pt_data.get("struc_work", 0.0)))),
                            ("MEP", _safe_float(get_val("u_mep", pt_data.get("mep", 0.0)))),
                            ("Utility", _safe_float(get_val("u_util", pt_data.get("utility", 0.0)))),
                        ]

                        for item_name, rate in gba_affected:
                            total = rate * gba_qty
                            st.info(
                                f"""
**{item_name}**
Hitungan: Rp {rate:,.0f} x GBA: {gba_qty:,.0f} m2
Total {item_name}: Rp {total:,.0f}
Terbilang: {n2w(total)}
                                """
                            )

                c_gfa_input, c_gfa_info = st.columns([5, 1])
                with c_gfa_input:
                    gfa = st.number_input("GFA (m2)", value=get_val("m_gfa", 0.0), step=100.0, key=f"m_gfa_{curr_id}")
                with c_gfa_info:
                    st.write("")
                    with st.popover(" "):
                        gfa_qty = _safe_float(gfa)
                        gfa_affected = [
                            ("Architecture", _safe_float(get_val("u_arch", pt_data.get("arch_base", 0.0)))),
                            ("Consultancy", _safe_float(get_val("sc_cons", pt_data.get("cons", 0.0)))),
                        ]

                        for item_name, rate in gfa_affected:
                            total = rate * gfa_qty
                            st.info(
                                f"""
**{item_name}**
Hitungan: Rp {rate:,.0f} x GFA: {gfa_qty:,.0f} m2
Total {item_name}: Rp {total:,.0f}
Terbilang: {n2w(total)}
                                """
                            )

                c_sgfa_input, c_sgfa_info = st.columns([5, 1])
                with c_sgfa_input:
                    sgfa = st.number_input("SGFA (m2)", value=get_val("m_sgfa", 0.0), step=100.0, key=f"m_sgfa_{curr_id}")
                with c_sgfa_info:
                    st.write("")
                    with st.popover(" "):
                        st.info(
                            f"""
**SGFA**
Used mainly for cost ratio/reporting reference.
Current SGFA: {_safe_float(sgfa):,.0f} m2
                            """
                        )

        with col_m2:
            with st.expander("Arsitektur", expanded=True):
                st.subheader("Interior")
                rooms = st.number_input(
                    "Ruang (unit)",
                    help="cth. 500 unit untuk 1 proyek Apartement A",
                    value=int(_safe_float(get_val("m_rooms", 0))),
                    step=1,
                    key=f"m_rooms_{curr_id}"
                )

                lobby_interior = st.number_input(
                    "Lobby Interior (m2)",
                    help="cth. 500 m2 lobby untuk 1 proyek Apartement A",
                    value=_safe_float(get_val("m_lobby", 0.0)),
                    step=10.0,
                    key=f"m_lobby_{curr_id}"
                )

                carpet_m2 = st.number_input("Karpet (m2)", value=get_val("m_carpet", 0.0), step=10.0, key=f"m_carpet_{curr_id}")
                glass_m2 = st.number_input("Kaca (m2)", value=get_val("m_glass", 0.0), step=10.0, key=f"m_glass_{curr_id}")
                st.subheader("Eksterior")
                facade = st.number_input("Facade (m2)", value=get_val("m_facade", 0.0), step=100.0, key=f"m_facade_{curr_id}")
                gondola_unit = st.number_input("Gondola (unit)", value=get_val("m_gondola", 0.0), step=1.0, key=f"m_gondola_{curr_id}")
                skylight_area = st.number_input("Skylight (m2)", value=get_val("m_skylight", 0.0), step=10.0, key=f"m_skylight_{curr_id}")
                railing_qty = st.number_input(
                "Railing Length (m'/room)",
                value=_safe_float(get_val("r_rail_qty", suggested_railing_qty)),
                step=1.0,
                key=f"r_rail_qty_{curr_id}",
                help="Can be synced from Area Analysis > Panjang Railing(m) using the main Use Area Analysis button."
                )
                st.caption(f"Total Railing: {railing_qty:.2f} m/room * {rooms:.0f} ruang = {railing_qty * rooms:,.2f} m'")    
                st.subheader("Pintu")
                glass_door = st.number_input(
                    "Glass Door (unit)",
                    value=_safe_float(get_val("m_door_g", suggested_door_glass)),
                    step=1.0,
                    key=f"m_door_g_{curr_id}"
                )

                wooden_door = st.number_input(
                    "Wooden Door (unit)",
                    value=_safe_float(get_val("m_door_w", suggested_door_wood)),
                    step=1.0,
                    key=f"m_door_w_{curr_id}"
                )

                steel_door = st.number_input(
                    "Steel Door (unit)",
                    value=_safe_float(get_val("m_door_s", suggested_door_steel)),
                    step=1.0,
                    key=f"m_door_s_{curr_id}"
                )

        with col_m3:
            with st.expander("Sanitari", expanded=True):
                st.subheader("Toilet Unit")
                san_qty_room = st.number_input("Toilet Private (unit/ruang)", help="Cth. 3 Toilet/1 Kamar (Apt)", value=get_val("r_san_qty", pt_data["san_room_qty"]), key=f"r_san_qty_{curr_type_key}")
                st.subheader("Toilet Umum")
                toilet_male = st.number_input("Toilet Umum - Pria (units)", value=get_val("m_toil_m", 0.0), step=1.0, key=f"m_toil_m_{curr_id}")
                toilet_female = st.number_input("Toilet Umum - Wanita (units)", value=get_val("m_toil_f", 0.0), step=1.0, key=f"m_toil_f_{curr_id}")
                disabled_toil = st.number_input("Toilet Difabel (units)", value=get_val("m_toil_d", 0.0), step=1.0, key=f"m_toil_d_{curr_id}")
                st.subheader("Mushola")
                mushola_unit = st.number_input("Mushola (units)", value=get_val("m_mushola", 0.0), step=1.0, key=f"m_mushola_{curr_id}")

        with col_m4:
            with st.expander("Fasilitas", expanded=True):
                st.subheader("Fasilitas")
                res_fac_m2 = st.number_input(
                    "Fasilitas Penghuni (m2)",
                    value=_safe_float(get_val("m_fac_res", suggested_res_fac_area)),
                    step=10.0,
                    key=f"m_fac_res_{curr_id}",
                    help="If synced from Area Analysis, this equals Residential Facility dataframe total divided by the current Resident Facility rate."
                )
                pub_fac_m2 = st.number_input("Fasilitas Publik (m2)", value=get_val("m_fac_pub", 0.0), step=10.0, key=f"m_fac_pub_{curr_id}")
                proj_fac_u = st.number_input("Fasilitas Proyek (unit)", value=get_val("m_fac_proj", 0.0), step=1.0, key=f"m_fac_proj_{curr_id}")
                land_m2 = st.number_input(
                    "Area Lanskap (m2)",
                    value=_safe_float(get_val("m_land_m2", suggested_landscape_area)),
                    step=100.0,
                    key=f"m_land_m2_{curr_id}",
                    help="If synced from Area Analysis, this equals External Works dataframe total divided by the current External Works rate."
                )
                st.subheader("Miscellaneous")
                m_status = st.radio("Ada Gym/Linen?", ["Tidak", "Ada"],
                                    index=1 if get_val("misc_switch", 0) == 1 else 0,
                                    key=f"misc_sw_{curr_id}", horizontal=True)
                misc_switch = 1 if m_status == "Ada" else 0
                st.session_state.projects[curr_id]["data"]["misc_switch"] = misc_switch

    # --- TAB 2: RATIOS & MULTIPLIERS ---
    with tab3:
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.subheader("Facade Ratio (%)")
            facade_precast_pct = st.number_input("Precast (%)", value=get_val("r_fac_pre", pt_data["facade_precast_pct"]), step=5.0, key=f"r_fac_pre_{curr_type_key}")
            facade_window_pct = st.number_input("Window Wall (%)", value=get_val("r_fac_win", pt_data["facade_window_pct"]), step=5.0, key=f"r_fac_win_{curr_type_key}")
            facade_double_pct = st.number_input("Double Skin (%)", value=get_val("r_fac_doub", pt_data["facade_double_pct"]), step=5.0, key=f"r_fac_doub_{curr_type_key}")
            t_fac_pct = facade_precast_pct + facade_window_pct + facade_double_pct

            if t_fac_pct != 100:
                st.warning(f"Warning: Total is **{t_fac_pct}%** (bukan 100%)")
        with col_r3:
            st.subheader("Waste & Skirting (%)")
            fl_skirt = st.number_input(
                "Skirting (%)", 
                value=get_val("s_floor", pt_data.get("fl_skirt", 20)), 
                key=f"s_floor{curr_type_key}"
            )
            fl_waste = st.number_input(
                "Floor Waste (%)", 
                value=get_val("w_floor", pt_data.get("fl_waste", 10)), # Use .get() with 1.1 as default
                key=f"w_floor{curr_type_key}"
            )        

            st.caption(f"Luas Lantai + Skirting + Waste: GFA: {gfa:.2f} x {1 + (fl_skirt/100):.2f} x {1 + (fl_waste/100):.2f} = {gfa*(1 + (fl_waste/100))*(1 + (fl_skirt/100)):.2f} m2")

            f_mult = (1 + (fl_waste/100)) * (1 + (fl_skirt/100))
        
        with col_r2:
            st.subheader("Flooring Ratio (%)")
            fl_ht_pct = st.number_input("HT/Ceramic Tile (%)", value=get_val("r_fl_ht", pt_data["fl_ht_pct"]), step=5.0, key=f"r_fl_ht_{curr_type_key}")
            fl_vinyl_pct = st.number_input("Vinyl (%)", value=get_val("r_fl_vin", pt_data["fl_vinyl_pct"]), step=5.0, key=f"r_fl_vin_{curr_type_key}")
            fl_marmer_pct = st.number_input("Marmer (%)", value=get_val("r_fl_mar", pt_data["fl_marmer_pct"]), step=5.0, key=f"r_fl_mar_{curr_type_key}")
            t_fl_pct = fl_ht_pct + fl_vinyl_pct + fl_marmer_pct

            if t_fl_pct != 100:
                st.warning(f"Warning: Total is **{t_fl_pct}%** (bukan 100%)")

    # --- TAB 3: UNIT RATES ---
    with tab4:
        with st.expander("Detail-Derived Rate Review", expanded=True):
            detail_rate_review_rows = build_detail_rate_review_rows(curr_proj, gba, gfa, rooms)
            detail_rate_review_df = pd.DataFrame(detail_rate_review_rows)

            st.caption("Review only. Detail-derived rates are not applied to Cost Analysis automatically.")
            st.dataframe(
                detail_rate_review_df,
                hide_index=True,
                width="stretch",
                column_config={
                    "Current Cost Rate": st.column_config.NumberColumn(
                        "Current Cost Rate",
                        format="Rp %.0f",
                    ),
                    "Detail-Derived Rate": st.column_config.NumberColumn(
                        "Detail-Derived Rate",
                        format="Rp %.0f",
                    ),
                    "Difference / Unit": st.column_config.NumberColumn(
                        "Difference / Unit",
                        format="Rp %.0f",
                    ),
                    "Detail Total": st.column_config.NumberColumn(
                        "Detail Total",
                        format="Rp %.0f",
                    ),
                    "Current Total": st.column_config.NumberColumn(
                        "Current Total",
                        format="Rp %.0f",
                    ),
                    "Difference Total": st.column_config.NumberColumn(
                        "Difference Total",
                        format="Rp %.0f",
                    ),
                },
            )

            earthwork_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "Earthworks"),
                {},
            )
            earthwork_apply_available = (
                earthwork_review_row.get("Status") == "Different"
                and _safe_float(earthwork_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(earthwork_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if earthwork_apply_available:
                st.warning(
                    "Earthworks detail-derived rate differs from the current Cost Analysis Earthwork Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply Earthworks Detail Rate",
                    key=f"apply_earthwork_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("earthwork_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("earthwork_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("Earthworks detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_earth"] = derived_rate
                        st.session_state[f"u_earth_{curr_type_key}"] = derived_rate
                        st.session_state.pop("earthwork_import_warning", None)

                        save_ok = save_after_user_action("Apply Earthworks Detail Rate")

                        if save_ok:
                            st.success("Earthworks detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "Earthworks detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            foundation_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "Foundation"),
                {},
            )
            foundation_apply_available = (
                foundation_review_row.get("Status") == "Different"
                and _safe_float(foundation_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(foundation_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if foundation_review_row:
                st.info(
                    "Foundation Detail-Derived Rate Review\n\n"
                    f"Current Foundation Rate: Rp {_safe_float(foundation_review_row.get('Current Cost Rate', 0.0)):,.0f}/m2\n\n"
                    f"Derived Foundation Rate: Rp {_safe_float(foundation_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/m2\n\n"
                    f"Difference: Rp {_safe_float(foundation_review_row.get('Difference / Unit', 0.0)):,.0f}/m2\n\n"
                    f"Status: {foundation_review_row.get('Status', '')}"
                )

            if foundation_apply_available:
                st.warning(
                    "Foundation detail-derived rate differs from the current Cost Analysis Foundation Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply Foundation Detail Rate",
                    key=f"apply_foundation_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("foundation_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("foundation_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("Foundation detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_found"] = derived_rate
                        st.session_state[f"u_found_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply Foundation Detail Rate")

                        if save_ok:
                            st.success("Foundation detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "Foundation detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            architectural_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "Architectural"),
                {},
            )
            architectural_apply_available = (
                architectural_review_row.get("Status") == "Different"
                and _safe_float(architectural_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(architectural_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if architectural_review_row:
                st.info(
                    "Architectural Detail-Derived Rate Review\n\n"
                    f"Current Architectural Rate: Rp {_safe_float(architectural_review_row.get('Current Cost Rate', 0.0)):,.0f}/m2\n\n"
                    f"Derived Architectural Rate: Rp {_safe_float(architectural_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/m2\n\n"
                    f"Difference: Rp {_safe_float(architectural_review_row.get('Difference / Unit', 0.0)):,.0f}/m2\n\n"
                    f"Status: {architectural_review_row.get('Status', '')}"
                )

            if architectural_apply_available:
                st.warning(
                    "Architectural detail-derived rate differs from the current Cost Analysis Architectural Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply Architectural Detail Rate",
                    key=f"apply_architectural_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("architectural_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("architectural_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("Architectural detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_arch"] = derived_rate
                        st.session_state[f"u_arch_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply Architectural Detail Rate")

                        if save_ok:
                            st.success("Architectural detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "Architectural detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            consultancy_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "Consultancy"),
                {},
            )
            consultancy_apply_available = (
                consultancy_review_row.get("Status") == "Different"
                and _safe_float(consultancy_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(consultancy_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if consultancy_review_row:
                st.info(
                    "Consultancy Detail-Derived Rate Review\n\n"
                    f"Current Consultancy Rate: Rp {_safe_float(consultancy_review_row.get('Current Cost Rate', 0.0)):,.0f}/m2\n\n"
                    f"Derived Consultancy Rate: Rp {_safe_float(consultancy_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/m2\n\n"
                    f"Difference: Rp {_safe_float(consultancy_review_row.get('Difference / Unit', 0.0)):,.0f}/m2\n\n"
                    f"Status: {consultancy_review_row.get('Status', '')}"
                )

            if consultancy_apply_available:
                st.warning(
                    "Consultancy detail-derived rate differs from the current Cost Analysis Consultancy Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply Consultancy Detail Rate",
                    key=f"apply_consultancy_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("consultancy_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("consultancy_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("Consultancy detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["sc_cons"] = derived_rate
                        st.session_state[f"sc_cons_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply Consultancy Detail Rate")

                        if save_ok:
                            st.success("Consultancy detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "Consultancy detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            ffe_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "FF&E"),
                {},
            )
            ffe_apply_available = (
                ffe_review_row.get("Status") == "Different"
                and _safe_float(ffe_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(ffe_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if ffe_review_row:
                st.info(
                    "FF&E Detail-Derived Rate Review\n\n"
                    f"Current FF&E Rate: Rp {_safe_float(ffe_review_row.get('Current Cost Rate', 0.0)):,.0f}/room\n\n"
                    f"Derived FF&E Rate: Rp {_safe_float(ffe_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/room\n\n"
                    f"Difference: Rp {_safe_float(ffe_review_row.get('Difference / Unit', 0.0)):,.0f}/room\n\n"
                    f"Status: {ffe_review_row.get('Status', '')}"
                )

            if ffe_apply_available:
                st.warning(
                    "FF&E detail-derived rate differs from the current Cost Analysis FF&E Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply FF&E Detail Rate",
                    key=f"apply_ffe_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("ffe_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("ffe_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("FF&E detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_ffe"] = derived_rate
                        st.session_state[f"u_ffe_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply FF&E Detail Rate")

                        if save_ok:
                            st.success("FF&E detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "FF&E detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            mep_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "MEP"),
                {},
            )
            mep_apply_available = (
                mep_review_row.get("Status") == "Different"
                and _safe_float(mep_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(mep_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if mep_review_row:
                st.info(
                    "MEP Detail-Derived Rate Review\n\n"
                    f"Current MEP Rate: Rp {_safe_float(mep_review_row.get('Current Cost Rate', 0.0)):,.0f}/m2\n\n"
                    f"Derived MEP Rate: Rp {_safe_float(mep_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/m2\n\n"
                    f"Difference: Rp {_safe_float(mep_review_row.get('Difference / Unit', 0.0)):,.0f}/m2\n\n"
                    f"Status: {mep_review_row.get('Status', '')}"
                )

            if mep_apply_available:
                st.warning(
                    "MEP detail-derived rate differs from the current Cost Analysis MEP Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply MEP Detail Rate",
                    key=f"apply_mep_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("mep_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("mep_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("MEP detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_mep"] = derived_rate
                        st.session_state[f"u_mep_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply MEP Detail Rate")

                        if save_ok:
                            st.success("MEP detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "MEP detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            utility_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "Utility"),
                {},
            )
            utility_apply_available = (
                utility_review_row.get("Status") == "Different"
                and _safe_float(utility_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(utility_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if utility_review_row:
                st.info(
                    "Utility Detail-Derived Rate Review\n\n"
                    f"Current Utility Rate: Rp {_safe_float(utility_review_row.get('Current Cost Rate', 0.0)):,.0f}/m2\n\n"
                    f"Derived Utility Rate: Rp {_safe_float(utility_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/m2\n\n"
                    f"Difference: Rp {_safe_float(utility_review_row.get('Difference / Unit', 0.0)):,.0f}/m2\n\n"
                    f"Status: {utility_review_row.get('Status', '')}"
                )

            if utility_apply_available:
                st.warning(
                    "Utility detail-derived rate differs from the current Cost Analysis Utility Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply Utility Detail Rate",
                    key=f"apply_utility_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("utility_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("utility_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("Utility detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_util"] = derived_rate
                        st.session_state[f"u_util_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply Utility Detail Rate")

                        if save_ok:
                            st.success("Utility detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "Utility detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

            structural_review_row = next(
                (row for row in detail_rate_review_rows if row.get("Section") == "Structural"),
                {},
            )
            structural_apply_available = (
                structural_review_row.get("Status") == "Different"
                and _safe_float(structural_review_row.get("Detail Total", 0.0)) > 0
                and _safe_float(structural_review_row.get("Detail-Derived Rate", 0.0)) > 0
            )

            if structural_review_row:
                st.info(
                    "Structural Detail-Derived Rate Review\n\n"
                    f"Current Structural Rate: Rp {_safe_float(structural_review_row.get('Current Cost Rate', 0.0)):,.0f}/m2\n\n"
                    f"Derived Structural Rate: Rp {_safe_float(structural_review_row.get('Detail-Derived Rate', 0.0)):,.0f}/m2\n\n"
                    f"Difference: Rp {_safe_float(structural_review_row.get('Difference / Unit', 0.0)):,.0f}/m2\n\n"
                    f"Status: {structural_review_row.get('Status', '')}"
                )

            if structural_apply_available:
                st.warning(
                    "Structural detail-derived rate differs from the current Cost Analysis Structural Rate. "
                    "Cost Analysis has not been updated automatically."
                )
                if st.button(
                    "Apply Structural Detail Rate",
                    key=f"apply_structural_detail_rate_{curr_id}",
                    type="primary",
                ):
                    derived_rate = _safe_float(curr_proj["data"].get("structural_derived_unit_price", 0.0))
                    detail_total = _safe_float(curr_proj["data"].get("structural_detail_total", 0.0))

                    if derived_rate <= 0 or detail_total <= 0:
                        st.error("Structural detail rate cannot be applied because the detail total or derived rate is zero.")
                    else:
                        curr_proj["data"]["u_struc"] = derived_rate
                        st.session_state[f"u_struc_{curr_type_key}"] = derived_rate

                        save_ok = save_after_user_action("Apply Structural Detail Rate")

                        if save_ok:
                            st.success("Structural detail rate applied to Cost Analysis.")
                            st.rerun()
                        else:
                            st.error(
                                "Structural detail rate was applied locally, but cloud save failed. Do not log out yet."
                            )

        with st.expander("Harga Fondasi & Struktur", expanded=True):
            c1, c2, c3 = st.columns(3)
            struc_earth = c1.number_input("Earthwork Rate (Rp)", value=get_val("u_earth", pt_data["struc_earth"]), key=f"u_earth_{curr_type_key}")
            struc_found = c2.number_input("Foundation Rate (Rp)", value=get_val("u_found", pt_data["struc_found"]), key=f"u_found_{curr_type_key}")
            struc_work = c3.number_input("Structural Work Rate (Rp)", value=get_val("u_struc", pt_data["struc_work"]), key=f"u_struc_{curr_type_key}")
            #caption
            c1.caption(f"""Hitungan: Rp {struc_earth:,.0f} x GBA: {gba:,.0f} m2  \n  Total Earthwork: Rp {struc_earth * gba:,.0f}  \n  Terbilang: {n2w(struc_earth * gba)}""")
            c2.caption(f"""Hitungan: Rp {struc_found:,.0f} x GBA: {gba:,.0f} m2  \n  Total Foundation: Rp {struc_found * gba:,.0f}  \n  Terbilang: {n2w(struc_found * gba)}""")
            c3.caption(f"""Hitungan: Rp {struc_work:,.0f} x GBA: {gba:,.0f} m2  \n  Total Structural Work: Rp {struc_work * gba:,.0f}  \n  Terbilang: {n2w(struc_work * gba)}""")

        with st.expander("Arsitektur & Fasad"):
            c1, c2 = st.columns(2)
            arch_base = c1.number_input("Architecture Base (Rp)", value=get_val("u_arch", pt_data["arch_base"]), key=f"u_arch_{curr_type_key}")
            lobby_rate = c2.number_input("Lobby Interior Rate (Rp)", value=get_val("u_lobby", pt_data["lobby"]), key=f"u_lobby_{curr_type_key}")
            c3, c4, c5 = st.columns(3)
            fac_precast_rate = c3.number_input("Precast Rate (Rp)", value=get_val("u_f_pre", pt_data["facade_precast_rate"]), key=f"u_f_pre_{curr_type_key}")
            fac_window_rate = c4.number_input("Window Wall Rate (Rp)", value=get_val("u_f_win", pt_data["facade_window_rate"]), key=f"u_f_win_{curr_type_key}")
            fac_double_rate = c5.number_input("Double Skin Rate (Rp)", value=get_val("u_f_doub", pt_data["facade_double_rate"]), key=f"u_f_doub_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {arch_base:,.0f} x GFA: {gfa:,.0f} m2  \n  Total Architecture Base: Rp {arch_base * gfa:,.0f}  \n  Terbilang: {n2w(arch_base * gfa)}""")
            c2.caption(f"""Hitungan: Rp {lobby_rate:,.0f} x Lobby Interior: {lobby_interior:,.0f} m2  \n  Total Lobby Interior: Rp {lobby_rate * lobby_interior:,.0f}  \n  Terbilang: {n2w(lobby_rate * lobby_interior)}""")
            c3.caption(f"""Hitungan: Rp {fac_precast_rate:,.0f} x Facade: {facade:,.0f} m2 x {facade_precast_pct}%  \n  Total Precast: Rp {fac_precast_rate * facade * (facade_precast_pct/100):,.0f}  \n  Terbilang: {n2w(fac_precast_rate * facade * (facade_precast_pct/100))}""")
            c4.caption(f"""Hitungan: Rp {fac_window_rate:,.0f} x Facade: {facade:,.0f} m2 x {facade_window_pct}%  \n  Total Window Wall: Rp {fac_window_rate * facade * (facade_window_pct/100):,.0f}  \n  Terbilang: {n2w(fac_window_rate * facade * (facade_window_pct/100))}""")
            c5.caption(f"""Hitungan: Rp {fac_double_rate:,.0f} x Facade: {facade:,.0f} m2 x {facade_double_pct}%  \n  Total Double Skin: Rp {fac_double_rate * facade * (facade_double_pct/100):,.0f}  \n  Terbilang: {n2w(fac_double_rate * facade * (facade_double_pct/100))}""")
        
        with st.expander("Pintu dan Hardware"):
            c1, c2, c3 = st.columns(3)
            door_wood = c1.number_input("Wooden Door Rate (Rp)", value=get_val("u_d_wood", pt_data["door_wood"]), key=f"u_d_wood_{curr_type_key}")
            door_glass = c3.number_input("Glass Door Rate (Rp)", value=get_val("u_d_glass", pt_data["door_glass"]), key=f"u_d_glass_{curr_type_key}")
            door_steel = c2.number_input("Steel Door Rate (Rp)", value=get_val("u_d_steel", pt_data["door_steel"]), key=f"u_d_steel_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {door_wood:,.0f} x  {wooden_door:,.0f} unit  \n  Total Pintu Kayu: Rp {door_wood * wooden_door:,.0f}  \n  Terbilang: {n2w(door_wood * wooden_door)}""")
            c3.caption(f"""Hitungan: Rp {door_glass:,.0f} x  {glass_door:,.0f} unit  \n  Total Pintu Kaca: Rp {door_glass * glass_door:,.0f}  \n  Terbilang: {n2w(door_glass * glass_door)}""")
            c2.caption(f"""Hitungan: Rp {door_steel:,.0f} x  {steel_door:,.0f} unit  \n  Total Pintu Baja: Rp {door_steel * steel_door:,.0f}  \n  Terbilang: {n2w(door_steel * steel_door)}""")
            hw_wood = c1.number_input("Hardware Wooden Door (Rp)", value=get_val("u_hw_wood", pt_data["hw_wood"]), key=f"u_hw_wood_{curr_type_key}")
            hw_steel = c2.number_input("Hardware Steel Door (Rp)", value=get_val("u_hw_steel", pt_data["hw_steel"]), key=f"u_hw_steel_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {hw_wood:,.0f} x  {wooden_door:,.0f} unit  \n  Total Hardware Pintu Kayu: Rp {hw_wood * wooden_door:,.0f}  \n  Terbilang: {n2w(hw_wood * wooden_door)}""")
            c2.caption(f"""Hitungan: Rp {hw_steel:,.0f} x  {steel_door:,.0f} unit  \n  Total Hardware Pintu Baja: Rp {hw_steel * steel_door:,.0f}  \n  Terbilang: {n2w(hw_steel * steel_door)}""")

        with st.expander("Sanitari"):
            c1, c2 = st.columns(2)
            san_room_rate = c1.number_input("Typical Unit Sanitary Rate (Rp)", value=get_val("u_s_room", pt_data["san_room_rate"]), key=f"u_s_room_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {san_room_rate:,.0f} x {rooms:,.0f} Rooms x {san_qty_room} Units / Room  \n  Total Private Bathroom: Rp {rooms * san_qty_room * san_room_rate:,.0f}  \n  Terbilang: {n2w(rooms * san_qty_room * san_room_rate)}""")
            san_pub_m = c2.number_input("Public Toilet Male Rate (Rp)", value=get_val("u_s_pub_m", pt_data["san_pub_m"]), key=f"u_s_pub_m_{curr_type_key}")
            c2.caption(f"""Hitungan: Rp {san_pub_m:,.0f} x {toilet_male:,.0f} Units  \n  Total Public Toilet Male: Rp {toilet_male * san_pub_m:,.0f}  \n  Terbilang: {n2w(toilet_male * san_pub_m)}""")
            san_pub_f = c2.number_input("Public Toilet Female Rate (Rp)", value=get_val("u_s_pub_f", pt_data["san_pub_f"]), key=f"u_s_pub_f_{curr_type_key}")
            c2.caption(f"""Hitungan: Rp {san_pub_f:,.0f} x {toilet_female:,.0f} Units  \n  Total Public Toilet Female: Rp {toilet_female * san_pub_f:,.0f}  \n  Terbilang: {n2w(toilet_female * san_pub_f)}""")
            san_dis = c1.number_input("Disabled Toilet Rate (Rp)", value=get_val("u_s_dis", pt_data["san_dis"]), key=f"u_s_dis_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {san_dis:,.0f} x {disabled_toil:,.0f} Units  \n  Total Toilet Difabel: Rp {disabled_toil * san_dis:,.0f}  \n  Terbilang: {n2w(disabled_toil * san_dis)}""")
            san_mushola = c1.number_input("Mushola Rate (Rp)", value=get_val("u_s_mushola", pt_data["san_mushola"]), key=f"u_s_mushola_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {san_mushola:,.0f} x {mushola_unit:,.0f} Units  \n  Total Mushola: Rp {mushola_unit * san_mushola:,.0f}  \n  Terbilang: {n2w(mushola_unit * san_mushola)}""")
        
        with st.expander("Lantai, Finishing, dan Interior"):
            st.subheader("Harga ")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.radio("Spek HT", ["Type1", "Type2"],
                    key=f"temp_spec_ht_{curr_type_key}",
                    horizontal=True,
                    on_change=update_price, args=("ht", "fl_ht_rate"))
                fl_ht_rate = st.number_input("HT Rate (Rp)",
                                            value=get_val("u_fl_ht", pt_data["fl_ht_rate"]["Type1"]),
                                            key=f"u_fl_ht_{curr_type_key}")
                c1.caption(f"""Hitungan: {fl_ht_pct}% of GFA x Rp {fl_ht_rate:,.0f} x {gfa:,.0f} m2 x {f_mult}  \n  Total HT: Rp {gfa * (fl_ht_pct / 100) * fl_ht_rate * f_mult:,.0f}  \n  Terbilang: {n2w(gfa * (fl_ht_pct / 100) * fl_ht_rate * f_mult)}""")

            with c2:
                st.radio("Spek Vinyl", ["Type1", "Type2"],
                    key=f"temp_spec_vin_{curr_type_key}",
                    horizontal=True,
                    on_change=update_price, args=("vin", "fl_vinyl_rate"))
                fl_vinyl_rate = st.number_input("Vinyl Rate (Rp)",
                                                value=get_val("u_fl_vin", pt_data["fl_vinyl_rate"]["Type1"]),
                                                key=f"u_fl_vin_{curr_type_key}")
                c2.caption(f"""Hitungan: {fl_vinyl_pct}% of GFA x Rp {fl_vinyl_rate:,.0f} x {gfa:,.0f} m2 x {f_mult}  \n  Total Vinyl: Rp {gfa * (fl_vinyl_pct / 100) * fl_vinyl_rate * f_mult:,.0f}  \n  Terbilang: {n2w(gfa * (fl_vinyl_pct / 100) * fl_vinyl_rate * f_mult)}""")
            
            with c3:
                st.radio("Spek Marmer", ["Type1", "Type2"],
                    key=f"temp_spec_mar_{curr_type_key}",
                    horizontal=True,
                    on_change=update_price, args=("mar", "fl_marmer_rate"))
                fl_marmer_rate = st.number_input("Marmer Rate (Rp)",
                                                value=get_val("u_fl_mar", pt_data["fl_marmer_rate"]["Type1"]),
                                                key=f"u_fl_mar_{curr_type_key}")
                c3.caption(f"""Hitungan: {fl_marmer_pct}% of GFA x Rp {fl_marmer_rate:,.0f} x {gfa:,.0f} m2 x {f_mult}  \n  Total Marmer: Rp {gfa * (fl_marmer_pct / 100) * fl_marmer_rate * f_mult:,.0f}  \n  Terbilang: {n2w(gfa * (fl_marmer_pct / 100) * fl_marmer_rate * f_mult)}""")
                
            carpet_rate = c1.number_input("Carpet Rate (Rp)", value=get_val("u_carpet", pt_data["carpet"]), key=f"u_carpet_{curr_type_key}")
            c1.caption(f"""Hitungan: {carpet_m2:,.0f} m2 x Rp {carpet_rate:,.0f}  \n  Total Carpet Work: Rp {carpet_m2 * carpet_rate:,.0f}  \n  Terbilang: {n2w(carpet_m2 * carpet_rate)}""")
            glass_rate = c2.number_input("Glass Work Rate (Rp)", value=get_val("u_glass", pt_data["glass"]), key=f"u_glass_{curr_type_key}")
            c2.caption(f"""Hitungan: {glass_m2:,.0f} m2 x Rp {glass_rate:,.0f}  \n  Total Glass Work: Rp {glass_m2 * glass_rate:,.0f}  \n  Terbilang: {n2w(glass_m2 * glass_rate)}""")            
            skylight_rate = c3.number_input("Skylight Rate (Rp)", value=get_val("u_sky", pt_data["skylight_rate"]), key=f"u_sky_{curr_type_key}")
            c3.caption(f"""Hitungan: {skylight_area:,.0f} m2 Total x Rp {skylight_rate:,.0f}  \n  Total Skylight Work: Rp {skylight_area * skylight_rate:,.0f}  \n  Terbilang: {n2w(skylight_area * skylight_rate)}""")
            gondola_rate = c1.number_input("Gondola Rate (Rp)", value=get_val("u_gondola", pt_data["gondola"]), key=f"u_gondola_{curr_type_key}")
            c1.caption(f"""Hitungan: {gondola_unit:,.0f} Units x Rp {gondola_rate:,.0f}  \n  Total Gondola: Rp {gondola_unit * gondola_rate:,.0f}  \n  Terbilang: {n2w(gondola_unit * gondola_rate)}""")
            railing_rate = c2.number_input("Railing Rate (Rp)", value=get_val("u_rail", pt_data["railing_rate"]), key=f"u_rail_{curr_type_key}")
            c2.caption(f"""Hitungan: {rooms * railing_qty:,.0f} m' Total x Rp {railing_rate:,.0f}  \n  Total Railing Work: Rp {rooms * railing_qty * railing_rate:,.0f}  \n  Terbilang: {n2w(rooms * railing_qty * railing_rate)}""")

        with st.expander("MEP, Dapur dan FF&E"):
            c1, c2 = st.columns(2)
            mep_rate = c1.number_input("MEP Works (Rp)", value=get_val("u_mep", pt_data["mep"]), key=f"u_mep_{curr_type_key}")
            c1.caption(f"""Hitungan: {gba:,.0f} m2 x Rp {mep_rate:,.0f}  \n  Total MEP Works: Rp {gba * mep_rate:,.0f}  \n  Terbilang: {n2w(gba * mep_rate)}""")
            utility_rate = c2.number_input("Utility Connection (Rp)", value=get_val("u_util", pt_data["utility"]), key=f"u_util_{curr_type_key}")
            c2.caption(f"""Hitungan: {gba:,.0f} m2 x Rp {utility_rate:,.0f}  \n  Total Utility Connection: Rp {gba * utility_rate:,.0f}  \n  Terbilang: {n2w(gba * utility_rate)}""")

            ffe_rate = c1.number_input("FF&E (Rp)", value=get_val("u_ffe", pt_data["ffe"]), key=f"u_ffe_{curr_type_key}")
            c1.caption(f"""Hitungan: {rooms:,.0f} Rooms x Rp {ffe_rate:,.0f}  \n  Total FF&E: Rp {rooms * ffe_rate:,.0f}  \n  Terbilang: {n2w(rooms * ffe_rate)}""")

            kitchen_rate = c2.number_input("Kitchen Equipment (Rp)", value=get_val("u_kit", pt_data["kitchen"]), key=f"u_kit_{curr_type_key}")
            c2.caption(f"""Hitungan: {rooms:,.0f} Rooms x Rp {kitchen_rate:,.0f}  \n  Total Kitchen Equipment: Rp {rooms * kitchen_rate:,.0f}  \n  Terbilang: {n2w(rooms * kitchen_rate)}""")

            misc_rate = c1.number_input("Misc (Linen/Gym Equipment) (Rp)", value=get_val("u_misc", pt_data["misc"]), key=f"u_misc_{curr_type_key}")
            c1.caption(f"""Hitungan: Rp {misc_rate if misc_switch else 0:,.0f}  \n  Total Misc. Costs: Rp {misc_rate * misc_switch:,.0f}  \n  Terbilang: {n2w(misc_rate * misc_switch)}""")
        
        with st.expander("External & Facility Rates"):
            c1, c2 = st.columns(2)
            ext_land_rate = c1.number_input("External Works (Landscape) (Rp)", value=get_val("u_ext", pt_data["ext_land"]), key=f"u_ext_{curr_type_key}")
            fac_pub_rate = c2.number_input("Public Facilities (Rp)", value=get_val("u_fac_p", pt_data["fac_pub"]), key=f"u_fac_p_{curr_type_key}")
            c1.caption(f"""Hitungan: {land_m2:,.0f} m2 x Rp {ext_land_rate:,.0f}  \n  Total External Works: Rp {land_m2 * ext_land_rate:,.0f}  \n  Terbilang: {n2w(land_m2 * ext_land_rate)}""")
            c2.caption(f"""Hitungan: {pub_fac_m2:,.0f} m2 x Rp {fac_pub_rate:,.0f}  \n  Total Public Facilities: Rp {pub_fac_m2 * fac_pub_rate:,.0f}  \n  Terbilang: {n2w(pub_fac_m2 * fac_pub_rate)}""")
            fac_res_rate = c1.number_input("Resident Facilities (Rp)", value=get_val("u_fac_r", pt_data["fac_res"]), key=f"u_fac_r_{curr_type_key}")
            fac_proj_rate = c2.number_input("Project Facilities (Rp)", value=get_val("u_fac_pr", pt_data["fac_proj"]), key=f"u_fac_pr_{curr_type_key}")
            c1.caption(f"""Hitungan: {res_fac_m2:,.0f} m2 x Rp {fac_res_rate:,.0f}  \n  Total Resident Facilities: Rp {res_fac_m2 * fac_res_rate:,.0f}  \n  Terbilang: {n2w(res_fac_m2 * fac_res_rate)}""")
            c2.caption(f"""Hitungan: {proj_fac_u:,.0f} Units x Rp {fac_proj_rate:,.0f}  \n  Total Project Facilities: Rp {proj_fac_u * fac_proj_rate:,.0f}  \n  Terbilang: {n2w(proj_fac_u * fac_proj_rate)}""")

    # --- TAB 4: SOFT COSTS ---
    with tab5:
        sc_col1, sc_col2, sc_col3 = st.columns(3)
        with sc_col1:
            with st.expander("QS", expanded=True):
                qs_months = st.number_input("Durasi QS (Bulan)", value=get_val("sc_qs_m", 0.0), step=1.0, key=f"sc_qs_m_{curr_id}")
                qs_rate = st.number_input("Harga QS (per Bulan) (Rp)", value=get_val("sc_qs_r", 0.0), step=1000000.0, key=f"sc_qs_r_{curr_id}")
                st.caption(f"""Hitungan: {qs_months} Months x Rp {qs_rate:,.0f}/Mo  \n  Total QS Services: Rp {qs_months * qs_rate:,.0f}  \n  Terbilang: {n2w(qs_months * qs_rate)}""")
        with sc_col2:
            with st.expander("PM", expanded=True):
                pm_months = st.number_input("Durasi PM (Bulan)", value=get_val("sc_pm_m", 0.0), step=1.0, key=f"sc_pm_m_{curr_id}")
                pm_rate = st.number_input("Harga PM (per Bulan) (Rp)", value=get_val("sc_pm_r", 0.0), step=1000000.0, key=f"sc_pm_r_{curr_id}")
                st.caption(f"""Hitungan: {pm_months} Months x Rp {pm_rate:,.0f}/Mo  \n  Total PM Services: Rp {pm_months * pm_rate:,.0f}  \n  Terbilang: {n2w(pm_months * pm_rate)}""")                
        with sc_col3:
            with st.expander("Lainnya", expanded=True):
                consultancy_rate = st.number_input("Biaya Konsultan (Rp) per m2 GFA", help="Biaya konsultan per m2 GFA", value=get_val("sc_cons", pt_data["cons"]), key=f"sc_cons_{curr_type_key}")
                st.caption(f"""Hitungan: {gfa:,.0f} m2 x Rp {consultancy_rate:,.0f}  \n  Total Consultancy Fee: Rp {gfa * consultancy_rate:,.0f}  \n  Terbilang: {n2w(gfa * consultancy_rate)}""")
                insurance_pct = st.number_input("Insurance (%)", help="Persentase premi asuransi", value=get_val("sc_ins", 0.12), step=0.01, key=f"sc_ins_{curr_id}")
    
    # --- TAB 5: CUSTOM ITEMS ---
    with tab6:
        st.subheader("Item Tambahan")

        default_smart_cc = [{"Item Description": "", "Rate (Rp)": 0.0, "Quantity": 1.0}]
        current_smart_cc = get_val("smart_custom_costs", default_smart_cc)

        edited_smart_cc = st.data_editor(
            pd.DataFrame(current_smart_cc),
            num_rows="dynamic",
            key=f"edit_smart_cc_{curr_id}",
            column_order=["Item Description", "Quantity", "Rate (Rp)"],
            column_config={
                "Item Description": st.column_config.TextColumn(
                    "Item Description",
                    width="large"
                ),
                "Quantity": st.column_config.NumberColumn(
                    "Quantity",
                    min_value=0,
                    default=1.0,
                    format="%.2f"
                ),
                "Rate (Rp)": st.column_config.NumberColumn(
                    "Rate (Rp)",
                    min_value=0.0,
                    default=0.0,
                    format="%.2f"
                ),
            },
            width="stretch"
        )

        smart_custom_inputs = {}
        smart_custom_costs = 0.0
        breakdown_details = []

        for index, row in edited_smart_cc.iterrows():
            suffix = index + 1

            desc = row.get("Item Description", "")
            rate = _safe_float(row.get("Rate (Rp)", 0.0))
            qty = _safe_float(row.get("Quantity", 1.0))

            smart_custom_inputs[f"input_name{suffix}"] = desc
            smart_custom_inputs[f"input_rate{suffix}"] = rate
            smart_custom_inputs[f"input_qty{suffix}"] = qty

            item_total = rate * qty
            smart_custom_costs += item_total

            if rate > 0:
                breakdown_details.append(
                    f"**{desc}**: Rp {rate:,.2f} x {qty} qty = **Rp {item_total:,.2f}**"
                )

        # Temporary current-rerun export/calculation state only
        st.session_state[f"smart_custom_data_{curr_id}"] = smart_custom_inputs

        st.markdown("---")
        st.markdown(f"### Total Harga Item Custom: Rp {smart_custom_costs:,.2f}")

        if breakdown_details:
            with st.expander("Custom Item Breakdown"):
                for detail in breakdown_details:
                    st.markdown(detail)

        st.caption("Edit the table, then click Save Change to apply and store the changes.")

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_smart_custom_clicked = st.button(
                "Save Change",
                key=f"save_smart_custom_to_cloud_{curr_id}",
                type="primary",
                width="stretch"
            )

        if save_smart_custom_clicked:
            set_data("smart_custom_costs", edited_smart_cc.to_dict("records"))

            save_ok = save_after_user_action("Save Smart Custom Costs")

            if save_ok:
                st.success("Saved to cloud.")
                st.rerun()
            else:
                st.error("Cloud save failed. Do not log out yet.")

    # --- LIVE AUTO-CALCULATIONS ---
    t_earth = gba * struc_earth
    t_found = gba * struc_found
    t_struc = gba * struc_work
    t_arch_base = gfa * arch_base
    t_precast = facade * (facade_precast_pct / 100) * fac_precast_rate
    t_window  = facade * (facade_window_pct / 100) * fac_window_rate
    t_double  = facade * (facade_double_pct / 100) * fac_double_rate
    t_w_door = wooden_door * door_wood
    t_g_door = glass_door * door_glass
    t_s_door = steel_door * door_steel
    t_lobby  = lobby_interior * lobby_rate
    t_gondola = gondola_unit * gondola_rate
    t_unit_san = rooms * san_qty_room * san_room_rate
    t_t_male   = toilet_male * san_pub_m
    t_t_female = toilet_female * san_pub_f
    t_t_dis    = disabled_toil * san_dis
    t_mushola  = mushola_unit * san_mushola
    t_kitchen = rooms * kitchen_rate
    t_hw_w    = wooden_door * hw_wood
    t_hw_s    = steel_door * hw_steel
    f_mult = (1 + (fl_waste/100)) * (1 + (fl_skirt/100))
    t_ht      = gfa * (fl_ht_pct / 100) * fl_ht_rate * f_mult
    t_vinyl   = gfa * (fl_vinyl_pct / 100) * fl_vinyl_rate * f_mult
    t_marmer  = gfa * (fl_marmer_pct / 100) * fl_marmer_rate * f_mult
    t_carpet     = carpet_m2 * carpet_rate
    t_glass_work = glass_m2 * glass_rate
    t_ffe        = rooms * ffe_rate
    t_misc       = misc_rate * misc_switch
    t_mep        = gba * mep_rate
    t_utility    = gba * utility_rate
    t_railing    = (rooms * railing_qty) * railing_rate
    t_skylight   = skylight_area * skylight_rate
    t_external = land_m2 * ext_land_rate
    t_pub_fac  = pub_fac_m2 * fac_pub_rate
    t_res_fac  = res_fac_m2 * fac_res_rate
    t_proj_fac = proj_fac_u * fac_proj_rate

    construction_subtotal = sum([
        t_earth, t_found, t_struc, t_arch_base, t_precast, t_window, t_double,
        t_w_door, t_g_door, t_s_door, t_lobby, t_gondola, t_unit_san, t_t_male,
        t_t_female, t_t_dis, t_mushola, t_kitchen, t_hw_w, t_hw_s, t_ht, t_vinyl,
        t_marmer, t_carpet, t_glass_work, t_ffe, t_misc, t_mep, t_utility,
        t_railing, t_skylight, t_external, t_pub_fac, t_res_fac, t_proj_fac,
        smart_custom_costs
    ])

    t_preliminary = (construction_subtotal) * 0.05
    t_contingency = (construction_subtotal + t_preliminary) * 0.03
    grand_total_hc = construction_subtotal + t_preliminary + t_contingency

    t_consultancy = gfa * consultancy_rate
    t_qs = qs_months * qs_rate
    t_pm = pm_months * pm_rate
    t_insurance = (grand_total_hc) * (insurance_pct / 100.0)

    total_soft_cost = t_consultancy + t_qs + t_pm + t_insurance
    grand_total_project = grand_total_hc + total_soft_cost

    group_earth = t_earth 
    group_found = t_found 
    group_struc = t_struc
    
    group_facade = t_precast + t_window + t_double
    group_sanitary = t_unit_san + t_t_male + t_t_female + t_t_dis + t_mushola
    group_floor =  t_ht + t_vinyl + t_marmer 
    group_door = t_w_door + t_g_door + t_s_door + t_hw_w + t_hw_s
    group_arch = (t_arch_base + + t_lobby + t_carpet + t_gondola 
                  + t_glass_work + t_kitchen  + t_railing + t_skylight 
                  + group_facade + group_sanitary + group_floor + group_door
                  + smart_custom_costs)
    
    group_ffe = t_ffe + t_misc 
    group_mep = t_mep 
    group_utility = t_utility
    group_ext = t_external
    group_misc = t_pub_fac + t_res_fac + t_proj_fac
    group_prelim = t_preliminary
    group_conting = t_contingency
    group_soft_cost = total_soft_cost
    group_hard_cost = grand_total_hc
    group_total = total_soft_cost + grand_total_hc

    with tab7:
        tab1, tab2, tab3 = st.tabs([
        "Hasil",
        "Tabel", "Chart"
        ])
        
        with tab1:
            box_base = "margin-bottom: 12px; padding: 8px; border-radius: 5px; background-color: #FFFFFF; border: 1px solid #E0E0E0;"
            label_style = "font-size: 12px; color: #666666; font-weight: bold;"
            val_style = "font-size: 14px; font-weight: bold; color: #000000; margin-top: 4px;"
            with st.expander("Detail Hard Cost", expanded=False):
                cs1, cs2, cs3, cs4, cs5, cs6 = st.columns(6)

                # 1. Preliminary (Army Green - Start)
                cs1.markdown(f"""<div style="{box_base} border-left: 5px solid #CDDC39;">
                    <div style="{label_style}">Preliminary</div>
                    <div style="{val_style}">{n2w(group_prelim)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_prelim:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 2. Earthwork (Lime)
                cs2.markdown(f"""<div style="{box_base} border-left: 5px solid #8BC34A;">
                    <div style="{label_style}">Earthwork</div>
                    <div style="{val_style}">{n2w(group_earth)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_earth:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 3. Utility (Sage Green)
                cs3.markdown(f"""<div style="{box_base} border-left: 5px solid #4CAF50;">
                    <div style="{label_style}">Utility</div>
                    <div style="{val_style}">{n2w(group_utility)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_utility:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 4. Foundation (Soft Green)
                cs4.markdown(f"""<div style="{box_base} border-left: 5px solid #689F38;">
                    <div style="{label_style}">Foundation</div>
                    <div style="{val_style}">{n2w(group_found)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_found:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 5. Structural (Success Green)
                cs5.markdown(f"""<div style="{box_base} border-left: 5px solid #388E3C;">
                    <div style="{label_style}">Structural</div>
                    <div style="{val_style}">{n2w(group_struc)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_struc:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 6. External Works (Olive)
                cs6.markdown(f"""<div style="{box_base} border-left: 5px solid #254E18;">
                    <div style="{label_style}">External Works</div>
                    <div style="{val_style}">{n2w(group_ext)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_ext:,.0f}</div>
                </div>""", unsafe_allow_html=True)


                # --- ROW 2: ct1 to ct6 ---
                ct1, ct2, ct3, ct4, ct5, ct6 = st.columns(6)

                # 7. Architecture (Medium Green)
                ct2.markdown(f"""<div style="{box_base} border-left: 5px solid #9CCC65;">
                    <div style="{label_style}">Architecture</div>
                    <div style="{val_style}">{n2w(group_arch)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_arch:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 8. Miscellaneous (Dark Olive)
                ct3.markdown(f"""<div style="{box_base} border-left: 5px solid #558B2F;">
                    <div style="{label_style}">Miscellaneous</div>
                    <div style="{val_style}">{n2w(group_misc)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_misc:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 9. FF & E (Dark Green)
                ct4.markdown(f"""<div style="{box_base} border-left: 5px solid #2E7D32;">
                    <div style="{label_style}">FF & E</div>
                    <div style="{val_style}">{n2w(group_ffe)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_ffe:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 10. Contingency (Deep Army)
                ct6.markdown(f"""<div style="{box_base} border-left: 5px solid #1B5E20;">
                    <div style="{label_style}">Contingency</div>
                    <div style="{val_style}">{n2w(group_conting)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_conting:,.0f}</div>
                </div>""", unsafe_allow_html=True)

                # 11. MEP Works (Forest Green)
                ct5.markdown(f"""<div style="{box_base} border-left: 5px solid #33691E;">
                    <div style="{label_style}">MEP Works</div>
                    <div style="{val_style}">{n2w(group_mep)}</div>
                    <div style="font-size: 10px; color: #888888; margin-top: 2px;">Rp {group_mep:,.0f}</div>
                </div>""", unsafe_allow_html=True)
            
        # --- TOTALS SUMMARY ---
            c1, c2 = st.columns(2)

            # Common styles
            summary_base = "margin-bottom: 20px; padding: 15px; border-radius: 8px; background-color: #FFFFFF; border: 1px solid #E0E0E0;"
            summary_label = "font-size: 16px; color: #666666; font-weight: bold;"
            summary_val = "font-size: 24px; font-weight: bold; color: #000000; line-height: 1.2; margin-top: 5px;"
            summary_n2w = "font-size: 14px; color: #888888; font-weight: normal; margin-top: 5px;"

            # Hard Cost - Starting with Lime Accent
            c1.markdown(f"""
                <div style="{summary_base} border-left: 8px solid #CDDC39; background-color: #F9F9F0">
                    <div style="{summary_label}">Total Hard Cost</div>
                    <div style="{summary_val}">Rp {grand_total_hc:,.2f}</div>
                    <div style="{summary_n2w}">Terbilang: {n2w(grand_total_hc)} Rupiah</div>
                </div>
            """, unsafe_allow_html=True)


            # Soft Cost - Moving to Success Green Accent
            c2.markdown(f"""
                <div style="{summary_base} border-left: 8px solid #4CAF50; background-color: #F1F8F1">
                    <div style="{summary_label}">Total Soft Cost</div>
                    <div style="{summary_val}">Rp {total_soft_cost:,.2f}</div>
                    <div style="{summary_n2w}">Terbilang: {n2w(total_soft_cost)} Rupiah</div>
                </div>
            """, unsafe_allow_html=True)

            # Grand Total - Final Army Green Accent (Thicker border and larger font)
            st.markdown(f"""
                <div style="{summary_base} border-left: 12px solid #254E18; padding: 20px; background-color: #F5F7F5">
                    <div style="font-size: 18px; color: #666666; font-weight: bold;">Project Cost</div>
                    <div style="font-size: 12px; color: #888888; margin-top: 8px;">
                        Rp {grand_total_hc:,.2f} + Rp {total_soft_cost:,.2f} =
                    </div>
                    <div style="font-size: clamp(24px, 4vw, 30px); font-weight: bold; color: #000000; line-height: 1.2; margin-top: 10px;">
                        Rp {grand_total_project:,.2f}
                    </div>
                    <div style="font-size: 16px; color: #888888; margin-top: 8px;">
                        Terbilang: {n2w(grand_total_project)} Rupiah
                    </div>
                </div>
            """, unsafe_allow_html=True)

        with tab2:
            if 'show_details' not in st.session_state:
                st.session_state.show_details = True

            # 2. Add the Toggle Button
            label = "Hide Details" if st.session_state.show_details else "Show Details"
            if st.button(label):
                st.session_state.show_details = not st.session_state.show_details
                st.rerun()

            # 3. Define the full dataset as a Dictionary
            data_dict = {
                # --- MAIN CONSTRUCTION ---
                "1. Preliminary Works": t_preliminary,
                "2. Earthwork": t_earth,
                "3. Foundation": t_found,
                "4. Structural Work": t_struc,
                
                # --- ARCHITECTURE ---
                "5. Total Architecture": group_arch,
                "5.1 Basic Architecture": t_arch_base,
                "5.2 Facade - Precast": t_precast,
                "5.3 Facade - Window Wall": t_window,
                "5.4 Facade - Double Skin": t_double,
                "5.5 Wooden Doors": t_w_door,
                "5.6 Glass Doors": t_g_door,
                "5.7 Steel Doors": t_s_door,
                "5.8 Lobby Interior": t_lobby,
                "5.9 Gondola": t_gondola,
                "5.10 Typical Unit Sanitary": t_unit_san,
                "5.11 Public Toilet Male": t_t_male,
                "5.12 Public Toilet Female": t_t_female,
                "5.13 Disabled Toilet": t_t_dis,
                "5.14 Mushola": t_mushola,
                "5.15 Kitchen Equipment": t_kitchen,
                "5.16 Hardware Pintu Kayu": t_hw_w,
                "5.17 Hardware Pintu Besi": t_hw_s,
                "5.18 HT/Ceramic Tile": t_ht,
                "5.19 Vinyl Flooring": t_vinyl,
                "5.20 Marmer Flooring": t_marmer,
                "5.21 Carpet Work": t_carpet,
                "5.22 Railing Work": t_railing,
                "5.23 Skylight Work": t_skylight,
                "5.24 Glass Work": t_glass_work,
                "5.25 Custom Item (Architecture)": smart_custom_costs,

                # --- FF&E & SERVICES ---
                "6. Total FF&E": group_ffe,
                "6.1 FF&E": t_ffe,
                "6.2 Misc. (Linen/Gym)": t_misc,
                "7. MEP Works": t_mep,
                "8. Utility Connection": t_utility,

                # --- EXTERNAL & FACILITIES ---
                "9. External Works": t_external,
                "10. Miscellanous/Facility": group_misc,
                "10.1 Public Facilities": t_pub_fac,
                "10.2 Resident Facilities": t_res_fac,
                "10.3 Project Facilities": t_proj_fac,

                # --- CONTINGENCY ---
                "11. Contingencies": t_contingency,

                # --- SOFT COSTS / CONSULTANTS ---
                "12. Consultancy Fee": t_consultancy,
                "13. Quantity Surveyor": t_qs,
                "14. Project Management": t_pm,
                "15. Insurance Coverage": t_insurance
            }

            # 4. Extract into lists for your DataFrame automatically
            original_descriptions = list(data_dict.keys())
            raw_amounts = list(data_dict.values())

            # 4. Filter and Indent Logic
            filtered_data = []
            major_numbers = [f"{i}. " for i in range(1, 16)]

            for desc, amt in zip(original_descriptions, raw_amounts):
                is_major = any(desc.startswith(num) for num in major_numbers)
                
                if st.session_state.show_details:
                    # If showing details, indent sub-items
                    display_desc = desc if is_major else f"        {desc}"
                    filtered_data.append({"Description": display_desc, "Amount": amt})
                else:
                    # If hiding details, only append major items
                    if is_major:
                        filtered_data.append({"Description": desc, "Amount": amt})

            # 5. Create DataFrame
            df = pd.DataFrame(filtered_data)
            df["Amount"] = df["Amount"].apply(lambda x: f"Rp {x:,.2f}")

            # 6. Styling Logic (Bold Major Items)
            def style_major_rows(row):
                clean_desc = row["Description"].strip()
                is_major = any(clean_desc.startswith(num) for num in major_numbers)
                return ['font-weight: bold' if is_major else '' for _ in row]

            # 7. Display
            styled_df = df.style.apply(style_major_rows, axis=1)
            st.dataframe(styled_df, width="stretch", hide_index=True)
        
        with tab3:
                st.subheader("Total Project Cost Breakdown")

                # 1. Define the dictionary first
                detailed_items = {
                    "Item": [
                        "Preliminary", "Earthwork", "Foundation", "Structural", 
                        "Architecture Work", "FF&E", "MEP Works", "Utility",
                        "External/Landscape", "Misc/Facility", "Contingency",
                        "Soft Cost/Consultancy & Insurance"
                    ],
                    "Amount": [
                        group_prelim, group_earth, group_found, group_struc, 
                        group_arch, group_ffe, group_mep, group_utility, 
                        group_ext, group_misc,
                        group_conting, group_soft_cost
                    ],
                    # Adding 'Type' helps with color coding the chart
                    "Type": ["Hard Cost"]*11 + ["Soft Cost"]*1 
                }

                # 2. Convert to DataFrame and FILTER OUT zeros
                df_detailed = pd.DataFrame(detailed_items)
                df_detailed = df_detailed[df_detailed["Amount"] > 0] 

                # 3. Dynamic height calculation
                chart_height = max(400, len(df_detailed) * 25)

                # 4. Create the Chart
                hover = alt.selection_point(on='mouseover', nearest=True, fields=['Item'], empty=False)

                detailed_chart = alt.Chart(df_detailed).mark_bar().encode(
                    x=alt.X("Amount:Q", title="Cost (Rp)"),
                    y=alt.Y("Item:N", sort="-x", title=""),
                    opacity=alt.condition(hover, alt.value(1), alt.value(0.7)),
                    color=alt.Color("Type:N", scale=alt.Scale(domain=['Hard Cost', 'Soft Cost'], range=["#1f77b4", "#c2a136"])),
                    tooltip=["Item", "Type", alt.Tooltip("Amount:Q", format=",.2f")]
                ).properties(height=chart_height).add_params(hover)

                st.altair_chart(detailed_chart, width="stretch")

# --- SAVE ALL METRICS TO SESSION STATE ---
    # We use .update() so we NEVER delete the Area Analysis's data!
    st.session_state.projects[curr_id]["data"].update({
        "ht_spec_type": get_val("ht_spec_type", "Type1"),
        "vin_spec_type": get_val("vin_spec_type", "Type1"),
        "mar_spec_type": get_val("mar_spec_type", "Type1"),
        "m_land": land_area, "m_gba": gba, "m_gfa": gfa, "m_sgfa": sgfa,
        "m_facade": facade, "m_rooms": rooms, "m_lobby": lobby_interior,
        "m_gondola": gondola_unit, "m_carpet": carpet_m2, "m_glass": glass_m2,
        "m_skylight": skylight_area, "m_door_g": glass_door, "m_door_w": wooden_door,
        "m_door_s": steel_door, "m_toil_m": toilet_male, "m_toil_f": toilet_female,
        "m_toil_d": disabled_toil, "m_mushola": mushola_unit, "m_fac_res": res_fac_m2,
        "m_fac_pub": pub_fac_m2, "m_fac_proj": proj_fac_u, "m_land_m2": land_m2,
        "r_fac_pre": facade_precast_pct, "r_fac_win": facade_window_pct, "r_fac_doub": facade_double_pct,
        "r_fl_ht": fl_ht_pct, "r_fl_vin": fl_vinyl_pct, "r_fl_mar": fl_marmer_pct,
        "r_san_qty": san_qty_room, "r_rail_qty": railing_qty,
        "u_earth": struc_earth, "u_found": struc_found, "u_struc": struc_work, "u_arch": arch_base,
        "u_lobby": lobby_rate, "u_f_pre": fac_precast_rate, "u_f_win": fac_window_rate, "u_f_doub": fac_double_rate,
        "u_d_wood": door_wood, "u_d_glass": door_glass, "u_d_steel": door_steel,
        "u_hw_wood": hw_wood, "u_hw_steel": hw_steel,
        "u_s_room": san_room_rate, "u_s_pub_m": san_pub_m, "u_s_pub_f": san_pub_f,
        "u_s_dis": san_dis, "u_s_mushola": san_mushola,
        "w_floor": fl_waste, "s_floor" :fl_skirt,
        "u_fl_ht": fl_ht_rate, "u_fl_vin": fl_vinyl_rate, "u_fl_mar": fl_marmer_rate,
        "u_carpet": carpet_rate, "u_glass": glass_rate, "u_sky": skylight_rate,
        "u_gondola": gondola_rate, "u_rail": railing_rate,
        "u_mep": mep_rate, "u_util": utility_rate,
        "u_ffe": ffe_rate, "u_kit": kitchen_rate, "u_misc": misc_rate,
        "u_ext": ext_land_rate, "u_fac_p": fac_pub_rate,
        "u_fac_r": fac_res_rate, "u_fac_pr": fac_proj_rate,
        "sc_cons": consultancy_rate, "sc_qs_m": qs_months, "sc_qs_r": qs_rate,
        "sc_pm_m": pm_months, "sc_pm_r": pm_rate, "sc_ins": insurance_pct
    })
    st.session_state.projects[curr_id]["data"]["grand_total_project"] = grand_total_project

    with tab8:
        st.header("Detail Pembuktian & Logika Perhitungan")

        # ==================================================
        # Group 1: Earthwork
        # group_earth = t_earth
        # ==================================================
        with st.expander("1. Earthwork", expanded=True):
            audit_earth = pd.DataFrame([
                {
                    "Item": "Earthwork",
                    "Formula": f"GBA ({gba:,.0f} m2) * Rate (Rp {struc_earth:,.0f})",
                    "Total": t_earth,
                },
            ])
            st.table(audit_earth.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 2: Foundation
        # group_found = t_found
        # ==================================================
        with st.expander("2. Foundation"):
            audit_found = pd.DataFrame([
                {
                    "Item": "Foundation",
                    "Formula": f"GBA ({gba:,.0f} m2) * Rate (Rp {struc_found:,.0f})",
                    "Total": t_found,
                },
            ])
            st.table(audit_found.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 3: Structure
        # group_struc = t_struc
        # ==================================================
        with st.expander("3. Structure"):
            audit_struc = pd.DataFrame([
                {
                    "Item": "Structural Work",
                    "Formula": f"GBA ({gba:,.0f} m2) * Rate (Rp {struc_work:,.0f})",
                    "Total": t_struc,
                },
            ])
            st.table(audit_struc.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 4: Architecture
        # group_arch =
        # t_arch_base + t_lobby + t_carpet + t_gondola
        # + t_glass_work + t_kitchen + t_railing + t_skylight
        # + group_facade + group_sanitary + group_floor + group_door
        # + smart_custom_costs
        # ==================================================
        with st.expander("4. Architecture"):
            audit_arch = pd.DataFrame([
                {
                    "Item": "Architecture Base",
                    "Formula": f"GFA ({gfa:,.0f} m2) * Rate (Rp {arch_base:,.0f})",
                    "Total": t_arch_base,
                },

                # Facade subgroup
                {
                    "Item": "Facade Precast",
                    "Formula": f"Facade ({facade:,.0f} m2) * {facade_precast_pct}% * Rate (Rp {fac_precast_rate:,.0f})",
                    "Total": t_precast,
                },
                {
                    "Item": "Facade Window Wall",
                    "Formula": f"Facade ({facade:,.0f} m2) * {facade_window_pct}% * Rate (Rp {fac_window_rate:,.0f})",
                    "Total": t_window,
                },
                {
                    "Item": "Facade Double Skin",
                    "Formula": f"Facade ({facade:,.0f} m2) * {facade_double_pct}% * Rate (Rp {fac_double_rate:,.0f})",
                    "Total": t_double,
                },

                # Sanitary subgroup
                {
                    "Item": "Typical Unit Sanitary",
                    "Formula": f"{rooms:,.0f} Rooms * {san_qty_room} Unit/Room * Rate (Rp {san_room_rate:,.0f})",
                    "Total": t_unit_san,
                },
                {
                    "Item": "Public Toilet Male",
                    "Formula": f"{toilet_male:,.0f} Unit * Rate (Rp {san_pub_m:,.0f})",
                    "Total": t_t_male,
                },
                {
                    "Item": "Public Toilet Female",
                    "Formula": f"{toilet_female:,.0f} Unit * Rate (Rp {san_pub_f:,.0f})",
                    "Total": t_t_female,
                },
                {
                    "Item": "Disabled Toilet",
                    "Formula": f"{disabled_toil:,.0f} Unit * Rate (Rp {san_dis:,.0f})",
                    "Total": t_t_dis,
                },
                {
                    "Item": "Mushola",
                    "Formula": f"{mushola_unit:,.0f} Unit * Rate (Rp {san_mushola:,.0f})",
                    "Total": t_mushola,
                },

                # Flooring subgroup
                {
                    "Item": "HT / Tile",
                    "Formula": f"GFA ({gfa:,.0f} m2) * {fl_ht_pct}% * {f_mult:.4f} * Rate (Rp {fl_ht_rate:,.0f})",
                    "Total": t_ht,
                },
                {
                    "Item": "Vinyl",
                    "Formula": f"GFA ({gfa:,.0f} m2) * {fl_vinyl_pct}% * {f_mult:.4f} * Rate (Rp {fl_vinyl_rate:,.0f})",
                    "Total": t_vinyl,
                },
                {
                    "Item": "Marmer",
                    "Formula": f"GFA ({gfa:,.0f} m2) * {fl_marmer_pct}% * {f_mult:.4f} * Rate (Rp {fl_marmer_rate:,.0f})",
                    "Total": t_marmer,
                },

                # Door subgroup
                {
                    "Item": "Wooden Door",
                    "Formula": f"{wooden_door:,.0f} Unit * Rate (Rp {door_wood:,.0f})",
                    "Total": t_w_door,
                },
                {
                    "Item": "Glass Door",
                    "Formula": f"{glass_door:,.0f} Unit * Rate (Rp {door_glass:,.0f})",
                    "Total": t_g_door,
                },
                {
                    "Item": "Steel Door",
                    "Formula": f"{steel_door:,.0f} Unit * Rate (Rp {door_steel:,.0f})",
                    "Total": t_s_door,
                },
                {
                    "Item": "Hardware Wooden Door",
                    "Formula": f"{wooden_door:,.0f} Unit * Rate (Rp {hw_wood:,.0f})",
                    "Total": t_hw_w,
                },
                {
                    "Item": "Hardware Steel Door",
                    "Formula": f"{steel_door:,.0f} Unit * Rate (Rp {hw_steel:,.0f})",
                    "Total": t_hw_s,
                },

                # Other architecture items
                {
                    "Item": "Lobby Interior",
                    "Formula": f"{lobby_interior:,.0f} m2 * Rate (Rp {lobby_rate:,.0f})",
                    "Total": t_lobby,
                },
                {
                    "Item": "Carpet Work",
                    "Formula": f"{carpet_m2:,.0f} m2 * Rate (Rp {carpet_rate:,.0f})",
                    "Total": t_carpet,
                },
                {
                    "Item": "Gondola",
                    "Formula": f"{gondola_unit:,.0f} Unit * Rate (Rp {gondola_rate:,.0f})",
                    "Total": t_gondola,
                },
                {
                    "Item": "Glass Work",
                    "Formula": f"{glass_m2:,.0f} m2 * Rate (Rp {glass_rate:,.0f})",
                    "Total": t_glass_work,
                },
                {
                    "Item": "Kitchen Equipment",
                    "Formula": f"{rooms:,.0f} Rooms * Rate (Rp {kitchen_rate:,.0f})",
                    "Total": t_kitchen,
                },
                {
                    "Item": "Railing",
                    "Formula": f"({rooms:,.0f} Rooms * {railing_qty} m'/room) * Rate (Rp {railing_rate:,.0f})",
                    "Total": t_railing,
                },
                {
                    "Item": "Skylight",
                    "Formula": f"{skylight_area:,.0f} m2 * Rate (Rp {skylight_rate:,.0f})",
                    "Total": t_skylight,
                },
                {
                    "Item": "Custom Items",
                    "Formula": "Total additional custom items",
                    "Total": smart_custom_costs,
                },
            ])

            st.markdown(fr"""
            **Rumus Pengali Lantai (f_mult):**  
            $(1 + \text{{Waste}} \%) \times (1 + \text{{Skirting}} \%) = (1 + {fl_waste/100}) \times (1 + {fl_skirt/100}) = **{f_mult:.4f}**$
            """)

            st.table(audit_arch.style.format({"Total": "Rp {:,.2f}"}))

            if smart_custom_costs > 0:
                st.markdown(f"**Detail Item Tambahan:** Rp {smart_custom_costs:,.2f}")
                for detail in breakdown_details:
                    st.markdown(f"- {detail}")

        # ==================================================
        # Group 5: FF&E
        # group_ffe = t_ffe + t_misc
        # ==================================================
        with st.expander("5. FF&E"):
            audit_ffe = pd.DataFrame([
                {
                    "Item": "FF&E",
                    "Formula": f"{rooms:,.0f} Rooms * Rate (Rp {ffe_rate:,.0f})",
                    "Total": t_ffe,
                },
                {
                    "Item": "Misc (Linen/Gym)",
                    "Formula": f"Switch ({misc_switch}) * Rate (Rp {misc_rate:,.0f})",
                    "Total": t_misc,
                },
            ])
            st.table(audit_ffe.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 6: MEP
        # group_mep = t_mep
        # ==================================================
        with st.expander("6. MEP"):
            audit_mep = pd.DataFrame([
                {
                    "Item": "MEP Works",
                    "Formula": f"GBA ({gba:,.0f} m2) * Rate (Rp {mep_rate:,.0f})",
                    "Total": t_mep,
                },
            ])
            st.table(audit_mep.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 7: Utility
        # group_utility = t_utility
        # ==================================================
        with st.expander("7. Utility"):
            audit_utility = pd.DataFrame([
                {
                    "Item": "Utility Connection",
                    "Formula": f"GBA ({gba:,.0f} m2) * Rate (Rp {utility_rate:,.0f})",
                    "Total": t_utility,
                },
            ])
            st.table(audit_utility.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 8: External Works
        # group_ext = t_external
        # ==================================================
        with st.expander("8. External Works"):
            audit_ext = pd.DataFrame([
                {
                    "Item": "External Works (Landscape)",
                    "Formula": f"{land_m2:,.0f} m2 * Rate (Rp {ext_land_rate:,.0f})",
                    "Total": t_external,
                },
            ])
            st.table(audit_ext.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 9: Misc / Facilities
        # group_misc = t_pub_fac + t_res_fac + t_proj_fac
        # ==================================================
        with st.expander("9. Misc / Facilities"):
            audit_misc = pd.DataFrame([
                {
                    "Item": "Public Facilities",
                    "Formula": f"{pub_fac_m2:,.0f} m2 * Rate (Rp {fac_pub_rate:,.0f})",
                    "Total": t_pub_fac,
                },
                {
                    "Item": "Resident Facilities",
                    "Formula": f"{res_fac_m2:,.0f} m2 * Rate (Rp {fac_res_rate:,.0f})",
                    "Total": t_res_fac,
                },
                {
                    "Item": "Project Facilities",
                    "Formula": f"{proj_fac_u:,.0f} Unit * Rate (Rp {fac_proj_rate:,.0f})",
                    "Total": t_proj_fac,
                },
            ])
            st.table(audit_misc.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 10: Preliminary
        # group_prelim = t_preliminary
        # ==================================================
        with st.expander("10. Preliminary"):
            audit_prelim = pd.DataFrame([
                {
                    "Item": "Preliminary Works",
                    "Formula": f"Construction Subtotal (Rp {construction_subtotal:,.0f}) * 5%",
                    "Total": t_preliminary,
                },
            ])
            st.table(audit_prelim.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 11: Contingency
        # group_conting = t_contingency
        # ==================================================
        with st.expander("11. Contingency"):
            audit_conting = pd.DataFrame([
                {
                    "Item": "Contingencies",
                    "Formula": f"(Subtotal + Prelim) (Rp {construction_subtotal + t_preliminary:,.0f}) * 3%",
                    "Total": t_contingency,
                },
            ])
            st.table(audit_conting.style.format({"Total": "Rp {:,.2f}"}))

        # ==================================================
        # Group 12: Soft Cost
        # group_soft_cost = total_soft_cost
        # ==================================================
        with st.expander("12. Soft Cost"):
            audit_soft = pd.DataFrame([
                {
                    "Item": "Consultancy Fee",
                    "Formula": f"GFA ({gfa:,.0f} m2) * Rate (Rp {consultancy_rate:,.0f})",
                    "Total": t_consultancy,
                },
                {
                    "Item": "QS Services",
                    "Formula": f"{qs_months} Months * Rate (Rp {qs_rate:,.0f})",
                    "Total": t_qs,
                },
                {
                    "Item": "PM Services",
                    "Formula": f"{pm_months} Months * Rate (Rp {pm_rate:,.0f})",
                    "Total": t_pm,
                },
                {
                    "Item": "Insurance Coverage",
                    "Formula": f"Total Hard Cost (Rp {grand_total_hc:,.0f}) * {insurance_pct}%",
                    "Total": t_insurance,
                },
            ])
            st.table(audit_soft.style.format({"Total": "Rp {:,.2f}"}))

#region --- DO NOT CHANGE#3 (OR GOD SO HELP ME) ---

def _parse_date_safe(value, fallback=None, dayfirst=True):
    if fallback is None:
        fallback = date.today()

    try:
        parsed = pd.to_datetime(value, dayfirst=dayfirst, errors="coerce")
        if pd.isna(parsed):
            return fallback
        return parsed.date()
    except Exception:
        return fallback

#endregion

def show_portfolio_summary():
    st.title("Summary")
    
    # ==========================================
    # 1. INITIALIZE EDITABLE SESSION STATE
    # ==========================================
    init_report_config()

    port_meta = get_port_meta()
    port_assumptions_df = get_port_assumptions_df()

    # ==========================================
    # 1B. SHARED SUMMARY UI + HEADER RENDERER
    # ==========================================
    from html import escape

    def safe_text(value):
        return escape(str(value if value is not None else ""))

    def render_portfolio_header(meta):
        title = safe_text(meta.get("title", ""))
        ref = safe_text(meta.get("ref", ""))
        version = safe_text(meta.get("version", ""))
        updated = safe_text(meta.get("updated", ""))
        created = safe_text(meta.get("created", ""))

        return f"""
<div class="asg-header">
<div class="asg-header-left">
ASG GROUP PROPERTY DEVELOPMENT<br>
QS & PROCUREMENT DIVISION<br>
{title}<br>
{ref}
</div>
<div class="asg-header-right">
VERSION &nbsp;&nbsp;: {version}<br>
UPDATED &nbsp;: {updated}<br>
CREATED &nbsp;: {created}
</div>
</div>
        """

    SUMMARY_CSS = """
    <style>
        /* ===============================
           CONFIG PAGE
        =============================== */
        .summary-config-panel {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 14px;
            padding: 16px 18px;
            margin-bottom: 14px;
            box-shadow: 0 1px 3px rgba(16,24,40,0.04);
        }

        .summary-config-title {
            font-size: 12px;
            font-weight: 750;
            color: #111827;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            margin-bottom: 8px;
        }

        .summary-config-desc {
            font-size: 13px;
            color: #6B7280;
            line-height: 1.55;
            margin-bottom: 12px;
        }

        .summary-kpi-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 12px;
            margin-bottom: 14px;
        }

        .summary-kpi-card {
            background: #F9FAFB;
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            padding: 12px 14px;
        }

        .summary-kpi-label {
            font-size: 11px;
            color: #6B7280;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            font-weight: 700;
            margin-bottom: 5px;
        }

        .summary-kpi-value {
            font-size: 16px;
            color: #111827;
            font-weight: 750;
            line-height: 1.35;
        }

        .summary-note-box {
            background: #F8F9FF;
            border: 1px solid #DDE1FF;
            border-left: 4px solid #3E4095;
            border-radius: 12px;
            padding: 12px 14px;
            font-size: 13px;
            color: #4B5563;
            line-height: 1.55;
            margin-bottom: 14px;
        }

        /* ===============================
           SHARED ASG REPORT PREVIEW
        =============================== */
        .asg-container {
            font-family: Calibri, Arial, sans-serif;
            font-size: 13px;
            color: #000000 !important;
            background-color: #FFFFFF !important;
            padding: 15px;
            border: 1px solid #D1D5DB;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(16,24,40,0.05);
            overflow-x: auto;
        }

        .asg-header {
            background-color: #0070C0 !important;
            color: #FFFFFF !important;
            padding: 7px 12px;
            font-weight: bold;
            font-size: 13px;
            display: flex;
            justify-content: space-between;
            gap: 16px;
            line-height: 1.45;
            margin-bottom: 15px;
            border: 1px solid #005A9C;
        }

        .asg-header-left {
            text-align: left;
        }

        .asg-header-right {
            text-align: right;
            min-width: 190px;
        }

        .asg-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background-color: #FFFFFF !important;
        }

        .asg-table th,
        .asg-table td {
            border: 2px solid #000000 !important;
            padding: 5px 8px;
            text-align: right;
            vertical-align: middle;
            color: #000000 !important;
        }

        .asg-table th {
            background-color: #F2F2F2 !important;
            text-align: center;
            font-weight: bold;
            color: #000000 !important;
        }

        .asg-table td.left {
            text-align: left;
            font-weight: bold;
        }

        .asg-table td.center {
            text-align: center;
            font-weight: bold;
        }

        .asg-table .bold-row td {
            font-weight: bold;
        }

        .asg-assumptions {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            background-color: #FFFFFF !important;
        }

        .asg-assumptions td {
            border: 1px solid #D9D9D9 !important;
            padding: 4px 8px;
            color: #000000 !important;
            vertical-align: top;
        }

        .asg-assumptions .yellow-header {
            background-color: #FFD966 !important;
            font-weight: bold;
            text-align: left;
            color: #000000 !important;
        }

        /* ===============================
           REKAP TABLE CONTAINER
        =============================== */
        .recap-wrapper {
            width: 100%;
            overflow-x: auto;
            font-family: Calibri, Arial, sans-serif;
            font-size: 11px;
        }

        .recap-table {
            border-collapse: separate;
            border-spacing: 0;
            white-space: nowrap;
        }

        .recap-table th {
            text-align: center !important;
            font-weight: bold;
            vertical-align: middle;
            border: 1px solid #000;
            padding: 4px 6px;
            background-color: #FFFFFF;
        }

        .recap-table td {
            border-right: 1px solid #000;
            border-bottom: 1px solid #000;
            border-left: 1px solid #000;
            padding: 4px 6px;
            background-color: #FFFFFF;
        }

        .sticky-col,
        .sticky-col2 {
            position: sticky;
            background-color: #F2F2F2 !important;
            z-index: 5;
        }

        .sticky-col3,
        .sticky-col4 {
            background-color: #F2F2F2 !important;
            z-index: 5;
        }

        .sticky-col {
            left: 0;
        }

        .sticky-col2 {
            left: 20px;
            text-align: left !important;
            min-width: 200px;
        }

        .bold-row {
            font-weight: bold;
            background-color: #F9F9F9;
        }

        @media (max-width: 900px) {
            .summary-kpi-grid {
                grid-template-columns: 1fr;
            }

            .asg-header {
                display: block;
            }

            .asg-header-right {
                text-align: left;
                margin-top: 8px;
            }
        }
    </style>
    """

    st.markdown(SUMMARY_CSS, unsafe_allow_html=True)

    # ==========================================
    # 2. TABS SETUP
    # ==========================================
    summary_list = ["Config", "FAD", "Rekap"]
    summary_tabs = st.tabs(summary_list)
    
    # --- TAB 1: EDITABLE NATIVE COMPONENTS ---
    # --- TAB 1: CONFIG ---
    with summary_tabs[0]:
        st.subheader("Config")
        st.caption("Configure the report header and assumptions used in both FAD and Rekap previews.")

        cfg = get_report_config()
        export_settings = cfg.setdefault("export_settings", {
            "prepared_by": "",
            "checked_by": ""
        })
        default_header_inputs = {
            "project_location": "PIK2.D2.",
            "project_name": "GINZA.MIDTOWN",
            "option_number": "2",
            "revision_number": "0",
            "drawing_date": "2026-02-02",
            "updated_date": date.today().strftime("%d-%m-%Y"),
            "created_date": date.today().strftime("%d-%m-%Y"),
        }
        header_inputs = cfg.setdefault("header_inputs", copy.deepcopy(default_header_inputs))
        for key, value in default_header_inputs.items():
            header_inputs.setdefault(key, value)

        st.markdown("##### Approval / Export Settings")

        c_prepared, c_checked = st.columns(2)

        export_settings["prepared_by"] = c_prepared.text_input(
            "Prepared By",
            value=export_settings.get("prepared_by", "")
        )

        export_settings["checked_by"] = c_checked.text_input(
            "Checked By",
            value=export_settings.get("checked_by", "")
        )

        project_count = len(st.session_state.projects)
        assumption_count = len(port_assumptions_df)

        st.markdown(
            f"""
            <div class="summary-kpi-grid">
                <div class="summary-kpi-card">
                    <div class="summary-kpi-label">Active Projects</div>
                    <div class="summary-kpi-value">{project_count}</div>
                </div>
                <div class="summary-kpi-card">
                    <div class="summary-kpi-label">Assumptions</div>
                    <div class="summary-kpi-value">{assumption_count} items</div>
                </div>
                <div class="summary-kpi-card">
                    <div class="summary-kpi-label">Output Format</div>
                    <div class="summary-kpi-value">FAD + Rekap</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        col_cfg, col_preview = st.columns([1.1, 1], gap="large")

        with col_cfg:
            with st.container(border=True):
                st.markdown("##### Header Configuration")
                st.caption("Input the variables only. The report header is generated automatically.")

                st.markdown(
                    """
                    <div class="summary-config-desc">
                        These fields control the blue report header. The same header will appear in both FAD and Rekap.
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                c1, c2 = st.columns(2)

                header_inputs["project_location"] = normalize_header_token(
                    c1.text_input(
                        "Project Location",
                        value=header_inputs.get("project_location", "PIK2.D2.")
                    ),
                    trailing_dot=True,
                )

                header_inputs["project_name"] = normalize_header_token(
                    c2.text_input(
                        "Project Name",
                        value=header_inputs.get("project_name", "GINZA.MIDTOWN")
                    )
                )

                c3, c4, c5 = st.columns(3)

                header_inputs["option_number"] = c3.text_input(
                    "Option Number",
                    value=header_inputs.get("option_number", "2")
                )

                header_inputs["revision_number"] = c4.text_input(
                    "Revision Number",
                    value=header_inputs.get("revision_number", "0")
                )

                drawing_date_value = c5.date_input(
                    "Drawing Date",
                    value=_parse_date_safe(
                        header_inputs.get("drawing_date", "2026-02-02"),
                        fallback=date(2026, 2, 2),
                        dayfirst=False
                    )
                )

                header_inputs["drawing_date"] = drawing_date_value.strftime("%Y-%m-%d")

                c6, c7 = st.columns(2)

                updated_date_value = c6.date_input(
                    "Updated Date",
                    value=_parse_date_safe(
                        header_inputs.get("updated_date", date.today().strftime("%d-%m-%Y")),
                        fallback=date.today(),
                        dayfirst=True
                    )
                )

                created_date_value = c7.date_input(
                    "Created Date",
                    value=_parse_date_safe(
                        header_inputs.get("created_date", date.today().strftime("%d-%m-%Y")),
                        fallback=date.today(),
                        dayfirst=True
                    )
                )

                header_inputs["updated_date"] = updated_date_value.strftime("%d-%m-%Y")
                header_inputs["created_date"] = created_date_value.strftime("%d-%m-%Y")

                generated_meta = build_portfolio_meta_from_inputs(header_inputs)
                port_meta.update(generated_meta)

        with col_preview:
            st.markdown("##### Header Preview")

            st.markdown(
                f"""
<div class="summary-note-box">
This preview uses the same header renderer used in FAD and Rekap.
Adjust the fields on the left and the preview will update automatically.
</div>
<div class="asg-container">
{render_portfolio_header(port_meta)}
</div>
                """,
                unsafe_allow_html=True
            )

        st.divider()

        st.markdown("##### Assumptions Configuration")
        st.caption("These assumptions are displayed below the FAD table. Keep the wording aligned with the required report format.")

        edited_assumptions = st.data_editor(
            get_port_assumptions_df(),
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
            column_config={
                "No.": st.column_config.TextColumn("No.", width="small"),
                "Assumption Description": st.column_config.TextColumn("Assumption Description", width="large"),
            }
        )

        save_c1, save_c2, save_c3 = st.columns([1, 2, 1])

        with save_c1:
            save_config_clicked = st.button(
                "Save Change",
                key="save_summary_config_to_cloud",
                type="primary",
                width="stretch"
            )

        if save_config_clicked:
            cfg["header_inputs"] = header_inputs
            port_meta.update(build_portfolio_meta_from_inputs(header_inputs))

            set_port_assumptions_df(edited_assumptions)

            save_ok = save_data_force()

            if save_ok:
                st.success("Saved to cloud.")
                st.rerun()
            else:
                st.error("Cloud save failed. Do not log out yet.")

    # --- TAB 2: EXACT FORMAT MIRROR (HTML/CSS) ---
    with summary_tabs[1]:
        st.subheader("Feasibility Analysis Data (FAD)")

        # 1. DATA PREPARATION (Define raw_data BEFORE anything else)
        raw_data = []
        tot_gba = tot_gfa = tot_sgfa = tot_budget = 0
        
        # Loop through projects to build the data list
        for sn, (pid, pdata) in enumerate(st.session_state.projects.items(), 1):
            d = pdata.get("data", {})
            ptype = pdata.get("type", "")

            gba = _safe_float(d.get("m_gba", 0.0))
            gfa = _safe_float(d.get("m_gfa", 0.0))
            sgfa = _safe_float(d.get("m_sgfa", 0.0))
            qty = _safe_float(d.get("m_rooms", 0.0))
            budget = _safe_float(d.get("grand_total_project", 0.0))
            
            raw_data.append({
                "SN": sn,
                "AREA": pdata.get("name", f"Project {sn}"),
                "GBA": gba, "GFA": gfa, "SGFA": sgfa,
                "QTY": qty, 
                "UNIT": "Units" if "Hotel" not in ptype else "RoomKey",
                "BUDGET": budget,
                "R_GBA": budget / gba if gba > 0 else 0,
                "R_GFA": budget / gfa if gfa > 0 else 0,
                "R_SGFA": budget / sgfa if sgfa > 0 else 0
            })
            tot_gba += gba; tot_gfa += gfa; tot_sgfa += sgfa; tot_budget += budget

        # Add the TOTAL row
        raw_data.append({
            "SN": "", "AREA": "TOTAL",
            "GBA": tot_gba, "GFA": tot_gfa, "SGFA": tot_sgfa,
            "QTY": None, "UNIT": "TOTAL",
            "BUDGET": tot_budget,
            "R_GBA": tot_budget / tot_gba if tot_gba > 0 else 0,
            "R_GFA": tot_budget / tot_gfa if tot_gfa > 0 else 0,
            "R_SGFA": tot_budget / tot_sgfa if tot_sgfa > 0 else 0
        })

        # 2. UI CONTROLS (Now they can safely see raw_data)
        col_btn1, col_btn2, _ = st.columns([1.5, 1.5, 3])
        
        with col_btn1:
            # This will now work because raw_data was defined in Step 1
            excel_output = generate_exact_portfolio_excel(
                port_meta,
                raw_data,
                get_port_assumptions_df()
            )
            
            st.download_button(
                label="Download Excel",
                data=excel_output,
                file_name="ASG_Portfolio_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                type="primary"
            )
                
        header_html = f"""
        <div class="asg-container">
            {render_portfolio_header(port_meta)}
        """

        # 3. Dynamic Data Table Core
        table_start = """
<table class="asg-table">
<thead>
<tr>
<th rowspan="2" style="width: 3%;">SN</th>
<th rowspan="2" style="width: 18%;">AREA</th>
<th colspan="3">BUILDING AREA (M2)</th>
<th colspan="2" style="width: 10%;">UNIT</th>
<th rowspan="2" style="width: 14%;">BUDGET ESTIMATE<br>RP</th>
<th colspan="3">COST RATIO RP/M2</th>
</tr>
<tr>
<th>GBA</th><th>GFA</th><th>SGFA</th>
<th></th><th></th>
<th>GBA</th><th>GFA</th><th>SGFA</th>
</tr>
</thead>
<tbody>
        """
        
        # 4. Generate Dynamic Rows from Active Projects
        table_rows = ""
        tot_gba = tot_gfa = tot_sgfa = tot_budget = 0
        
        for sn, (pid, pdata) in enumerate(st.session_state.projects.items(), 1):
            d = pdata.get("data", {})
            name = pdata.get("name", f"Project {sn}")
            ptype = pdata.get("type", "")
            
            gba = _safe_float(d.get("m_gba", 0.0))
            gfa = _safe_float(d.get("m_gfa", 0.0))
            sgfa = _safe_float(d.get("m_sgfa", 0.0))
            qty = _safe_float(d.get("m_rooms", 0.0))
            budget = _safe_float(d.get("grand_total_project", 0.0))
            
            if "Hotel" in ptype: unit_lbl = "RoomKey"
            elif "Parking" in ptype: unit_lbl = "lots"
            else: unit_lbl = "Units"
            
            r_gba = budget / gba if gba > 0 else 0
            r_gfa = budget / gfa if gfa > 0 else 0
            r_sgfa = budget / sgfa if sgfa > 0 else 0
            
            table_rows += f"""<tr class="bold-row">
<td class="center">{sn}</td>
<td class="left">{name}</td>
<td>{gba:,.2f}</td><td>{gfa:,.2f}</td><td>{sgfa:,.2f}</td>
<td class="center">{qty:,.0f}</td><td class="center">{unit_lbl}</td>
<td>{budget:,.0f}</td>
<td>{r_gba:,.0f}</td><td>{r_gfa:,.0f}</td><td>{r_sgfa:,.0f}</td>
</tr>"""
            
            tot_gba += gba
            tot_gfa += gfa
            tot_sgfa += sgfa
            tot_budget += budget

        tot_r_gba = tot_budget / tot_gba if tot_gba > 0 else 0
        tot_r_gfa = tot_budget / tot_gfa if tot_gfa > 0 else 0
        tot_r_sgfa = tot_budget / tot_sgfa if tot_sgfa > 0 else 0

        table_end = f"""<tr class="bold-row" style="background-color: #F2F2F2;">
<td class="center" colspan="2">TOTAL</td>
<td>{tot_gba:,.2f}</td><td>{tot_gfa:,.2f}</td><td>{tot_sgfa:,.2f}</td>
<td class="center" colspan="2">TOTAL</td>
<td>{tot_budget:,.0f}</td>
<td>{tot_r_gba:,.0f}</td><td>{tot_r_gfa:,.0f}</td><td>{tot_r_sgfa:,.0f}</td>
</tr>
</tbody>
</table>"""

        assumptions_html = """<table class="asg-assumptions">
<tr>
    <td class="yellow-header" style="width: 3%;">I.</td>
    <td class="yellow-header">ASSUMPTIONS</td>
</tr>"""
        for _, row in get_port_assumptions_df().iterrows():
            num = row.get("No.", "")
            desc = row.get("Assumption Description", "")
            if pd.notna(desc) and str(desc).strip() != "":
                assumptions_html += f"""<tr>
<td style="text-align: center;">{num}</td>
<td>{desc}</td>
</tr>"""
        assumptions_html += """</table>"""

        full_html = header_html + table_start + table_rows + table_end + assumptions_html + "</div>"

        st.markdown(full_html, unsafe_allow_html=True)

# --- TAB 3: WIDE RECAP COST ---
    with summary_tabs[2]:
        st.subheader("Comprehensive Recap Matrix (Cost & Ratios)")

        recap_excel_data = generate_recap_excel(port_meta, st.session_state.projects)
        st.session_state.recap_math_engine = get_recap_values
        math_engine = st.session_state.recap_math_engine
        
        col_btn, col_info = st.columns([1.5, 4.5])
        with col_btn:
            st.download_button(
                label="Download Excel",
                data=recap_excel_data,
                file_name="ASG_Recap_Cost_Wide.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                type="primary"
            )

        with col_info:
            curr_id, curr_proj = get_current_project()
            curr_data = curr_proj.get("data", {}) if isinstance(curr_proj, dict) else {}
            stored_total = _safe_float(curr_data.get("grand_total_project", 0))
            recap_values = math_engine(curr_proj) if isinstance(curr_proj, dict) else {}
            recap_total = _safe_float(recap_values.get("TOTAL, EXCLD PPN", 0))
            diff_total = recap_total - stored_total
            diff_pct = (diff_total / stored_total * 100) if stored_total > 0 else 0
            tolerance = max(1000, stored_total * 0.0001)
            control_msg = (
                f"Stored Cost Analysis Total: Rp {stored_total:,.0f} | "
                f"Recap Calculated Total: Rp {recap_total:,.0f} | "
                f"Difference: Rp {diff_total:,.0f} ({diff_pct:,.4f}%)"
            )

            if abs(diff_total) <= tolerance:
                st.success(f"Recap matches Cost Analysis control total. {control_msg}")
            else:
                st.warning(
                    "Recap differs from stored Cost Analysis total. Check whether Cost Analysis was saved or formulas changed. "
                    + control_msg
                )

        # --- GENERATE HTML PREVIEW ---
        bg_colors = ["#EAEAEA", "#FCE4D6", "#F2DCDB", "#E1D5E7", "#DDEBF7", "#E2EFDA", "#D9E1F2", "#F4B084", "#FFF2CC"]
        project_list = [("TOTAL", {"name": "TOTAL"})] + list(st.session_state.projects.items())

        tot_cache = {}; global_cost = {}; tot_gba = tot_gfa = tot_sgfa = 0
        for pid, pdata in st.session_state.projects.items():
            vals = math_engine(pdata)
            tot_cache[pid] = vals
            for k, v in vals.items(): global_cost[k] = global_cost.get(k, 0) + v
            d = pdata.get("data", {})
            tot_gba += _safe_float(d.get("m_gba", 0)); tot_gfa += _safe_float(d.get("m_gfa", 0)); tot_sgfa += _safe_float(d.get("m_sgfa", 0))

        global_hc = global_cost.get("HARDCOST", 0); safe_hc = global_hc if global_hc > 0 else 1
        global_sc = global_cost.get("SOFTCOST", 0); safe_sc = global_sc if global_sc > 0 else 1

        html_str = f"""
        <div class="asg-container">
        {render_portfolio_header(port_meta)}

<div class="recap-wrapper">
<table class="recap-table">
        """

        html_str += """
<tr><th rowspan='4' class='sticky-col'>SN</th>
<th rowspan='4' class='sticky-col2' style='min-width:200px;'>DESCRIPTION</th>
<th rowspan='4' class='sticky-col3'>COA</th><th rowspan='4' class='sticky-col4'>%</th>
        """
        for i, (pid, pdata) in enumerate(project_list):
            color = bg_colors[i % len(bg_colors)]
            html_str += f"<th colspan='5' style='background-color:{color}; color:#000;'>{pdata.get('name', 'PROJECT').upper()}</th>"
        html_str += "</tr><tr>"
        for i in range(len(project_list)):
            color = bg_colors[i % len(bg_colors)]
            html_str += f"<th style='background-color:{color};'>ESTIMATE</th><th colspan='4' style='background-color:{color};'>Cost Ratio (Rp/m2)</th>"
        html_str += "</tr><tr>"
        for i in range(len(project_list)):
            color = bg_colors[i % len(bg_colors)]
            html_str += f"<th style='background-color:{color};'>TOTAL</th><th style='background-color:{color};'>GBA</th><th style='background-color:{color};'>GFA</th><th style='background-color:{color};'>SGFA</th><th style='background-color:{color};'>NFA</th>"
        html_str += "</tr><tr>"
        for i, (pid, pdata) in enumerate(project_list):
            color = bg_colors[i % len(bg_colors)]
            if pid == "TOTAL":
                gba, gfa, sgfa, nfa = tot_gba, tot_gfa, tot_sgfa, tot_gfa * 0.82
            else:
                d = pdata.get("data", {})
                gba, gfa, sgfa = _safe_float(d.get("m_gba", 0)), _safe_float(d.get("m_gfa", 0)), _safe_float(d.get("m_sgfa", 0))
                nfa = gfa * 0.82
            html_str += f"<th style='background-color:{color};'>Rp</th><th style='background-color:{color};'>{gba:,.0f}</th><th style='background-color:{color};'>{gfa:,.0f}</th><th style='background-color:{color};'>{sgfa:,.0f}</th><th style='background-color:{color};'>{nfa:,.0f}</th>"
        html_str += "</tr>"

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

        for sn, desc, coa, is_bold, cat in row_mapping:
            global_val = global_cost.get(desc, 0)
            if cat == "HC": pct = global_val / safe_hc
            elif cat == "SC": pct = global_val / safe_sc
            elif cat in ["SC_TOTAL", "TOTAL"]: pct = global_val / safe_hc
            else: pct = 0

            tr_class = " class='bold-row'" if is_bold else ""
            html_str += f"<tr{tr_class}>"
            # Sticky Columns (Anchor columns remain neutral grey)
            html_str += f"<td class='sticky-col' style='text-align:center;'>{sn}</td>"
            html_str += f"<td class='sticky-col2'>{desc}</td>"
            html_str += f"<td class='sticky-col3'>{coa}</td>"
            html_str += f"<td class='sticky-col4'>{pct*100:.2f}%</td>"
            
            # Project Data Columns (Colored to match headers)
            for i, (pid, pdata) in enumerate(project_list):
                val = global_cost.get(desc, 0) if pid == "TOTAL" else tot_cache[pid].get(desc, 0)
                color = bg_colors[i % len(bg_colors)] # Get the header's color
                
                # Calculate divisors
                if pid == "TOTAL":
                    gba_f, gfa_f, sgfa_f = tot_gba or 1, tot_gfa or 1, tot_sgfa or 1
                    nfa_f = (tot_gfa * 0.82) or 1
                else:
                    d = pdata.get("data", {})
                    gba_f = _safe_float(d.get("m_gba") or 1)
                    gfa_f = _safe_float(d.get("m_gfa") or 1)
                    sgfa_f = _safe_float(d.get("m_sgfa") or 1)
                    nfa_f = gfa_f * 0.82 or 1
                
                # Apply the background color style to every <td> in this column
                c_style = f"style='background-color:{color};'"
                
                html_str += f"<td {c_style}>{val:,.0f}</td>"
                html_str += f"<td {c_style}>{val/gba_f:,.0f}</td>"
                html_str += f"<td {c_style}>{val/gfa_f:,.0f}</td>"
                html_str += f"<td {c_style}>{val/sgfa_f:,.0f}</td>"
                html_str += f"<td {c_style}>{val/nfa_f:,.0f}</td>"
            
            # Spacer for desktop 'over-scroll' comparison
            html_str += "<td style='border:none; background:transparent; min-width:600px;'></td>"
            html_str += "</tr>"

        html_str += "</table></div></div>"
        st.markdown(html_str, unsafe_allow_html=True)

def show_snapshots():
    st.title("Archive")
    curr_id, curr_proj = get_current_project()

    atab1, atab2, atab3= st.tabs([
        "Online Backup", "Offline Backup", "Notes"
    ])

    with atab3:
        st.subheader("Notes")

        curr_id, pdata = get_current_project()
        pdata.setdefault("data", {})
        pdata["data"].setdefault("simple_notes", [])

        note_key = f"simple_note_input_{curr_id}"

        def add_simple_note():
            curr_id_cb, pdata_cb = get_current_project()
            pdata_cb.setdefault("data", {})
            pdata_cb["data"].setdefault("simple_notes", [])

            clean_note = str(st.session_state.get(note_key, "")).strip()

            if clean_note == "":
                st.session_state["simple_note_warning"] = "Please write a note first."
                return

            pdata_cb["data"]["simple_notes"].insert(0, {
                "created_at": datetime.utcnow().isoformat(),
                "note": clean_note
            })

            st.session_state[note_key] = ""
            st.session_state["simple_note_warning"] = ""
            save_ok = save_after_user_action("Add Simple Note")

        st.text_area(
            "Write a note",
            placeholder="Example: Check area lanskap and rate later...",
            height=120,
            key=note_key
        )

        st.button(
            "Add Note",
            type="primary",
            width="stretch",
            key=f"add_simple_note_btn_{curr_id}",
            on_click=add_simple_note
        )

        if st.session_state.get("simple_note_warning"):
            st.warning(st.session_state["simple_note_warning"])

        st.divider()
        st.markdown("### Saved Notes")

        notes = pdata["data"].get("simple_notes", [])

        if not notes:
            st.info("No notes yet.")
        else:
            for idx, note in enumerate(notes):
                created_at = note.get("created_at", "")

                try:
                    created_dt = datetime.fromisoformat(created_at)
                    created_display = (created_dt + timedelta(hours=7)).strftime("%d %b %Y, %H:%M WIB")
                except Exception:
                    created_display = created_at

                with st.container(border=True):
                    col_note, col_delete = st.columns([5, 1])

                    with col_note:
                        st.caption(created_display)
                        st.write(note.get("note", ""))

                    with col_delete:
                        if st.button(
                            "Delete",
                            key=f"delete_simple_note_{curr_id}_{idx}",
                            width="stretch"
                        ):
                            pdata["data"]["simple_notes"].pop(idx)
                            save_ok = save_after_user_action("Delete Simple Note")

                            if save_ok:
                                st.rerun()

    with atab1:
        # --- CREATE LOCAL STUDY / SAVE NEW SNAPSHOT ---
        st.header("Online Backup")
        st.subheader("Create New Study")

        snapshots = load_snapshots()

        def unique_snapshot_name(base_name):
            clean_base = str(base_name).strip()
            existing_names = {
                str(snap.get("snapshot_name", "")).strip()
                for snap in snapshots
                if isinstance(snap, dict)
            }

            if clean_base not in existing_names:
                return clean_base

            idx = 1
            while f"{clean_base} ({idx})" in existing_names:
                idx += 1

            return f"{clean_base} ({idx})"

        projects = st.session_state.get("projects", {})
        curr_id = st.session_state.get("current_proj_id")
        has_active_working_study = (
            isinstance(projects, dict)
            and len(projects) > 0
            and curr_id in projects
        )

        col1, _ = st.columns([4, 3])
        snapshot_name = col1.text_input(
            "Name", 
            placeholder="e.g. Project X - Opt 1 - Rev 0"
        )
        col_create, col_save, _ = st.columns([1, 1, 5])

        if col_create.button("Create New", width="stretch", icon=icon_safe("create_new_folder")):
            if snapshot_name.strip() == "":
                col_create.warning("Please enter Project name.")
            else:
                if create_new_feasibility_study(snapshot_name):
                    st.success(f"Started local study **{snapshot_name.strip()}**. Use Save to archive it.")
                    st.rerun()

        if col_save.button(
            "Save",
            width="stretch",
            icon=icon_safe("save_as"),
            disabled=not has_active_working_study,
        ):
            if snapshot_name.strip() == "":
                col_save.warning("Please enter Project name.")
            else:
                save_name = unique_snapshot_name(snapshot_name)
                if save_snapshot(save_name):
                    st.success(f"Study **{save_name}** saved to archive.")
                    st.rerun()

        st.divider()

        # --- LIST EXISTING SNAPSHOTS ---
        st.subheader("Load File")

        if not snapshots:
            st.info("No saved projects yet.")
        else:
            # ==================================================
            # PAGINATION SETUP
            # ==================================================
            PAGE_SIZE = 10

            if "archive_page" not in st.session_state:
                st.session_state.archive_page = 0

            total_items = len(snapshots)
            total_pages = max(1, (total_items + PAGE_SIZE - 1) // PAGE_SIZE)

            # Safety clamp if files were deleted / changed
            if st.session_state.archive_page >= total_pages:
                st.session_state.archive_page = total_pages - 1

            if st.session_state.archive_page < 0:
                st.session_state.archive_page = 0

            start_idx = st.session_state.archive_page * PAGE_SIZE
            end_idx = start_idx + PAGE_SIZE
            page_snapshots = snapshots[start_idx:end_idx]

            st.caption(
                f"Showing {start_idx + 1}-{min(end_idx, total_items)} of {total_items} saved files"
            )

            st.divider()

            # ==================================================
            # PAGINATED SNAPSHOT LIST
            # ==================================================
            for snap in page_snapshots:
                snap_id = snap["id"]

                rename_key = f"renaming_{snap_id}"
                delete_key = f"deleting_{snap_id}"

                if rename_key not in st.session_state:
                    st.session_state[rename_key] = False

                if delete_key not in st.session_state:
                    st.session_state[delete_key] = False

                col1, col2, col3, col4 = st.columns([4, 1, 1, 1])

                from datetime import datetime, timedelta

                created_utc = datetime.fromisoformat(snap["created_at"].replace("Z", "+00:00"))
                created_local = created_utc + timedelta(hours=7)
                formatted_date = created_local.strftime("%d %b %Y, %H:%M")

                is_active = st.session_state.get("loaded_snapshot_id") == snap_id
                active_label = " - ACTIVE" if is_active else ""

                # ==================================================
                # NORMAL DISPLAY MODE
                # ==================================================
                if not st.session_state[rename_key] and not st.session_state[delete_key]:
                    with col1:
                        st.markdown(f"**{snap['snapshot_name']}**")
                        st.caption(f"Saved: {formatted_date} WIB{active_label}")

                    if col2.button("Rename", key=f"rename_start_{snap_id}", width="stretch"):
                        st.session_state[rename_key] = True
                        st.session_state[delete_key] = False
                        st.rerun()

                    if col3.button("Load", key=f"load_{snap_id}", type="primary", width="stretch"):
                        data = load_snapshot_data(snap_id)

                        if data:
                            restore_app_payload(data)

                            st.session_state.loaded_snapshot_id = snap_id
                            st.session_state.loaded_snapshot_name = snap["snapshot_name"]
                            st.session_state.current_study_name = snap["snapshot_name"]

                            save_data()
                            st.success(f"Loaded **{snap['snapshot_name']}**.")
                            st.rerun()

                    if col4.button("Delete", key=f"delete_start_{snap_id}", width="stretch"):
                        st.session_state[delete_key] = True
                        st.session_state[rename_key] = False
                        st.rerun()

                # ==================================================
                # RENAME MODE
                # ==================================================
                elif st.session_state[rename_key]:
                    with col1:
                        new_archive_name = st.text_input(
                            "New file name",
                            value=snap["snapshot_name"],
                            key=f"rename_input_{snap_id}",
                            label_visibility="collapsed"
                        )
                        st.caption(f"Saved: {formatted_date} WIB{active_label}")

                    if col2.button("Save Name", key=f"rename_save_{snap_id}", type="primary", width="stretch"):
                        if rename_snapshot(snap_id, new_archive_name):
                            st.session_state[rename_key] = False
                            st.success("File renamed.")
                            st.rerun()

                    if col3.button("Cancel", key=f"rename_cancel_{snap_id}", width="stretch"):
                        st.session_state[rename_key] = False
                        st.rerun()

                    with col4:
                        st.empty()

                # ==================================================
                # DELETE CONFIRMATION MODE
                # ==================================================
                elif st.session_state[delete_key]:
                    with col1:
                        st.warning(f"Delete **{snap['snapshot_name']}**?")
                        st.caption("This archived file will be removed from the library.")

                    if col2.button("Confirm", key=f"delete_confirm_{snap_id}", type="primary", width="stretch"):
                        was_active = st.session_state.get("loaded_snapshot_id") == snap_id

                        if delete_snapshot(snap_id):
                            if was_active:
                                if not st.session_state.get("current_study_name"):
                                    st.session_state.current_study_name = resolve_current_study_name()
                                save_data()

                            st.session_state[delete_key] = False
                            st.success("File deleted.")
                            st.rerun()

                    if col3.button("Cancel", key=f"delete_cancel_{snap_id}", width="stretch"):
                        st.session_state[delete_key] = False
                        st.rerun()

                    with col4:
                        st.empty()

                st.divider()

            # ==================================================
            # PAGINATION CONTROLS
            # ==================================================
            c1, col_prev, col_page, col_next, c2 = st.columns([5, 1, 2, 1, 5])

            with col_prev:
                if st.button(
                    "Previous",
                    key="archive_prev_page",
                    width="stretch",
                    disabled=st.session_state.archive_page <= 0
                ):
                    st.session_state.archive_page -= 1
                    st.rerun()

            with col_page:
                st.markdown(
                    f"<div style='text-align:center; padding-top: 0.45rem;'>"
                    f"Page {st.session_state.archive_page + 1} of {total_pages}"
                    f"</div>",
                    unsafe_allow_html=True
                )

            with col_next:
                if st.button(
                    "Next",
                    key="archive_next_page",
                    width="stretch",
                    disabled=st.session_state.archive_page >= total_pages - 1
                ):
                    st.session_state.archive_page += 1
                    st.rerun()  

    with atab2:
        st.header("Offline Backup")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Import")
            uploaded_file = st.file_uploader("Upload CSV Database:", type=["csv"])

            if uploaded_file is not None:
                file_key = getattr(uploaded_file, 'file_id', uploaded_file.name)
                
                if "last_loaded_file" not in st.session_state or st.session_state.last_loaded_file != file_key:
                    try:
                        import ast  # Needed to turn strings back into tables
                        df_import = pd.read_csv(uploaded_file)
                        df_import = df_import.replace({np.nan: None})
                        
                        if df_import is not None and not df_import.empty:
                            global_temp_custom = {}
                            import_pid_map = {} # Track mapped IDs to prevent row-by-row renaming

                            for index, row in df_import.iterrows():
                                # Get raw ID
                                raw_pid = str(row.get("Project_ID", curr_id)).strip()
                                key = str(row.get("Metric_Key", "")).strip()
                                val = row.get("Value", "")

                                if not key or pd.isna(val): continue 

                                # 1. Handle Project ID Mapping (Renaming if duplicate exists)
                                if raw_pid not in import_pid_map:
                                    target_pid = raw_pid
                                    if target_pid in st.session_state.projects:
                                        # Unique suffix for this file import
                                        target_pid = f"{target_pid}_imported_{file_key[:4]}"
                                    import_pid_map[raw_pid] = target_pid

                                pid = import_pid_map[raw_pid]

                                if pid not in st.session_state.projects:
                                    st.session_state.projects[pid] = {
                                        "name": f"Imported Project {pid}",
                                        "type": "Hotel", 
                                        "data": {}
                                    }
                                
                                # 2. Handle Metadata
                                if key == "proj_name":
                                    st.session_state.projects[pid]["name"] = str(val)
                                elif key == "proj_type":
                                    st.session_state.projects[pid]["type"] = str(val)
                                
                                # 3. Handle Custom Item Reconstruction
                                elif key.startswith(("input_name", "input_rate", "input_qty")):
                                    if pid not in global_temp_custom:
                                        global_temp_custom[pid] = {}
                                    try:
                                        idx = int(''.join(filter(str.isdigit, key)))
                                        if idx not in global_temp_custom[pid]:
                                            global_temp_custom[pid][idx] = {"Item Description": "", "Rate (Rp)": 0.0, "Quantity": 1.0}
                                        
                                        if "name" in key: global_temp_custom[pid][idx]["Item Description"] = str(val)
                                        elif "rate" in key: global_temp_custom[pid][idx]["Rate (Rp)"] = _safe_float(val)
                                        elif "qty" in key: global_temp_custom[pid][idx]["Quantity"] = _safe_float(val)
                                    except: continue

                                # 4. Handle Standard Metrics & Nested Tables (Area Analysis)
                                else:
                                    try:
                                        str_val = str(val).strip()

                                        # Strip surrounding quotes that CSV may add
                                        if str_val.startswith('"') and str_val.endswith('"'):
                                            str_val = str_val[1:-1]

                                        if str_val.startswith("[{") or str_val.startswith("['"): 
                                            # Try JSON first (clean), then fall back to Python literal
                                            try:
                                                st.session_state.projects[pid]["data"][key] = _json.loads(str_val)
                                            except Exception:
                                                try:
                                                    st.session_state.projects[pid]["data"][key] = ast.literal_eval(str_val)
                                                except Exception:
                                                    st.session_state.projects[pid]["data"][key] = str_val
                                        else:
                                            try:
                                                st.session_state.projects[pid]["data"][key] = _safe_float(val)
                                            except (ValueError, TypeError):
                                                st.session_state.projects[pid]["data"][key] = str(val)
                                    except (ValueError, TypeError, SyntaxError):
                                        st.session_state.projects[pid]["data"][key] = str(val)

                            # 5. Finalize Custom Item Lists
                            for pid_key, items_dict in global_temp_custom.items():
                                sorted_custom = [items_dict[i] for i in sorted(items_dict.keys())]
                                st.session_state.projects[pid_key]["data"]["smart_custom_costs"] = sorted_custom

                            # 6. CLEAR UI CACHE ANCHORS (Crucial for Area Analysis reload)
                            clear_project_ui_cache()

                            st.session_state.last_loaded_file = file_key
                            save_data()
                            st.success(f"OK Import Successful! New projects created to avoid overwrites.")
                            st.rerun()
                        else:
                            st.warning("Warning: The uploaded CSV is empty.")
                    except Exception as e:
                        st.error(f"Error: Error during import: {e}")
                        
        with c2:
            st.subheader("Export")
            # --- 1. CURRENT PROJECT ONLY ---
            current_project_csv = []
            current_project_csv.append({"Project_ID": curr_id, "Metric_Key": "proj_name", "Value": curr_proj["name"]})
            current_project_csv.append({"Project_ID": curr_id, "Metric_Key": "proj_type", "Value": curr_proj["type"]})
            
            for k, v in st.session_state.projects[curr_id]["data"].items():
                if k not in ("header_info", "assumptions"):
                    serialized_v = _json.dumps(v) if isinstance(v, list) else v
                    current_project_csv.append({"Project_ID": curr_id, "Metric_Key": k, "Value": serialized_v})

            df_curr = pd.DataFrame(current_project_csv)
            csv_buffer = df_curr.to_csv(index=False).encode("utf-8")

            # --- 2. GLOBAL DATABASE ---
            all_projects_csv = []
            for pid, pdata in st.session_state.projects.items():
                all_projects_csv.append({"Project_ID": pid, "Metric_Key": "proj_name", "Value": pdata["name"]})
                all_projects_csv.append({"Project_ID": pid, "Metric_Key": "proj_type", "Value": pdata["type"]})
                for k, v in pdata["data"].items():
                    if k not in ("header_info", "assumptions"):
                        serialized_v = _json.dumps(v) if isinstance(v, list) else v
                        all_projects_csv.append({
                            "Project_ID": pid,
                            "Metric_Key": k,
                            "Value": serialized_v
                        })

            df_all = pd.DataFrame(all_projects_csv)
            csv_all_buffer = df_all.to_csv(index=False).encode("utf-8")

            # --- 3. DOWNLOAD BUTTONS ---
            st.download_button(
                label=f"Download {curr_proj['name']} only",
                data=csv_buffer,
                file_name=f"Project_{curr_id}.csv",
                mime="text/csv",
                width="stretch"
            )

            active_file_name = st.session_state.get("loaded_snapshot_name")

            st.download_button(
                label=f"Download {active_file_name} File",
                data=csv_all_buffer,
                file_name="ProCalc_Global_Database.csv",
                mime="text/csv",
                width="stretch",
                type="primary",
                icon=icon_safe("download")
            )

#region --- LOGIN SCREEN AND SIDE BAR(INSIDE MAIN APP) ---
# 1. SETUP & SESSION CHECK
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

import time

def login_screen():
    st.markdown("""
        <style>
            /* Hide the default Streamlit sidebar and top header */
            [data-testid="stSidebar"] {display: none;}
            [data-testid="stHeader"] {display: none;}
            
            /* --- RESPONSIVE SPACING FIX --- */
            /* 1. Default spacing for Desktop (PC) */
            .block-container {
                padding-top: 4rem !important; 
            }
            
            /* 2. Overrides for Mobile (Screens smaller than 768px) */
            @media (max-width: 768px) {
                .block-container {
                    padding-top: 1rem !important; /* Pulls everything up on mobile */
                }
                /* Hides the empty left column on mobile so it doesn't push the form down */
                [data-testid="column"]:nth-of-type(1) {
                    display: none !important;
                }
            }
            
            /* Typography & Alignment */
            .login-header {
                text-align: center;
                margin-top: 1rem;
                margin-bottom: 1.5rem;
            }
            .login-title {
                color: #1E3A8A;
                font-size: 26px;
                font-weight: 700;
                margin-bottom: 5px;
            }
            .login-subtitle {
                color: #6B7280;
                font-size: 15px;
            }
            
            /* Form Card Styling */
            [data-testid="stForm"] {
                border: 1px solid rgba(49, 51, 63, 0.1);
                border-radius: 12px;
                padding: 2rem;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
                background-color: var(--background-color);
            }
            
            /* HTML Logo Styling */
            .logo-container {
                display: flex;
                justify-content: center;
                margin-bottom: 0.5rem;
            }
            .logo-container img {
                width: 45%; /* Keeps the logo scaled similarly to the old column layout */
                max-width: 200px;
            }
        </style>
    """, unsafe_allow_html=True)
        
    # 2. Adjust Layout Proportions (creates a clean, focused center column)
    col1, center_col, col3 = st.columns([1.5, 2, 1.5])
    
    with center_col:
        # 3. Nesting the logo right above the text for perfect alignment
        logo_col_1, logo_col_2, logo_col_3 = st.columns([1, 1.5, 1])
        with logo_col_2:
            st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTQWTNsgu_c6WzJAehb4zQ3qdTKNauleAXe4w&s", width="stretch")
            
        st.markdown("""
            <div class="login-header">
                <div class="login-title">Project Feasibility Study</div>
                <div class="login-subtitle">Please sign in to your account</div>
            </div>
        """, unsafe_allow_html=True)
        
        # 4. The Login Form
        with st.form("login_gate", clear_on_submit=False):
            # Added placeholders for better UX
            email = st.text_input("Email", placeholder="name@agungsedayu.com")
            password = st.text_input("Password", type="password", placeholder="********")
            
            st.write("") # Small gap before button
            submit = st.form_submit_button("Sign In", width="stretch", type="primary")
            
            if submit:
                # Basic input validation
                if not email or not password:
                    st.warning("Please enter both email and password.", icon="Warning:")
                else:
                    # Added a spinner so the UI doesn't freeze during API calls
                    with st.spinner("Authenticating securely..."):
                        try:
                            res = supabase.auth.sign_in_with_password({
                                "email": email, 
                                "password": password
                            })
                            
                            supabase.postgrest.auth(res.session.access_token)
                            st.session_state.logged_in = True
                            st.session_state.user = res.user
                            st.session_state.access_token = res.session.access_token
                            # Register this device session fingerprint
                            register_device_session()

                            # Clear projects so main_app() re-loads from Supabase
                            for key in ["projects", "storage_loaded"]:
                                if key in st.session_state:
                                    del st.session_state[key]
                            
                            # Format username to look cleaner (e.g., john.doe -> John Doe)
                            raw_username = res.user.email.split("@")[0]
                            clean_username = raw_username.replace('.', ' ').title()
                            
                            st.success(f"Identity Verified: **{clean_username}**", icon=icon_safe("person_check"))
                            time.sleep(0.8)  # Slightly faster transition
                            st.rerun()
                            
                        except Exception as e:
                            # Catch generic invalid credential messages and make them user-friendly
                            error_msg = str(e)
                            if "Invalid login credentials" in error_msg or "400" in error_msg:
                                st.error("Invalid email or password. Please try again.", icon="Error:")
                            else:
                                st.error(f"Authentication error: {error_msg}", icon="Error")

    # 5. Professional Footer
    st.markdown(f"""
        <hr style="border: none; border-top: 1px solid #E5E7EB; margin-top: 50px; margin-bottom: 20px;">
        <div style='text-align: center; color: #9CA3AF; font-size: 12px; font-family: sans-serif; line-height: 1.6;'>
            v{APP_VERSION} | &copy; 2026 QS & Procurement - ASG. All rights reserved.<br>
            <span style="letter-spacing: 1px; font-weight: 500;">INTERNAL AGUNG SEDAYU GROUP USE ONLY</span>
        </div>
    """, unsafe_allow_html=True)

# 3. THE ACTUAL APPLICATION
def main_app():
    # The 'Assembler'
    ensure_app_state_loaded()

    enforce_single_device()  # <- add this

    curr_id, curr_proj = get_current_project()

    #region --- SIDEBAR ----
    st.sidebar.title("Main Navigation")

    user = st.session_state.get("user")
    user_email = getattr(user, "email", "user@example.com")
    username = user_email.split("@")[0]
    clean_name = username.replace('.', ' ').title()
    st.sidebar.markdown(f"Hello, **{clean_name}**!")


    page_choice = st.sidebar.radio(
        "Pilih Pekerjaan:",
        ["Start", "Area Analysis", "Cost Analysis", "Database", "Summary", "Archive"]
    )

    # Always build sidebar list AFTER project repair
    curr_id, curr_proj = get_current_project()

    proj_ids = list(st.session_state.projects.keys())
    proj_labels = [
        f"{st.session_state.projects[pid]['name']} ({st.session_state.projects[pid]['type']})"
        for pid in proj_ids
    ]

    current_index = proj_ids.index(curr_id) if curr_id in proj_ids else 0
    
    # ==================================================
    # SIDEBAR CURRENT FILE / QUICK SAVE
    # ==================================================
    active_archive_id = st.session_state.get("loaded_snapshot_id")
    active_archive_name = st.session_state.get("loaded_snapshot_name")
    active_file_name = st.session_state.get("loaded_snapshot_name")
    current_study_name = resolve_current_study_name()
    st.session_state.current_study_name = current_study_name

    st.sidebar.caption(f"Current study: {current_study_name}")
    if active_archive_id and active_archive_name:
        st.sidebar.caption("Status: Linked to archive")
    else:
        st.sidebar.caption("Status: Unsaved / not archived")

    on = st.sidebar.toggle("Show detailed control.")
    if on:
        if active_archive_id:
            if st.sidebar.button(
                "Quick Save",
                key="sidebar_save_current_archive",
                type="primary",
                width="stretch",
                icon=mi("save") if "mi" in globals() else None,
                help=f"{active_file_name} is currently loaded."
            ):
                snapshot_ok = overwrite_current_snapshot()
                storage_ok = save_data()

                if snapshot_ok and storage_ok:
                    st.sidebar.success("Saved.")
                else:
                    st.sidebar.error("Quick Save failed - check errors above.")
                st.rerun()
        else:
            st.sidebar.info("No file linked yet.")
            st.sidebar.button(
                "Quick Save",
                key="sidebar_save_disabled_no_archive",
                width="stretch",
                disabled=True,
                icon=mi("save") if "mi" in globals() else None,
            )



        # ==================================================
        # ACTIVE COMPONENT SELECTOR
        # ==================================================
        st.sidebar.subheader("Project List")

        st.sidebar.radio(
            "Active Component:",
            options=proj_labels,
            index=current_index,
            key="project_selector",
            on_change=cb_switch_project,
            label_visibility="collapsed"
        )

        curr_id, curr_proj = get_current_project()

        # ==================================================
        # SIDEBAR COMPONENT ACTION MODE
        # ==================================================
        if "sidebar_component_mode" not in st.session_state:
            st.session_state.sidebar_component_mode = None

        project_type_options = list(PROJECT_DATABASE.keys())


        # ==================================================
        # DEFAULT ACTION BUTTONS
        # ==================================================
        if st.session_state.sidebar_component_mode is None:
            c1, c2 = st.sidebar.columns(2)

            with c1:
                if st.button(
                    "Add",
                    key="sidebar_component_add_start",
                    type="primary",
                    width="stretch",
                    icon=icon_safe("add")
                ):
                    st.session_state.sidebar_component_mode = "add"
                    st.rerun()

            with c2:
                if st.button(
                    "Edit",
                    key="sidebar_component_edit_start",
                    width="stretch",
                    icon=icon_safe("edit")
                ):
                    st.session_state.sidebar_component_mode = "edit"
                    st.rerun()

            if st.sidebar.button(
                "Delete",
                key="sidebar_component_delete_start",
                width="stretch",
                icon=icon_safe("delete")
            ):
                st.session_state.sidebar_component_mode = "delete"
                st.rerun()


        # ==================================================
        # ADD COMPONENT MODE
        # ==================================================
        elif st.session_state.sidebar_component_mode == "add":
            with st.sidebar.form("sidebar_add_component_form", clear_on_submit=False):
                st.markdown("**Add Component**")

                new_component_name = st.text_input(
                    "Component Name",
                    placeholder="e.g. Apartment Tower A"
                )

                new_component_type = st.selectbox(
                    "Component Type",
                    options=project_type_options,
                    index=project_type_options.index("Hotel") if "Hotel" in project_type_options else 0
                )

                c_create, c_cancel = st.columns(2)

                with c_create:
                    create_clicked = st.form_submit_button(
                        "Create",
                        type="primary",
                        width="stretch"
                    )

                with c_cancel:
                    cancel_clicked = st.form_submit_button(
                        "Cancel",
                        width="stretch"
                    )

                if create_clicked:
                    clean_name = new_component_name.strip()

                    if clean_name == "":
                        st.warning("Component name cannot be empty.")

                    elif is_duplicate_project_name(clean_name):
                        st.error("Project name already exists.")

                    else:
                        st.session_state.proj_counter += 1
                        new_id = f"proj_{st.session_state.proj_counter}"

                        st.session_state.projects[new_id] = {
                            "name": clean_name,
                            "type": new_component_type,
                            "data": {}
                        }

                        st.session_state.current_proj_id = new_id
                        st.session_state.sidebar_component_mode = None

                        if save_after_user_action(
                            success_message="Component added and saved to cloud.",
                            fail_message="Component added locally, but cloud save failed. Do not log out yet."
                        ):
                            st.rerun()

                if cancel_clicked:
                    st.session_state.sidebar_component_mode = None
                    st.rerun()


        # ==================================================
        # EDIT COMPONENT MODE
        # ==================================================
        elif st.session_state.sidebar_component_mode == "edit":
            curr_id, curr_proj = get_current_project()

            current_name = curr_proj.get("name", "")
            current_type = curr_proj.get("type", "Hotel")

            if current_type not in project_type_options:
                current_type = "Hotel"

            with st.sidebar.form("sidebar_edit_component_form", clear_on_submit=False):
                st.markdown("**Edit Component**")

                edited_component_name = st.text_input(
                    "Component Name",
                    value=current_name
                )

                edited_component_type = st.selectbox(
                    "Component Type",
                    options=project_type_options,
                    index=project_type_options.index(current_type)
                )

                c_save, c_cancel = st.columns(2)

                with c_save:
                    save_clicked = st.form_submit_button(
                        "Save",
                        type="primary",
                        width="stretch"
                    )

                with c_cancel:
                    cancel_clicked = st.form_submit_button(
                        "Cancel",
                        width="stretch"
                    )

                if cancel_clicked:
                    st.session_state.sidebar_component_mode = None
                    st.rerun()
                    
                if save_clicked:
                    clean_name = edited_component_name.strip()

                    if clean_name == "":
                        st.warning("Component name cannot be empty.")

                    elif is_duplicate_project_name(clean_name, current_id=curr_id):
                        st.error("Component name already exists. Use a unique name.")

                    else:
                        old_type = st.session_state.projects[curr_id].get("type", "Hotel")

                        st.session_state.projects[curr_id]["name"] = clean_name

                        if edited_component_type != old_type:
                            st.session_state.projects[curr_id]["type"] = edited_component_type

                            # Clear old cost/input values so new database defaults can take effect
                            st.session_state.projects[curr_id]["data"] = {}

                            # Clear possible widget/session cache tied to old type
                            keys_to_clear = [
                                k for k in st.session_state.keys()
                                if str(curr_id) in str(k)
                                or "temp_spec_" in str(k)
                                or "u_fl_" in str(k)
                            ]

                            for k in keys_to_clear:
                                if k not in ["projects", "current_proj_id", "proj_counter"]:
                                    del st.session_state[k]
                        else:
                            st.session_state.projects[curr_id]["type"] = edited_component_type

                        st.session_state.sidebar_component_mode = None

                        if save_after_user_action(
                            success_message="Component saved to cloud.",
                            fail_message="Component changed locally, but cloud save failed. Do not log out yet."
                        ):
                            st.rerun()

        # ==================================================
        # DELETE COMPONENT MULTI-SELECT MODE
        # ==================================================
        elif st.session_state.sidebar_component_mode == "delete":
            projects = st.session_state.get("projects", {})

            st.sidebar.warning("Delete selected components?")
            st.sidebar.caption("Select one or more components below, then confirm delete.")

            if not isinstance(projects, dict) or len(projects) <= 1:
                st.sidebar.warning("At least one component must remain.")

                if st.sidebar.button(
                    "Cancel",
                    key="sidebar_component_delete_cancel_single",
                    width="stretch"
                ):
                    st.session_state.sidebar_component_mode = None
                    st.rerun()

            else:
                selected_delete_ids = []

                st.sidebar.markdown("**Select components to delete:**")

                for pid, proj in projects.items():
                    proj_name = proj.get("name", pid)
                    proj_type = proj.get("type", "")

                    checked = st.sidebar.checkbox(
                        f"{proj_name} ({proj_type})",
                        key=f"delete_component_check_{pid}"
                    )

                    if checked:
                        selected_delete_ids.append(pid)

                remaining_count = len(projects) - len(selected_delete_ids)

                if selected_delete_ids:
                    st.sidebar.caption(
                        f"{len(selected_delete_ids)} selected. "
                        f"{remaining_count} component(s) will remain."
                    )

                if remaining_count < 1:
                    st.sidebar.error("You cannot delete all components. At least one must remain.")

                c_confirm, c_cancel = st.sidebar.columns(2)

                with c_confirm:
                    if st.button(
                        "Confirm",
                        key="sidebar_component_delete_confirm",
                        type="primary",
                        width="stretch",
                        disabled=(len(selected_delete_ids) == 0 or remaining_count < 1)
                    ):
                        for pid in selected_delete_ids:
                            if pid in projects:
                                del projects[pid]

                        clear_project_ui_cache_for_ids(selected_delete_ids)

                        # If active project was deleted, switch to first remaining project
                        if st.session_state.get("current_proj_id") not in projects:
                            st.session_state.current_proj_id = next(iter(projects.keys()))

                        # Clear delete checkbox states
                        for k in list(st.session_state.keys()):
                            if str(k).startswith("delete_component_check_"):
                                del st.session_state[k]

                        repair_projects_state(save=True)
                        st.session_state.sidebar_component_mode = None

                        if save_after_user_action(
                            success_message="Component deleted and saved to cloud.",
                            fail_message="Component deleted locally, but cloud save failed. Do not log out yet."
                        ):
                            st.rerun()

                with c_cancel:
                    if st.button(
                        "Cancel",
                        key="sidebar_component_delete_cancel",
                        width="stretch"
                    ):
                        # Clear delete checkbox states
                        for k in list(st.session_state.keys()):
                            if str(k).startswith("delete_component_check_"):
                                del st.session_state[k]

                        st.session_state.sidebar_component_mode = None
                        st.rerun()

        st.sidebar.markdown("---")

    if st.sidebar.button("Logout", type="primary", icon=icon_safe("logout")):
        st.session_state.logged_in = False
        st.session_state.access_token = None
        st.session_state.user = None

        for key_to_clear in [
            "projects",
            "storage_loaded",
            "report_config",
            "port_meta",
            "port_assumptions"
        ]:
            if key_to_clear in st.session_state:
                del st.session_state[key_to_clear]

        st.rerun()
        
    st.sidebar.caption(f"v{APP_VERSION} | (c) 2026 QS & Procurement - ASG")
    #endregion

    #debugcode
    if st.session_state.get("current_page") == "App Debugger":
        if st.sidebar.button(
            "Back to App",
            key="close_app_debugger",
            width="stretch",
            icon=mi("arrow_back") if "mi" in globals() else None
        ):
            st.session_state.current_page = None
            st.rerun()
    elif st.sidebar.button(
        "Debug",
        key="open_app_debugger",
        width="stretch",
        icon=mi("bug_report") if "mi" in globals() else None
    ):
        st.session_state.current_page = "App Debugger"
        st.rerun()

    if st.session_state.get("current_page") == "App Debugger":
        render_app_debugger()
        return
    #enddebugcode


    if page_choice == "Area Analysis":
        show_area_calculator()
    elif page_choice == "Database":
        show_project_database()
    elif page_choice == "Summary":
        show_portfolio_summary()
    elif page_choice == "Archive":
        show_snapshots()
    elif page_choice == "Cost Analysis":
        show_cost_estimator()
    else:
        render_feasibility_study_landing()

# 4. THE GATEKEEPER LOGIC
# Replace your entire gatekeeper section at the bottom with this:

if not st.session_state.logged_in:
    login_screen()
else:
    # Re-apply token on EVERY script run, not just after login
    token = st.session_state.get("access_token")
    if token:
        supabase.postgrest.auth(token)
    main_app()
#endregion

#region --- THANK YOU ---
# Version: 1.1.0
# Environment: Streamlit 1.56.0, Python 3.13
# for the future me or any IT person that might look at this code,
# this code is made in 2026, by a totally newbie programmer wannabe, 
# but with over 8 years of work experience and Architecture Bachelor 
# (architure as in construction, not that architecture)
# if someday this might not work, know that I (Boris Prilyan Sidabutar, B. Arch) 
# make this alone (many thanks especially to Jesus and for Gemini, ChatGPT and Claude too)
#endregion
