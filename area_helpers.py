import pandas as pd


def safe_float(v, default=0.0):
    if v is None:
        return default

    try:
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


_safe_float = safe_float


def calculate_area_totals_from_table(area_table_data):
    unit_area_col = "Unit Area"
    breakdown_cols = [
        "Parkir", "Roof/Deck", "MEP Outdoor",
        "Koridor/Lobby", "Stair, MEP, Etc",
        unit_area_col, "Office"
    ]

    if not isinstance(area_table_data, list) or len(area_table_data) == 0:
        return {
            "gba": 0.0,
            "gfa": 0.0,
            "sgfa": 0.0,
            "nfa": 0.0
        }

    df = pd.DataFrame(area_table_data)

    if unit_area_col not in df.columns and "Unit" in df.columns:
        df[unit_area_col] = df["Unit"]

    for col in breakdown_cols:
        if col not in df.columns:
            df[col] = 0.0

        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    total = df[breakdown_cols].sum(axis=1)

    gba = _safe_float(total.sum())
    gfa = _safe_float((total - df[["Parkir", "Roof/Deck", "MEP Outdoor"]].sum(axis=1)).sum())
    sgfa = _safe_float(df[[unit_area_col, "Office", "Koridor/Lobby"]].sum(axis=1).sum())
    nfa = _safe_float(df[[unit_area_col, "Office"]].sum(axis=1).sum())

    return {
        "gba": gba,
        "gfa": gfa,
        "sgfa": sgfa,
        "nfa": nfa
    }


def generate_area_rows(upper_floors, basements):
    rows = [
        {
            "FL": "Roof Machine",
            "Space Type": "Roof",
            "Floor to Floor Height (m)": 0.0,
            "Typical Unit": 0,
            "Parkir": 0.0,
            "Roof/Deck": 0.0,
            "MEP Outdoor": 0.0,
            "Koridor/Lobby": 0.0,
            "Stair, MEP, Etc": 0.0,
            "Unit Area": 0.0,
            "Office": 0.0,
        },
        {
            "FL": "Roof",
            "Space Type": "Roof",
            "Floor to Floor Height (m)": 0.0,
            "Typical Unit": 0,
            "Parkir": 0.0,
            "Roof/Deck": 0.0,
            "MEP Outdoor": 0.0,
            "Koridor/Lobby": 0.0,
            "Stair, MEP, Etc": 0.0,
            "Unit Area": 0.0,
            "Office": 0.0,
        },
    ]

    for i in range(int(upper_floors), 1, -1):
        rows.append(
            {
                "FL": f"{i}F",
                "Space Type": "Unit",
                "Floor to Floor Height (m)": 0.0,
                "Typical Unit": 0,
                "Parkir": 0.0,
                "Roof/Deck": 0.0,
                "MEP Outdoor": 0.0,
                "Koridor/Lobby": 0.0,
                "Stair, MEP, Etc": 0.0,
                "Unit Area": 0.0,
                "Office": 0.0,
            }
        )

    rows.append(
        {
            "FL": "1F",
            "Space Type": "Lobby",
            "Floor to Floor Height (m)": 0.0,
            "Typical Unit": 0,
            "Parkir": 0.0,
            "Roof/Deck": 0.0,
            "MEP Outdoor": 0.0,
            "Koridor/Lobby": 0.0,
            "Stair, MEP, Etc": 0.0,
            "Unit Area": 0.0,
            "Office": 0.0,
        }
    )

    for i in range(1, int(basements) + 1):
        fl_name = "LG" if i == 1 else f"B{i - 1}"
        rows.append(
            {
                "FL": fl_name,
                "Space Type": "Carpark",
                "Floor to Floor Height (m)": 0.0,
                "Typical Unit": 0,
                "Parkir": 0.0,
                "Roof/Deck": 0.0,
                "MEP Outdoor": 0.0,
                "Koridor/Lobby": 0.0,
                "Stair, MEP, Etc": 0.0,
                "Unit Area": 0.0,
                "Office": 0.0,
            }
        )

    return rows


def guess_area_f2f_height(row):
    fl = str(row.get("FL", "")).strip()
    space_type = str(row.get("Space Type", "")).strip()

    if fl == "1F" or space_type == "Lobby":
        return 4.5
    if fl == "LG" or fl.startswith("B") or space_type == "Carpark":
        return 3.2
    if fl in ["Roof", "Roof Machine"] or space_type == "Roof":
        return 3.0
    if space_type == "Office":
        return 3.6
    if space_type == "Facility":
        return 4.0

    return 3.2


def normalize_none_records(records):
    if not isinstance(records, list):
        return []

    safe_records = []

    for row in records:
        if not isinstance(row, dict):
            continue

        safe_row = {}

        for k, v in row.items():
            safe_row[k] = "" if v is None else v

        safe_records.append(safe_row)

    return safe_records


def clean_area_records(records):
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

    required_cols = ["FL", "Space Type", F2F_COL, TYPICAL_UNIT_COL] + breakdown_cols

    if not isinstance(records, list) or len(records) == 0:
        records = generate_area_rows(5, 1)

    records = normalize_none_records(records)

    df = pd.DataFrame(records)

    # Old migration: "Unit" used to mean area.
    if UNIT_AREA_COL not in df.columns and "Unit" in df.columns:
        df[UNIT_AREA_COL] = df["Unit"]

    for col in required_cols:
        if col not in df.columns:
            if col in ["FL", "Space Type"]:
                df[col] = ""
            elif col == F2F_COL:
                df[col] = df.apply(lambda r: guess_area_f2f_height(r), axis=1)
            else:
                df[col] = 0.0

    df["FL"] = df["FL"].astype(str).str.strip()
    df["Space Type"] = df["Space Type"].astype(str).str.strip()

    df[F2F_COL] = pd.to_numeric(df[F2F_COL], errors="coerce").fillna(0.0)

    df[TYPICAL_UNIT_COL] = (
        pd.to_numeric(df[TYPICAL_UNIT_COL], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    for col in breakdown_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return df[required_cols].to_dict("records")


def calculate_area_dataframe(records):
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

    cleaned_records = clean_area_records(records)
    df = pd.DataFrame(cleaned_records)

    for col in breakdown_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df[TYPICAL_UNIT_COL] = (
        pd.to_numeric(df[TYPICAL_UNIT_COL], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    df[F2F_COL] = pd.to_numeric(df[F2F_COL], errors="coerce").fillna(0.0)

    df["TOTAL"] = df[breakdown_cols].sum(axis=1)
    df["GBA"] = df["TOTAL"]
    df["GFA"] = df["TOTAL"] - df[["Parkir", "Roof/Deck", "MEP Outdoor"]].sum(axis=1)
    df["SGFA"] = df[[UNIT_AREA_COL, "Office", "Koridor/Lobby"]].sum(axis=1)
    df["NFA"] = df[[UNIT_AREA_COL, "Office"]].sum(axis=1)

    return df


def clean_door_records(records):
    F2F_COL = "Floor to Floor Height (m)"
    TYPICAL_UNIT_COL = "Typical Unit"

    DOOR_WOOD_COL = "Pintu Kayu"
    DOOR_STEEL_COL = "Pintu Besi"
    DOOR_GLASS_COL = "Pintu Kaca"

    cols = [
        "FL",
        "Space Type",
        F2F_COL,
        TYPICAL_UNIT_COL,
        DOOR_WOOD_COL,
        DOOR_STEEL_COL,
        DOOR_GLASS_COL,
    ]

    if not isinstance(records, list):
        records = []

    records = normalize_none_records(records)

    df = pd.DataFrame(records)

    for col in cols:
        if col not in df.columns:
            if col in ["FL", "Space Type"]:
                df[col] = ""
            else:
                df[col] = 0

    df["FL"] = df["FL"].astype(str).str.strip()
    df["Space Type"] = df["Space Type"].astype(str).str.strip()

    df[F2F_COL] = pd.to_numeric(df[F2F_COL], errors="coerce").fillna(0.0)

    for col in [TYPICAL_UNIT_COL, DOOR_WOOD_COL, DOOR_STEEL_COL, DOOR_GLASS_COL]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df[cols].to_dict("records")


def area_records_to_input_view(records, input_mode):
    """
    Converts stored detailed category records into either:
    - Detailed Category Input
    - Consultant Summary Input
    """
    clean_records = clean_area_records(records)
    df = pd.DataFrame(clean_records)

    for col in [
        "Parkir",
        "Roof/Deck",
        "MEP Outdoor",
        "Koridor/Lobby",
        "Stair, MEP, Etc",
        "Unit Area",
        "Office",
    ]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    if input_mode == "Consultant Summary":
        # Cumulative logic:
        # NFA = Unit + Office
        # SGFA = NFA + Koridor/Lobby
        # GFA = SGFA + Stair, MEP, Etc
        df["NFA Input"] = df["Unit Area"] + df["Office"]
        df["SGFA Input"] = df["NFA Input"] + df["Koridor/Lobby"]
        df["GFA Input"] = df["SGFA Input"] + df["Stair, MEP, Etc"]

        view_cols = [
            "FL",
            "Space Type",
            "Floor to Floor Height (m)",
            "Typical Unit",
            "Parkir",
            "Roof/Deck",
            "MEP Outdoor",
            "GFA Input",
            "SGFA Input",
            "NFA Input",
        ]

        return df[view_cols].copy()

    view_cols = [
        "FL",
        "Space Type",
        "Floor to Floor Height (m)",
        "Typical Unit",
        "Parkir",
        "Roof/Deck",
        "MEP Outdoor",
        "Koridor/Lobby",
        "Stair, MEP, Etc",
        "Unit Area",
        "Office",
    ]

    return df[view_cols].copy()


def input_view_to_area_records(view_df, input_mode):
    """
    Converts editor output back into stored detailed category records.
    Backend remains detailed even when user edits consultant summary mode.
    """
    F2F_COL = "Floor to Floor Height (m)"
    TYPICAL_UNIT_COL = "Typical Unit"
    UNIT_AREA_COL = "Unit Area"

    df = view_df.copy()

    base_cols = ["FL", "Space Type", F2F_COL, TYPICAL_UNIT_COL]

    for col in base_cols:
        if col not in df.columns:
            df[col] = "" if col in ["FL", "Space Type"] else 0

    df["FL"] = df["FL"].astype(str).str.strip()
    df["Space Type"] = df["Space Type"].astype(str).str.strip()
    df[F2F_COL] = pd.to_numeric(df[F2F_COL], errors="coerce").fillna(0.0)
    df[TYPICAL_UNIT_COL] = (
        pd.to_numeric(df[TYPICAL_UNIT_COL], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    for col in ["Parkir", "Roof/Deck", "MEP Outdoor"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    if input_mode == "Consultant Summary":
        for col in ["GFA Input", "SGFA Input", "NFA Input"]:
            if col not in df.columns:
                df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        # Convert cumulative consultant values back to detailed categories.
        df[UNIT_AREA_COL] = df["NFA Input"]
        df["Office"] = 0.0

        df["Koridor/Lobby"] = df["SGFA Input"] - df["NFA Input"]
        df["Stair, MEP, Etc"] = df["GFA Input"] - df["SGFA Input"]

        # Prevent negative areas from breaking calculation.
        df["Koridor/Lobby"] = df["Koridor/Lobby"].clip(lower=0.0)
        df["Stair, MEP, Etc"] = df["Stair, MEP, Etc"].clip(lower=0.0)

    else:
        for col in ["Koridor/Lobby", "Stair, MEP, Etc", UNIT_AREA_COL, "Office"]:
            if col not in df.columns:
                df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    stored_cols = [
        "FL",
        "Space Type",
        F2F_COL,
        TYPICAL_UNIT_COL,
        "Parkir",
        "Roof/Deck",
        "MEP Outdoor",
        "Koridor/Lobby",
        "Stair, MEP, Etc",
        UNIT_AREA_COL,
        "Office",
    ]

    return clean_area_records(df[stored_cols].to_dict("records"))


def validate_consultant_summary_input(view_df):
    """
    Warns if consultant cumulative inputs are logically inconsistent:
    GFA should be >= SGFA should be >= NFA.
    """
    warnings = []

    required = ["FL", "GFA Input", "SGFA Input", "NFA Input"]

    for col in required:
        if col not in view_df.columns:
            return warnings

    check_df = view_df.copy()

    for col in ["GFA Input", "SGFA Input", "NFA Input"]:
        check_df[col] = pd.to_numeric(check_df[col], errors="coerce").fillna(0.0)

    bad_sgfa = check_df[check_df["SGFA Input"] < check_df["NFA Input"]]
    bad_gfa = check_df[check_df["GFA Input"] < check_df["SGFA Input"]]

    if len(bad_sgfa) > 0:
        floors = ", ".join(bad_sgfa["FL"].astype(str).tolist())
        warnings.append(f"SGFA is smaller than NFA on: {floors}")

    if len(bad_gfa) > 0:
        floors = ", ".join(bad_gfa["FL"].astype(str).tolist())
        warnings.append(f"GFA is smaller than SGFA on: {floors}")

    return warnings
