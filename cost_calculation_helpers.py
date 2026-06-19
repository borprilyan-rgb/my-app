def _safe_float(value, default=0.0):
    if value is None:
        return default

    try:
        if value != value:
            return default
    except Exception:
        pass

    if isinstance(value, str):
        value = (
            value.strip()
            .replace("Rp", "")
            .replace("rp", "")
            .replace(",", "")
            .replace("%", "")
        )

        if value == "" or value.lower() in ["none", "nan", "null", "-"]:
            return default

    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def build_cost_raw_from_project_data(project_data, project_type_data):
    """Map saved project data keys into calculate_live_costs(raw) inputs."""
    data = project_data if isinstance(project_data, dict) else {}
    type_data = project_type_data if isinstance(project_type_data, dict) else {}

    def get_val(key, default_db_key=None, default_val=0.0):
        value = data.get(key)
        if value is not None and value != "":
            return _safe_float(value)
        if default_db_key and default_db_key in type_data:
            return _safe_float(type_data.get(default_db_key))
        return _safe_float(default_val)

    smart_custom_costs = sum(
        _safe_float(item.get("Rate (Rp)", 0)) * _safe_float(item.get("Quantity", 1))
        for item in data.get("smart_custom_costs", [])
        if isinstance(item, dict)
    )

    return {
        "gba": get_val("m_gba"),
        "gfa": get_val("m_gfa"),
        "struc_earth": get_val("u_earth", "struc_earth"),
        "struc_found": get_val("u_found", "struc_found"),
        "struc_work": get_val("u_struc", "struc_work"),
        "arch_base": get_val("u_arch", "arch_base"),
        "facade": get_val("m_facade"),
        "facade_precast_pct": get_val("r_fac_pre", "facade_precast_pct"),
        "fac_precast_rate": get_val("u_f_pre", "facade_precast_rate"),
        "facade_window_pct": get_val("r_fac_win", "facade_window_pct"),
        "fac_window_rate": get_val("u_f_win", "facade_window_rate"),
        "facade_double_pct": get_val("r_fac_doub", "facade_double_pct"),
        "fac_double_rate": get_val("u_f_doub", "facade_double_rate"),
        "wooden_door": get_val("m_door_w"),
        "door_wood": get_val("u_d_wood", "door_wood"),
        "glass_door": get_val("m_door_g"),
        "door_glass": get_val("u_d_glass", "door_glass"),
        "steel_door": get_val("m_door_s"),
        "door_steel": get_val("u_d_steel", "door_steel"),
        "lobby_interior": get_val("m_lobby"),
        "lobby_rate": get_val("u_lobby", "lobby"),
        "gondola_unit": get_val("m_gondola"),
        "gondola_rate": get_val("u_gondola", "gondola"),
        "rooms": get_val("m_rooms"),
        "san_qty_room": get_val("r_san_qty", "san_room_qty"),
        "san_room_rate": get_val("u_s_room", "san_room_rate"),
        "toilet_male": get_val("m_toil_m"),
        "san_pub_m": get_val("u_s_pub_m", "san_pub_m"),
        "toilet_female": get_val("m_toil_f"),
        "san_pub_f": get_val("u_s_pub_f", "san_pub_f"),
        "disabled_toil": get_val("m_toil_d"),
        "san_dis": get_val("u_s_dis", "san_dis"),
        "mushola_unit": get_val("m_mushola"),
        "san_mushola": get_val("u_s_mushola", "san_mushola"),
        "kitchen_rate": get_val("u_kit", "kitchen"),
        "hw_wood": get_val("u_hw_wood", "hw_wood"),
        "hw_steel": get_val("u_hw_steel", "hw_steel"),
        "fl_waste": get_val("w_floor", "fl_waste", 10),
        "fl_skirt": get_val("s_floor", "fl_skirt", 20),
        "fl_ht_pct": get_val("r_fl_ht", "fl_ht_pct"),
        "fl_ht_rate": get_val("u_fl_ht"),
        "fl_vinyl_pct": get_val("r_fl_vin", "fl_vinyl_pct"),
        "fl_vinyl_rate": get_val("u_fl_vin"),
        "fl_marmer_pct": get_val("r_fl_mar", "fl_marmer_pct"),
        "fl_marmer_rate": get_val("u_fl_mar"),
        "carpet_m2": get_val("m_carpet"),
        "carpet_rate": get_val("u_carpet", "carpet"),
        "glass_m2": get_val("m_glass"),
        "glass_rate": get_val("u_glass", "glass"),
        "ffe_rate": get_val("u_ffe", "ffe"),
        "misc_rate": get_val("u_misc", "misc"),
        "misc_switch": get_val("misc_switch"),
        "mep_rate": get_val("u_mep", "mep"),
        "utility_rate": get_val("u_util", "utility"),
        "railing_qty": get_val("r_rail_qty", "railing_qty"),
        "railing_rate": get_val("u_rail", "railing_rate"),
        "skylight_area": get_val("m_skylight"),
        "skylight_rate": get_val("u_sky", "skylight_rate"),
        "land_m2": get_val("m_land_m2"),
        "ext_land_rate": get_val("u_ext", "ext_land"),
        "pub_fac_m2": get_val("m_fac_pub"),
        "fac_pub_rate": get_val("u_fac_p", "fac_pub"),
        "res_fac_m2": get_val("m_fac_res"),
        "fac_res_rate": get_val("u_fac_r", "fac_res"),
        "proj_fac_u": get_val("m_fac_proj"),
        "fac_proj_rate": get_val("u_fac_pr", "fac_proj"),
        "consultancy_rate": get_val("sc_cons", "cons"),
        "qs_months": get_val("sc_qs_m"),
        "qs_rate": get_val("sc_qs_r", "qs_rate"),
        "pm_months": get_val("sc_pm_m"),
        "pm_rate": get_val("sc_pm_r", "pm_rate"),
        "insurance_pct": get_val("sc_ins", "insurance_pct", 0.12),
        "prelim_pct": get_val("sc_prelim_pct", "prelim_pct", 5.0),
        "contingency_pct": get_val("sc_contingency_pct", "contingency_pct", 3.0),
        "smart_custom_costs": smart_custom_costs,
    }


def calculate_live_costs(raw):
    """Pure Cost Analysis calculations for the live project page."""
    def n(key, default=0.0):
        return _safe_float(raw.get(key, default), default)

    costs = {}

    costs["t_earth"] = n("gba") * n("struc_earth")
    costs["t_found"] = n("gba") * n("struc_found")
    costs["t_struc"] = n("gba") * n("struc_work")
    costs["t_arch_base"] = n("gfa") * n("arch_base")
    costs["t_precast"] = n("facade") * (n("facade_precast_pct") / 100) * n("fac_precast_rate")
    costs["t_window"] = n("facade") * (n("facade_window_pct") / 100) * n("fac_window_rate")
    costs["t_double"] = n("facade") * (n("facade_double_pct") / 100) * n("fac_double_rate")
    costs["t_w_door"] = n("wooden_door") * n("door_wood")
    costs["t_g_door"] = n("glass_door") * n("door_glass")
    costs["t_s_door"] = n("steel_door") * n("door_steel")
    costs["t_lobby"] = n("lobby_interior") * n("lobby_rate")
    costs["t_gondola"] = n("gondola_unit") * n("gondola_rate")
    costs["t_unit_san"] = n("rooms") * n("san_qty_room") * n("san_room_rate")
    costs["t_t_male"] = n("toilet_male") * n("san_pub_m")
    costs["t_t_female"] = n("toilet_female") * n("san_pub_f")
    costs["t_t_dis"] = n("disabled_toil") * n("san_dis")
    costs["t_mushola"] = n("mushola_unit") * n("san_mushola")
    costs["t_kitchen"] = n("rooms") * n("kitchen_rate")
    costs["t_hw_w"] = n("wooden_door") * n("hw_wood")
    costs["t_hw_s"] = n("steel_door") * n("hw_steel")
    costs["f_mult"] = (1 + (n("fl_waste") / 100)) * (1 + (n("fl_skirt") / 100))
    costs["t_ht"] = n("gfa") * (n("fl_ht_pct") / 100) * n("fl_ht_rate") * costs["f_mult"]
    costs["t_vinyl"] = n("gfa") * (n("fl_vinyl_pct") / 100) * n("fl_vinyl_rate") * costs["f_mult"]
    costs["t_marmer"] = n("gfa") * (n("fl_marmer_pct") / 100) * n("fl_marmer_rate") * costs["f_mult"]
    costs["t_carpet"] = n("carpet_m2") * n("carpet_rate")
    costs["t_glass_work"] = n("glass_m2") * n("glass_rate")
    costs["t_ffe"] = n("rooms") * n("ffe_rate")
    costs["t_misc"] = n("misc_rate") * n("misc_switch")
    costs["t_mep"] = n("gba") * n("mep_rate")
    costs["t_utility"] = n("gba") * n("utility_rate")
    costs["t_railing"] = (n("rooms") * n("railing_qty")) * n("railing_rate")
    costs["t_skylight"] = n("skylight_area") * n("skylight_rate")
    costs["t_external"] = n("land_m2") * n("ext_land_rate")
    costs["t_pub_fac"] = n("pub_fac_m2") * n("fac_pub_rate")
    costs["t_res_fac"] = n("res_fac_m2") * n("fac_res_rate")
    costs["t_proj_fac"] = n("proj_fac_u") * n("fac_proj_rate")
    costs["group_misc"] = costs["t_pub_fac"] + costs["t_res_fac"] + costs["t_proj_fac"]

    costs["construction_subtotal"] = sum([
        costs["t_earth"], costs["t_found"], costs["t_struc"], costs["t_arch_base"],
        costs["t_precast"], costs["t_window"], costs["t_double"], costs["t_w_door"],
        costs["t_g_door"], costs["t_s_door"], costs["t_lobby"], costs["t_gondola"],
        costs["t_unit_san"], costs["t_t_male"], costs["t_t_female"], costs["t_t_dis"],
        costs["t_mushola"], costs["t_kitchen"], costs["t_hw_w"], costs["t_hw_s"],
        costs["t_ht"], costs["t_vinyl"], costs["t_marmer"], costs["t_carpet"],
        costs["t_glass_work"], costs["t_ffe"], costs["t_misc"], costs["t_mep"],
        costs["t_utility"], costs["t_railing"], costs["t_skylight"], costs["t_external"],
        costs["group_misc"], n("smart_custom_costs"),
    ])

    costs["prelim_pct"] = n("prelim_pct", 5.0)
    costs["contingency_pct"] = n("contingency_pct", 3.0)

    costs["t_preliminary"] = costs["construction_subtotal"] * (costs["prelim_pct"] / 100.0)
    costs["t_contingency"] = (
        costs["construction_subtotal"] + costs["t_preliminary"]
    ) * (costs["contingency_pct"] / 100.0)
    costs["grand_total_hc"] = (
        costs["construction_subtotal"] + costs["t_preliminary"] + costs["t_contingency"]
    )

    costs["t_consultancy"] = n("gfa") * n("consultancy_rate")
    costs["t_qs"] = n("qs_months") * n("qs_rate")
    costs["t_pm"] = n("pm_months") * n("pm_rate")
    costs["t_insurance"] = costs["grand_total_hc"] * (n("insurance_pct") / 100.0)

    costs["total_soft_cost"] = (
        costs["t_consultancy"] + costs["t_qs"] + costs["t_pm"] + costs["t_insurance"]
    )
    costs["grand_total_project"] = costs["grand_total_hc"] + costs["total_soft_cost"]

    costs["group_earth"] = costs["t_earth"]
    costs["group_found"] = costs["t_found"]
    costs["group_struc"] = costs["t_struc"]
    costs["group_facade"] = costs["t_precast"] + costs["t_window"] + costs["t_double"]
    costs["group_sanitary"] = (
        costs["t_unit_san"] + costs["t_t_male"] + costs["t_t_female"]
        + costs["t_t_dis"] + costs["t_mushola"]
    )
    costs["group_floor"] = costs["t_ht"] + costs["t_vinyl"] + costs["t_marmer"]
    costs["group_door"] = (
        costs["t_w_door"] + costs["t_g_door"] + costs["t_s_door"]
        + costs["t_hw_w"] + costs["t_hw_s"]
    )
    costs["group_arch"] = (
        costs["t_arch_base"] + costs["t_lobby"] + costs["t_carpet"] + costs["t_gondola"]
        + costs["t_glass_work"] + costs["t_kitchen"] + costs["t_railing"]
        + costs["t_skylight"] + costs["group_facade"] + costs["group_sanitary"]
        + costs["group_floor"] + costs["group_door"] + n("smart_custom_costs")
    )

    costs["group_ffe"] = costs["t_ffe"] + costs["t_misc"]
    costs["group_mep"] = costs["t_mep"]
    costs["group_utility"] = costs["t_utility"]
    costs["group_ext"] = costs["t_external"]
    costs["group_prelim"] = costs["t_preliminary"]
    costs["group_conting"] = costs["t_contingency"]
    costs["group_soft_cost"] = costs["total_soft_cost"]
    costs["group_hard_cost"] = costs["grand_total_hc"]
    costs["group_total"] = costs["total_soft_cost"] + costs["grand_total_hc"]

    return costs
