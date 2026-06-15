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
