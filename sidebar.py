# =========================================================
# sidebar.py — Streamlit Sidebar Form
# =========================================================

import streamlit as st
import pandas as pd


# =========================================================
# INVERTER COLUMN MAP (รองรับชื่อ column หลายแบบ)
# =========================================================
_INV_MAP = {
    "power_ac":    ["Power_kW", "Power_W", "AC Power (W)", "Power"],
    "v_dc_max":    ["Max_DC_Voltage_V", "V_max_V", "Max DC Voltage (V)"],
    "i_mppt":      ["Max_DC_Current_A", "I_max_A", "Max Input Current (A)"],
    "pv_max":      ["Max_PV_Power_W", "Max_PV_Power", "Max PV Power (W)"],
    "mppt_min":    ["MPPT_min_V", "MPPT Min (V)", "Vmppt_min"],
    "mppt_max":    ["MPPT_max_V", "MPPT Max (V)", "Vmppt_max"],
    "mppt_count":  ["MPPT_Count", "MPPT", "MPPT_count", "Num MPPT", "Number of MPPT"],
    "model":       ["Model", "model", "Model Name"],
    "brand":       ["Brand", "brand"],
}

_PANEL_MAP = {
    "Pm":  ["Pmax_W", "Power_W", "Power (W)", "Pm (W)", "Pm"],
    "Vmp": ["Vmp_V", "Vmp (V)", "Vmp"],
    "Voc": ["Voc_V", "Voc (V)", "Voc"],
    "Imp": ["Imp_A", "Imp (A)", "Imp"],
    "Isc": ["Isc_A", "Isc (A)", "Isc"],
    "model": ["Model", "model", "Model Name"],
    "brand": ["Brand", "brand"],
}


def _find_col(df, aliases):
    for a in aliases:
        if a in df.columns:
            return a
    return None


def _get_val(row, aliases, default=0.0):
    for a in aliases:
        if a in row.index and pd.notna(row[a]):
            return row[a]
    return default


# =========================================================
# AUTOFILL HELPERS
# =========================================================
def _autofill_inverter(df: pd.DataFrame, label: str):
    """ดึงค่า spec จาก row ที่ user เลือก แล้วเขียนลง session_state"""
    if df.empty or label == "— พิมพ์เอง —":
        return

    model_col = _find_col(df, _INV_MAP["model"])
    brand_col = _find_col(df, _INV_MAP["brand"])

    # หา row ที่ตรงกับ label
    if model_col and brand_col:
        row_mask = (df[brand_col].astype(str) + " " + df[model_col].astype(str)) == label
    elif model_col:
        row_mask = df[model_col].astype(str) == label
    else:
        return

    matches = df[row_mask]
    if matches.empty:
        return
    row = matches.iloc[0]

    # Power_kW vs Power_W
    power_col = _find_col(df, _INV_MAP["power_ac"])
    if power_col:
        pwr = float(_get_val(row, [power_col], 0))
        # ถ้าชื่อ column มี 'kw' (case-insensitive) → คูณ 1000
        if "kw" in power_col.lower():
            pwr *= 1000
        st.session_state["inv_power_ac"] = int(pwr)

    v_col = _find_col(df, _INV_MAP["v_dc_max"])
    if v_col:
        st.session_state["inv_v_dc_max"] = float(_get_val(row, [v_col], 1100))

    i_col = _find_col(df, _INV_MAP["i_mppt"])
    if i_col:
        st.session_state["inv_i_sc_max"] = float(_get_val(row, [i_col], 25.0))

    pv_col = _find_col(df, _INV_MAP["pv_max"])
    if pv_col:
        st.session_state["inv_pv_power_max"] = int(_get_val(row, [pv_col], 13000))

    mn_col = _find_col(df, _INV_MAP["mppt_min"])
    if mn_col:
        st.session_state["v_mppt_min"] = float(_get_val(row, [mn_col], 200))

    mx_col = _find_col(df, _INV_MAP["mppt_max"])
    if mx_col:
        st.session_state["v_mppt_max"] = float(_get_val(row, [mx_col], 850))

    mc_col = _find_col(df, _INV_MAP["mppt_count"])
    if mc_col:
        st.session_state["mppt_count"] = int(_get_val(row, [mc_col], 1))

    # autofill MPPT_Count (new DB column name)
    for mppt_col in ["MPPT_Count", "MPPT_count", "MPPT"]:
        if mppt_col in row.index and pd.notna(row[mppt_col]):
            st.session_state["mppt_count"] = int(row[mppt_col])
            break

    # autofill inverter price
    if "Price_THB" in row.index and pd.notna(row["Price_THB"]):
        st.session_state["inv_price_thb"] = float(row["Price_THB"])

    # autofill inverter efficiency
    for eff_col in ["Efficiency_%", "Efficiency_pct", "Efficiency"]:
        if eff_col in row.index and pd.notna(row[eff_col]):
            st.session_state["inv_efficiency"] = float(row[eff_col])
            break

    # autofill datasheet URL
    if "Datasheet_URL" in row.index and pd.notna(row["Datasheet_URL"]):
        st.session_state["inv_datasheet_url"] = str(row["Datasheet_URL"])


def _autofill_panel(df: pd.DataFrame, label: str):
    """ดึงค่า spec แผงจาก row ที่ user เลือก"""
    if df.empty or label == "— พิมพ์เอง —":
        return

    model_col = _find_col(df, _PANEL_MAP["model"])
    brand_col = _find_col(df, _PANEL_MAP["brand"])

    if model_col and brand_col:
        row_mask = (df[brand_col].astype(str) + " " + df[model_col].astype(str)) == label
    elif model_col:
        row_mask = df[model_col].astype(str) == label
    else:
        return

    matches = df[row_mask]
    if matches.empty:
        return
    row = matches.iloc[0]

    for key, aliases in [("Pm","Pm"),("Vmp","Vmp"),("Voc","Voc"),("Imp","Imp"),("Isc","Isc")]:
        col = _find_col(df, _PANEL_MAP[key])
        if col:
            val = float(_get_val(row, [col], st.session_state.get(key, 0)))
            st.session_state[key] = val
    # autofill panel price
    if "Price_THB" in row.index and pd.notna(row["Price_THB"]):
        st.session_state["panel_price_thb"] = float(row["Price_THB"])

    # autofill panel datasheet URL
    if "Datasheet_URL" in row.index and pd.notna(row["Datasheet_URL"]):
        st.session_state["panel_datasheet_url"] = str(row["Datasheet_URL"])


# =========================================================
# BUILD DROPDOWN LABELS
# =========================================================
def _inv_labels(df: pd.DataFrame) -> list:
    if df.empty:
        return []
    model_col = _find_col(df, _INV_MAP["model"])
    brand_col = _find_col(df, _INV_MAP["brand"])
    if model_col and brand_col:
        return (df[brand_col].astype(str) + " " + df[model_col].astype(str)).tolist()
    elif model_col:
        return df[model_col].astype(str).tolist()
    return []


def _panel_labels(df: pd.DataFrame) -> list:
    if df.empty:
        return []
    model_col = _find_col(df, _PANEL_MAP["model"])
    brand_col = _find_col(df, _PANEL_MAP["brand"])
    if model_col and brand_col:
        return (df[brand_col].astype(str) + " " + df[model_col].astype(str)).tolist()
    elif model_col:
        return df[model_col].astype(str).tolist()
    return []


# =========================================================
# MAIN RENDER
# =========================================================
def render_sidebar():
    panels_db    = st.session_state.get("panels_db",    pd.DataFrame())
    inverters_db = st.session_state.get("inverters_db", pd.DataFrame())

    # -------------------------------------------------------
    # AUTOFILL ทำนอก form (on_change ใน selectbox)
    # ทำก่อน form render เพื่อให้ number_input อ่านค่าใหม่
    # -------------------------------------------------------
    # --- Panel dropdown (นอก form, ใช้ st.sidebar โดยตรง) ---
    st.sidebar.subheader("🔍 เลือกแผงจาก Database")
    panel_opts = ["— พิมพ์เอง —"] + _panel_labels(panels_db)
    prev_panel = st.session_state.get("_selected_panel", panel_opts[0])
    if prev_panel not in panel_opts:
        prev_panel = panel_opts[0]
    selected_panel = st.sidebar.selectbox(
        "แผงโซลาร์", panel_opts,
        index=panel_opts.index(prev_panel), key="_selected_panel",
    )
    if selected_panel != "— พิมพ์เอง —":
        _autofill_panel(panels_db, selected_panel)
        st.sidebar.caption(f"✅ Autofill: {selected_panel}")

    # --- Inverter dropdown (นอก form) ---
    st.sidebar.subheader("🔍 เลือก Inverter จาก Database")
    inv_opts = ["— พิมพ์เอง —"] + _inv_labels(inverters_db)
    prev_inv = st.session_state.get("_selected_inv", inv_opts[0])
    if prev_inv not in inv_opts:
        prev_inv = inv_opts[0]
    selected_inv = st.sidebar.selectbox(
        "Inverter", inv_opts,
        index=inv_opts.index(prev_inv), key="_selected_inv",
    )
    if selected_inv != "— พิมพ์เอง —":
        _autofill_inverter(inverters_db, selected_inv)
        st.sidebar.caption(f"✅ Autofill: {selected_inv}")

    # -------------------------------------------------------
    # FORM (ค่าที่ autofill จะปรากฏใน number_input อัตโนมัติ)
    # -------------------------------------------------------
    with st.sidebar.form("pv_design_form"):

        # ---------- LOAD & RESOURCE ----------
        st.header("ข้อมูลโหลดไฟฟ้า")
        st.number_input("พลังงานไฟฟ้าต่อวัน (kWh/day)", min_value=0.0, value=30.0, step=1.0, key="E_day")
        st.number_input("ชั่วโมงแสงอาทิตย์ (Peak Sun Hours)", min_value=1.0, max_value=7.0, value=4.5, step=0.1, key="H_sun")
        st.slider("Performance Ratio (PR)", 0.6, 0.9, 0.8, 0.01, key="PR")

        # ---------- ROOF AREA ----------
        st.header("พื้นที่ติดตั้ง")
        st.number_input("พื้นที่หลังคาใช้งานได้ (m²)", min_value=1.0, value=50.0, step=1.0, key="area")

        # ---------- PV MODULE ----------
        st.header("สเปคแผงโซลาร์ (ต่อแผง)")
        st.number_input("Vmp (V)", 10.0, value=float(st.session_state.get("Vmp", 41.0)), step=0.1, key="Vmp")
        st.number_input("Voc (V)", 10.0, value=float(st.session_state.get("Voc", 50.0)), step=0.1, key="Voc")
        st.number_input("Imp (A)", 1.0,  value=float(st.session_state.get("Imp", 13.0)), step=0.1, key="Imp")
        st.number_input("Isc (A)", 1.0,  value=float(st.session_state.get("Isc", 13.5)), step=0.1, key="Isc")
        st.number_input("กำลังแผง (Pm, W)", 100, value=int(st.session_state.get("Pm", 550)), step=5, key="Pm")

        # ---------- INVERTER ----------
        st.header("สเปคอินเวอร์เตอร์")
        st.number_input("AC Rated Power (W)", min_value=1000, value=int(st.session_state.get("inv_power_ac", 10000)), step=500, key="inv_power_ac")
        st.number_input("DC Max Voltage (V)", min_value=100, value=max(100, int(st.session_state.get("inv_v_dc_max", 1100))), step=50, key="inv_v_dc_max")
        st.number_input("Max Input Current / MPPT (A)", min_value=1.0, value=float(st.session_state.get("inv_i_sc_max", 25.0)), step=0.5, key="inv_i_sc_max")
        st.number_input("Max PV Power (W)", min_value=1000, value=int(st.session_state.get("inv_pv_power_max", 13000)), step=500, key="inv_pv_power_max")

        st.subheader("MPPT Setting")
        st.number_input("MPPT Min Voltage (V)", min_value=10, value=max(10, int(st.session_state.get("v_mppt_min", 200))), step=10, key="v_mppt_min")
        st.number_input("MPPT Max Voltage (V)", min_value=50, value=max(50, int(st.session_state.get("v_mppt_max", 850))), step=10, key="v_mppt_max")
        _mppt_val = int(st.session_state.get("mppt_count", 1))
        st.number_input("จำนวน MPPT (autofill จาก DB)", min_value=1, max_value=12,
                        value=_mppt_val, step=1, key="mppt_count",
                        help="ค่านี้จะถูกเติมอัตโนมัติเมื่อเลือก Inverter จาก Database")

        # ---------- ECONOMICS ----------
        st.header("เศรษฐศาสตร์โครงการ")
        st.caption("💡 ต้นทุนคำนวณจากราคาอุปกรณ์ใน Database อัตโนมัติ")
        st.number_input("ราคาแผงโซลาร์ต่อแผง (บาท/แผง)", min_value=0, value=int(st.session_state.get("panel_price_thb", 4500)), step=100, key="panel_price_thb")
        st.number_input("ราคา Inverter (บาท)", min_value=0, value=int(st.session_state.get("inv_price_thb", 20000)), step=1000, key="inv_price_thb")
        st.slider("ค่าอุปกรณ์เสริม + ติดตั้ง (%)", 10, 60, 30, 5, key="accessories_pct",
                  help="สายไฟ, โครงเหล็ก, ค่าแรง, ค่าดำเนินการ ฯลฯ")
        st.number_input("ค่าไฟฟ้า (Tariff, บาท/kWh)", min_value=0.0, value=4.0, step=0.1, key="tariff")
        st.number_input("ค่าไฟส่งออก (Export Tariff, บาท/kWh)", min_value=0.0, value=0.0, step=0.1, key="export_tariff")
        st.slider("สัดส่วนใช้ไฟเอง (Self-use ratio)", 0.0, 1.0, 0.6, 0.05, key="self_use")
        st.number_input("อายุโครงการ (ปี)", min_value=1, value=25, step=1, key="years")

        submitted = st.form_submit_button("⚡ Calculate PV System")

    return submitted