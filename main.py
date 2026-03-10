# =========================================================
# main.py — Solar Rooftop Designer | Entry Point
# =========================================================

import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
try:
    from serpapi import GoogleSearch          # serpapi package
except ImportError:
    from serpapi.google_search import GoogleSearch  # google-search-results package

# ---- Local modules ----
from config import (
    SERPAPI_KEY, GEMINI_KEY, OPENAI_KEY, SPREADSHEET_KEY,
    SF_VOC_COLD, SF_VMP_HOT, SF_CURRENT,
    V_MPPT_MIN_DEFAULT, V_MPPT_MAX_DEFAULT, MPPT_COUNT_DEFAULT,
    DISCOUNT_RATE, OM_RATIO,
    INV_REPLACEMENT_YEAR, INV_REPLACEMENT_COST,
)
from sidebar   import render_sidebar
from sheets    import connect_spreadsheet, load_db, load_db_by_name, append_to_sheet, detect_worksheet_from_text
from design    import ss, validate_design_inputs, calc_pv_capacity, validate_module, calc_string_design
from financial import calc_financials
from ai_engine import ai_select_from_database
from export    import build_ieee_paper


# =========================================================
# APP CONFIG
# =========================================================
st.set_page_config(
    page_title="Solar Rooftop Designer",
    page_icon="🔆",
    layout="wide",
)
st.title(" Solar Rooftop Designer ")


# =========================================================
# SIDEBAR
# =========================================================
if "run_design" not in st.session_state:
    st.session_state.run_design = False

submitted = render_sidebar()

if submitted:
    st.session_state.run_design = True


# =========================================================
# DATABASE VIEW (MULTI-TAB)
# =========================================================
spreadsheet = connect_spreadsheet()

st.header("Equipment Database ")

tabs = {
    "Solar Panels": "Panels_DB",
    "Inverters":    "Inverters_DB",
    "Accessories":  "Accessories",
}

if "panels_db"     not in st.session_state:
    st.session_state["panels_db"]     = pd.DataFrame()
if "inverters_db"  not in st.session_state:
    st.session_state["inverters_db"]  = pd.DataFrame()
if "accessories_db" not in st.session_state:
    st.session_state["accessories_db"] = pd.DataFrame()

tab_ui = st.tabs(list(tabs.keys()))

for ui_tab, sheet_name in zip(tab_ui, tabs.values()):
    with ui_tab:
        try:
            # ใช้ load_db_by_name — cache key = sheet_name string ป้องกัน cross-tab contamination
            df = load_db_by_name(SPREADSHEET_KEY, sheet_name)

            if df.empty:
                st.info(f"{sheet_name} ยังไม่มีข้อมูล")
            else:
                st.dataframe(df, use_container_width=True)

                if sheet_name == "Panels_DB":
                    st.session_state["panels_db"] = df
                elif sheet_name == "Inverters_DB":
                    st.session_state["inverters_db"] = df
                elif sheet_name == "Accessories":
                    st.session_state["accessories_db"] = df

        except Exception as e:
            st.error(f"❌ ไม่สามารถโหลดแท็บ {sheet_name}")
            st.caption(str(e))


# =========================================================
# SERPAPI SEARCH & SAVE
# =========================================================
st.header("🔍 ค้นหาและบันทึกอุปกรณ์")

eq_type = st.selectbox("ประเภทอุปกรณ์ (Type)", ["Panels_DB", "Inverters_DB"], key="eq_type_sel")

c1, c2 = st.columns(2)
with c1:
    brand = st.text_input("ยี่ห้อ (Brand)", key="search_brand")
    model = st.text_input("รุ่น (Model)",   key="search_model")

with c2:
    query = st.text_input(
        "คำค้นหา (Search query)",
        value=f"{brand} {model} datasheet filetype:pdf".strip(),
        key="search_query",
    )

# --- สเปคอุปกรณ์ตาม type ---
st.markdown("#### กรอกสเปค (ใส่ข้อมูลจาก Datasheet)")
if eq_type == "Panels_DB":
    sc1, sc2, sc3, sc4 = st.columns(4)
    s_pmax = sc1.number_input("Pmax (W)",         min_value=0,   value=0,    step=5,   key="s_pmax")
    s_voc  = sc1.number_input("Voc (V)",           min_value=0.0, value=0.0,  step=0.1, key="s_voc")
    s_isc  = sc2.number_input("Isc (A)",           min_value=0.0, value=0.0,  step=0.1, key="s_isc")
    s_vmp  = sc2.number_input("Vmp (V)",           min_value=0.0, value=0.0,  step=0.1, key="s_vmp")
    s_imp  = sc3.number_input("Imp (A)",           min_value=0.0, value=0.0,  step=0.1, key="s_imp")
    s_eff  = sc3.number_input("Efficiency (%)",    min_value=0.0, value=0.0,  step=0.1, key="s_eff")
    s_price= sc4.number_input("ราคา (Price, บาท)", min_value=0,   value=0,    step=100, key="s_price_panel")
else:
    sc1, sc2, sc3, sc4 = st.columns(4)
    s_pkw    = sc1.number_input("AC Power (kW)",       min_value=0.0, value=0.0,  step=0.5, key="s_pkw")
    s_pvmax  = sc1.number_input("Max PV Power (W)",    min_value=0,   value=0,    step=500, key="s_pvmax")
    s_idcmax = sc2.number_input("Max DC Current (A)",  min_value=0.0, value=0.0,  step=0.5, key="s_idcmax")
    s_vdcmax = sc2.number_input("Max DC Voltage (V)",  min_value=0,   value=0,    step=50,  key="s_vdcmax")
    s_vmpmin = sc3.number_input("MPPT Min (V)",        min_value=0,   value=0,    step=10,  key="s_vmpmin")
    s_vmpmax = sc3.number_input("MPPT Max (V)",        min_value=0,   value=0,    step=10,  key="s_vmpmax")
    s_price_inv = sc4.number_input("ราคา (Price, บาท)",min_value=0,   value=0,    step=500, key="s_price_inv")

col_btn1, col_btn2 = st.columns([1, 3])
do_search = col_btn1.button("🔍 Search Datasheet", key="btn_search")
do_save   = col_btn2.button("💾 Save to Database", key="btn_save", type="primary")

# ---- SEARCH ----
if do_search:
    if not SERPAPI_KEY:
        st.error("❌ ยังไม่ได้ตั้งค่า SERPAPI_KEY"); st.stop()
    if not brand or not model:
        st.warning("⚠️ กรุณากรอก Brand และ Model"); st.stop()

    params = {"engine": "google", "q": query, "api_key": SERPAPI_KEY, "num": 10}
    res = GoogleSearch(params).get_dict()

    pdf_candidates = []
    for r in res.get("organic_results", []):
        link  = r.get("link", "")
        title = r.get("title", "").lower()
        if link.lower().endswith(".pdf"):
            score = 0
            if "datasheet" in title or "data sheet" in title: score += 2
            if "specification" in title: score += 1
            if brand.lower() in title:  score += 1
            if model.lower() in title:  score += 2
            pdf_candidates.append({"title": r.get("title",""), "link": link,
                                    "score": score, "source": r.get("source","Google")})

    pdf_candidates = sorted(pdf_candidates, key=lambda x: x["score"], reverse=True)
    st.session_state["_pdf_candidates"] = pdf_candidates
    st.session_state["_search_done"]    = True

if st.session_state.get("_search_done"):
    pdf_candidates = st.session_state.get("_pdf_candidates", [])
    st.markdown("**Datasheet ที่พบ:**")
    if pdf_candidates:
        for i, p in enumerate(pdf_candidates[:3], start=1):
            st.markdown(f"**{i}. {p['title']}** | [เปิด PDF]({p['link']}) | Source: {p['source']}")
    else:
        st.warning("⚠️ ไม่พบ Datasheet PDF ที่ชัดเจน")

# ---- SAVE ----
if do_save:
    if not brand or not model:
        st.warning("⚠️ กรุณากรอก Brand และ Model"); st.stop()

    try:
        ws = spreadsheet.worksheet(eq_type)
    except Exception:
        st.error(f"❌ ไม่พบแท็บ {eq_type} ใน Google Sheets"); st.stop()

    records  = ws.get_all_records()
    df_exist = pd.DataFrame(records) if records else pd.DataFrame()

    # Duplicate check
    if not df_exist.empty and {"Brand", "Model"}.issubset(df_exist.columns):
        dup = df_exist[
            (df_exist["Brand"].str.lower() == brand.lower()) &
            (df_exist["Model"].str.lower() == model.lower())
        ]
        if not dup.empty:
            st.warning("⚠️ อุปกรณ์นี้มีอยู่แล้วในฐานข้อมูล")
            st.dataframe(dup); st.stop()

    pdf_candidates = st.session_state.get("_pdf_candidates", [])
    datasheet = pdf_candidates[0]["link"]   if pdf_candidates else ""
    source    = pdf_candidates[0]["source"] if pdf_candidates else "Manual"
    now_str   = datetime.now().strftime("%Y-%m-%d %H:%M")

    # --- บันทึกตาม column order ของแต่ละ sheet ---
    if eq_type == "Panels_DB":
        # Brand, Model, Pmax_W, Voc_V, Isc_A, Vmp_V, Imp_A, Efficiency_pct,
        # Price_THB, Datasheet_URL, Source, Last_Update
        row_data = [
            brand, model,
            s_pmax, s_voc, s_isc, s_vmp, s_imp, s_eff,
            s_price, datasheet, source, now_str
        ]
    else:
        # Brand, Model, Power_kW, Max_PV_Power_W, Max_DC_Current_A, Max_DC_Voltage_V,
        # MPPT_min_V, MPPT_max_V, Type, Phase, Electrical_Check, Price_THB
        row_data = [
            brand, model,
            s_pkw, s_pvmax, s_idcmax, s_vdcmax, s_vmpmin, s_vmpmax,
            "On-Grid", 1, "", s_price_inv
        ]

    append_to_sheet(ws, row_data)
    st.session_state["_search_done"] = False
    st.success(f"✅ บันทึก {brand} {model} ลงแท็บ {eq_type} เรียบร้อย")
    st.rerun()


# =========================================================
# PV SYSTEM DESIGN
# =========================================================
st.header("🔆 ออกแบบระบบผลิตไฟฟ้าพลังงานแสงอาทิตย์ (Solar Rooftop Designer)")

if not st.session_state.get("run_design", False):
    st.info("⬅️ กรุณากรอกข้อมูลทาง Sidebar แล้วกด **Calculate PV System**")
    st.stop()

# ---- รวบรวมข้อมูลตั้งต้น ----
E_day = ss("E_day"); H_sun = ss("H_sun"); PR = ss("PR"); area = ss("area")
if min(E_day, H_sun, PR, area) <= 0:
    st.error("❌ ข้อมูล Load / PSH / PR / Area ต้องมากกว่า 0"); st.stop()
for w in validate_design_inputs(E_day, H_sun, PR, area):
    st.warning(f"⚠️ {w}")

P_pv_load, P_pv_area, P_pv_design, E_est_day = calc_pv_capacity(E_day, H_sun, PR, area)

Pm=ss("Pm"); Vmp=ss("Vmp"); Voc=ss("Voc"); Imp=ss("Imp"); Isc=ss("Isc")
if min(Pm, Vmp, Voc, Imp, Isc) <= 0:
    st.error("❌ สเปคแผงไม่ครบหรือมีค่าติดลบ"); st.stop()
for err in validate_module(Pm, Vmp, Voc, Imp, Isc):
    if "ต้องมากกว่า" in err: st.error(f"❌ {err}"); st.stop()
    else: st.warning(f"⚠️ {err}")

inv_ac=ss("inv_power_ac"); inv_v=ss("inv_v_dc_max")
inv_i=ss("inv_i_sc_max");  inv_pv=ss("inv_pv_power_max")
if min(inv_ac, inv_v, inv_i, inv_pv) <= 0:
    st.error("❌ สเปค Inverter ไม่ถูกต้อง"); st.stop()

d = calc_string_design(
    P_pv_design, Pm, Vmp, Voc, Imp, Isc,
    inv_ac, inv_v, inv_i, inv_pv,
    v_mppt_min=float(ss("v_mppt_min") or V_MPPT_MIN_DEFAULT),
    v_mppt_max=float(ss("v_mppt_max") or V_MPPT_MAX_DEFAULT),
    mppt_count=int(ss("mppt_count")   or MPPT_COUNT_DEFAULT),
)
if d["panels_per_string"] < d["n_min_mppt"]:
    st.error("❌ ไม่สามารถจัด String ให้อยู่ใน MPPT window"); st.stop()

dc_capacity       = d["dc_capacity"]
dc_ac_ratio       = d["dc_ac_ratio"]
panels_per_string = d["panels_per_string"]
strings_used      = d["strings_used"]
Voc_str  = d.get("Voc_string", 0) or 0
Vmp_str  = d.get("Vmp_string", 0) or 0
I_str    = d.get("I_string",   0) or 0
mppt_alloc = d.get("mppt_allocation", [])

# แจ้งเตือน
if d.get("auto_reduced") and d.get("I_op", 0) > inv_i * 1.02:
    st.warning(f"⚠️ Imp = {d['I_op']:.2f} A เกิน Max Input Current/MPPT = {inv_i:.1f} A → ปรับ 1 string/MPPT อัตโนมัติ")
elif d.get("I_op", 0) > inv_i:
    st.info(f"ℹ️ Imp = {d['I_op']:.2f} A เกิน {inv_i:.1f} A เล็กน้อย (tolerance 2%) → ใช้งานได้")
if d.get("isc_warning"):
    st.warning(f"⚠️ Isc_string = {d['I_string']:.2f} A เกิน Inverter Isc limit ≈ {d['inv_i_sc']:.1f} A")
if d["strings_used"] < d["strings_required"]:
    st.warning("⚠️ จำนวน String ถูกจำกัดด้วยกระแส Inverter → ระบบอาจผลิตไฟได้ไม่เต็มตาม Design PV")
if d["dc_power_installed"] > inv_pv:
    st.warning(f"⚠️ DC Power ติดตั้ง = {d['dc_power_installed']/1000:.2f} kWp เกิน Inverter PV Max ({inv_pv/1000:.2f} kWp)")
if d["string_clamped"]:
    st.info("ℹ️ ปรับจำนวนแผงต่อ string ให้ไม่เกินความต้องการจริง (engineering clamp)")

# ตรวจสอบเงื่อนไข
v_min_val = float(ss("v_mppt_min") or V_MPPT_MIN_DEFAULT)
voc_ok  = bool(Voc_str and Voc_str <= inv_v)
isc_ok  = bool(I_str   and I_str   <= inv_i * 1.25)
vmpp_ok = bool(Vmp_str and Vmp_str >= v_min_val)
pps_ok  = bool(panels_per_string >= d["n_min_mppt"])
str_ok  = bool(strings_used <= d.get("strings_max", 999))
dc_ok   = bool(d["dc_power_installed"] <= inv_pv)

def _ok(c):
    if c:
        return '<span style="color:#375623;font-weight:bold;font-size:13px">✅ ผ่าน</span>'
    return '<span style="color:#C00000;font-weight:bold;font-size:13px">❌ ไม่ผ่าน</span>'

HDR1 = "background:#1F5C8B;color:white;padding:8px 12px;border-radius:6px 6px 0 0;font-weight:bold;font-size:14px"
HDR2 = "background:#2E75B6;color:white;padding:8px 12px;border-radius:6px 6px 0 0;font-weight:bold;font-size:14px"
TR_Y = "background:#FFF2CC"
TR_G = "background:#F2F2F2"
TD   = "padding:5px 9px;border:1px solid #9DC3E6"
TDC  = "padding:5px 9px;border:1px solid #9DC3E6;font-weight:bold;text-align:center"
TDU  = "padding:5px 9px;border:1px solid #9DC3E6;color:#777;text-align:center;font-size:11px"

# =================================================================
# ส่วนที่ 1: สเปคอุปกรณ์ (Equipment Specifications)
# =================================================================
st.markdown("---")
st.markdown("### 📋 สเปคอุปกรณ์ (Equipment Specifications)")

v_min_d = ss("v_mppt_min") or V_MPPT_MIN_DEFAULT
v_max_d = ss("v_mppt_max") or V_MPPT_MAX_DEFAULT
n_mppt  = int(ss("mppt_count") or MPPT_COUNT_DEFAULT)

col_inv, col_pv = st.columns(2)

with col_inv:
    rows_inv = [
        ("กำลังไฟฟ้า AC (Rated AC Power)",            f"{inv_ac/1000:.1f}", "kW"),
        ("แรงดันสูงสุดที่รับได้ (Max. DC Voltage)",    f"{inv_v}",            "V"),
        ("แรงดัน MPPT ต่ำสุด (MPPT Min. Voltage)",    f"{v_min_d}",          "V"),
        ("แรงดัน MPPT สูงสุด (MPPT Max. Voltage)",    f"{v_max_d}",          "V"),
        ("กระแสลัดวงจรสูงสุด (Max. SC Current)",      f"{inv_i}",            "A"),
        ("จำนวน MPPT Tracker",                         f"{n_mppt}",           "ชุด"),
        ("กำลัง PV สูงสุดที่รับได้ (Max. PV Power)",  f"{inv_pv/1000:.1f}", "kWp"),
    ]
    rows_html = "".join([
        f'<tr style="{TR_Y}"><td style="{TD}">{r[0]}</td>'
        f'<td style="{TDC}">{r[1]}</td><td style="{TDU}">{r[2]}</td></tr>'
        for r in rows_inv
    ])
    st.markdown(
        f'<div style="{HDR1}">⚡ Inverter (อินเวอร์เตอร์)</div>'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px">''<tr><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:left">รายการ</th><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:center">ค่า</th><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:center">หน่วย</th></tr>'f'{rows_html}</table>',
        unsafe_allow_html=True)

with col_pv:
    rows_pv = [
        ("กำลังไฟฟ้าสูงสุด (Rated Power, Pm)",         f"{Pm:.0f}",         "W"),
        ("แรงดัน ณ กำลังสูงสุด (Vmp)",                 f"{Vmp:.2f}",        "V"),
        ("แรงดันวงจรเปิด (Open Circuit Voltage, Voc)",  f"{Voc:.2f}",        "V"),
        ("กระแส ณ กำลังสูงสุด (Imp)",                   f"{Imp:.2f}",        "A"),
        ("กระแสลัดวงจร (Short Circuit Current, Isc)",   f"{Isc:.2f}",        "A"),
        ("ขนาดระบบที่ออกแบบ (Design PV Capacity)",      f"{P_pv_design:.2f}","kWp"),
        ("พลังงานที่ผลิตได้ (Estimated Energy/Day)",     f"{E_est_day:.1f}",  "kWh/วัน"),
    ]
    rows_html2 = "".join([
        f'<tr style="{TR_Y}"><td style="{TD}">{r[0]}</td>'
        f'<td style="{TDC}">{r[1]}</td><td style="{TDU}">{r[2]}</td></tr>'
        for r in rows_pv
    ])
    st.markdown(
        f'<div style="{HDR2}">☀️ PV Module (แผงโซลาร์) at STC</div>'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px">''<tr><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:left">รายการ</th><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:center">ค่า</th><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:center">หน่วย</th></tr>'f'{rows_html2}</table>',
        unsafe_allow_html=True)

# =================================================================
# ส่วนที่ 2: ผลการออกแบบ String + ตรวจสอบ OK/FAIL
# =================================================================
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 🔗 ผลการออกแบบ String (String Design Results)")

col_sum, col_chk = st.columns([55, 45])

with col_sum:
    sum_rows = [
        ("จำนวนแผงโซลาร์ใน 1 String (Panels/String)",         str(panels_per_string), "แผง"),
        ("Vmp รวม String อุณหภูมิสูง (Vmpp,string hot)",       f"{Vmp_str:.1f}",       "V"),
        ("Voc รวม String อุณหภูมิต่ำ (Voc,string cold)",       f"{Voc_str:.1f}",       "V"),
        ("กระแสลัดวงจรของ String (Isc_string)",                 f"{I_str:.2f}",         "A"),
        ("จำนวน String ที่ใช้งาน (Strings Used)",               str(strings_used),      "string"),
        ("กำลัง DC ติดตั้งรวม (Total DC Installed Capacity)",   f"{dc_capacity:.2f}",   "kWp"),
    ]
    s_rows_html = "".join([
        f'<tr style="{TR_Y}"><td style="{TD}">{r[0]}</td>'
        f'<td style="{TDC}">{r[1]}</td><td style="{TDU}">{r[2]}</td></tr>'
        for r in sum_rows
    ])
    st.markdown(
        f'<div style="{HDR2}">📊 ตารางสรุปการออกแบบ</div>'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px">''<tr><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:left">รายการ</th><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:center">ค่า</th><th style="background:#2E75B6;color:white;padding:5px 9px;border:1px solid #9DC3E6;text-align:center">หน่วย</th></tr>'f'{s_rows_html}</table>',
        unsafe_allow_html=True)

with col_chk:
    # (label, actual_val, unit, limit_val, condition_ok)
    chk_rows = [
        ("จำนวนแผงต่อ String ≥ จำนวนแผงขั้นต่ำ MPPT",
         f"{panels_per_string}", "แผง", f"≥ {d['n_min_mppt']}", pps_ok),
        ("แรงดัน Vmp รวม String (ร้อน) ≥ แรงดัน MPPT ต่ำสุด",
         f"{Vmp_str:.1f}", "V", f"≥ {v_min_val:.0f} V", vmpp_ok),
        ("แรงดัน Voc รวม String (เย็น) ≤ แรงดันสูงสุด Inverter",
         f"{Voc_str:.1f}", "V", f"≤ {inv_v:.0f} V", voc_ok),
        ("กระแสลัดวงจร String ≤ กระแส Inverter × 1.25",
         f"{I_str:.2f}", "A", f"≤ {inv_i*1.25:.2f} A", isc_ok),
        ("จำนวน String ≤ จำนวน String สูงสุดของ Inverter",
         f"{strings_used}", "string", f"≤ {d.get('strings_max','-')}", str_ok),
        ("กำลัง DC รวม ≤ กำลัง PV สูงสุดที่ Inverter รับได้",
         f"{dc_capacity:.2f}", "kWp", f"≤ {inv_pv/1000:.1f} kWp", dc_ok),
    ]
    chk_rows_html = "".join([
        f'<tr style="{TR_G}">'
        f'<td style="{TD};font-size:12px">{r[0]}</td>'
        f'<td style="padding:5px 6px;border:1px solid #9DC3E6;font-weight:bold;text-align:center">{r[1]}</td>'
        f'<td style="padding:5px 4px;border:1px solid #9DC3E6;color:#777;text-align:center;font-size:11px">{r[2]}</td>'
        f'<td style="padding:5px 6px;border:1px solid #9DC3E6;color:#555;text-align:center;font-size:11px">{r[3]}</td>'
        f'<td style="padding:5px 6px;border:1px solid #9DC3E6;text-align:center;background:#E2EFDA">{_ok(r[4])}</td>'
        f'</tr>'
        for r in chk_rows
    ])
    st.markdown(
        f'<div style="{HDR2}">✅ ตรวจสอบเงื่อนไขความปลอดภัย</div>'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px">'
        f'<tr>'
        f'<th style="background:#2E75B6;color:white;padding:5px 8px;border:1px solid #9DC3E6;text-align:left">เงื่อนไข</th>'
        f'<th style="background:#2E75B6;color:white;padding:5px 6px;border:1px solid #9DC3E6;text-align:center">ค่า</th>'
        f'<th style="background:#2E75B6;color:white;padding:5px 4px;border:1px solid #9DC3E6;text-align:center">หน่วย</th>'
        f'<th style="background:#2E75B6;color:white;padding:5px 6px;border:1px solid #9DC3E6;text-align:center">เกณฑ์</th>'
        f'<th style="background:#2E75B6;color:white;padding:5px 6px;border:1px solid #9DC3E6;text-align:center">ผล</th>'
        f'</tr>'
        f'{chk_rows_html}</table>',
        unsafe_allow_html=True)

# =================================================================
# ส่วนที่ 3: MPPT Allocation Table
# =================================================================
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### ⚙️ การจัด MPPT Allocation")

n_mppt_show = int(ss("mppt_count") or MPPT_COUNT_DEFAULT)
alloc_show  = (list(mppt_alloc) + [0]*n_mppt_show)[:n_mppt_show]
dc_total_w  = dc_capacity * 1000
dc_ac_bg    = "#E2EFDA" if 1.10 <= dc_ac_ratio <= 1.25 else ("#FCE4D6" if dc_ac_ratio > 1.25 else "#FFF2CC")

mppt_th = "".join([
    f'<th style="background:#2E75B6;color:white;padding:6px 8px;border:1px solid #9DC3E6">MPPT.{i+1}</th>'
    for i in range(n_mppt_show)
])
mppt_td = "".join([
    f'<td style="text-align:center;padding:5px 8px;border:1px solid #9DC3E6;background:#F2F2F2">{v}</td>'
    for v in alloc_show
])
st.markdown(
    f'<table style="width:100%;border-collapse:collapse;font-size:12px;text-align:center">'
    f'<tr>'
    f'<th style="background:#1F5C8B;color:white;padding:6px 8px;border:1px solid #9DC3E6">Inverter</th>'
    f'{mppt_th}'
    f'<th style="background:#1F5C8B;color:white;padding:6px 8px;border:1px solid #9DC3E6">จำนวนแผง<br>ทั้งหมด</th>'
    f'<th style="background:#1F5C8B;color:white;padding:6px 8px;border:1px solid #9DC3E6">วัตต์แผง<br>(Wp/Module)</th>'
    f'<th style="background:#1F5C8B;color:white;padding:6px 8px;border:1px solid #9DC3E6">Inverter<br>(W)</th>'
    f'<th style="background:#1F5C8B;color:white;padding:6px 8px;border:1px solid #9DC3E6">วัตต์รวม<br>ของแผง (W)</th>'
    f'<th style="background:#1F5C8B;color:white;padding:6px 8px;border:1px solid #9DC3E6">DC to AC<br>Ratio</th>'
    f'</tr>'
    f'<tr>'
    f'<td style="padding:5px 8px;border:1px solid #9DC3E6;background:#F2F2F2">1</td>'
    f'{mppt_td}'
    f'<td style="padding:5px 8px;border:1px solid #9DC3E6;background:#F2F2F2">{d.get("panels_required","-")}</td>'
    f'<td style="padding:5px 8px;border:1px solid #9DC3E6;background:#F2F2F2">{Pm:.0f}</td>'
    f'<td style="padding:5px 8px;border:1px solid #9DC3E6;background:#F2F2F2">{inv_ac:.0f}</td>'
    f'<td style="padding:5px 8px;border:1px solid #9DC3E6;background:#F2F2F2">{dc_total_w:.0f}</td>'
    f'<td style="padding:5px 8px;border:1px solid #9DC3E6;background:{dc_ac_bg};font-weight:bold;font-size:14px">{dc_ac_ratio:.2f}</td>'
    f'</tr></table>',
    unsafe_allow_html=True)

# =================================================================
# ส่วนที่ 4: DC/AC Ratio Banner
# =================================================================
if 1.10 <= dc_ac_ratio <= 1.25:
    dc_status = "✅ อยู่ในช่วงที่แนะนำ (1.10 – 1.25) | ระบบมีประสิทธิภาพดี"
    dc_col = "#E2EFDA"; dc_bdr = "#375623"
elif dc_ac_ratio > 1.25:
    dc_status = "⚠️ สูงกว่าช่วงแนะนำ (>1.25) | เสี่ยงต่อ Clipping Loss"
    dc_col = "#FCE4D6"; dc_bdr = "#C00000"
else:
    dc_status = "⚠️ ต่ำกว่าช่วงแนะนำ (<1.10) | Inverter ใหญ่เกินไป"
    dc_col = "#FFF2CC"; dc_bdr = "#7F6000"

st.markdown(
    f'<div style="display:flex;align-items:center;border:2px solid {dc_bdr};'
    f'border-radius:8px;overflow:hidden;margin:10px 0">'
    f'<div style="background:#1F5C8B;color:white;padding:14px 20px;font-weight:bold;'
    f'font-size:14px;min-width:150px;text-align:center">⚡ DC/AC Ratio</div>'
    f'<div style="background:#F4B942;padding:14px 24px;font-size:36px;font-weight:bold;'
    f'color:#1F5C8B;min-width:110px;text-align:center">{dc_ac_ratio:.2f}</div>'
    f'<div style="background:{dc_col};padding:14px 20px;font-size:14px;'
    f'color:{dc_bdr};flex:1;font-weight:bold">{dc_status}</div>'
    f'</div>',
    unsafe_allow_html=True)

# =================================================================
# FINANCIAL — คำนวณ CAPEX จาก Database
# =================================================================
project_life   = int(st.session_state.get("years", 25))
tariff_self    = float(st.session_state.get("tariff", 4.0))
tariff_export  = float(st.session_state.get("export_tariff", 0.0))
self_use_ratio = float(st.session_state.get("self_use", 0.6))
accessories_pct = float(st.session_state.get("accessories_pct", 30)) / 100.0

# ราคาจาก DB (autofill หรือกรอกเอง)
panel_price     = float(st.session_state.get("panel_price_thb", 4500))
inv_price       = float(st.session_state.get("inv_price_thb",  20000))

# จำนวนแผงและ inverter ที่ต้องใช้จริง
n_panels_total  = d.get("panels_required", panels_per_string * strings_used)
n_inverters     = 1   # 1 inverter per design

# คำนวณต้นทุนแยกรายการ
cost_panels     = panel_price  * n_panels_total
cost_inverter   = inv_price    * n_inverters
cost_equipment  = cost_panels  + cost_inverter
cost_accessories = cost_equipment * accessories_pct
CAPEX           = cost_equipment + cost_accessories

if E_est_day <= 0 or CAPEX <= 0:
    st.warning("⚠️ Financial calculation not possible"); st.stop()

fin = calc_financials(
    E_est_day=E_est_day, CAPEX=CAPEX, project_life=project_life,
    tariff_self=tariff_self, tariff_export=tariff_export, self_use_ratio=self_use_ratio,
)
simple_payback     = fin["simple_payback"]
discounted_payback = fin["discounted_payback"]
npv                = fin["npv"]
irr_val            = fin["irr_val"]

e_yr    = fin.get("E_year_1", E_est_day * 365)
co2_yr  = e_yr * 0.4715 / 1000
co2_25  = co2_yr * 25
pb_str  = f"{simple_payback:.1f} ปี"     if simple_payback     else f">{project_life} ปี"
dpb_str = f"{discounted_payback:.1f} ปี" if discounted_payback else f">{project_life} ปี"
irr_str = f"{irr_val*100:.1f}%"          if irr_val is not None else "-"
npv_str = f"{npv:,.0f} ฿"               if npv is not None    else "-"

# =================================================================
# ส่วนที่ 5: ตารางต้นทุนรวม (CAPEX Breakdown)
# =================================================================
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 💰 ผลการวิเคราะห์เศรษฐศาสตร์ (Financial Analysis)")

st.markdown(
    f'''<div style="background:#1F5C8B;color:white;padding:6px 12px;border-radius:6px 6px 0 0;font-weight:bold">
    💵 สรุปต้นทุนการติดตั้ง (CAPEX Breakdown)
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr>
        <th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:left">รายการ</th>
        <th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:center">จำนวน</th>
        <th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:center">ราคาต่อหน่วย</th>
        <th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:center">รวม (บาท)</th>
      </tr>
      <tr style="background:#FFF2CC">
        <td style="padding:5px 10px;border:1px solid #9DC3E6">🔆 แผงโซลาร์เซลล์ (PV Modules)</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;text-align:center">{n_panels_total} แผง</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;text-align:center">{panel_price:,.0f} ฿/แผง</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;font-weight:bold;text-align:center">{cost_panels:,.0f}</td>
      </tr>
      <tr style="background:#F2F2F2">
        <td style="padding:5px 10px;border:1px solid #9DC3E6">⚡ อินเวอร์เตอร์ (Inverter)</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;text-align:center">{n_inverters} เครื่อง</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;text-align:center">{inv_price:,.0f} ฿/เครื่อง</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;font-weight:bold;text-align:center">{cost_inverter:,.0f}</td>
      </tr>
      <tr style="background:#FFF2CC">
        <td style="padding:5px 10px;border:1px solid #9DC3E6">🔧 อุปกรณ์เสริม + ค่าติดตั้ง ({accessories_pct*100:.0f}%)</td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;text-align:center" colspan="2">
          สายไฟ, โครงเหล็ก, Protection, ค่าแรงติดตั้ง
        </td>
        <td style="padding:5px 10px;border:1px solid #9DC3E6;font-weight:bold;text-align:center">{cost_accessories:,.0f}</td>
      </tr>
      <tr style="background:#1F5C8B">
        <td style="padding:7px 10px;border:1px solid #9DC3E6;color:white;font-weight:bold" colspan="3">
          💰 ต้นทุนรวมทั้งหมด (Total CAPEX)
        </td>
        <td style="padding:7px 10px;border:1px solid #9DC3E6;color:#F4B942;font-weight:bold;text-align:center;font-size:16px">
          {CAPEX:,.0f} ฿
        </td>
      </tr>
    </table>''',
    unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =================================================================
# ส่วนที่ 5: Metric Cards
# =================================================================
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("💵 เงินลงทุน (CAPEX)",       f"{CAPEX:,.0f} ฿")
m2.metric("📅 คืนทุนธรรมดา (Payback)", pb_str)
m3.metric("📅 คืนทุนคิดลด (Disc. PB)", dpb_str)
m4.metric("📈 IRR",                      irr_str)
m5.metric("💹 NPV",                      npv_str)
m6.metric("🌿 ลด CO₂/ปี",               f"{co2_yr:.2f} t")

# =================================================================
# ส่วนที่ 6: Financial Detail Table
# =================================================================
fin_rows = [
    (TR_Y, "เงินลงทุนเริ่มต้น (Capital Expenditure: CAPEX)",              f"{CAPEX:,.0f}",          "บาท"),
    (TR_G, "พลังงานที่ผลิตได้ปีแรก (Year-1 Energy Production)",           f"{e_yr:,.0f}",           "kWh/ปี"),
    (TR_Y, "อัตราค่าไฟที่ประหยัดได้ (Self-use Electricity Tariff)",       f"{tariff_self:.2f}",     "บาท/kWh"),
    (TR_G, "สัดส่วนการใช้ไฟเอง (Self-consumption Ratio)",                  f"{self_use_ratio*100:.0f}", "%"),
    (TR_Y, "ระยะเวลาคืนทุนแบบธรรมดา (Simple Payback Period)",             pb_str,                   "ปี"),
    (TR_G, "ระยะเวลาคืนทุนแบบคิดลด (Discounted Payback Period)",          dpb_str,                  "ปี"),
    (TR_Y, "มูลค่าปัจจุบันสุทธิ (Net Present Value: NPV)",                f"{npv:,.0f}",            "บาท"),
    (TR_G, "อัตราผลตอบแทนภายใน (Internal Rate of Return: IRR)",           irr_str,                  "%"),
    ("#E2EFDA", "การลดการปล่อย CO₂ ต่อปี (Annual CO₂ Reduction)",         f"{co2_yr:.2f}",          "tCO₂/ปี"),
    ("#E2EFDA", "การลดการปล่อย CO₂ ตลอด 25 ปี (Lifetime CO₂ Reduction)", f"{co2_25:.1f}",          "tCO₂ (25 ปี)"),
]
fin_html = "".join([
    f'<tr style="background:{r[0]}"><td style="{TD}">{r[1]}</td>'
    f'<td style="{TDC};color:{"#375623" if "CO" in r[1] else "black"}">{r[2]}</td>'
    f'<td style="{TDU}">{r[3]}</td></tr>'
    for r in fin_rows
])
note_yr = fin.get("inv_replacement_year", 12)
st.markdown(
    f'<div style="{HDR1}">📊 รายละเอียดการวิเคราะห์เศรษฐศาสตร์และสิ่งแวดล้อม</div>'
    f'<table style="width:100%;border-collapse:collapse;font-size:13px">'
    f'<tr><th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:left">รายการ</th>'
    f'<th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:center">ค่า</th>'
    f'<th style="background:#2E75B6;color:white;padding:6px 10px;border:1px solid #9DC3E6;text-align:center">หน่วย</th></tr>'
    f'{fin_html}'
    f'<tr style="background:#F2F2F2"><td colspan="3" style="{TD};color:#666;font-size:11px">'
    f'หมายเหตุ: PV Degradation = 0.5%/ปี | O&amp;M = 1.5%/ปี | '
    f'Inverter Replacement ปีที่ {note_yr} | Grid Emission Factor = 0.4715 kgCO₂/kWh'
    f'</td></tr></table>',
    unsafe_allow_html=True)



# =========================================================
# AI RECOMMENDATION
# =========================================================
if "ai_result"  not in st.session_state:
    st.session_state["ai_result"]  = None
if "ai_loading" not in st.session_state:
    st.session_state["ai_loading"] = False

if st.button("Generate AI Recommendation", disabled=st.session_state["ai_loading"]):

    panels_df   = st.session_state.get("panels_db",   pd.DataFrame())
    inverters_df = st.session_state.get("inverters_db", pd.DataFrame())

    if panels_df.empty or inverters_df.empty:
        st.warning("⚠️ Equipment database not loaded.")
        st.stop()

    st.session_state["ai_loading"] = True
    try:
        with st.spinner("AI selecting optimal equipment..."):
            ai_result = ai_select_from_database(
                panels_df=panels_df,
                inverters_df=inverters_df,
                dc_capacity=dc_capacity,
                dc_ac_ratio=dc_ac_ratio,
                area=area,
                GEMINI_KEY=GEMINI_KEY,
                OPENAI_KEY=OPENAI_KEY,
                # ส่ง sidebar spec เพื่อให้ AI ใช้ reference ที่ถูกต้อง
                sidebar_Imp=ss("Imp"),
                sidebar_Isc=ss("Isc"),
                sidebar_Vmp=ss("Vmp"),
                sidebar_Voc=ss("Voc"),
                sidebar_Pm=ss("Pm"),
                sidebar_inv_i=ss("inv_i_sc_max"),
                sidebar_string_design=d,
            )
            st.session_state["ai_result"] = ai_result
        st.success("✅ AI recommendation generated successfully.")
    except Exception as e:
        st.session_state["ai_result"] = "AI execution failed."
        st.error(f"❌ AI Error: {str(e)}")
    finally:
        st.session_state["ai_loading"] = False

if st.session_state.get("ai_result"):
    st.markdown("## AI Recommendation Result")
    st.code(st.session_state["ai_result"])


# =========================================================
# EXPORT IEEE ENGINEERING PAPER
# =========================================================
st.header(" Export IEEE Engineering Paper")

if st.button(" Generate IEEE Paper", key="ieee_export_btn"):
    pdf_bytes = build_ieee_paper(
        # design basis
        dc_capacity=dc_capacity,
        dc_ac_ratio=dc_ac_ratio,
        H_sun=H_sun,
        PR=PR,
        area=area,
        E_day=E_day,
        E_est_day=E_est_day,
        # string design dict
        d=d,
        # inverter specs
        inv_ac=inv_ac,
        inv_v=inv_v,
        inv_i=inv_i,
        inv_pv=inv_pv,
        v_mppt_min=float(ss("v_mppt_min") or 0),
        v_mppt_max=float(ss("v_mppt_max") or 0),
        mppt_count=int(ss("mppt_count") or 1),
        # module specs
        Pm=Pm, Vmp=Vmp, Voc=Voc, Imp=Imp, Isc=Isc,
        # financial
        CAPEX=CAPEX,
        simple_payback=simple_payback,
        discounted_payback=discounted_payback,
        npv=npv,
        irr_val=irr_val,
        project_life=project_life,
        tariff_self=tariff_self,
        tariff_export=tariff_export,
        E_year_1=fin.get("E_year_1"),
        # legacy
        panels_per_string=panels_per_string,
        strings_used=strings_used,
        # AI
        ai_result=st.session_state.get("ai_result", None),
    )

    st.download_button(
        "⬇ Download PDF",
        data=pdf_bytes,
        file_name="IEEE_Solar_PV_Paper.pdf",
        mime="application/pdf",
        key="download_ieee_btn",
    )