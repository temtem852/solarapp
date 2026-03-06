# =========================================================
# main.py — Solar Rooftop Designer | Entry Point
# =========================================================

import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
from serpapi import GoogleSearch

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
# SERPAPI SEARCH
# =========================================================
st.header(" ค้นหาอุปกรณ์ ")

c1, c2 = st.columns(2)

with c1:
    eq_type = st.selectbox("ประเภทอุปกรณ์ (Type)", ["Panels_DB", "Inverters_DB"])
    brand   = st.text_input("ยี่ห้อ (Brand)")
    model   = st.text_input("รุ่น (Model)")
    power   = st.number_input("กำลังไฟฟ้า (Power, W)", min_value=0)

with c2:
    query = st.text_input(
        "คำค้นหา (Search query)",
        value=f"{brand} {model} datasheet filetype:pdf".strip(),
    )

if st.button(" Search & Save"):

    if not SERPAPI_KEY:
        st.error("❌ ยังไม่ได้ตั้งค่า SERPAPI_KEY")
        st.stop()

    if not brand or not model:
        st.warning("⚠️ กรุณากรอก Brand และ Model")
        st.stop()

    try:
        ws = spreadsheet.worksheet(eq_type)
    except Exception:
        st.error(f"❌ ไม่พบแท็บ {eq_type} ใน Google Sheets")
        st.stop()

    records  = ws.get_all_records()
    df_exist = pd.DataFrame(records) if records else pd.DataFrame()

    params = {
        "engine":  "google",
        "q":       query,
        "api_key": SERPAPI_KEY,
        "num":     10,
    }
    res = GoogleSearch(params).get_dict()

    pdf_candidates = []
    for r in res.get("organic_results", []):
        link    = r.get("link", "")
        title   = r.get("title", "").lower()

        if link.lower().endswith(".pdf"):
            score = 0
            if "datasheet" in title or "data sheet" in title:
                score += 2
            if "specification" in title:
                score += 1
            if brand.lower() in title:
                score += 1
            if model.lower() in title:
                score += 2
            pdf_candidates.append({
                "title":  r.get("title", ""),
                "link":   link,
                "score":  score,
                "source": r.get("source", "Google"),
            })

    pdf_candidates = sorted(pdf_candidates, key=lambda x: x["score"], reverse=True)

    st.markdown("### Datasheet ที่พบ ")
    if pdf_candidates:
        for i, p in enumerate(pdf_candidates[:3], start=1):
            st.markdown(
                f"**{i}. {p['title']}**  \n"
                f" [เปิด Datasheet PDF]({p['link']})  \n"
                f"แหล่งที่มา (Source): {p['source']}"
            )
    else:
        st.warning("⚠️ ไม่พบ Datasheet PDF ที่ชัดเจน")

    datasheet = pdf_candidates[0]["link"]   if pdf_candidates else ""
    source    = pdf_candidates[0]["source"] if pdf_candidates else "Google"

    # Duplicate check
    if not df_exist.empty and {"Brand", "Model"}.issubset(df_exist.columns):
        dup = df_exist[
            (df_exist["Brand"].str.lower() == brand.lower()) &
            (df_exist["Model"].str.lower() == model.lower())
        ]
        if not dup.empty:
            st.warning("⚠️ อุปกรณ์นี้มีอยู่แล้วในฐานข้อมูล")
            st.dataframe(dup)
            st.stop()

    append_to_sheet(ws, [
        brand, model, power, "",
        datasheet, source,
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        query,
    ])

    st.success(f"บันทึกอุปกรณ์ลงแท็บ {eq_type} เรียบร้อย")
    st.rerun()


# =========================================================
# PV SYSTEM DESIGN
# =========================================================
st.header(" PV System Design | การออกแบบระบบผลิตไฟฟ้าพลังงานแสงอาทิตย์")

if not st.session_state.get("run_design", False):
    st.info("⬅️ กรุณากรอกข้อมูลทาง Sidebar แล้วกด **Calculate PV System**")
    st.stop()


# --- Design Basis ---
st.markdown("## Design Basis | ข้อมูลตั้งต้น")

E_day = ss("E_day")
H_sun = ss("H_sun")
PR    = ss("PR")
area  = ss("area")

if min(E_day, H_sun, PR, area) <= 0:
    st.error("❌ ข้อมูล Load / PSH / PR / Area ต้องมากกว่า 0")
    st.stop()

for w in validate_design_inputs(E_day, H_sun, PR, area):
    st.warning(f"⚠️ {w}")

st.info(
    f"**Design Inputs Summary**\n"
    f"- Daily energy demand: **{E_day:.1f} kWh/day**\n"
    f"- Peak Sun Hours (PSH): **{H_sun:.2f} h/day**\n"
    f"- Performance Ratio (PR): **{PR:.2f}**\n"
    f"- Available area: **{area:.1f} m²**"
)


# --- PV Capacity Sizing ---
st.markdown("## PV Capacity Sizing | คำนวณขนาดระบบ")

P_pv_load, P_pv_area, P_pv_design, E_est_day = calc_pv_capacity(E_day, H_sun, PR, area)

st.markdown(
    f"- PV from load: **{P_pv_load:.2f} kWp**\n"
    f"- PV from area: **{P_pv_area:.2f} kWp**\n\n"
    f"✅ **Design PV Capacity: {P_pv_design:.2f} kWp**  \n"
    f"Estimated Energy: **{E_est_day:.2f} kWh/day**"
)


# --- PV Module ---
st.markdown("## PV Module | สเปคแผงจากผู้ใช้")

Pm  = ss("Pm")
Vmp = ss("Vmp")
Voc = ss("Voc")
Imp = ss("Imp")
Isc = ss("Isc")

if min(Pm, Vmp, Voc, Imp, Isc) <= 0:
    st.error("❌ สเปคแผงไม่ครบหรือมีค่าติดลบ")
    st.stop()

for err in validate_module(Pm, Vmp, Voc, Imp, Isc):
    if "ต้องมากกว่า" in err:
        st.error(f"❌ {err}")
        st.stop()
    else:
        st.warning(f"⚠️ {err} → ตรวจสอบ datasheet อีกครั้ง")

st.info(
    f"**Module Electrical Summary**\n"
    f"- Rated Power (Pm): **{Pm:.0f} W**\n"
    f"- Vmp / Imp: **{Vmp:.1f} V / {Imp:.1f} A**\n"
    f"- Voc / Isc: **{Voc:.1f} V / {Isc:.1f} A**"
)


# --- Inverter ---
st.markdown("## Inverter | สเปคอินเวอร์เตอร์จากผู้ใช้")

inv_ac = ss("inv_power_ac")
inv_v  = ss("inv_v_dc_max")
inv_i  = ss("inv_i_sc_max")
inv_pv = ss("inv_pv_power_max")

if min(inv_ac, inv_v, inv_i, inv_pv) <= 0:
    st.error("❌ สเปค Inverter ไม่ถูกต้อง")
    st.stop()

dc_ac_actual = P_pv_design * 1000 / inv_ac
if dc_ac_actual < 1.0:
    st.warning("⚠️ Inverter ใหญ่เกินไป → Efficiency ต่ำ")
elif dc_ac_actual > 1.35:
    st.warning("⚠️ DC/AC ratio สูง → เสี่ยง clipping")
else:
    st.info("✅ ขนาด Inverter เหมาะสม")


# --- String Design ---
st.markdown("## String Design | ออกแบบจำนวนแผงต่อ String")

d = calc_string_design(
    P_pv_design,
    Pm, Vmp, Voc, Imp, Isc,
    inv_ac, inv_v, inv_i, inv_pv,
    v_mppt_min=float(ss("v_mppt_min") or V_MPPT_MIN_DEFAULT),
    v_mppt_max=float(ss("v_mppt_max") or V_MPPT_MAX_DEFAULT),
    mppt_count=int(ss("mppt_count") or MPPT_COUNT_DEFAULT),
)

if d["string_clamped"]:
    st.info("ℹ️ ปรับจำนวนแผงต่อ string ให้ไม่เกินความต้องการจริง (engineering clamp)")

if d["panels_per_string"] < d["n_min_mppt"]:
    st.error("❌ ไม่สามารถจัด String ให้อยู่ใน MPPT window")
    st.stop()

st.info(f"✔ แผงต่อ String: **{d['panels_per_string']} แผง**")


# --- String Quantity ---
st.markdown("## String Quantity | คำนวณจำนวน String")

# แสดง warning เฉพาะถ้าเกิน tolerance 2%
if d.get("auto_reduced") and d.get("I_op", 0) > inv_i * 1.02:
    st.warning(
        f"⚠️ Imp = {d['I_op']:.2f} A เกิน Max Input Current/MPPT = {inv_i:.1f} A\n"
        f"→ ปรับเป็น 1 string/MPPT อัตโนมัติ ควรเลือก Inverter ที่ Max Input Current >= {d['I_op']:.1f} A"
    )
elif d.get("I_op", 0) > inv_i:
    st.info(
        f"ℹ️ Imp = {d['I_op']:.2f} A เกิน {inv_i:.1f} A เล็กน้อย (อยู่ใน tolerance 2%) → ใช้งานได้"
    )

if d.get("isc_warning"):
    st.warning(
        f"⚠️ Isc_string = {d['I_string']:.2f} A เกิน Inverter Isc limit ≈ {d['inv_i_sc']:.1f} A\n"
        f"→ ตรวจสอบ Max Short-Circuit Current ใน Datasheet Inverter"
    )

st.write(
    f"- Panels required: **{d['panels_required']} แผง**\n"
    f"- Strings required (ตามโหลด): **{d['strings_required']} string**\n"
    f"- Inverter รองรับได้สูงสุด: **{d['strings_max']} string**"
)

if d["strings_used"] < d["strings_required"]:
    st.warning(
        "⚠️ จำนวน String ถูกจำกัดด้วยกระแส Inverter\n"
        "→ ระบบอาจผลิตไฟได้ไม่เต็มตาม Design PV"
    )
else:
    st.success("✅ จำนวน String เพียงพอตาม Design PV")

if d["dc_power_installed"] > inv_pv:
    st.warning(
        f"⚠️ DC Power ติดตั้ง = {d['dc_power_installed']/1000:.2f} kWp "
        f"เกิน Inverter PV Max ({inv_pv/1000:.2f} kWp)"
    )


# --- MPPT Allocation ---
st.markdown("## MPPT Allocation | การกระจาย String")
for i, s in enumerate(d["mppt_allocation"], start=1):
    st.write(f"- MPPT {i}: **{s} string(s)**")


# --- Final Electrical Check ---
st.markdown("## Final Electrical Check | ตรวจสอบขั้นสุดท้าย")

dc_capacity = d["dc_capacity"]
dc_ac_ratio = d["dc_ac_ratio"]
panels_per_string = d["panels_per_string"]
strings_used = d["strings_used"]

st.success(
    f"### ✅ Final System Configuration\n"
    f"- DC Capacity: **{dc_capacity:.2f} kWp**\n"
    f"- DC/AC Ratio: **{dc_ac_ratio:.2f}**\n"
    f"- Voc,string (cold): **{d['Voc_string']:.0f} V**\n"
    f"- Vmpp,string (hot): **{d['Vmp_string']:.0f} V**"
)

st.write(st.session_state.get("ai_result", "ยังไม่ได้เรียก AI"))


# =========================================================
# FINANCIAL PERFORMANCE
# =========================================================
st.header("Financial Performance | PVsyst-grade Analysis")

CAPEX        = float(st.session_state.get("CAPEX", 480_000))
project_life = int(st.session_state.get("years", 25))
tariff_self  = float(st.session_state.get("tariff", 4.0))
tariff_export = float(st.session_state.get("export_tariff", 0.0))
self_use_ratio = float(st.session_state.get("self_use", 0.6))

if E_est_day <= 0 or CAPEX <= 0:
    st.warning("⚠️ Financial calculation not possible")
    st.stop()

fin = calc_financials(
    E_est_day=E_est_day,
    CAPEX=CAPEX,
    project_life=project_life,
    tariff_self=tariff_self,
    tariff_export=tariff_export,
    self_use_ratio=self_use_ratio,
)

simple_payback     = fin["simple_payback"]
discounted_payback = fin["discounted_payback"]
npv                = fin["npv"]
irr_val            = fin["irr_val"]

st.markdown(
    f"### ผลการวิเคราะห์ทางการเงิน (Financial Results – PVsyst-grade)\n\n"
    f"**เศรษฐศาสตร์ของระบบ (System Economics)**\n"
    f"- เงินลงทุนเริ่มต้น (CAPEX): **{CAPEX:,.0f} THB**\n"
    f"- พลังงานปีแรก (Year-1 Energy): **{fin['E_year_1']:,.0f} kWh/year**\n"
    f"- อัตราการใช้ไฟเอง (Self-consumption): **{self_use_ratio*100:.0f} %**\n\n"
    f"**ตัวชี้วัดทางการเงิน (Financial Indicators)**\n"
    f"- ระยะเวลาคืนทุนแบบธรรมดา (Simple Payback):  \n"
    f"  **{simple_payback if simple_payback else '>' + str(project_life)} ปี (years)**\n\n"
    f"- ระยะเวลาคืนทุนแบบคิดลด (Discounted Payback):  \n"
    f"  **{discounted_payback if discounted_payback else '>' + str(project_life)} ปี (years)**\n\n"
    f"- มูลค่าปัจจุบันสุทธิ (NPV) @ {fin['discount_rate']*100:.0f}%:  \n"
    f"  **{npv:,.0f} THB**\n\n"
    f"- อัตราผลตอบแทนภายใน (IRR):  \n"
    f"  **{irr_val*100:.1f} %**\n\n"
    f"**หมายเหตุเชิงวิศวกรรม (Engineering Notes)**\n"
    f"- คิดค่าการเสื่อมสภาพของแผง PV (PV degradation) = **0.5 %/year**\n"
    f"- ค่าบำรุงรักษาระบบ (O&M) = **1.5 % ของ CAPEX ต่อปี**\n"
    f"- ค่าทดแทนอินเวอร์เตอร์ (Inverter replacement) ปีที่ **{fin['inv_replacement_year']}**\n"
    f"- รายได้แยกการใช้ไฟเอง (Self-use) และไฟส่งออก (Export)"
)


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
        dc_capacity=dc_capacity,
        dc_ac_ratio=dc_ac_ratio,
        H_sun=H_sun,
        PR=PR,
        area=area,
        E_day=E_day,
        E_est_day=E_est_day,
        panels_per_string=panels_per_string,
        strings_used=strings_used,
        simple_payback=simple_payback,
        irr_val=irr_val,
        project_life=project_life,
        ai_result=st.session_state.get("ai_result", "No AI result available."),
    )

    st.download_button(
        "⬇ Download PDF",
        data=pdf_bytes,
        file_name="IEEE_Solar_PV_Paper.pdf",
        mime="application/pdf",
        key="download_ieee_btn",
    )
