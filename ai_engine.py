# =========================================================
# ai_engine.py — IEEE MCDM + LLM Equipment Selection
# =========================================================

import os
import numpy as np
import pandas as pd

from design import inverter_feasible
from config import SF_CURRENT


# =========================================================
# COLUMN NAME ALIASES
# =========================================================
PANEL_COL_ALIASES = {
    "Power_W":        ["Pmax_W", "Power_W", "Power (W)", "Pm (W)", "Pm", "Watt", "Power"],
    "Efficiency_pct": ["Efficiency_pct", "Efficiency (%)", "Eff (%)", "Efficiency", "eff_pct"],
    "Vmp_V":          ["Vmp_V", "Vmp (V)", "Vmp"],
    "Voc_V":          ["Voc_V", "Voc (V)", "Voc"],
    "Imp_A":          ["Imp_A", "Imp (A)", "Imp"],
    "Model":          ["Model", "model", "Model Name", "Product"],
}

INV_COL_ALIASES = {
    "Power_kW":    ["Power_kW", "Power_W", "Power (W)", "AC Power (W)", "Rated Power (W)", "Power"],
    "I_max_A":     ["Max_DC_Current_A", "I_max_A", "I_max (A)", "Max Input Current (A)", "Idc_max", "I_max"],
    "V_max_V":     ["Max_DC_Voltage_V", "V_max_V", "V_max (V)", "Max DC Voltage (V)", "Vdc_max", "V_max"],
    "MPPT_min_V":  ["MPPT_min_V", "MPPT Min (V)", "MPPT_min", "Vmppt_min"],
    "MPPT_max_V":  ["MPPT_max_V", "MPPT Max (V)", "MPPT_max", "Vmppt_max"],
    "MPPT_Count":  ["MPPT_Count", "MPPT", "MPPT_count", "Num MPPT", "Number of MPPT"],
    "Efficiency":  ["Efficiency_%", "Efficiency_pct", "Efficiency (%)", "Eff_%", "eff"],
    "Max_PV_Power":["Max_PV_Power_W", "Max_PV_Power", "Max PV Power (W)", "PV_Power_max", "inv_pv_power_max"],
    "Model":       ["Model", "model", "Model Name", "Product"],
}


def _resolve_col(df: pd.DataFrame, aliases: list, field: str) -> str:
    for alias in aliases:
        if alias in df.columns:
            return alias
    raise KeyError(
        f"ไม่พบคอลัมน์ '{field}' ใน sheet\n"
        f"คอลัมน์ที่มี: {list(df.columns)}\n"
        f"คอลัมน์ที่รองรับ: {aliases}"
    )


def _map_cols(df: pd.DataFrame, alias_map: dict) -> dict:
    return {
        field: _resolve_col(df, aliases, field)
        for field, aliases in alias_map.items()
    }


# =========================================================
# LOAD WEIGHTS CONFIG
# =========================================================
def load_weights_config(csv_path: str = "Weights_Config.csv") -> dict:
    """
    โหลด Weights_Config.csv → dict
    ข้ามบรรทัดที่ขึ้นต้นด้วย # (comment)
    ถ้าไม่พบไฟล์ → ใช้ค่า default
    """
    defaults = {
        "weight_power_score":               0.6,
        "weight_eff_score":                 0.4,
        "weight_engineering":               0.7,
        "weight_ratio":                     0.3,
        "panel_power_mode":                 1.0,   # 0=Gaussian, 1=Linear
        "panel_power_mu":                   550.0,
        "panel_power_sigma":                120.0,
        "panel_power_min":                  400.0,
        "panel_power_max":                  700.0,
        "panel_eff_mu":                     21.0,
        "panel_eff_sigma":                  2.0,
        "dc_ac_ratio_sigma":                0.25,
        "hard_limit_total_panel_vs_max_pv": 1.0,
        "efficiency_target_min":            1.1,
        "efficiency_target_max":            1.3,
    }

    if not os.path.exists(csv_path):
        return defaults

    try:
        df = pd.read_csv(
            csv_path,
            comment="#",
            header=0,
            names=["Parameter", "Value", "Unit", "Description"],
        )
        cfg = defaults.copy()
        for _, row in df.iterrows():
            key = str(row["Parameter"]).strip()
            if key in cfg:
                try:
                    cfg[key] = float(row["Value"])
                except (ValueError, TypeError):
                    pass
        return cfg
    except Exception:
        return defaults


# =========================================================
# GAUSSIAN PENALTY (IEEE style)
# =========================================================
def gaussian_score(x, mu, sigma):
    if sigma == 0:
        return 0
    return np.exp(-0.5 * ((x - mu) / sigma) ** 2)


# =========================================================
# LLM EXPLANATION LAYER
# =========================================================
def generate_llm_explanation(prompt, GEMINI_KEY=None, OPENAI_KEY=None):
    openai_error = None
    gemini_error = None

    if OPENAI_KEY:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a professional solar PV engineer."},
                    {"role": "user",   "content": prompt},
                ],
                temperature=0.3,
                max_tokens=500,
            )
            content = response.choices[0].message.content
            if content:
                return content.strip()
        except Exception as e:
            openai_error = str(e)

    if GEMINI_KEY:
        try:
            import google.generativeai as genai
            genai.configure(api_key=GEMINI_KEY)
            for model_name in ["models/gemini-1.5-flash", "models/gemini-1.5-pro", "gemini-pro", "models/gemini-2.5-flash"]:
                try:
                    model    = genai.GenerativeModel(model_name)
                    response = model.generate_content(prompt)
                    if hasattr(response, "text") and response.text:
                        return response.text.strip()
                except Exception as inner:
                    gemini_error = str(inner)
        except Exception as e:
            gemini_error = str(e)

    return (
        f"AI explanation unavailable.\n"
        f"OpenAI error: {openai_error}\n"
        f"Gemini error: {gemini_error}\n"
        f"System proceeds with deterministic MCDM result only."
    )


# =========================================================
# MAIN AI SELECTOR (IEEE FULL MCDM)
# =========================================================
def ai_select_from_database(
    panels_df,
    inverters_df,
    dc_capacity,
    dc_ac_ratio,
    area,
    GEMINI_KEY=None,
    OPENAI_KEY=None,
    weights_csv: str = "Weights_Config.csv",
    # --- sidebar spec (optional) เพื่อใช้เป็น reference ---
    sidebar_Imp: float = None,
    sidebar_Isc: float = None,
    sidebar_Vmp: float = None,
    sidebar_Voc: float = None,
    sidebar_Pm:  float = None,
    sidebar_inv_i: float = None,
    sidebar_string_design: dict = None,
):
    """
    IEEE MCDM equipment selection with:
    - Hard Limit: Total_Panel_Watt <= Max_PV_Power_W
    - Efficiency Target: DC/AC ratio ระหว่าง efficiency_target_min–max
    - Weights/thresholds อ่านจาก Weights_Config.csv
    """
    if panels_df.empty or inverters_df.empty:
        return "Database empty."

    panels_df    = panels_df.copy()
    inverters_df = inverters_df.copy()

    # =====================================================
    # STEP 0 — Load config + resolve columns
    # =====================================================
    cfg = load_weights_config(weights_csv)

    try:
        p_cols = _map_cols(panels_df, PANEL_COL_ALIASES)
    except KeyError as e:
        return f"❌ Panel DB column error:\n{e}"

    # Max_PV_Power เป็น optional — inverter บางรุ่นอาจไม่มี column นี้
    inv_alias_required = {k: v for k, v in INV_COL_ALIASES.items() if k != "Max_PV_Power"}
    try:
        i_cols = _map_cols(inverters_df, inv_alias_required)
    except KeyError as e:
        return f"❌ Inverter DB column error:\n{e}"

    # หา Max_PV_Power column (optional)
    max_pv_col = None
    for alias in INV_COL_ALIASES["Max_PV_Power"]:
        if alias in inverters_df.columns:
            max_pv_col = alias
            break

    hard_limit_enabled = cfg["hard_limit_total_panel_vs_max_pv"] == 1.0
    eff_min = cfg["efficiency_target_min"]
    eff_max = cfg["efficiency_target_max"]

    # =====================================================
    # STEP 1 — Panel MCDM scoring (Power + Efficiency + Imp compatibility)
    # =====================================================
    power_series = pd.to_numeric(panels_df[p_cols["Power_W"]], errors="coerce")
    imp_series   = pd.to_numeric(panels_df[p_cols["Imp_A"]],   errors="coerce")

    # --- Power score ---
    if cfg["panel_power_mode"] == 1:
        p_min = cfg["panel_power_min"]
        p_max = cfg["panel_power_max"]
        denom = p_max - p_min if p_max != p_min else 1.0
        panels_df["power_score"] = ((power_series - p_min) / denom).clip(0.0, 1.0)
    else:
        panels_df["power_score"] = power_series.apply(
            lambda x: gaussian_score(x, cfg["panel_power_mu"], cfg["panel_power_sigma"])
        )

    # --- Efficiency score ---
    panels_df["eff_score"] = pd.to_numeric(
        panels_df[p_cols["Efficiency_pct"]], errors="coerce"
    ).apply(lambda x: gaussian_score(x, cfg["panel_eff_mu"], cfg["panel_eff_sigma"]))

    # --- Imp compatibility score ---
    # ถ้ามี sidebar_inv_i → ใช้เป็น reference แทนค่าสูงสุดจาก DB
    inv_i_col    = i_cols["I_max_A"]
    inv_i_values = pd.to_numeric(inverters_df[inv_i_col], errors="coerce")
    inv_i_ref    = float(sidebar_inv_i) if sidebar_inv_i else float(inv_i_values.max())
    isc_headroom = inv_i_ref * (SF_CURRENT - 1.0)

    def imp_compat_score(imp_val):
        if pd.isna(imp_val):
            return 0.0
        # Hard rule: Imp > inv_i_ref → กระแสเกิน operating limit → ตัดออกทันที
        if imp_val > inv_i_ref:
            return 0.0
        # ผ่าน operating limit → ยิ่ง margin มากยิ่งดี
        margin = inv_i_ref - imp_val
        return min(1.0, 0.5 + 0.5 * (margin / inv_i_ref))

    panels_df["imp_score"] = imp_series.apply(imp_compat_score)

    # --- Hard filter: ตัดแผงที่ Imp > inv_i_ref ออกก่อน scoring ---
    # ถ้าตัดแล้วไม่มีเหลือ → แสดง warning แต่ไม่กรอง (fallback)
    panels_eligible = panels_df[panels_df["imp_score"] > 0].copy()
    if panels_eligible.empty:
        panels_eligible = panels_df.copy()   # fallback: ใช้ทั้งหมด แต่จะได้ warning
        imp_no_match = True
    else:
        imp_no_match = False
    panels_df = panels_eligible

    # --- Combined panel score (re-normalized weights) ---
    w_power = cfg["weight_power_score"]
    w_eff   = cfg["weight_eff_score"]
    w_imp   = cfg.get("weight_imp_score", 0.5)
    w_total = w_power + w_eff + w_imp
    panels_df["panel_score"] = (
        (w_power / w_total) * panels_df["power_score"]
        + (w_eff   / w_total) * panels_df["eff_score"]
        + (w_imp   / w_total) * panels_df["imp_score"]
    )

    # ถ้าไม่มีแผงผ่าน Imp filter → แจ้ง error พร้อมคำแนะนำ
    if imp_no_match:
        imp_needed    = inv_i_ref
        min_imp_in_db = float(panels_df[p_cols["Imp_A"]].min())
        return (
            "ERR: No panel with Imp <= " + str(round(imp_needed, 1)) + " A\n\n"
            + "All panels in DB exceed Max Input Current of Inverter.\n\n"
            + "Solutions:\n"
            + "  1. Add panels with Imp <= " + str(round(imp_needed, 1)) + " A to Database\n"
            + "  2. Use Inverter with Max Input Current >= " + str(round(min_imp_in_db, 1)) + " A"
        )

    best_panel  = panels_df.sort_values("panel_score", ascending=False).iloc[0]
    panel_power = float(best_panel[p_cols["Power_W"]])
    Vmp         = float(best_panel[p_cols["Vmp_V"]])
    Voc         = float(best_panel[p_cols["Voc_V"]])
    Imp         = float(best_panel[p_cols["Imp_A"]])
    Isc_panel   = float(best_panel[p_cols["Imp_A"]]) * 1.05   # estimate Isc ≈ Imp*1.05 ถ้าไม่มี

    # =====================================================
    # STEP 2 — String parameters
    # ถ้ามี sidebar_string_design → ใช้ค่าจาก sidebar (ออกแบบแล้ว)
    # ถ้าไม่มี → คำนวณใหม่จาก best_panel
    # =====================================================
    if sidebar_string_design and sidebar_Pm:
        # ใช้ผลการออกแบบจาก sidebar โดยตรง
        modules_per_string = sidebar_string_design.get("panels_per_string", 8)
        n_strings          = sidebar_string_design.get("strings_used", 1)
        n_panels           = modules_per_string * n_strings
        # คำนวณ electrical ด้วยสเปค best_panel จาก DB
        Vmp_string       = modules_per_string * Vmp
        Voc_string       = modules_per_string * Voc
        I_string         = Isc_panel * SF_CURRENT
        total_panel_watt = n_panels * panel_power
    else:
        # fallback: คำนวณใหม่จาก best_panel
        n_panels_needed    = max(1, int(dc_capacity * 1000 / panel_power))
        modules_per_string = max(1, int(600 / Vmp))
        n_strings          = max(1, int(np.ceil(n_panels_needed / modules_per_string)))
        n_panels           = modules_per_string * n_strings
        Vmp_string         = modules_per_string * Vmp
        Voc_string         = modules_per_string * Voc
        I_string           = Isc_panel * SF_CURRENT
        total_panel_watt   = n_panels * panel_power

    # =====================================================
    # STEP 3 — Inverter evaluation
    # =====================================================
    power_col  = i_cols["Power_kW"]
    is_kw_unit = "kw" in power_col.lower()

    inv_scores     = []
    inv_ac_watts   = []
    hard_fail_flags = []
    eff_flags      = []

    for _, inv in inverters_df.iterrows():
        try:
            raw_power  = float(inv[power_col])
            ac_power_w = raw_power * 1000 if is_kw_unit else raw_power

            # --------------------------------------------------
            # HARD LIMIT: Total_Panel_Watt <= Max_PV_Power_W
            # --------------------------------------------------
            hard_fail = False
            if hard_limit_enabled and max_pv_col is not None:
                max_pv_w = float(inv[max_pv_col])
                if total_panel_watt > max_pv_w:
                    hard_fail = True

            # --------------------------------------------------
            # ENGINEERING FEASIBILITY (electrical)
            # --------------------------------------------------
            feasible = inverter_feasible(
                float(inv[i_cols["I_max_A"]]),
                float(inv[i_cols["V_max_V"]]),
                I_string,
                Voc_string,
                float(inv[i_cols["MPPT_min_V"]]),
                float(inv[i_cols["MPPT_max_V"]]),
                Vmp_string,
            )

            # Hard fail → score = 0 ทันที
            if hard_fail:
                engineering_penalty = 0.0
            else:
                engineering_penalty = 1.0 if feasible else 0.0

            # --------------------------------------------------
            # EFFICIENCY TARGET: DC/AC ratio ระหว่าง eff_min–eff_max
            # อยู่นอก band = Engineering Fail ทันที (score = 0)
            # --------------------------------------------------
            ratio       = (dc_capacity * 1000) / ac_power_w
            in_eff_band = eff_min <= ratio <= eff_max

            # ถ้าอยู่นอก efficiency band → fail เหมือน hard limit
            if not in_eff_band:
                engineering_penalty = 0.0
                ratio_score         = 0.0
            else:
                # Gaussian score ภายใน band — center = midpoint
                ratio_mu    = (eff_min + eff_max) / 2
                ratio_score = gaussian_score(ratio, mu=ratio_mu, sigma=cfg["dc_ac_ratio_sigma"])

            total_score = (
                cfg["weight_engineering"] * engineering_penalty
                + cfg["weight_ratio"]     * ratio_score
            )

        except Exception:
            total_score  = 0.0
            ac_power_w   = 0.0
            hard_fail    = False
            in_eff_band  = False
            ratio        = 0.0

        inv_scores.append(total_score)
        inv_ac_watts.append(ac_power_w)
        hard_fail_flags.append(hard_fail)
        eff_flags.append(in_eff_band)

    inverters_df["ai_score"]    = inv_scores
    inverters_df["_ac_power_w"] = inv_ac_watts
    inverters_df["_hard_fail"]  = hard_fail_flags
    inverters_df["_eff_ok"]     = eff_flags

    # =====================================================
    # GUARD: ถ้าทุก inverter ได้ score=0 → ไม่มีตัวผ่านเกณฑ์
    # ห้ามเลือกแบบสุ่ม → แจ้ง error พร้อมสาเหตุ
    # =====================================================
    passed_df = inverters_df[inverters_df["ai_score"] > 0]

    if passed_df.empty:
        # วิเคราะห์สาเหตุ
        n_hard_fail = sum(hard_fail_flags)
        n_eff_fail  = sum(1 for ok in eff_flags if not ok)
        reasons = []
        if n_hard_fail > 0:
            reasons.append(f"Hard Limit FAIL {n_hard_fail}/{len(inverters_df)} ตัว (Total Panel Watt เกิน Max PV Power)")
        if n_eff_fail > 0:
            reasons.append(f"Efficiency Target FAIL {n_eff_fail}/{len(inverters_df)} ตัว (DC/AC ratio อยู่นอกช่วง {eff_min}–{eff_max})")

        reason_str = "\n".join(f"  • {r}" for r in reasons)
        needed_ac  = (dc_capacity * 1000) / ((eff_min + eff_max) / 2)

        return (
            f"❌ ไม่พบ Inverter ที่ผ่านเกณฑ์ทั้งหมด (0/{len(inverters_df)} ตัว)\n\n"
            f"สาเหตุ:\n{reason_str}\n\n"
            f"คำแนะนำ:\n"
            f"  • ขนาด Inverter ที่เหมาะสมสำหรับระบบ {dc_capacity:.2f} kWp:\n"
            f"    AC Power ≈ {needed_ac/1000:.1f} kW (DC/AC target = {(eff_min+eff_max)/2:.2f})\n"
            f"  • เพิ่ม Inverter ขนาด {needed_ac/1000:.0f}–{dc_capacity*1000/eff_min/1000:.0f} kW ใน Database\n"
            f"  • หรือปรับ dc_capacity ใน Sidebar ให้ตรงกับ Inverter ที่มี"
        )

    best_inv   = passed_df.sort_values("ai_score", ascending=False).iloc[0]
    best_ratio = (dc_capacity * 1000) / best_inv["_ac_power_w"]

    # =====================================================
    # STEP 4 — Engineering verdict
    # =====================================================
    hard_fail_count  = sum(hard_fail_flags)
    eff_ok_count     = sum(eff_flags)
    total_inv        = len(inverters_df)

    hard_limit_status = (
        f"⚠️ Hard Limit FAIL ({hard_fail_count}/{total_inv} inverters ถูกตัดออก)"
        if hard_fail_count > 0
        else "✅ Hard Limit PASS (ทุก inverter รองรับ Total Panel Watt)"
    )

    eff_status = (
        f"✅ Efficiency Target PASS — DC/AC = {best_ratio:.2f} (เป้า {eff_min}–{eff_max})"
        if best_inv["_eff_ok"]
        else f"⚠️ Efficiency Target WARNING — DC/AC = {best_ratio:.2f} อยู่นอกช่วง {eff_min}–{eff_max}"
    )

    # =====================================================
    # STEP 4 — Build deterministic summary
    # =====================================================
    # ส่งกลับเป็น dict แทน text เพื่อให้ main.py แสดงผลสวยได้
    import json as _json
    ai_dict = {
        "panel_model":       str(best_panel[p_cols["Model"]]),
        "inv_model":         str(best_inv[i_cols["Model"]]),
        "total_panel_watt":  round(total_panel_watt, 0),
        "n_panels":          int(n_panels),
        "modules_per_string":int(modules_per_string),
        "n_strings":         int(n_strings),
        "Vmp_string":        round(Vmp_string, 1),
        "Voc_string":        round(Voc_string, 1),
        "I_string":          round(I_string, 2),
        "dc_ac_ratio":       round(best_ratio, 2),
        "ai_score":          round(float(best_inv["ai_score"]), 3),
        "hard_fail_count":   int(hard_fail_count),
        "total_inv":         int(total_inv),
        "eff_ok_count":      int(eff_ok_count),
        "eff_min":           eff_min,
        "eff_max":           eff_max,
        "hard_limit_pass":   hard_fail_count == 0,
        "eff_band_pass":     bool(best_inv["_eff_ok"]),
    }
    deterministic_summary = "AI_RESULT_JSON:" + _json.dumps(ai_dict, ensure_ascii=False)

    # =====================================================
    # STEP 5 — LLM Engineering Verdict (OpenAI / Gemini)
    # =====================================================
    llm_prompt = f"""
คุณเป็น Solar PV Engineer ผู้เชี่ยวชาญ วิเคราะห์ผลการออกแบบระบบนี้และให้คำแนะนำเชิงวิศวกรรมเป็นภาษาไทย

=== ข้อมูลระบบ ===
แผงโซลาร์ที่เลือก : {best_panel[p_cols['Model']]}
อินเวอร์เตอร์ที่เลือก : {best_inv[i_cols['Model']]}
กำลังแผงรวม (Total Panel Watt) : {total_panel_watt:,.0f} W
จำนวนแผงทั้งหมด : {n_panels} แผง ({modules_per_string} แผง/string × {n_strings} string)
String Vmp : {Vmp_string:.1f} V | String Voc : {Voc_string:.1f} V | String I : {I_string:.1f} A
DC/AC Ratio : {best_ratio:.2f}
AI Score : {best_inv['ai_score']:.3f}

=== ผลการตรวจสอบ ===
Hard Limit (Total Panel <= Max PV Power) : {"PASS" if hard_fail_count == 0 else f"FAIL — {hard_fail_count} inverter ถูกตัดออก"}
Efficiency Target (DC/AC {eff_min}–{eff_max}) : {"PASS" if best_inv['_eff_ok'] else "WARNING — อยู่นอกช่วงเป้าหมาย"}
Inverters ผ่านเกณฑ์ : {eff_ok_count}/{total_inv}

=== คำถาม ===
1. อุปกรณ์ที่เลือกเหมาะสมกับสภาพแวดล้อมประเทศไทยหรือไม่? เพราะอะไร?
2. DC/AC ratio ที่ได้มีผลต่อประสิทธิภาพและการ clipping อย่างไร?
3. มีข้อควรระวังหรือข้อแนะนำเพิ่มเติมสำหรับการติดตั้งจริงหรือไม่?

ตอบสั้น กระชับ เป็นข้อๆ ไม่เกิน 150 คำ
"""

    llm_verdict = generate_llm_explanation(
        llm_prompt,
        GEMINI_KEY=GEMINI_KEY,
        OPENAI_KEY=OPENAI_KEY,
    )

    result = deterministic_summary + "|||LLM|||" + llm_verdict
    return result


# =========================================================
# DATASHEET EXTRACTOR
# วิธี: ดาวน์โหลด PDF → แปลงเป็น text → ส่งให้ Gemini/OpenAI
# รองรับ free tier เพราะส่งแค่ text ไม่ส่ง PDF file
# =========================================================
import requests, json, re, io

def _pdf_to_text(url: str, timeout: int = 20) -> str:
    """ดาวน์โหลด PDF แล้วแปลงเป็น text ด้วย pdfplumber"""
    import pdfplumber
    resp = requests.get(url, timeout=timeout,
                        headers={"User-Agent": "Mozilla/5.0",
                                 "Accept": "application/pdf"})
    resp.raise_for_status()
    text_pages = []
    with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
        for page in pdf.pages[:6]:          # อ่านแค่ 6 หน้าแรก
            t = page.extract_text()
            if t:
                text_pages.append(t)
    return "\n".join(text_pages)[:6000]     # จำกัด 6000 chars


def _ask_gemini_text(prompt: str, GEMINI_KEY: str) -> str:
    """เรียก Gemini text API (ฟรี tier รองรับ)"""
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 512},
    }
    resp = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"gemini-1.5-flash:generateContent?key={GEMINI_KEY}",
        json=payload, timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["candidates"][0]["content"]["parts"][0]["text"]


def _ask_openai_text(prompt: str, OPENAI_KEY: str) -> str:
    """Fallback: เรียก OpenAI text API"""
    headers = {"Authorization": f"Bearer {OPENAI_KEY}",
               "Content-Type": "application/json"}
    payload = {
        "model": "gpt-4o-mini",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0, "max_tokens": 512,
    }
    resp = requests.post("https://api.openai.com/v1/chat/completions",
                         headers=headers, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def extract_specs_from_datasheet(
    pdf_url: str, eq_type: str, GEMINI_KEY: str = "", OPENAI_KEY: str = ""
) -> dict:
    """
    PDF URL → text → LLM → dict of specs
    ทำงานได้บน Gemini free tier เพราะไม่ส่ง PDF file
    """
    if not pdf_url:
        return {"_error": "ไม่มี PDF URL"}

    # 1. แปลง PDF เป็น text
    try:
        pdf_text = _pdf_to_text(pdf_url)
    except Exception as e:
        return {"_error": f"อ่าน PDF ไม่ได้: {e}"}

    if not pdf_text.strip():
        return {"_error": "PDF ไม่มีข้อความ (อาจเป็นรูปภาพ)"}

    # 2. สร้าง prompt
    if eq_type == "Panels_DB":
        fields = (
            "Pmax_W (Maximum Power Wp), "
            "Voc_V (Open Circuit Voltage V), "
            "Isc_A (Short Circuit Current A), "
            "Vmp_V (Maximum Power Voltage V), "
            "Imp_A (Maximum Power Current A), "
            "Efficiency_pct (Module Efficiency %)"
        )
        example = '{"Pmax_W":580,"Voc_V":49.8,"Isc_A":14.2,"Vmp_V":41.5,"Imp_A":13.6,"Efficiency_pct":22.4}'
    else:
        fields = (
            "Power_kW (Rated AC Output Power kW — if W divide by 1000), "
            "Max_PV_Power_W (Max DC Input Power W), "
            "Max_DC_Current_A (Max Input Current per MPPT A), "
            "Max_DC_Voltage_V (Max DC Voltage V), "
            "MPPT_min_V (MPPT min voltage V), "
            "MPPT_max_V (MPPT max voltage V)"
        )
        example = '{"Power_kW":10,"Max_PV_Power_W":13000,"Max_DC_Current_A":25,"Max_DC_Voltage_V":1100,"MPPT_min_V":200,"MPPT_max_V":850}'

    prompt = f"""You are a solar equipment datasheet parser.
From the following datasheet text, extract ONLY these numeric values: {fields}

Return ONLY a valid JSON object like this example: {example}
Use null for any value not found. No explanation, no markdown, just raw JSON.

--- DATASHEET TEXT ---
{pdf_text}
--- END ---"""

    # 3. เรียก LLM (Gemini ก่อน fallback OpenAI)
    raw = None
    for fn, key, name in [
        (_ask_gemini_text, GEMINI_KEY, "Gemini"),
        (_ask_openai_text, OPENAI_KEY, "OpenAI"),
    ]:
        if not key:
            continue
        try:
            raw = fn(prompt, key)
            break
        except Exception as e:
            last_err = f"{name}: {e}"

    if raw is None:
        return {"_error": last_err if "last_err" in dir() else "ไม่มี API key"}

    # 4. parse JSON
    try:
        clean = re.sub(r"```[a-z]*", "", raw).strip().strip("`").strip()
        return json.loads(clean)
    except Exception:
        return {"_error": f"แปลง JSON ไม่ได้: {raw[:200]}"}


# =========================================================
# PRICE SEARCHER — ค้นราคาอุปกรณ์ ด้วย SerpAPI
# =========================================================
def search_price(brand: str, model: str, SERPAPI_KEY: str) -> dict:
    """
    ค้นราคาจาก Google Shopping ผ่าน SerpAPI
    Returns {"price_thb": float | None, "source": str, "url": str}
    """
    if not SERPAPI_KEY:
        return {"price_thb": None, "source": "", "url": ""}

    try:
        from serpapi import GoogleSearch
    except ImportError:
        from serpapi.google_search import GoogleSearch

    queries = [
        f"{brand} {model} ราคา บาท",
        f"{brand} {model} price THB",
        f"{brand} {model} solar panel price",
    ]

    for q in queries:
        try:
            params = {
                "engine":   "google",
                "q":        q,
                "api_key":  SERPAPI_KEY,
                "num":      10,
                "gl":       "th",
                "hl":       "th",
            }
            res = GoogleSearch(params).get_dict()

            # 1. ลองดู shopping_results ก่อน
            for item in res.get("shopping_results", []):
                price_str = item.get("price", "")
                nums = re.findall(r"[\d,]+\.?\d*", price_str.replace(",", ""))
                if nums:
                    price = float(nums[0])
                    # ถ้าน้อยกว่า 100 อาจเป็นดอลลาร์ ×33
                    if price < 100:
                        price *= 33
                    return {
                        "price_thb": price,
                        "source": item.get("source", "Google Shopping"),
                        "url":    item.get("link", ""),
                    }

            # 2. ลอง organic_results — หาตัวเลขที่น่าจะเป็นราคาไทย
            for r in res.get("organic_results", []):
                snippet = r.get("snippet", "") + " " + r.get("title", "")
                # หา pattern เช่น 4,500 บาท / ฿4500 / THB 4500
                m = re.search(r"(?:฿|THB|บาท)[^\d]*(\d[\d,]+)|(\d[\d,]+)[^\d]*(?:฿|บาท|THB)", snippet)
                if m:
                    num_str = (m.group(1) or m.group(2)).replace(",", "")
                    price = float(num_str)
                    if 500 < price < 5_000_000:
                        return {
                            "price_thb": price,
                            "source": r.get("source", r.get("displayed_link", "")),
                            "url":    r.get("link", ""),
                        }
        except Exception:
            continue

    return {"price_thb": None, "source": "ไม่พบราคา", "url": ""}