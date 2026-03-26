# =========================================================
# design.py — PV System Design & Engineering Logic
# =========================================================

import numpy as np
import streamlit as st

from config import (
    SF_VOC_COLD,
    SF_VMP_HOT,
    SF_CURRENT,
    V_MPPT_MIN_DEFAULT,
    V_MPPT_MAX_DEFAULT,
    MPPT_COUNT_DEFAULT,
)


# =========================================================
# HELPER: Safe session_state read
# =========================================================
def ss(key, default=0.0):
    try:
        return float(st.session_state.get(key, default))
    except Exception:
        return default


# =========================================================
# ENGINEERING RANGE WARNINGS (non-blocking)
# =========================================================
def validate_design_inputs(E_day, H_sun, PR, area):
    """Returns list of warning strings."""
    warnings = []
    if not (1.0 <= H_sun <= 7.0):
        warnings.append("PSH อยู่นอกช่วงปกติ (1–7 h/day)")
    if not (0.65 <= PR <= 0.90):
        warnings.append("PR อยู่นอกช่วงที่พบได้ทั่วไป (0.65–0.90)")
    if E_day < 5:
        warnings.append("โหลดไฟฟ้าค่อนข้างต่ำ อาจไม่คุ้มค่าทางเศรษฐศาสตร์")
    if area < 10:
        warnings.append("พื้นที่ติดตั้งจำกัด อาจจำกัดขนาดระบบ")
    return warnings


# =========================================================
# PV CAPACITY SIZING
# =========================================================
def calc_pv_capacity(E_day, H_sun, PR, area):
    """
    Returns (P_pv_design_kWp, E_est_day_kWh).
    Takes minimum of load-based and area-based sizing.
    """
    P_pv_load = E_day / (H_sun * PR)
    P_pv_area = area * 0.20          # ≈ 200 W/m² packing density
    P_pv_design = min(P_pv_load, P_pv_area)
    E_est_day = P_pv_design * H_sun * PR
    return P_pv_load, P_pv_area, P_pv_design, E_est_day


# =========================================================
# PV MODULE VALIDATION
# =========================================================
def validate_module(Pm, Vmp, Voc, Imp, Isc):
    """
    Returns list of error strings.
    Empty list = module specs are consistent.
    """
    errors = []
    Pm_calc = Vmp * Imp
    if Pm_calc < 0.9 * Pm or Pm_calc > 1.1 * Pm:
        errors.append(
            f"ความไม่สอดคล้องของสเปคแผง: "
            f"Pm datasheet = {Pm:.0f} W, Vmp × Imp = {Pm_calc:.0f} W"
        )
    if Voc <= Vmp:
        errors.append("Voc ต้องมากกว่า Vmp")
    if Isc <= Imp:
        errors.append("Isc ต้องมากกว่า Imp")
    return errors


# =========================================================
# STRING DESIGN
# =========================================================
def calc_string_design(
    P_pv_design,
    Pm, Vmp, Voc, Imp, Isc,
    inv_ac, inv_v, inv_i, inv_pv,
    v_mppt_min=V_MPPT_MIN_DEFAULT,
    v_mppt_max=V_MPPT_MAX_DEFAULT,
    mppt_count=MPPT_COUNT_DEFAULT,
):
    # defensive cast
    mppt_count = int(mppt_count)
    v_mppt_min = float(v_mppt_min)
    v_mppt_max = float(v_mppt_max)
    """
    Full IEEE string design calculation.
    Returns a result dict with all engineering values.
    """
    # --- String sizing bounds ---
    n_max_voc  = int(inv_v / (Voc * SF_VOC_COLD))
    n_max_mppt = int(v_mppt_max / (Vmp * SF_VMP_HOT))  # Vmpp,hot = Vmp × SF_VMP_HOT
    n_min_mppt = int(np.ceil(v_mppt_min / (Vmp * SF_VMP_HOT * 0.90)))  # Vmpp,hot × 0.90

    panels_required_est = int(np.ceil(P_pv_design * 1000 / Pm))
    dc_ac_max = 1.30

    # คำนวณ strings_max ก่อน (ขึ้นกับ Imp vs inv_i)
    _I_string_pre = Isc * SF_CURRENT              # Isc×1.25 ตาม spec 3.4.5
    _spm_pre      = max(1, int(inv_i // _I_string_pre)) if _I_string_pre > 0 else 1
    _strings_max_pre = _spm_pre * mppt_count

    # -------------------------------------------------------
    # pps Algorithm:
    # ขั้น 1: หา pps_elec = ขีดจำกัดทางไฟฟ้า (Voc, MPPT)
    # ขั้น 2: หา strings_req ถ้าใช้ pps_elec
    # ขั้น 3: ถ้า strings_req <= strings_max → ลด pps ให้พอดีกับโหลด
    #         ถ้า strings_req > strings_max  → ใช้ strings_max แล้วหา pps
    # ขั้น 4: DC/AC clamp
    # -------------------------------------------------------
    _pps_elec    = max(1, min(n_max_voc, n_max_mppt))
    _str_req_at_elec = int(np.ceil(panels_required_est / _pps_elec))

    if _str_req_at_elec <= _strings_max_pre:
        # strings ที่ต้องการ ไม่เกิน capacity → ลด pps ให้พอดีโหลด
        _str_target   = _str_req_at_elec
        pps_load      = int(np.ceil(panels_required_est / _str_target))
        panels_per_string = max(1, min(_pps_elec, pps_load))
    else:
        # strings ไม่พอ → จำกัดที่ strings_max แล้ว pps ตามนั้น
        pps_load      = int(np.ceil(panels_required_est / _strings_max_pre)) if _strings_max_pre > 0 else panels_required_est
        panels_per_string = max(1, min(_pps_elec, pps_load))

    clamped = panels_per_string < _pps_elec

    # DC/AC safety check
    for _ in range(panels_per_string):
        _su    = min(int(np.ceil(panels_required_est / panels_per_string)), _strings_max_pre)
        _dc_ac = (panels_per_string * _su * Pm / 1000) / (inv_ac / 1000)
        if _dc_ac <= dc_ac_max or panels_per_string <= 1:
            break
        panels_per_string -= 1
        clamped = True

    # --- String quantity ---
    panels_required = int(np.ceil(P_pv_design * 1000 / Pm))
    strings_required = int(np.ceil(panels_required / panels_per_string))

    # ---------------------------------------------------------
    # IEEE / IEC 62548: แยก 2 ค่ากระแส
    #   I_op  = Imp  → เทียบ Max Input Current per MPPT (operating)
    #   I_sc  = Isc × SF_CURRENT → เทียบ Max Short-Circuit Current
    #
    # inv_i (sidebar) = Max input current per MPPT
    # inv_i_sc_max    = Isc limit = inv_i × SF_CURRENT (ถ้าไม่มี field แยก)
    # ---------------------------------------------------------
    I_op        = Imp                     # operating current (1 string)
    I_sc_string = Isc * SF_CURRENT        # short-circuit current NEC-derated
    inv_i_sc    = inv_i * SF_CURRENT      # inverter Isc limit (estimated)

    # strings/MPPT จาก Imp (operating) ไม่ใช่ Isc
    # Tolerance 2% — Imp เกิน inv_i ไม่เกิน 2% ถือว่าผ่าน (rounding/datasheet tolerance)
    IMP_TOLERANCE = 0.02
    # อยู่ใน tolerance → ใช้ inv_i เป็น cap เพื่อให้ floor division ได้ 1
    # strings_per_mppt_max ตาม spec 3.4.5: I_string = Isc × 1.25
    strings_per_mppt_max = int(inv_i // I_sc_string) if I_sc_string > 0 else 0

    auto_reduced = False
    if strings_per_mppt_max < 1:
        strings_per_mppt_max = 1
        auto_reduced = True

    # warning ถ้า Isc string เกิน Isc limit ของ inverter
    isc_warning = I_sc_string > inv_i_sc

    # I_string ใช้ Isc*SF สำหรับ display (convention เดิม)
    I_string = I_sc_string

    strings_max  = strings_per_mppt_max * mppt_count
    strings_used = min(strings_required, strings_max)

    # --- ถ้า strings_used < mppt_count → ลองแบ่งแผงใหม่ให้เต็ม MPPT ---
    # เช่น 8 แผง 1 string + 2 MPPT → ลองทำ 2 string × 4 แผง
    if strings_used < mppt_count:
        _new_pps = panels_per_string // mppt_count
        if _new_pps >= 1:
            _new_vmp = _new_pps * Vmp * SF_VMP_HOT
            _new_voc = _new_pps * Voc * SF_VOC_COLD
            # ตรวจ electrical ว่าผ่านหรือไม่
            _vmp_ok = _new_vmp >= v_mppt_min
            _voc_ok = _new_voc <= inv_v
            if _vmp_ok and _voc_ok:
                panels_per_string = _new_pps
                strings_used      = mppt_count
                # คำนวณ string params ใหม่
                I_sc_string = Isc * strings_per_mppt_max  # ยังคง 1 string/MPPT
                Voc_string  = panels_per_string * Voc * SF_VOC_COLD
                Vmp_string  = panels_per_string * Vmp * SF_VMP_HOT

    # --- MPPT allocation: แบ่ง string ให้กระจายเท่าๆ กัน ---
    mppt_allocation = []
    base   = strings_used // mppt_count       # string ต่อ MPPT ขั้นต่ำ
    extra  = strings_used % mppt_count        # MPPT แรกๆ ได้ +1
    for i in range(mppt_count):
        mppt_allocation.append(base + (1 if i < extra else 0))

    # --- Final electrical ---
    dc_capacity = panels_per_string * strings_used * Pm / 1000
    dc_ac_ratio = dc_capacity / (inv_ac / 1000)

    Voc_string = panels_per_string * Voc * SF_VOC_COLD
    Vmp_string = panels_per_string * Vmp * SF_VMP_HOT
    dc_power_installed = panels_per_string * strings_used * Pm

    return {
        # String sizing
        "n_max_voc":          n_max_voc,
        "n_max_mppt":         n_max_mppt,
        "n_min_mppt":         n_min_mppt,
        "panels_per_string":  panels_per_string,
        "string_clamped":     clamped,
        # String quantity
        "panels_required":    panels_required,
        "strings_required":   strings_required,
        "I_string":           I_string,
        "strings_per_mppt_max": strings_per_mppt_max,
        "auto_reduced":       auto_reduced,
        "isc_warning":        isc_warning,
        "I_op":               I_op,
        "inv_i_sc":           inv_i_sc,
        "strings_max":        strings_max,
        "strings_used":       strings_used,
        "mppt_allocation":    mppt_allocation,
        # Final electrical
        "dc_capacity":        dc_capacity,
        "dc_ac_ratio":        dc_ac_ratio,
        "Voc_string":         Voc_string,
        "Vmp_string":         Vmp_string,
        "dc_power_installed": dc_power_installed,
    }


# =========================================================
# IEEE ENGINEERING FEASIBILITY CHECK
# =========================================================
def inverter_feasible(
    inv_i_max, inv_v_max,
    I_string, Voc_string,
    v_mppt_min, v_mppt_max,
    Vmp_string,
):
    if I_string > inv_i_max:
        return False
    if Voc_string > inv_v_max:
        return False
    if not (v_mppt_min <= Vmp_string <= v_mppt_max):
        return False
    return True