# =========================================================
# financial.py — Financial Performance Analysis
# =========================================================

from config import (
    DISCOUNT_RATE,
    DEGRADATION_RATE,
    OM_RATIO,
    INV_REPLACEMENT_YEAR,
    INV_REPLACEMENT_COST,
)


# =========================================================
# IRR (Newton-Raphson)
# =========================================================
def irr(cashflows, guess=0.1):
    """Newton-Raphson IRR — returns None ถ้าไม่ converge หรือค่าผิดปกติ"""
    # Guard: ถ้า cashflow บวกทั้งหมด หรือลบทั้งหมด → ไม่มี IRR
    pos = any(cf > 0 for cf in cashflows[1:])
    neg = any(cf < 0 for cf in cashflows[:1])
    if not pos or not neg:
        return None

    r = guess
    for _ in range(200):
        try:
            f  = sum(cf / ((1 + r) ** i) for i, cf in enumerate(cashflows))
            df = sum(-i * cf / ((1 + r) ** (i + 1)) for i, cf in enumerate(cashflows))
        except (OverflowError, ZeroDivisionError):
            return None
        if abs(df) < 1e-12:
            break
        step = f / df
        r -= step
        if r <= -1:          # r ไม่สามารถน้อยกว่า -100%
            r = -0.9999
        if abs(step) < 1e-8: # converged
            break

    # ตรวจผลลัพธ์สมเหตุสมผล: -100% ถึง +500%
    if r is None or not (-1.0 < r < 5.0):
        return None
    return r


# =========================================================
# FULL FINANCIAL ANALYSIS (PVsyst-grade)
# =========================================================
def calc_financials(
    E_est_day,
    CAPEX,
    project_life,
    tariff_self,
    tariff_export=0.0,
    self_use_ratio=0.6,
    discount_rate=DISCOUNT_RATE,
    degradation=DEGRADATION_RATE,
    om_ratio=OM_RATIO,
    inv_replacement_year=INV_REPLACEMENT_YEAR,
    inv_replacement_cost=INV_REPLACEMENT_COST,
):
    """
    Returns dict with cashflows and all financial metrics.
    """
    E_year_1 = E_est_day * 365   # kWh/year

    cashflows = [-CAPEX]
    discounted_cum = -CAPEX

    simple_payback     = None
    discounted_payback = None

    for y in range(1, project_life + 1):
        # PV degradation
        E_y = E_year_1 * ((1 - degradation) ** (y - 1))

        # Revenue split
        revenue = (
            E_y * self_use_ratio * tariff_self
            + E_y * (1 - self_use_ratio) * tariff_export
        )

        # O&M — คิดเป็น % ของรายรับ ไม่ใช่ % ของ CAPEX (เหมาะกว่าสำหรับระบบเล็ก)
        # om_ratio ใช้เป็น % ของ CAPEX → ยังคงเดิม แต่ต้องไม่เกินรายรับ
        om_cost = min(CAPEX * om_ratio, revenue * 0.5)

        # Inverter replacement
        replacement = inv_replacement_cost if y == inv_replacement_year else 0

        net_cf = revenue - om_cost - replacement
        cashflows.append(net_cf)

        # Simple payback — interpolate เศษปี (ไม่รวม replacement ใน cumulative)
        if simple_payback is None:
            cum = sum(cashflows[1:])
            if cum >= CAPEX:
                prev_cum = cum - net_cf
                if net_cf > 0:
                    frac = max(0.0, min(1.0, (CAPEX - prev_cum) / net_cf))
                    simple_payback = (y - 1) + frac
                else:
                    simple_payback = float(y)

        # Discounted payback — interpolate เศษปี
        discounted_cf = net_cf / ((1 + discount_rate) ** y)
        prev_disc_cum = discounted_cum
        discounted_cum += discounted_cf
        if discounted_payback is None and discounted_cum >= 0:
            if discounted_cf > 0:
                frac = max(0.0, min(1.0, -prev_disc_cum / discounted_cf))
                discounted_payback = (y - 1) + frac
            else:
                discounted_payback = float(y)

    npv     = sum(cf / ((1 + discount_rate) ** i) for i, cf in enumerate(cashflows))
    irr_val = irr(cashflows)

    # ถ้าไม่คืนทุนภายใน project_life → ให้ค่าเป็น project_life+  (ไม่ใช่ None)
    if simple_payback is None:
        simple_payback = float(project_life + 1)
    if discounted_payback is None:
        discounted_payback = float(project_life + 1)

    return {
        "E_year_1":             E_year_1,
        "cashflows":            cashflows,
        "simple_payback":       simple_payback,
        "discounted_payback":   discounted_payback,
        "npv":                  npv,
        "irr_val":              irr_val,
        "discount_rate":        discount_rate,
        "inv_replacement_year": inv_replacement_year,
        "project_life":         project_life,
    }