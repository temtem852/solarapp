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
    r = guess
    for _ in range(100):
        f  = sum(cf / ((1 + r) ** i) for i, cf in enumerate(cashflows))
        df = sum(-i * cf / ((1 + r) ** (i + 1)) for i, cf in enumerate(cashflows))
        if abs(df) < 1e-9:
            break
        r -= f / df
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

        # O&M
        om_cost = CAPEX * om_ratio

        # Inverter replacement
        replacement = inv_replacement_cost if y == inv_replacement_year else 0

        net_cf = revenue - om_cost - replacement
        cashflows.append(net_cf)

        # Simple payback
        if simple_payback is None:
            if sum(cashflows[1:]) >= CAPEX:
                simple_payback = y

        # Discounted payback
        discounted_cf = net_cf / ((1 + discount_rate) ** y)
        discounted_cum += discounted_cf
        if discounted_payback is None and discounted_cum >= 0:
            discounted_payback = y

    npv     = sum(cf / ((1 + discount_rate) ** i) for i, cf in enumerate(cashflows))
    irr_val = irr(cashflows)

    return {
        "E_year_1":           E_year_1,
        "cashflows":          cashflows,
        "simple_payback":     simple_payback,
        "discounted_payback": discounted_payback,
        "npv":                npv,
        "irr_val":            irr_val,
        "discount_rate":      discount_rate,
        "inv_replacement_year": inv_replacement_year,
    }
