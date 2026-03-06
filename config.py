# =========================================================
# config.py — Constants, Keywords, Safety Factors
# =========================================================

import os
from dotenv import load_dotenv

load_dotenv()

# =========================================================
# API KEYS
# =========================================================
SERPAPI_KEY        = os.getenv("SERPAPI_KEY")
GEMINI_KEY         = os.getenv("GEMINI_API_KEY")
OPENAI_KEY         = os.getenv("OPENAI_API_KEY")
SPREADSHEET_KEY    = os.getenv("SPREADSHEET_KEY")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")

# =========================================================
# LLM PROVIDER (Auto-detect)
# =========================================================
if GEMINI_KEY:
    LLM_PROVIDER = "gemini"
elif OPENAI_KEY:
    LLM_PROVIDER = "openai"
else:
    LLM_PROVIDER = None

# =========================================================
# GOOGLE SHEETS — TAB MAPPING
# =========================================================
TAB_KEYWORDS = {
    "Panels_DB": [
        "panel", "solar panel", "pv module", "module",
        "mono", "perc", "topcon", "bifacial", "vertex", "tiger"
    ],
    "Inverters_DB": [
        "inverter", "string inverter", "hybrid inverter",
        "on-grid", "off-grid", "mppt", "sungrow", "growatt", "huawei"
    ],
    "Batteries": [
        "battery", "lithium", "lifepo4", "storage", "bms"
    ],
    "Accessories": [
        "mount", "rail", "clamp", "mc4",
        "dc cable", "ac cable", "combiner"
    ]
}

DEFAULT_TAB = "Inverters_DB"

# =========================================================
# ENGINEERING SAFETY FACTORS (IEEE / IEC 62548)
# =========================================================
SF_VOC_COLD   = 1.20   # Voc cold-temperature correction
SF_VMP_HOT    = 0.90   # Vmp hot-temperature correction
SF_CURRENT    = 1.25   # NEC 690.8(A) string current factor

# MPPT window assumptions (when not read from DB)
V_MPPT_MIN_DEFAULT = 200   # V
V_MPPT_MAX_DEFAULT = 850   # V
MPPT_COUNT_DEFAULT = 1

# =========================================================
# FINANCIAL DEFAULTS
# =========================================================
DISCOUNT_RATE          = 0.08    # WACC
DEGRADATION_RATE       = 0.005   # 0.5 %/year PV degradation
OM_RATIO               = 0.015   # 1.5 % of CAPEX / year
INV_REPLACEMENT_YEAR   = 12
INV_REPLACEMENT_COST   = 80_000  # THB

# =========================================================
# AI / MCDM SCORING DEFAULTS
# =========================================================
PANEL_POWER_MU    = 550    # W  — Gaussian centre for panel scoring
PANEL_POWER_SIGMA = 120    # W
PANEL_EFF_MU      = 21     # %
PANEL_EFF_SIGMA   = 2      # %

DC_AC_RATIO_SIGMA = 0.25
