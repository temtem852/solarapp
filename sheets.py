# =========================================================
# sheets.py — Google Sheets Helpers
# =========================================================

import time
import pandas as pd
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from datetime import datetime

from config import (
    SPREADSHEET_KEY,
    SERVICE_ACCOUNT_FILE,
    TAB_KEYWORDS,
    DEFAULT_TAB,
)

# =========================================================
# RETRY CONFIG (Exponential Backoff)
# =========================================================
MAX_RETRIES  = 4
RETRY_DELAYS = [2, 5, 10, 20]
RETRY_CODES  = {500, 503, 429}


def _with_retry(fn, *args, label="API call", **kwargs):
    last_error = None
    for attempt, delay in enumerate(RETRY_DELAYS, start=1):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.args[0].get("code", 0) if e.args else 0
            if status in RETRY_CODES:
                last_error = e
                st.warning(
                    f"⚠️ {label} — Google API {status}, "
                    f"retry {attempt}/{MAX_RETRIES} ใน {delay}s..."
                )
                time.sleep(delay)
            else:
                raise
        except Exception:
            raise

    st.error(f"❌ {label} ล้มเหลวหลังลอง {MAX_RETRIES} ครั้ง: {last_error}")
    raise last_error


# =========================================================
# CONNECT GOOGLE SHEETS
# =========================================================
@st.cache_resource
def connect_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes,
    )
    client = gspread.authorize(creds)
    return _with_retry(client.open_by_key, SPREADSHEET_KEY, label="connect_spreadsheet")


# =========================================================
# LOAD DATABASE — cache by sheet_name string (not object)
# =========================================================
@st.cache_data(ttl=300)
def load_db_by_name(spreadsheet_key: str, sheet_name: str) -> pd.DataFrame:
    """
    Cache key = (spreadsheet_key, sheet_name) — ไม่มีโอกาส cross-contaminate.
    ต้องเรียกผ่าน load_db() wrapper ด้านล่าง
    """
    spreadsheet = connect_spreadsheet()
    try:
        ws      = _with_retry(spreadsheet.worksheet, sheet_name, label=f"worksheet:{sheet_name}")
        records = _with_retry(ws.get_all_records, label=f"get_all_records:{sheet_name}")
    except Exception as e:
        st.error(f"❌ Load sheet '{sheet_name}' failed: {e}")
        return pd.DataFrame()

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def load_db(_worksheet) -> pd.DataFrame:
    """
    Backward-compatible wrapper.
    ดึง sheet title จาก worksheet object แล้วเรียก load_db_by_name
    เพื่อให้ cache key ถูกต้องเสมอ
    """
    try:
        sheet_name = _worksheet.title
    except Exception:
        sheet_name = str(_worksheet)
    return load_db_by_name(SPREADSHEET_KEY, sheet_name)


# =========================================================
# AUTO DETECT WORKSHEET
# =========================================================
def detect_worksheet_from_text(text: str, spreadsheet):
    if not text:
        return _with_retry(spreadsheet.worksheet, DEFAULT_TAB, label="worksheet")

    text = str(text).lower()
    for sheet_name, keywords in TAB_KEYWORDS.items():
        for kw in keywords:
            if kw and kw.lower() in text:
                try:
                    return _with_retry(spreadsheet.worksheet, sheet_name, label=f"worksheet:{sheet_name}")
                except Exception as e:
                    st.warning(f"⚠️ ไม่พบ tab: {sheet_name} ({e})")

    try:
        return _with_retry(spreadsheet.worksheet, DEFAULT_TAB, label="worksheet:default")
    except Exception:
        sheets = _with_retry(spreadsheet.worksheets, label="worksheets")
        return sheets[0]


# =========================================================
# APPEND ROW (RAW — prevent formula injection)
# =========================================================
def append_to_sheet(worksheet, row: list):
    try:
        _with_retry(worksheet.append_row, row, label="append_to_sheet", value_input_option="RAW")
    except Exception as e:
        st.error(f"❌ Append failed: {e}")
        raise


# =========================================================
# HIGH-LEVEL HELPER
# =========================================================
def save_search_result_to_sheet(search_query, brand, model, power, datasheet_url, source="Google"):
    spreadsheet = connect_spreadsheet()
    worksheet   = detect_worksheet_from_text(f"{search_query} {brand} {model}", spreadsheet)
    append_to_sheet(worksheet, [
        brand, model, power, datasheet_url, source,
        datetime.now().strftime("%Y-%m-%d %H:%M"), search_query,
    ])
    return worksheet.title


# =========================================================
# DETECT INVERTER AC COLUMN
# =========================================================
def find_ac_column(df: pd.DataFrame):
    if df is None or df.empty:
        return None
    candidates = ["Power_kW", "AC_kW", "Rated Power", "AC Power (kW)", "AC Power", "Nominal AC Power", "Output Power"]
    for col in candidates:
        if col in df.columns:
            return col
    for col in df.columns:
        c = col.lower()
        if "ac" in c and ("kw" in c or "power" in c):
            return col
    return None
