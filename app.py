# app.py
# Dynamic Restock v12 – Streamlit app
# Vendor / Vendor Code / Color mapped EXCLUSIVELY from STOCK
#   Vendor  (A) <- Stock['Brand'] (per Our Code)
#   Vendor Code (B) <- Stock['Vendor Code'] OR Stock['Vendors/Vendor Product Code'] (per Our Code)
#   Color  (C) <- από ρητή στήλη Color/Χρώμα ή robust parse από κείμενο (Variant Values/Options/Attributes/Title)
# SALES χρησιμοποιείται μόνο για ποσότητες (Qty Ordered).
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st
from collections import Counter

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Upload Stock + Sales → dynamic restock. Vendor / Code / Color are mapped from STOCK.")

# ---------------- Helpers ----------------
def to_int_safe(x):
    try:
        if pd.isna(x): return 0
        return int(float(str(x).strip()))
    except Exception:
        return 0

def clean_our_code(x):
    """Normalize to 8-digit numeric string (strip .0, keep digits)."""
    if pd.isna(x): return None
    s = str(x).strip()
    if s.endswith(".0"): s = s[:-2]
    s = re.sub(r"\D", "", s)
    if not s: return None
    return s.zfill(8)[:8]

def extract_size_from_variant_values(text):
    """Detect EU size 36–42 in free text."""
    if pd.isna(text): return None
    m = re.search(r"(3[6-9]|4[0-2])\b", str(text))
    return int(m.group(1)) if m else None

def extract_color_from_text(text):
    """
    Επιστρέφει το χρώμα από κείμενο τύπου:
      "Χρώμα: Ταμπά  Μεγέθη γυναικεία παπούτσια: 36"
    Ανθεκτικό σε:
      - Χρώμα/ΧΡΩΜΑ/Color με ή χωρίς ':', με '-', '–', '—'
      - έξτρα κενά
      - διαχωριστικά (',', ';', '|')
      - στοπ πριν από 'Μεγ...' (Μεγέθη/Μέγεθος) ή 'Size/Sizes/Taille' ή διαχωριστικό ή EOL
    """
    if pd.isna(text):
        return None
    s = re.sub(r"\s+", " ", str(text)).strip()
    # Κύριο regex
    m = re.search(
        r"(?:Χρώμα|ΧΡΩΜΑ|Color)\s*[:：\-–—]?\s*"
        r"(.+?)"
        r"(?=\s*(?:Μεγ[\wΆ-ώ]+|Sizes?|Size|Taille|,|;|\||$))",
        s,
        flags=re.IGNORECASE
    )
    color = m.group(1).strip() if m else None
    if color:
        color = re.sub(r"[\s,;|]+$", "", color).strip()
        color = color.strip(' "\'“”‘’')
    return color if color else None

def build_variant_sku(our_code8, size):
    """11ψήφιο SKU: OurCode(8) + (Size-34).zfill(3)"""
    if our_code8 is None or pd.isna(size): return None
    return f"{our_code8}{str(int(size)-34).zfill(3)}"

def base_target_for_size(size):
    try: s = int(size)
    except: return 0
    if s in (38,39): return 6
    if s in (37,40): return 4
    if s == 41: return 2
    if s in (36,42): return 1
    return 0

def clip(x, lo, hi):
    try: return max(lo, min(hi, x))
    except: return lo

def _norm(s):
    return re.sub(r"[\s/_\-]+", "", str(s).strip().lower())

def find_col(df, tokens, *, exclude_tokens=None):
    """Find column whose normalized name contains ALL tokens (and none of excluded)."""
    toks = [t.lower() for t in (tokens if isinstance(tokens, (list,tuple)) else [tokens])]
    excl = [t.lower() for t in (exclude_tokens or [])]
    for c in df.columns:
        nc = _norm(c)
        if all(t in nc for t in toks) and all(t not in nc for t in excl):
            return c
    return None

def find_any_col(df, list_of_token_sets, *, exclude_tokens=None):
    for tokens in list_of_token_sets:
        col = find_col(df, tokens, exclude_tokens=exclude_tokens)
        if col: return col
    return None

def first_non_null(s):
    s = s.dropna()
    return s.iloc[0] if not s.empty else np.nan

def mode_non_null(series):
    vals = [str(x).strip() for x in series if pd.notna(x) and str(x).strip() != ""]
    if not vals: return np.nan
    return Counter(vals).most_common(1)[0][0]

# ---------------- Sidebar ----------------
st.sidebar.header("⚙️ Settings")
stock_sheet = st.sidebar.text_input("Stock sheet name", value="Sheet1")
sales_sheet = st.sidebar.text_input("Sales sheet name", value="Sales Analysis")

# ---------------- Uploads ----------------
c1, c2 = st.columns(2)
with c1: stock_file = st.file_uploader("📂 Upload STOCK Excel", type=["xlsx","xls"])
with c2: sales_file = st.file_uploader("📂 Upload SALES Excel", type=["xlsx","xls"])
run_btn = st.button("🚀 Run Dynamic Restock")

if run_btn:
    if not stock_file or not sales_file:
        st.error("Please upload both STOCK and SALES files."); st.stop()

    # ---------- Read ----------
    try:
        stock_raw = pd.read_excel(stock_file, sheet_name=stock_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read STOCK sheet '{stock_sheet}': {e}"); st.stop()
    try:
        sales_raw = pd.read_excel(sales_file, sheet_name=sales_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read SALES sheet '{sales_sheet}': {e}"); st.stop()

    # ---------- STOCK parsing (authoritative for Vendor fields) ----------
    stock = stock_raw.copy()

    # Our Code: από 'Color SKU' ή 'Our Code'
    color_sku_col_stock = "Color SKU" if "Color SKU" in stock.columns else find_col(stock, ["color","sku"])
    our_code_col_stock = "Our Code" if "Our Code" in stock.columns else None
    if color_sku_col_stock:
        stock["Our Code"] = stock[color_sku_col_stock].apply(clean_our_code)
    elif our_code_col_stock:
        stock["Our Code"] = stock[our_code_col_stock].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' or 'Our Code'."); st.stop()

    # Size από Variant Values ή Size
    vv_col_stock = "Variant Values" if "Variant Values" in stock.columns else find_any_col(stock, [["variant","]()]()
