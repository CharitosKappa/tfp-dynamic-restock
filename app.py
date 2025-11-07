# app.py
# Dynamic Restock v12 – Streamlit app (robust Color)
# Vendor & Vendor Code: από STOCK
# Color: 1) από STOCK (Color / Χρώμα / parsing "Χρώμα: ..."), 2) fallback από SALES ανά Variant SKU
# Sales: μόνο για ποσότητες
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st
from collections import Counter

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Color κυρίως από STOCK (Color/Χρώμα ή parsing), fallback από SALES ανά Variant SKU")

# ---------------- Helpers ----------------
def to_int_safe(x):
    try:
        if pd.isna(x): return 0
        return int(float(str(x).strip()))
    except Exception:
        return 0

def to_float_safe(x):
    try:
        if pd.isna(x): return 0.0
        s = str(x)
        # αν έρχεται με ευρωπαϊκή μορφή "1.234,56"
        if re.search(r"\d[.,]\d{3}[.,]\d{2}$", s):
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        try: return float(x)
        except Exception: return 0.0

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

# --- πολύ ανεκτικό parsing χρώματος από stock κείμενα ---
COLOR_PATTERNS = [
    r"(?:Χρώμα|ΧΡΩΜΑ|Color|Colour)\s*[:：\-–—]?\s*(.+?)\s*(?=(?:Μεγ|Sizes?|Size|Taille|Μέγεθ|Νούμερο|,|;|\||/|$))",
    # με παρενθέσεις στο τέλος πχ " ... (Χρώμα: Ταμπά) "
    r"\((?:\s*(?:Χρώμα|ΧΡΩΜΑ|Color|Colour)\s*[:：\-–—]?\s*)(.+?)\)",
]

def extract_color_from_stock_text(text):
    if pd.isna(text): return None
    s = re.sub(r"\s+", " ", str(text)).strip()
    for pat in COLOR_PATTERNS:
        m = re.search(pat, s, flags=re.IGNORECASE)
        if m:
            c = m.group(1).strip().strip(' "\'“”‘’').rstrip(",;|/")
            # καθάρισε καταλήξεις τύπου "Ταμπά 36"
            c = re.sub(r"\s*(?:EU)?\d{2}\b.*$", "", c).strip()
            if c: return c
    return None

def extract_color_from_sales_line(text):
    """
    Από SALES γραμμή τύπου "[###########] ... (Μαύρο, L/XL)" -> "Μαύρο".
    Επιλέγει την παρένθεση όπου το 2ο κομμάτι μοιάζει με size, αλλιώς παίρνει το πρώτο κομμάτι μιας παρένθεσης με κόμμα.
    """
    if pd.isna(text): return None
    s = str(text)
    parens = re.findall(r"\(([^)]*)\)", s)
    if not parens: return None
    SIZE_HINT_RE = re.compile(
        r"\b(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|ONE\s*SIZE|ONESIZE|OS|EU\s?\d{2}|[3-5]\d(?:/[3-5]\d)?|[A-Z]/[A-Z])\b",
        flags=re.IGNORECASE
    )
    cands = [p for p in parens if "," in p]
    for p in cands:
        left, right = p.split(",", 1)
        if SIZE_HINT_RE.search(right) or re.search(r"\d|/", right):
            color = left.strip().strip(' "\'“”‘’')
            if color: return color
    if cands:
        return cands[0].split(",",1)[0].strip().strip(' "\'“”‘’') or None
    return None

def extract_variant_sku_from_text(text):
    """11ψήφιο από '[###########]' ή σκέτο 11ψήφιο."""
    if pd.isna(text): return None
    s = str(text)
    m = re.search(r"\[(\d{11})\]", s)
    if m: return m.group(1)
    m = re.search(r"(^|\D)(\d{11})(\D|$)", s)
    return m.group(2) if m else None

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
    """Find column whose normalized name contains ALL tokens, none of excluded."""
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

def coalesce(*vals):
    for v in vals:
        if pd.notna(v) and str(v).strip() != "":
            return v
    return np.nan

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

    # ---------- STOCK parsing (authoritative Vendor/Vendor Code/Color) ----------
    stock = stock_raw.copy()

    # Our Code (8-digit)
    color_sku_col_stock = "Color SKU" if "Color SKU" in stock.columns else find_col(stock, ["color","sku"])
    our_code_col_stock = "Our Code" if "Our Code" in stock.columns else None
    if color_sku_col_stock:
        stock["Our Code"] = stock[color_sku_col_stock].apply(clean_our_code)
    elif our_code_col_stock:
        stock["Our Code"] = stock[our_code_col_stock].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' or 'Our Code'."); st.stop()

    # Size
    vv_col_stock = None
    for tokens in [
        ["variant","values"],
        ["attribute","values"],
        ["variant","options"],
        ["options"],
        ["attributes"],
        ["χαρακτηρισ"],  # characteristics
        ["ιδιότη"],      # attributes (greek stems)
        ["παραλλαγ"],
        ["title"],
        ["περιγραφ"],    # description
    ]:
        vv_col_stock = find_col(stock, tokens)
        if vv_col_stock: break

    if vv_col_stock:
        stock["Size"] = stock[vv_col_stock].apply(extract_size_from_variant_values)
    elif "Size" in stock.columns:
        stock["Size"] = stock["Size"]
    else:
        st.error("Stock must have 'Variant Values' (ή παρόμοιο) ή μια στήλη 'Size'."); st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # On Hand / Forecasted
    onhand_col = "On Hand" if "On Hand" in stock.columns else find_any_col(stock, [["on","hand"]])
    forecast_col = "Forecasted" if "Forecasted" in stock.columns else find_any_col(stock, [["forecast"]])
    stock["On Hand"] = stock[onhand_col].apply(to_int_safe) if onhand_col else 0
    stock["Forecasted"] = stock[forecast_col].apply(to_int_safe) if forecast_col else 0

    # Vendor / Vendor Code
    brand_col = "Brand" if "Brand" in stock.columns else find_col(stock, ["brand"])
    vendor_code_col_stock = (
        "Vendor Code" if "Vendor Code" in stock.columns else
        ("Vendors/Vendor Product Code" if "Vendors/Vendor Product Code" in stock.columns else
         find_any_col(stock, [["vendors","vendor","product","code"],["vendor","product","code"],["vendorcode"]]))
    )

    # Color direct column (αποφεύγουμε Color SKU/Code)
    color_direct_col = None
    for tokens in [["χρώμα"],["color"],["colour"]]:
        cand = find_col(stock, tokens, exclude_tokens=["sku","code"])
        if cand:
            color_direct_col = cand
            break

    # FFill σε πιθανές πηγές κειμένου
    text_color_sources = []
    for cand in [vv_col_stock, "Variant Options", "Options", "Attributes", "Title", "Description"]:
        if cand in stock.columns and cand not in text_color_sources:
            text_color_sources.append(cand)
    for c in [brand_col, vendor_code_col_stock, color_direct_col] + text_color_sources:
        if c and c in stock.columns:
            stock[c] = stock[c].ffill()

    # --- Ανά γραμμή COLOR από STOCK: direct -> parsing από text sources ---
    if color_direct_col:
        stock["_Color_line"] = stock[color_direct_col]
    else:
        stock["_Color_line"] = np.nan

    if text_color_sources:
        def pick_color_row(row):
            v = row.get(color_direct_col) if color_direct_col else None
            if pd.notna(v) and str(v).strip() != "":
                return v
            for col in text_color_sources:
                c = extract_color_from_stock_text(row.get(col))
                if c: return c
            return np.nan
        stock["_Color_line"] = stock.apply(pick_color_row, axis=1)

    # --- Συγκεντρωτικοί χάρτες ανά Our Code ---
    tmp_map = pd.DataFrame({"Our Code": stock["Our Code"]})
    tmp_map["__VendorTmp"] = stock[brand_col] if brand_col else np.nan
    tmp_map["__VendorCodeTmp"] = stock[vendor_code_col_stock] if vendor_code_col_stock else np.nan
    tmp_map["__ColorTmp"] = stock["_Color_line"]

    stock_info_map = (
        tmp_map.groupby("Our Code", as_index=False)
               .agg({"__VendorTmp": mode_non_null,
                     "__VendorCodeTmp": mode_non_null,
                     "__ColorTmp": mode_non_null})
               .rename(columns={"__VendorTmp":"Vendor_from_stock",
                                "__VendorCodeTmp":"Vendor Code_from_stock",
                                "__ColorTmp":"Color_from_stock"})
    )

    # Variant SKU & stock levels per variant
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # ---------- SALES parsing (μόνο quantities + fallback color by Variant) ----------
    sales = sales_raw.copy()

    # Βρες πιθανή στήλη με γραμμές "[###########] ... (Color, Size)"
    sales_line_col = None
    for c in sales.columns:
        try:
            s = sales[c].astype(str)
            if s.str.contains(r"\(", regex=True).any():
                sales_line_col = c; break
        except Exception:
            pass

    color_by_variant = pd.DataFrame(columns=["Variant SKU","Color_from_sales_by_variant"])

    if sales_line_col:
        sl = sales[sales_line_col].astype(str)
        sales["Variant SKU (from line)"] = sl.apply(extract_variant_sku_from_text)
        sales["Color (from line)"] = sl.apply(extract_color_from_sales_line)
        color_by_variant = (
            sales.dropna(subset=["Variant SKU (from line)","Color (from line)"])
                 .groupby("Variant SKU (from line)", as_index=False)["Color (from line)"]
                 .agg(lambda s: Counter([x for x in s if pd.notna(x) and str(x).strip()!=""]).most_common(1)[0][0])
                 .rename(columns={"Variant SKU (from line)":"Variant SKU",
                                  "Color (from line)":"Color_from_sales_by_variant"})
        )

    # Πωλήσεις (για targets)
    sku_col_qty = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d{11}\]").any():
                sku_col_qty = c; break
        except Exception:
            pass
    if sku_col_qty is None:
        for c in sales.columns:
            try:
                if sales[c].astype(str).str.fullmatch(r"\d{11}").any():
                    sku_col_qty = c; break
            except Exception:
                pass

    total_col = "Total" if "Total" in sales.columns else find_col(sales, ["total"])
    if sku_col_qty is not None and total_col is not None:
        sales["Variant SKU"] = sales[sku_col_qty].astype(str).str.extract(r"\[(\d{11})\]").iloc[:,0]
        mask_no_br = sales["Variant SKU"].isna() & sales[sku_col_qty].astype(str).str.fullmatch(r"\d{11}")
        sales.loc[mask_no_br, "Variant SKU"] = sales.loc[mask_no_br, sku_col_qty].astype(str)
        sales["Qty Ordered"] = sales[total_col].apply(to_int_safe)

        sales_by_variant = (
            sales.dropna(subset=["Variant SKU"])
                 .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
        )
        sales_by_variant["Our Code"] = sales_by_variant["Variant SKU"].str.slice(0,8)
        sales_by_color = (
            sales_by_variant.groupby("Our Code", as_index=False)["Qty Ordered"].sum()
                            .rename(columns={"Qty Ordered":"Sales Color Total"})
        )
    else:
        sales_by_variant = pd.DataFrame(columns=["Variant SKU","Qty Ordered"])
        sales_by_color = pd.DataFrame(columns=["Our Code","Sales Color Total"])

    # ---------- Merge ----------
    df = stock_grp.copy()
    # Vendor/Code/Color από STOCK
    df = df.merge(stock_info_map, on="Our Code", how="left")

    # fallback Color από SALES ανά Variant (μόνο όπου λείπει από STOCK)
    if not color_by_variant.empty:
        df = df.merge(color_by_variant, on="Variant SKU", how="left")
    else:
        df["Color_from_sales_by_variant"] = np.nan

    df["Color"] = df.apply(lambda r: coalesce(r.get("Color_from_stock"),
                                              r.get("Color_from_sales_by_variant")), axis=1)

    # Join sales quantities
    if not sales_by_variant.empty:
        df = df.merge(sales_by_variant[["Variant SKU","Qty Ordered"]], on="Variant SKU", how="left")
    if
