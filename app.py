# app.py
# Dynamic Restock v12 – Streamlit web app (with Vendor, Vendor Code, Vendor Color fixed)
# Requirements: streamlit, pandas, numpy, openpyxl
# Run: streamlit run app.py

import io
import re
import math
import numpy as np
import pandas as pd
import streamlit as st

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Upload 2 Excel files (Stock + Sales) → get dynamic restock recommendations.")

# ---------------- Helpers ----------------
def to_int_safe(x):
    try:
        if pd.isna(x):
            return 0
        return int(float(str(x).strip()))
    except Exception:
        return 0

def clean_our_code(x):
    """Normalize to 8-digit string (strip .0, keep digits only)."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    if len(s) == 0:
        return None
    if len(s) < 8:
        s = s.zfill(8)
    elif len(s) > 8:
        s = s[:8]
    return s

def extract_size_from_variant_values(text):
    """Extract EU size 36–42 from 'Variant Values' or free text."""
    if pd.isna(text):
        return None
    s = str(text)
    m = re.search(r"(3[6-9]|4[0-2])\b", s)
    if m:
        return int(m.group(1))
    return None

def extract_color_from_variant_values(text):
    """Extract color after 'Χρώμα:' or 'Color:' from 'Variant Values' text."""
    if pd.isna(text):
        return None
    s = str(text)
    m = re.search(r"(?:Χρώμα|Color)\s*:\s*([^|\n\r]+)", s, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None

def build_variant_sku(our_code8, size):
    """11-digit SKU: OurCode(8) + (Size-34).zfill(3)."""
    if our_code8 is None or pd.isna(size):
        return None
    suffix = str(int(size) - 34).zfill(3)
    return f"{our_code8}{suffix}"

def base_target_for_size(size):
    """Base Target rules."""
    try:
        s = int(size)
    except Exception:
        return 0
    if s in (38, 39):
        return 6
    if s in (37, 40):
        return 4
    if s == 41:
        return 2
    if s in (36, 42):
        return 1
    return 0

def clip(x, lo, hi):
    try:
        return max(lo, min(hi, x))
    except Exception:
        return lo

def pick_existing_col(df, candidates):
    """Return first existing column name from candidates list, else None."""
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ---------------- Sidebar Inputs ----------------
st.sidebar.header("⚙️ Settings")
stock_sheet = st.sidebar.text_input("Stock sheet name", value="Sheet1")
sales_sheet = st.sidebar.text_input("Sales sheet name", value="Sales Analysis")

st.sidebar.markdown("---")
st.sidebar.caption("Expected columns (auto-fallbacks exist):")
st.sidebar.code("""Stock:
- 'Variant Values' includes size/color info
- 'Color SKU' (8-digit), 'Vendor', 'Vendor Code', 'Vendor Color'
- 'On Hand', 'Forecasted'

Sales:
- A column containing text like '... [12345678901]' (SKU in brackets)
- 'Total' with quantity
""")

# ---------------- File Upload ----------------
col1, col2 = st.columns(2)
with col1:
    stock_file = st.file_uploader("📂 Upload STOCK Excel", type=["xlsx", "xls"])
with col2:
    sales_file = st.file_uploader("📂 Upload SALES Excel", type=["xlsx", "xls"])

run_btn = st.button("🚀 Run Dynamic Restock")

if run_btn:
    if not stock_file or not sales_file:
        st.error("Please upload both STOCK and SALES files.")
        st.stop()

    # ---------- 1) Read inputs ----------
    try:
        stock_raw = pd.read_excel(stock_file, sheet_name=stock_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read STOCK sheet '{stock_sheet}': {e}")
        st.stop()

    try:
        sales_raw = pd.read_excel(sales_file, sheet_name=sales_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read SALES sheet '{sales_sheet}': {e}")
        st.stop()

    # ---------- 2) Clean Stock ----------
    stock = stock_raw.copy()

    # Detect vendor-related columns (robust to variants)
    vendor_col = pick_existing_col(stock, ["Vendor", "vendor", "Manufacturer"])
    vendor_code_col = pick_existing_col(stock, ["Vendor Code", "VendorCode", "vendor_code", "Manufacturer Code"])
    vendor_color_col = pick_existing_col(stock, ["Vendor Color", "VendorColour", "Vendor colour", "Χρώμα", "Color", "colour"])

    # Forward-fill the known columns to propagate parent rows into size rows
    for c in [vendor_col, vendor_code_col, vendor_color_col, "Color SKU"]:
        if c and c in stock.columns:
            stock[c] = stock[c].ffill()

    # Ensure 'Vendor' and 'Vendor Code' columns exist
    if not vendor_col:
        stock["Vendor"] = None
        vendor_col = "Vendor"
    if not vendor_code_col:
        stock["Vendor Code"] = None
        vendor_code_col = "Vendor Code"

    # Ensure 'Vendor Color' column exists and is filled
    if "Vendor Color" not in stock.columns:
        stock["Vendor Color"] = None
        if vendor_color_col and vendor_color_col != "Vendor Color":
            stock["Vendor Color"] = stock[vendor_color_col]
        # fill from Variant Values parsing where missing
        if "Variant Values" in stock.columns:
            vv_color = stock["Variant Values"].apply(extract_color_from_variant_values)
            stock["Vendor Color"] = stock["Vendor Color"].fillna(vv_color)
        stock["Vendor Color"] = stock["Vendor Color"].ffill()
    else:
        # top up Vendor Color from Variant Values if empty
        if "Variant Values" in stock.columns:
            vv_color = stock["Variant Values"].apply(extract_color_from_variant_values)
            stock["Vendor Color"] = stock["Vendor Color"].fillna(vv_color)
        stock["Vendor Color"] = stock["Vendor Color"].ffill()

    # Sizes
    if "Variant Values" in stock.columns:
        stock["Size"] = stock["Variant Values"].apply(extract_size_from_variant_values)
    elif "Size" not in stock.columns:
        st.error("Stock sheet must contain 'Variant Values' or a usable 'Size' column.")
        st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36, 37, 38, 39, 40, 41, 42])].copy()

    # Our Code (Color SKU) normalize → 8-digit
    if "Color SKU" in stock.columns:
        stock["Our Code"] = stock["Color SKU"].apply(clean_our_code)
    else:
        st.warning("Column 'Color SKU' not found in Stock. Attempting to derive from 'Our Code' or similar...")
        if "Our Code" in stock.columns:
            stock["Our Code"] = stock["Our Code"].apply(clean_our_code)
        else:
            st.error("Need either 'Color SKU' or 'Our Code' in Stock data.")
            st.stop()

    # On Hand / Forecasted to int
    for c in ["On Hand", "Forecasted"]:
        if c in stock.columns:
            stock[c] = stock[c].apply(to_int_safe)
        else:
            stock[c] = 0

    # Build Variant SKU (11-digit)
    stock["Variant SKU"] = stock.apply(
        lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1
    )

    # Groupby to one row per variant
    stock_grp = (
        stock.groupby(["Our Code", "Variant SKU", "Size"], as_index=False)
        .agg({
            "Vendor": "first",
            "Vendor Code": "first",
            "Vendor Color": "first",
            "On Hand": "max",
            "Forecasted": "max",
        })
    )

    # ---------- 3) Clean Sales ----------
    sales = sales_raw.copy()

    # Find the column containing '[123456...]'
    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d+\]").any():
                sku_col = c
                break
        except Exception:
            continue
    if sku_col is None:
        if "Unnamed: 0" in sales.columns:
            sku_col = "Unnamed: 0"
        else:
            st.error("Could not find a sales column containing '[<digits>]'.")
            st.stop()

    # Extract Variant SKU from square brackets
    sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d+)\]").iloc[:, 0]

    # Qty Ordered from 'Total'
    if "Total" not in sales.columns:
        st.error("Sales sheet must contain a 'Total' column with ordered quantities.")
        st.stop()
    sales["Qty Ordered"] = sales["Total"].apply(to_int_safe)

    # Sum per variant
    sales_by_variant = (
        sales.dropna(subset=["Variant SKU"])
        .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
    )

    # Sales by color (first 8 of Variant SKU)
    sales_by_variant["ColorSKU"] = sales_by_variant["Variant SKU"].str.slice(0, 8)
    sales_by_color = (
        sales_by_variant.groupby("ColorSKU", as_index=False)["Qty Ordered"].sum()
        .rename(columns={"Qty Ordered": "Sales Color Total"})
    )

    # ---------- 4) Merge Stock + Sales ----------
    df = stock_grp.merge(sales_by_variant, on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)

    df = df.merge(
        sales_by_color,
        left_on="Our Code",
        right_on="ColorSKU",
        how="left"
    )
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)
    df = df.drop(columns=["ColorSKU"], errors="ignore")

    # ---------- 5) Targets ----------
    # 5.1 Base Target
    df["Base Target"] = df["Size"].apply(base_target_for_size)

    # 5.2 Global Multiplier = clip(SalesColorTotal / BaseSumColor, 0.5, 5.0)
    base_sum_per_color = (
        df.groupby("Our Code", as_index=False)["Base Target"].sum()
        .rename(columns={"Base Target": "BaseSumColor"})
    )
    df = df.merge(base_sum_per_color, on="Our Code", how="left")
    df["BaseSumColor"] = df["BaseSumColor"].replace(0, np.nan)
    df["GlobalMult"] = (df["Sales Color Total"] / df["BaseSumColor"])
    df["GlobalMult"] = df["GlobalMult"].fillna(0)
    df["GlobalMult"] = df["GlobalMult"].apply(lambda x: clip(x, 0.5, 5.0))

    # 5.3 Size Multiplier
    avg_sales_per_color = (
        df.groupby("Our Code", as_index=False)["Qty Ordered"].mean()
        .rename(columns={"Qty Ordered": "AvgSalesPerSize"})
    )
    df = df.merge(avg_sales_per_color, on="Our Code", how="left")
    df["AvgSalesPerSize"] = df["AvgSalesPerSize"].replace(0, np.nan)

    def compute_size_mult(row):
        if row["Sales Color Total"] == 0:
            return 0.0
        q = row["Qty Ordered"]
        avg = row["AvgSalesPerSize"]
        if pd.isna(avg) or avg == 0:
            return 1.0
        val = q / avg
        return clip(val, 0.5, 2.0)

    df["SizeMult"] = df.apply(compute_size_mult, axis=1)

    # 5.4 Adjusted Target (aggressive + refined rule)
    df["AdjRaw"] = df["Base Target"] * df["GlobalMult"] * df["SizeMult"]

    # Boost +20% if Sales > 2 × BaseTarget
    df["AdjRaw"] = np.where(
        df["Qty Ordered"] > (2 * df["Base Target"]),
        df["AdjRaw"] * 1.2,
        df["AdjRaw"]
    )

    # ceil and floor by Base Target
    df["AdjCeil"] = df["AdjRaw"].apply(lambda x: int(math.ceil(x)) if pd.notna(x) else 0)
    df["Adjusted Target"] = df[["AdjCeil", "Base Target"]].max(axis=1)

    # Zero-sales rule: if Sales == 0 → Adjusted = 0
    df.loc[df["Qty Ordered"] == 0, "Adjusted Target"] = 0

    # Core sizes refinement: 38–39–40 if color sells but zero stock/forecast/sales on this variant
    core_mask = (
        (df["Qty Ordered"] == 0) &
        (df["On Hand"] == 0) &
        (df["Forecasted"] == 0) &
        (df["Size"].isin([38, 39, 40])) &
        (df["Sales Color Total"] > 0)
    )
    df.loc[core_mask, "Adjusted Target"] = df.loc[core_mask, "Base Target"]

    # ---------- 6) Restock Quantity ----------
    df["Restock Quantity"] = (df["Adjusted Target"] - df["Forecasted"]).clip(lower=0)

    # ---------- 7) Final file ----------
    final_cols = [
        "Vendor", "Vendor Code", "Vendor Color",     # first 3 columns (fixed)
        "Our Code", "Variant SKU", "Size",
        "On Hand", "Forecasted", "Qty Ordered", "Sales Color Total",
        "Base Target", "GlobalMult", "SizeMult", "Adjusted Target",
        "Restock Quantity"
    ]
    for c in final_cols:
        if c not in df.columns:
            df[c] = np.nan

    out = (
        df[final_cols]
        .drop_duplicates(subset=["Variant SKU", "Size"], keep="first")
        .sort_values(["Our Code", "Size"])
        .reset_index(drop=True)
    )

    st.success("Done! Preview below ↓")
    st.dataframe(out, use_container_width=True)

    # Download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Restock v12")
    st.download_button(
        label="⬇️ Download dynamic_restock_order_v12.xlsx",
        data=buffer.getvalue(),
        file_name="dynamic_restock_order_v12.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- Help panel ----------
with st.expander("ℹ️ How to run"):
    st.markdown("""
1) Install Python 3.10+  
2) `pip install streamlit pandas numpy openpyxl`  
3) Save as `app.py`  
4) Run: `streamlit run app.py`  
5) Upload:
   - **Stock**: e.g. `Product Variant (2).xlsx` (sheet `Sheet1`)
   - **Sales**: e.g. `Pivot Sales Analysis (3).xlsx` (sheet `Sales Analysis`)
6) Adjust sheet names from the sidebar if needed.
""")
