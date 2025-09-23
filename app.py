# app.py
# Dynamic Restock v12 – Streamlit web app (robust Vendor/Vendor Code/Vendor Color)
# Requirements: streamlit, pandas, numpy, openpyxl
# Run: streamlit run app.py

import io
import re
import math
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Upload 2 Excel files (Stock + Sales) → get dynamic restock recommendations.")

# ---------------- Helpers ----------------
def to_int_safe(x):
    try:
        if pd.isna(x): return 0
        return int(float(str(x).strip()))
    except Exception:
        return 0

def clean_our_code(x):
    if pd.isna(x): return None
    s = str(x).strip()
    if s.endswith(".0"): s = s[:-2]
    s = re.sub(r"\D", "", s)
    if not s: return None
    if len(s) < 8: s = s.zfill(8)
    elif len(s) > 8: s = s[:8]
    return s

def extract_size_from_variant_values(text):
    if pd.isna(text): return None
    m = re.search(r"(3[6-9]|4[0-2])\b", str(text))
    return int(m.group(1)) if m else None

def extract_color_from_variant_values(text):
    if pd.isna(text): return None
    s = str(text)
    m = re.search(r"(?:Χρώμα|Color)\s*:\s*([^|\n\r]+)", s, flags=re.IGNORECASE)
    return m.group(1).strip() if m else None

def build_variant_sku(our_code8, size):
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

def pick_existing_col(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    return None

# ---------------- Sidebar ----------------
st.sidebar.header("⚙️ Settings")
stock_sheet = st.sidebar.text_input("Stock sheet name", value="Sheet1")
sales_sheet = st.sidebar.text_input("Sales sheet name", value="Sales Analysis")

st.sidebar.markdown("---")
st.sidebar.caption("We’ll try to auto-detect Vendor/Vendor Code/Vendor Color.")
st.sidebar.code("""Stock:
- 'Variant Values' (size/color text) OR 'Size'
- 'Color SKU' (or 'Our Code')
- Optional: 'Vendor', 'Vendor Code', 'Vendor Color' (or synonyms)
- 'On Hand', 'Forecasted'

Sales:
- Column with text like '... [12345678901]' (SKU in brackets)
- 'Total' = quantity
""")

# ---------------- Uploads ----------------
c1, c2 = st.columns(2)
with c1: stock_file = st.file_uploader("📂 Upload STOCK Excel", type=["xlsx","xls"])
with c2: sales_file = st.file_uploader("📂 Upload SALES Excel", type=["xlsx","xls"])
run_btn = st.button("🚀 Run Dynamic Restock")

if run_btn:
    if not stock_file or not sales_file:
        st.error("Please upload both STOCK and SALES files.")
        st.stop()

    # 1) Read
    try:
        stock_raw = pd.read_excel(stock_file, sheet_name=stock_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read STOCK sheet '{stock_sheet}': {e}"); st.stop()

    try:
        sales_raw = pd.read_excel(sales_file, sheet_name=sales_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read SALES sheet '{sales_sheet}': {e}"); st.stop()

    # 2) Clean Stock
    stock = stock_raw.copy()

    # Detect a variety of vendor column names
    vendor_col = pick_existing_col(stock, ["Vendor","vendor","Manufacturer","Brand","Supplier"])
    vendor_code_col = pick_existing_col(stock, ["Vendor Code","VendorCode","vendor_code","Manufacturer Code","Brand Code","Supplier Code","Code"])
    vendor_color_col_hint = pick_existing_col(stock, ["Vendor Color","VendorColour","Vendor colour","Χρώμα Vendor","ΧρώμαVendor"])
    color_generic_col = pick_existing_col(stock, ["Color","Colour","Χρώμα","Vendor Colour","VendorColor","Vendor Colour"])

    # Forward fill base columns (useful for parent rows → size rows)
    for c in [vendor_col, vendor_code_col, vendor_color_col_hint, color_generic_col, "Color SKU"]:
        if c and c in stock.columns: stock[c] = stock[c].ffill()

    # Guarantee canonical columns exist
    if not vendor_col:
        stock["Vendor"] = None; vendor_col = "Vendor"
    if not vendor_code_col:
        stock["Vendor Code"] = None; vendor_code_col = "Vendor Code"
    if "Vendor Color" not in stock.columns:
        stock["Vendor Color"] = None
        # seed from explicit vendor-color-like col or from generic Color
        seed_col = vendor_color_col_hint or color_generic_col
        if seed_col and seed_col != "Vendor Color":
            stock["Vendor Color"] = stock[seed_col]
    # top-up Vendor Color from Variant Values, then ffill
    if "Variant Values" in stock.columns:
        vv_color = stock["Variant Values"].apply(extract_color_from_variant_values)
        stock["Vendor Color"] = stock["Vendor Color"].fillna(vv_color)
    stock["Vendor Color"] = stock["Vendor Color"].ffill()

    # Size detection
    if "Variant Values" in stock.columns:
        stock["Size"] = stock["Variant Values"].apply(extract_size_from_variant_values)
    elif "Size" not in stock.columns:
        st.error("Stock sheet must have 'Variant Values' or a usable 'Size' column."); st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # Our Code normalization
    if "Color SKU" in stock.columns:
        stock["Our Code"] = stock["Color SKU"].apply(clean_our_code)
    elif "Our Code" in stock.columns:
        stock["Our Code"] = stock["Our Code"].apply(clean_our_code)
    else:
        st.error("Need either 'Color SKU' or 'Our Code' in Stock data."); st.stop()

    # On Hand / Forecasted to int
    for c in ["On Hand","Forecasted"]:
        stock[c] = stock[c].apply(to_int_safe) if c in stock.columns else 0

    # Variant SKU
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)

    # Build a **color-level vendor map** so we never lose vendor fields
    vendor_map = (
        stock.groupby("Our Code", as_index=False)
             .agg({
                 vendor_col: "first",
                 vendor_code_col: "first",
                 "Vendor Color": "first"
             })
    )
    vendor_map = vendor_map.rename(columns={
        vendor_col: "Vendor",
        vendor_code_col: "Vendor Code"
    })

    # Group to one row per variant (we won't rely on vendor fields here anymore)
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # 3) Clean Sales
    sales = sales_raw.copy()

    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d+\]").any():
                sku_col = c; break
        except Exception: pass
    if sku_col is None:
        if "Unnamed: 0" in sales.columns: sku_col="Unnamed: 0"
        else:
            st.error("Could not find a sales column containing '[<digits>]'."); st.stop()

    sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d+)\]").iloc[:,0]
    if "Total" not in sales.columns:
        st.error("Sales sheet must contain a 'Total' column with ordered quantities."); st.stop()
    sales["Qty Ordered"] = sales["Total"].apply(to_int_safe)

    sales_by_variant = (
        sales.dropna(subset=["Variant SKU"])
             .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
    )
    sales_by_variant["ColorSKU"] = sales_by_variant["Variant SKU"].str.slice(0,8)
    sales_by_color = (
        sales_by_variant.groupby("ColorSKU", as_index=False)["Qty Ordered"].sum()
                        .rename(columns={"Qty Ordered":"Sales Color Total"})
    )

    # 4) Merge
    df = stock_grp.merge(sales_by_variant, on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)
    df["Our Code"] = df["Variant SKU"].str.slice(0,8)  # ensure present even if coming only from variant

    df = df.merge(sales_by_color, left_on="Our Code", right_on="ColorSKU", how="left").drop(columns=["ColorSKU"], errors="ignore")
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # **Bring back vendor fields from color-level map**
    df = df.merge(vendor_map, on="Our Code", how="left")

    # 5) Targets
    df["Base Target"] = df["Size"].apply(base_target_for_size)

    base_sum_per_color = (
        df.groupby("Our Code", as_index=False)["Base Target"].sum()
          .rename(columns={"Base Target":"BaseSumColor"})
    )
    df = df.merge(base_sum_per_color, on="Our Code", how="left")
    df["BaseSumColor"] = df["BaseSumColor"].replace(0, np.nan)
    df["GlobalMult"] = (df["Sales Color Total"] / df["BaseSumColor"]).fillna(0).apply(lambda x: clip(x, 0.5, 5.0))

    avg_sales_per_color = (
        df.groupby("Our Code", as_index=False)["Qty Ordered"].mean()
          .rename(columns={"Qty Ordered":"AvgSalesPerSize"})
    )
    df = df.merge(avg_sales_per_color, on="Our Code", how="left")
    df["AvgSalesPerSize"] = df["AvgSalesPerSize"].replace(0, np.nan)

    def compute_size_mult(row):
        if row["Sales Color Total"] == 0: return 0.0
        avg = row["AvgSalesPerSize"]
        if pd.isna(avg) or avg == 0: return 1.0
        return clip(row["Qty Ordered"]/avg, 0.5, 2.0)

    df["SizeMult"] = df.apply(compute_size_mult, axis=1)

    df["AdjRaw"] = df["Base Target"] * df["GlobalMult"] * df["SizeMult"]
    df["AdjRaw"] = np.where(df["Qty Ordered"] > (2*df["Base Target"]), df["AdjRaw"]*1.2, df["AdjRaw"])
    df["AdjCeil"] = df["AdjRaw"].apply(lambda x: int(math.ceil(x)) if pd.notna(x) else 0)
    df["Adjusted Target"] = df[["AdjCeil","Base Target"]].max(axis=1)

    zero_sales_mask = (df["Qty Ordered"] == 0)
    df.loc[zero_sales_mask, "Adjusted Target"] = 0

    core_mask = (
        (df["Qty Ordered"] == 0) &
        (df["On Hand"] == 0) &
        (df["Forecasted"] == 0) &
        (df["Size"].isin([38,39,40])) &
        (df["Sales Color Total"] > 0)
    )
    df.loc[core_mask, "Adjusted Target"] = df.loc[core_mask, "Base Target"]

    # 6) Restock
    df["Restock Quantity"] = (df["Adjusted Target"] - df["Forecasted"]).clip(lower=0)

    # 7) Final file
    final_cols = [
        "Vendor", "Vendor Code", "Vendor Color",  # first 3 columns (fixed)
        "Our Code", "Variant SKU", "Size",
        "On Hand", "Forecasted", "Qty Ordered", "Sales Color Total",
        "Base Target", "GlobalMult", "SizeMult", "Adjusted Target",
        "Restock Quantity"
    ]
    for c in final_cols:
        if c not in df.columns: df[c] = np.nan

    out = (
        df[final_cols]
        .drop_duplicates(subset=["Variant SKU","Size"], keep="first")
        .sort_values(["Our Code","Size"])
        .reset_index(drop=True)
    )

    # Quick diagnostics (to βεβαιωθείς ότι γεμίζουν)
    with st.expander("🔎 Diagnostics – Non-null counts"):
        diag = out[["Vendor","Vendor Code","Vendor Color"]].notna().sum()
        st.write(diag.to_frame("non_null").T)

    st.success("Done! Preview below ↓")
    st.dataframe(out, use_container_width=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Restock v12")
    st.download_button(
        "⬇️ Download dynamic_restock_order_v12.xlsx",
        data=buf.getvalue(),
        file_name="dynamic_restock_order_v12.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with st.expander("ℹ️ How to run"):
    st.markdown("""
1) Python 3.10+  
2) `pip install streamlit pandas numpy openpyxl`  
3) `streamlit run app.py`  
4) Upload Stock (sheet `Sheet1`) + Sales (sheet `Sales Analysis`)  
5) Αν τα vendor πεδία έχουν άλλη ονομασία, το app τα αναγνωρίζει από συνώνυμα.  
""")
