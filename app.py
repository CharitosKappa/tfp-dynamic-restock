# app.py
# Dynamic Restock v12 – Streamlit web app
# Sales-driven mapping for Vendor / Vendor Code / Vendor Color
# Requirements: streamlit, pandas, numpy, openpyxl

import io
import re
import math
import numpy as np
import pandas as pd
import streamlit as st

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Upload 2 Excel files (Stock + Sales) → dynamic restock recommendations + vendor mapping from Sales.")

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

def coalesce(pref, fallback):
    """Prefer pref unless it's NaN/None/empty; else use fallback."""
    if pd.notna(pref) and str(pref).strip() != "":
        return pref
    return fallback

# ---------------- Sidebar Inputs ----------------
st.sidebar.header("⚙️ Settings")
stock_sheet = st.sidebar.text_input("Stock sheet name", value="Sheet1")
sales_sheet = st.sidebar.text_input("Sales sheet name", value="Sales Analysis")

st.sidebar.markdown("---")
st.sidebar.caption("Vendor fields are mapped from the Sales file (Display Name, Vendor Product Code, Variant Values → Χρώμα:).")
st.sidebar.code("""Sales columns used:
- Vendors/Display Name        → Vendor
- Vendors/Vendor Product Code → Vendor Code
- Variant Values (Χρώμα: ...) → Vendor Color
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

    # ---------- 2) Clean Stock (only what's needed for sizes/Our Code/stock levels) ----------
    stock = stock_raw.copy()

    # Size detection
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
    elif "Our Code" in stock.columns:
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

    # Variant SKU (11-digit)
    stock["Variant SKU"] = stock.apply(
        lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1
    )

    # One row per variant (keep only stock levels here)
    stock_grp = (
        stock.groupby(["Our Code", "Variant SKU", "Size"], as_index=False)
        .agg({
            "On Hand": "max",
            "Forecasted": "max",
        })
    )

    # ---------- 3) Clean Sales + build maps ----------
    sales = sales_raw.copy()

    # Find the column containing '[123456...]' to extract Variant SKU
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

    # Extract Variant SKU from square brackets and Our Code (first 8 digits)
    sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d+)\]").iloc[:, 0]
    sales["Our Code"] = sales["Variant SKU"].str.slice(0, 8)

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

    # ---------- Sales-driven Vendor mapping (as requested) ----------
    vendor_display_col = "Vendors/Display Name"
    vendor_code_col = "Vendors/Vendor Product Code"
    variant_values_col = "Variant Values"

    # Prepare a sales map per Our Code
    sales_map_cols = []
    if vendor_display_col in sales.columns:
        sales_map_cols.append(vendor_display_col)
    if vendor_code_col in sales.columns:
        sales_map_cols.append(vendor_code_col)
    if variant_values_col in sales.columns:
        sales["__SalesVendorColor"] = sales[variant_values_col].apply(extract_color_from_variant_values)
        sales_map_cols.append("__SalesVendorColor")

    # Build map only if we have at least one of the needed columns
    if sales_map_cols:
        sales_vendor_map = (
            sales.dropna(subset=["Our Code"])
                 .groupby("Our Code", as_index=False)[sales_map_cols].agg(lambda s: s.dropna().iloc[0] if s.dropna().size else np.nan)
        )
        # Rename to canonical export names
        rename_map = {}
        if vendor_display_col in sales_map_cols:
            rename_map[vendor_display_col] = "Vendor"
        if vendor_code_col in sales_map_cols:
            rename_map[vendor_code_col] = "Vendor Code"
        if "__SalesVendorColor" in sales_map_cols:
            rename_map["__SalesVendorColor"] = "Vendor Color"
        sales_vendor_map = sales_vendor_map.rename(columns=rename_map)
    else:
        # Empty map if none of the columns exist
        sales_vendor_map = pd.DataFrame(columns=["Our Code", "Vendor", "Vendor Code", "Vendor Color"])

    # ---------- 4) Merge Stock + Sales ----------
    df = stock_grp.merge(sales_by_variant, on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)

    df = df.merge(
        sales_by_color,
        left_on="Our Code",
        right_on="ColorSKU",
        how="left"
    ).drop(columns=["ColorSKU"], errors="ignore")
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # Merge Sales vendor map (preferred source)
    df = df.merge(sales_vendor_map, on="Our Code", how="left", suffixes=("", "_from_sales"))

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

    # ---------- 7) Finalize Vendor fields (ensure presence) ----------
    # If any of Vendor / Vendor Code / Vendor Color missing, create them
    for c in ["Vendor", "Vendor Code", "Vendor Color"]:
        if c not in df.columns:
            df[c] = np.nan

    # ---------- 8) Final file ----------
    final_cols = [
        "Vendor", "Vendor Code", "Vendor Color",  # first 3 columns from Sales mapping
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

    # Diagnostics panel
    with st.expander("🔎 Diagnostics – Non-null counts for vendor fields"):
        diag = out[["Vendor", "Vendor Code", "Vendor Color"]].notna().sum()
        st.write(diag.to_frame("non_null").T)

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

# ---------- Help ----------
with st.expander("ℹ️ Notes"):
    st.markdown("""
- **Vendor**, **Vendor Code**, **Vendor Color** στο export προέρχονται από το **Sales**:
  - `Vendors/Display Name` → Vendor  
  - `Vendors/Vendor Product Code` → Vendor Code  
  - `Variant Values` (εξαγωγή μετά το `Χρώμα:`) → Vendor Color  
- Το mapping γίνεται σε **color-level** (ανά `Our Code` = 8 πρώτα ψηφία του Variant SKU).
- Αν κάποια από τις παραπάνω στήλες λείπει στο Sales, η αντίστοιχη έξοδος θα είναι κενή.
""")
