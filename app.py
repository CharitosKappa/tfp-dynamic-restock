# app.py
# Dynamic Restock v12 – Streamlit app
# Explicit mapping (from SALES):
#   Our Code (export)  <-> Sales['Color SKU']  (8ψήφιο)
#   Vendor (export A)  <- Sales['Brand']
#   Vendor Code (B)    <- Sales['Vendors/Vendor Product Code']
#   Color (C)          <- κείμενο μετά το "Χρώμα:" από Sales['Variant Values']
# Υπόλοιπη λογική: targets & restock όπως στο v12
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v1", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v1")
st.caption("Upload Stock + Sales → dynamic restock. Vendor/Code/Color mapped directly from Sales via Color SKU ↔ Our Code.")

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
    """EU sizes 36–42 from free text (e.g. Variant Values)."""
    if pd.isna(text): return None
    m = re.search(r"(3[6-9]|4[0-2])\b", str(text))
    return int(m.group(1)) if m else None

def extract_color_after_keyword(text):
    """Return text after 'Χρώμα:' ή 'Color:' από Variant Values."""
    if pd.isna(text): return None
    s = str(text)
    m = re.search(r"(?:Χρώμα|Color)\s*:\s*([^|\n\r]+)", s, flags=re.IGNORECASE)
    return m.group(1).strip() if m else None

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
        st.error("Please upload both STOCK and SALES files.")
        st.stop()

    # ---------- Read ----------
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

    # ---------- STOCK essentials ----------
    stock = stock_raw.copy()

    # Size (από Variant Values ή υπάρχουσα Size)
    if "Variant Values" in stock.columns:
        stock["Size"] = stock["Variant Values"].apply(extract_size_from_variant_values)
    elif "Size" in stock.columns:
        stock["Size"] = stock["Size"]
    else:
        st.error("Stock must have 'Variant Values' or a 'Size' column.")
        st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # Our Code (8-digit) από Stock ('Color SKU' ή 'Our Code')
    if "Color SKU" in stock.columns:
        stock["Our Code"] = stock["Color SKU"].apply(clean_our_code)
    elif "Our Code" in stock.columns:
        stock["Our Code"] = stock["Our Code"].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' or 'Our Code'.")
        st.stop()

    # On Hand / Forecasted
    stock["On Hand"] = stock["On Hand"].apply(to_int_safe) if "On Hand" in stock.columns else 0
    stock["Forecasted"] = stock["Forecasted"].apply(to_int_safe) if "Forecasted" in stock.columns else 0

    # Variant SKU (11 digits)
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)

    # Συνοπτικός πίνακας αποθέματος ανά variant
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # ---------- SALES explicit mapping ----------
    sales = sales_raw.copy()

    # ΑΠΑΙΤΟΥΜΕΝΕΣ στήλες στο Sales:
    required_sales_cols = ["Color SKU", "Brand", "Vendors/Vendor Product Code", "Variant Values", "Total"]
    missing = [c for c in required_sales_cols if c not in sales.columns]
    if missing:
        st.error(f"Sales is missing required columns: {missing}")
        st.stop()

    # Κανονικοποίηση Our Code από Sales['Color SKU']
    sales["OurCode_from_sales"] = sales["Color SKU"].apply(clean_our_code)

    # Color από Variant Values μετά το 'Χρώμα:'
    sales["Color_from_sales"] = sales["Variant Values"].apply(extract_color_after_keyword)

    # Χτίζουμε color-level map από Sales (key = OurCode_from_sales)
    sales_map = (
        sales.dropna(subset=["OurCode_from_sales"])
             .groupby("OurCode_from_sales", as_index=False)
             .agg({
                 "Brand": lambda s: s.dropna().iloc[0] if s.dropna().size else np.nan,
                 "Vendors/Vendor Product Code": lambda s: s.dropna().iloc[0] if s.dropna().size else np.nan,
                 "Color_from_sales": lambda s: s.dropna().iloc[0] if s.dropna().size else np.nan,
             })
             .rename(columns={
                 "OurCode_from_sales": "Our Code",
                 "Brand": "Vendor",
                 "Vendors/Vendor Product Code": "Vendor Code",
                 "Color_from_sales": "Color",
             })
    )

    # Πωλήσεις ανά variant (θα βρούμε Variant SKU από οποιαδήποτε στήλη έχει [11ψήφιο])
    # Προσπαθούμε να εντοπίσουμε στήλη με '[\d+]'
    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d{11}\]").any():
                sku_col = c
                break
        except Exception:
            pass
    if sku_col is None:
        # fallback: καθαρό 11ψήφιο
        for c in sales.columns:
            if sales[c].astype(str).str.fullmatch(r"\d{11}").any():
                sku_col = c
                break

    if sku_col is not None:
        sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d{11})\]").iloc[:,0]
        mask_no_br = sales["Variant SKU"].isna() & sales[sku_col].astype(str).str.fullmatch(r"\d{11}")
        sales.loc[mask_no_br, "Variant SKU"] = sales.loc[mask_no_br, sku_col].astype(str)
        sales["Qty Ordered"] = sales["Total"].apply(to_int_safe)

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
        # Αν δεν βρούμε Variant SKU στο Sales, συνεχίζουμε με μηδενικές πωλήσεις
        sales_by_variant = pd.DataFrame(columns=["Variant SKU","Qty Ordered"])
        sales_by_color = pd.DataFrame(columns=["Our Code","Sales Color Total"])

    # ---------- Merge: φέρνουμε Vendor / Vendor Code / Color από Sales μέσω Our Code ----------
    df = stock_grp.merge(sales_map, on="Our Code", how="left")

    # Πωλήσεις ανά variant + ανά χρώμα (αν βρέθηκαν)
    if not sales_by_variant.empty:
        df = df.merge(sales_by_variant, on="Variant SKU", how="left")
    if "Qty Ordered" not in df.columns:
        df["Qty Ordered"] = 0
    if not sales_by_color.empty:
        df = df.merge(sales_by_color, on="Our Code", how="left")
    if "Sales Color Total" not in df.columns:
        df["Sales Color Total"] = 0

    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # ---------- Targets ----------
    df["Base Target"] = df["Size"].apply(base_target_for_size)

    base_sum_per_color = (
        df.groupby("Our Code", as_index=False)["Base Target"].sum()
          .rename(columns={"Base Target":"BaseSumColor"})
    )
    df = df.merge(base_sum_per_color, on="Our Code", how="left")
    df["BaseSumColor"] = df["BaseSumColor"].replace(0, np.nan)

    df["GlobalMult"] = (df["Sales Color Total"] / df["BaseSumColor"]).fillna(0)
    df["GlobalMult"] = df["GlobalMult"].apply(lambda x: clip(x, 0.5, 5.0))

    avg_sales_per_color = (
        df.groupby("Our Code", as_index=False)["Qty Ordered"].mean()
          .rename(columns={"Qty Ordered":"AvgSalesPerSize"})
    )
    df = df.merge(avg_sales_per_color, on="Our Code", how="left")
    df["AvgSalesPerSize"] = df["AvgSalesPerSize"].replace(0, np.nan)

    def compute_size_mult(row):
        if row["Sales Color Total"] == 0:
            return 0.0
        avg = row["AvgSalesPerSize"]
        if pd.isna(avg) or avg == 0:
            return 1.0
        return clip(row["Qty Ordered"]/avg, 0.5, 2.0)

    df["SizeMult"] = df.apply(compute_size_mult, axis=1)

    df["AdjRaw"] = df["Base Target"] * df["GlobalMult"] * df["SizeMult"]
    df["AdjRaw"] = np.where(df["Qty Ordered"] > (2*df["Base Target"]), df["AdjRaw"]*1.2, df["AdjRaw"])
    df["AdjCeil"] = df["AdjRaw"].apply(lambda x: int(math.ceil(x)) if pd.notna(x) else 0)
    df["Adjusted Target"] = df[["AdjCeil","Base Target"]].max(axis=1)

    # Zero-sales rule
    df.loc[df["Qty Ordered"] == 0, "Adjusted Target"] = 0

    # Core sizes refinement
    core_mask = (
        (df["Qty Ordered"] == 0) &
        (df["On Hand"] == 0) &
        (df["Forecasted"] == 0) &
        (df["Size"].isin([38,39,40])) &
        (df["Sales Color Total"] > 0)
    )
    df.loc[core_mask, "Adjusted Target"] = df.loc[core_mask, "Base Target"]

    # Restock
    df["Restock Quantity"] = (df["Adjusted Target"] - df["Forecasted"]).clip(lower=0)

    # ---------- Export ----------
    final_cols = [
        "Vendor", "Vendor Code", "Color",   # A, B, C (όπως ζητήθηκε)
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

    # ---------- Diagnostics ----------
    with st.expander("🔎 Diagnostics"):
        st.write({
            "Sales required columns present": {c: (c in sales.columns) for c in ["Color SKU","Brand","Vendors/Vendor Product Code","Variant Values","Total"]},
            "Non-null counts (Vendor/Code/Color)": out[["Vendor","Vendor Code","Color"]].notna().sum().to_dict(),
        })
        st.write("Sample (first 10):", out[["Our Code","Vendor","Vendor Code","Color"]].head(10))

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
