# app.py
# Dynamic Restock v12 – Streamlit app
# A: Vendor      <- STOCK['Brand'] (ανά Our Code)
# B: Vendor Code <- STOCK['Vendor Code'] ή STOCK['Vendors/Vendor Product Code'] (ανά Our Code)
# C: Color       <- STOCK['Variant Values'] με regex "Χρώμα: ..." + forward-fill (ανά γραμμή/Variant)
# Sales μόνο για Qty/Targets
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st
from collections import Counter

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Vendor & Vendor Code από STOCK • Color από STOCK (Variant Values → 'Χρώμα:' → forward fill)")

# ---------------- Helpers ----------------
def to_int_safe(x):
    try:
        if pd.isna(x): return 0
        return int(float(str(x).strip()))
    except Exception:
        return 0

def clean_our_code(x):
    """Normalize to 8-digit numeric string (strip non-digits)."""
    if pd.isna(x): return None
    s = re.sub(r"\D", "", str(x).strip())
    if not s: return None
    return s[:8].zfill(8)

def extract_size_from_variant_values(text):
    """Detect EU size 36–42 in free text."""
    if pd.isna(text): return None
    m = re.search(r"(3[6-9]|4[0-2])\b", str(text))
    return int(m.group(1)) if m else None

def extract_color_from_variant_values(text):
    """
    Πάρε χρώμα μετά το 'Χρώμα:' (ή 'Color:') και πριν από 'Μεγ..', 'Size', κόμμα/ελ/pipe ή τέλος.
    """
    if pd.isna(text): return np.nan
    s = re.sub(r"\s+", " ", str(text)).strip()
    m = re.search(
        r"(?:Χρώμα|ΧΡΩΜΑ|Color)\s*[:：\-–—]?\s*(.+?)(?=\s*(?:Μεγ[\wΆ-ώ]+|Sizes?|Size|Taille|,|;|\||$))",
        s, flags=re.IGNORECASE
    )
    color = m.group(1).strip() if m else None
    if color:
        color = re.sub(r"[\s,;|]+$", "", color).strip().strip(' "\'“”‘’')
    return color if color else np.nan

def build_variant_sku(our_code8, size):
    """11ψήφιο SKU: OurCode(8) + (Size-34).zfill(3)"""
    if our_code8 is None or pd.isna(size): return None
    return f"{our_code8}{str(int(size)-34).zfill(3)}"

def base_target_for_size(size):
    try: s = int(size)
    except Exception: return 0
    if s in (38, 39): return 6
    if s in (37, 40): return 4
    if s == 41: return 2
    if s in (36, 42): return 1
    return 0

def clip(x, lo, hi):
    try: return max(lo, min(hi, x))
    except Exception: return lo

def _norm(s): return re.sub(r"[\s/_\-]+", "", str(s).strip().lower())

def find_col(df, tokens):
    toks = [t.lower() for t in (tokens if isinstance(tokens, (list, tuple)) else [tokens])]
    for c in df.columns:
        nc = _norm(c)
        if all(t in nc for t in toks): return c
    return None

def find_any_col(df, token_sets):
    for tokens in token_sets:
        col = find_col(df, tokens)
        if col: return col
    return None

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

    # 1) Read
    try:
        stock = pd.read_excel(stock_file, sheet_name=stock_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read STOCK sheet '{stock_sheet}': {e}"); st.stop()
    try:
        sales = pd.read_excel(sales_file, sheet_name=sales_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read SALES sheet '{sales_sheet}': {e}"); st.stop()

    # 2) STOCK essentials
    # Our Code
    color_sku_col = "Color SKU" if "Color SKU" in stock.columns else find_col(stock, ["color","sku"])
    our_code_col = "Our Code" if "Our Code" in stock.columns else None
    if color_sku_col:
        stock["Our Code"] = stock[color_sku_col].apply(clean_our_code)
    elif our_code_col:
        stock["Our Code"] = stock[our_code_col].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' ή 'Our Code'."); st.stop()

    # Variant Values (για Size & Color)
    vv_col = "Variant Values" if "Variant Values" in stock.columns else find_any_col(stock, [["variant","values"]])
    if not vv_col:
        st.error("Το STOCK πρέπει να έχει στήλη 'Variant Values' (ή αντίστοιχη) με 'Χρώμα:' & 'Μεγέθη:'."); st.stop()

    # ---- Color από STOCK (ffill) ----
    stock = stock.fillna(method="ffill")
    stock["ColorName"] = stock[vv_col].apply(extract_color_from_variant_values)
    stock["ColorName"] = stock["ColorName"].fillna(method="ffill")

    # Size filter
    stock["Size"] = stock[vv_col].apply(extract_size_from_variant_values) if "Size" not in stock.columns else stock["Size"]
    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # On Hand / Forecasted
    onhand_col = "On Hand" if "On Hand" in stock.columns else find_any_col(stock, [["on","hand"]])
    forecast_col = "Forecasted" if "Forecasted" in stock.columns else find_any_col(stock, [["forecast"]])
    stock["On Hand"] = stock[onhand_col].apply(to_int_safe) if onhand_col else 0
    stock["Forecasted"] = stock[forecast_col].apply(to_int_safe) if forecast_col else 0

    # Vendor & Vendor Code από STOCK (ffill)
    brand_col = "Brand" if "Brand" in stock.columns else find_col(stock, ["brand"])
    vendor_code_col = (
        "Vendor Code" if "Vendor Code" in stock.columns else
        ("Vendors/Vendor Product Code" if "Vendors/Vendor Product Code" in stock.columns else
         find_any_col(stock, [["vendors","vendor","product","code"],["vendor","product","code"],["vendorcode"]]))
    )
    for c in [brand_col, vendor_code_col]:
        if c and c in stock.columns:
            stock[c] = stock[c].ffill()

    # Build Variant SKU & group (κρατάμε ColorName από STOCK)
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max","ColorName": mode_non_null})
    )

    # Vendor maps (ανά Our Code) από STOCK
    tmp_vendor = pd.DataFrame({"Our Code": stock["Our Code"]})
    tmp_vendor["__Vendor"] = stock[brand_col] if brand_col else np.nan
    tmp_vendor["__VendorCode"] = stock[vendor_code_col] if vendor_code_col else np.nan
    stock_vendor_map = (
        tmp_vendor.groupby("Our Code", as_index=False)
                  .agg({"__Vendor": mode_non_null, "__VendorCode": mode_non_null})
                  .rename(columns={"__Vendor":"Vendor", "__VendorCode":"Vendor Code"})
    )

    # 3) SALES parsing (για πωλήσεις μόνο)
    # Εντοπισμός στήλης με [11ψήφιο] ή καθαρό 11ψήφιο
    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d{11}\]").any():
                sku_col = c; break
        except Exception:
            pass
    if sku_col is None:
        for c in sales.columns:
            try:
                if sales[c].astype(str).str.fullmatch(r"\d{11}").any():
                    sku_col = c; break
            except Exception:
                pass

    total_col = "Total" if "Total" in sales.columns else find_col(sales, ["total"])

    if sku_col is not None and total_col is not None:
        sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d{11})\]").iloc[:,0]
        mask_no_br = sales["Variant SKU"].isna() & sales[sku_col].astype(str).str.fullmatch(r"\d{11}")
        sales.loc[mask_no_br, "Variant SKU"] = sales.loc[mask_no_br, sku_col].astype(str)
        sales["Qty Ordered"] = sales[total_col].apply(to_int_safe)

        sales_by_variant = (
            sales.dropna(subset=["Variant SKU"])
                 .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
        )
        # --- ΣΗΜΑΝΤΙΚΟ: Our Code για το color-level aggregation ---
        sales_by_variant["Our Code"] = sales_by_variant["Variant SKU"].astype(str).str.slice(0,8)
        sales_by_color = (
            sales_by_variant.groupby("Our Code", as_index=False)["Qty Ordered"].sum()
                            .rename(columns={"Qty Ordered":"Sales Color Total"})
        )
    else:
        sales_by_variant = pd.DataFrame(columns=["Variant SKU","Qty Ordered"])
        sales_by_color = pd.DataFrame(columns=["Our Code","Sales Color Total"])

    # 4) Merge
    df = stock_grp.merge(sales_by_variant[["Variant SKU","Qty Ordered"]], on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)

    # ---- Αλεξίσφαιρη ύπαρξη "Our Code" πριν το merge ----
    if "Our Code" not in df.columns or df["Our Code"].isna().all():
        # Φτιάξε από το Variant SKU (8 πρώτα)
        df["Our Code"] = df["Variant SKU"].astype(str).str.slice(0,8)

    if "Our Code" not in sales_by_color.columns and not sales_by_color.empty:
        # Σπανίως, αν κάτι πήγε στραβά, ανακατασκεύασε από Variant SKU (αν υπάρχει)
        if "Variant SKU" in sales_by_color.columns:
            sales_by_color["Our Code"] = sales_by_color["Variant SKU"].astype(str).str.slice(0,8)

    # Τώρα το merge στο Our Code θα είναι ασφαλές
    df = df.merge(sales_by_color, on="Our Code", how="left")
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # Vendor από STOCK
    df = df.merge(stock_vendor_map, on="Our Code", how="left")

    # Color από STOCK ήδη στο stock_grp ως ColorName
    df.rename(columns={"ColorName":"Color"}, inplace=True)

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

    # 6) Restock
    df["Restock Quantity"] = (df["Adjusted Target"] - df["Forecasted"]).clip(lower=0)

    # 7) Export
    final_cols = [
        "Vendor", "Vendor Code", "Color",
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

    # Diagnostics
    with st.expander("🔎 Diagnostics"):
        st.write({
            "Columns in df before color/qty merges": list(stock_grp.columns),
            "Columns in sales_by_variant": list(sales_by_variant.columns),
            "Columns in sales_by_color": list(sales_by_color.columns),
            "Has 'Our Code' in df?": "Our Code" in df.columns,
            "Non-null (Vendor/Vendor Code/Color)": out[["Vendor","Vendor Code","Color"]].notna().sum().to_dict(),
        })
        sample = stock[[vv_col]].head(10).copy()
        sample["Extracted Color"] = sample[vv_col].apply(extract_color_from_variant_values)
        st.write("Stock → Extracted Color samples:", sample)

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
