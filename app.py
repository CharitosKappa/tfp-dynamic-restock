# app.py
# Dynamic Restock v12 – Streamlit app
# Explicit mappings requested:
#   Vendor  (export col A) <- Sales['Brand'] by Our Code <-> Sales['Color SKU']
#   Vendor Code (export col B) <- Sales['Vendors/Vendor Product Code']  (unchanged)
#   Color  (export col C) <- parse after "Χρώμα:" from Sales['Variant Values']
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v1", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Upload Stock + Sales → dynamic restock. Vendor/Code/Color mapped explicitly from Sales.")

# ---------------- Helpers ----------------
def to_int_safe(x):
    try:
        if pd.isna(x): return 0
        return int(float(str(x).strip()))
    except Exception:
        return 0

def clean_our_code(x):
    """Normalize to 8-digit numeric string."""
    if pd.isna(x): return None
    s = str(x).strip()
    if s.endswith(".0"): s = s[:-2]
    s = re.sub(r"\D", "", s)
    if not s: return None
    return s.zfill(8)[:8]

def extract_size_from_variant_values(text):
    if pd.isna(text): return None
    m = re.search(r"(3[6-9]|4[0-2])\b", str(text))
    return int(m.group(1)) if m else None

def extract_color_after_keyword(text):
    """Return text after 'Χρώμα:' or 'Color:' inside Variant Values."""
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

def _norm(s):
    return re.sub(r"[\s/_\-]+", "", str(s).strip().lower())

def find_col(df, tokens, *, exclude_tokens=None):
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

    # ---------- STOCK essentials ----------
    stock = stock_raw.copy()

    vv_col_stock = find_any_col(stock, [["variant","values"]]) or ("Variant Values" if "Variant Values" in stock.columns else None)
    if vv_col_stock:
        stock["Size"] = stock[vv_col_stock].apply(extract_size_from_variant_values)
    elif "Size" in stock.columns:
        stock["Size"] = stock["Size"]
    else:
        st.error("Stock must have 'Variant Values' or a usable 'Size' column."); st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # Our Code (8-digit) from Stock
    color_sku_col_stock = find_col(stock, ["color","sku"]) or ("Color SKU" if "Color SKU" in stock.columns else None)
    our_code_col_stock = "Our Code" if "Our Code" in stock.columns else None
    if color_sku_col_stock:
        stock["Our Code"] = stock[color_sku_col_stock].apply(clean_our_code)
    elif our_code_col_stock:
        stock["Our Code"] = stock[our_code_col_stock].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' or 'Our Code'."); st.stop()

    # On Hand / Forecasted
    onhand_col = find_any_col(stock, [["on","hand"]]) or ("On Hand" if "On Hand" in stock.columns else None)
    forecast_col = find_any_col(stock, [["forecast"]]) or ("Forecasted" if "Forecasted" in stock.columns else None)
    stock["On Hand"] = stock[onhand_col].apply(to_int_safe) if onhand_col else 0
    stock["Forecasted"] = stock[forecast_col].apply(to_int_safe) if forecast_col else 0

    # Variant SKU (11-digit rule)
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)

    # Keep one row per variant for stock levels
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # ---------- SALES parsing & explicit mapping ----------
    sales = sales_raw.copy()

    # Exact columns per your instructions
    sales_color_sku_col = "Color SKU" if "Color SKU" in sales.columns else find_col(sales, ["color","sku"])
    sales_brand_col = "Brand" if "Brand" in sales.columns else find_col(sales, ["brand"])
    sales_vendor_code_col = "Vendors/Vendor Product Code" if "Vendors/Vendor Product Code" in sales.columns else find_any_col(sales, [["vendors","vendor","product","code"],["vendor","product","code"]])
    sales_variant_values_col = "Variant Values" if "Variant Values" in sales.columns else find_any_col(sales, [["variant","values"]])

    # Normalize Sales Color SKU to 8-digit to match Our Code
    if not sales_color_sku_col:
        st.error("Sales must contain a 'Color SKU' column for the mapping."); st.stop()
    sales["ColorSKU_norm"] = sales[sales_color_sku_col].apply(clean_our_code)

    # Parse Color from Variant Values (after 'Χρώμα:')
    if sales_variant_values_col:
        sales["Color_from_sales"] = sales[sales_variant_values_col].apply(extract_color_after_keyword)
    else:
        sales["Color_from_sales"] = np.nan

    # Build color-level map from Sales: key = ColorSKU_norm
    cols_for_sales_map = ["ColorSKU_norm"]
    if sales_brand_col: cols_for_sales_map.append(sales_brand_col)
    if sales_vendor_code_col: cols_for_sales_map.append(sales_vendor_code_col)
    cols_for_sales_map.append("Color_from_sales")

    sales_color_map = (
        sales.dropna(subset=["ColorSKU_norm"])[cols_for_sales_map]
             .groupby("ColorSKU_norm", as_index=False)
             .agg(lambda s: s.dropna().iloc[0] if s.dropna().size else np.nan)
    )

    # Rename to canonical export names
    rename_cols = {}
    if sales_brand_col: rename_cols[sales_brand_col] = "Vendor_from_sales"
    if sales_vendor_code_col: rename_cols[sales_vendor_code_col] = "Vendor Code_from_sales"
    rename_cols["Color_from_sales"] = "Color_from_sales"
    sales_color_map = sales_color_map.rename(columns=rename_cols)

    # Ensure canonical columns exist
    for c in ["Vendor_from_sales","Vendor Code_from_sales","Color_from_sales"]:
        if c not in sales_color_map.columns: sales_color_map[c] = np.nan

    # ---------- Sales by Variant / by Color for targets ----------
    # Try to find a column with [digits] for Variant SKU (optional, only for sales qtys)
    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d+\]").any():
                sku_col = c; break
        except Exception:
            pass
    if sku_col is None:
        for c in sales.columns:
            if sales[c].astype(str).str.fullmatch(r"\d{11}").any():
                sku_col = c; break

    if sku_col is not None:
        sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d+)\]").iloc[:,0]
        mask_no_br = sales["Variant SKU"].isna() & sales[sku_col].astype(str).str.fullmatch(r"\d{11}")
        sales.loc[mask_no_br, "Variant SKU"] = sales.loc[mask_no_br, sku_col].astype(str)
        sales["Our Code from Variant"] = sales["Variant SKU"].str.slice(0,8)
    else:
        sales["Variant SKU"] = np.nan
        sales["Our Code from Variant"] = np.nan

    total_col = "Total" if "Total" in sales.columns else find_col(sales, ["total"])
    if total_col is None:
        st.error("Sales must contain a 'Total' column with ordered quantities."); st.stop()
    sales["Qty Ordered"] = sales[total_col].apply(to_int_safe)

    # Aggregations
    sales_by_variant = (
        sales.dropna(subset=["Variant SKU"])
             .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
    )
    sales_by_variant["ColorSKU_from_variant"] = sales_by_variant["Variant SKU"].str.slice(0,8)
    sales_by_color_qty = (
        sales_by_variant.groupby("ColorSKU_from_variant", as_index=False)["Qty Ordered"].sum()
                        .rename(columns={"Qty Ordered":"Sales Color Total"})
    )

    # ---------- Merge: bring Sales Vendor/Color fields to Stock via Our Code <-> Color SKU ----------
    df = stock_grp.merge(
        sales_color_map,
        left_on="Our Code",
        right_on="ColorSKU_norm",
        how="left"
    ).drop(columns=["ColorSKU_norm"], errors="ignore")

    # Bring Qty Ordered & Sales Color Total (if available)
    df = df.merge(sales_by_variant, on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)
    df = df.merge(sales_by_color_qty, left_on="Our Code", right_on="ColorSKU_from_variant", how="left") \
           .drop(columns=["ColorSKU_from_variant"], errors="ignore")
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # ---------- Final Vendor/Color columns for export ----------
    df.rename(columns={
        "Vendor_from_sales": "Vendor",
        "Vendor Code_from_sales": "Vendor Code",
        "Color_from_sales": "Color"
    }, inplace=True)

    # ---------- Targets logic ----------
    df["Base Target"] = df["Size"].apply(base_target_for_size)
    base_sum_per_color = df.groupby("Our Code", as_index=False)["Base Target"].sum().rename(columns={"Base Target":"BaseSumColor"})
    df = df.merge(base_sum_per_color, on="Our Code", how="left")
    df["BaseSumColor"] = df["BaseSumColor"].replace(0, np.nan)
    df["GlobalMult"] = (df["Sales Color Total"] / df["BaseSumColor"]).fillna(0).apply(lambda x: clip(x, 0.5, 5.0))

    avg_sales_per_color = df.groupby("Our Code", as_index=False)["Qty Ordered"].mean().rename(columns={"Qty Ordered":"AvgSalesPerSize"})
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

    df.loc[df["Qty Ordered"] == 0, "Adjusted Target"] = 0
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

    # ---------- Final export ----------
    final_cols = [
        "Vendor", "Vendor Code", "Color",   # first three per your spec
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

    # Diagnostics (to verify mapping)
    with st.expander("🔎 Diagnostics"):
        det = {
            "Sales columns found": {
                "Color SKU": sales_color_sku_col,
                "Brand": sales_brand_col,
                "Vendors/Vendor Product Code": sales_vendor_code_col,
                "Variant Values": sales_variant_values_col,
            },
            "Non-null in export": out[["Vendor","Vendor Code","Color"]].notna().sum().to_dict(),
            "Example rows": out[["Our Code","Vendor","Vendor Code","Color"]].head(10).to_dict(orient="records")
        }
        st.write(det)

    st.success("Done! Preview below ↓")
    st.dataframe(out, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Restock v12")
    st.download_button(
        label="⬇️ Download dynamic_restock_order_v12.xlsx",
        data=buffer.getvalue(),
        file_name="dynamic_restock_order_v12.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
