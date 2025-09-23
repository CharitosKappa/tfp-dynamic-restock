# app.py
# Dynamic Restock v12 – Streamlit app
# Robust vendor mapping from Sales (Vendors/Display Name, Vendors/Vendor Product Code, Variant Values -> Χρώμα:)
# + fallback από Stock αν λείπουν
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Upload Stock + Sales → dynamic restock with vendor fields mapped from Sales (robust column detection).")

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
    return s.zfill(8)[:8]

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

# ---------- Column detection (robust) ----------
def _norm(s):
    return re.sub(r"[\s/_\-]+", "", str(s).strip().lower())

def find_col(df, tokens_or_patterns, mode="all"):
    """
    tokens_or_patterns: list of strings (tokens) or regex patterns.
    mode="all": all tokens present in normalized column name
    mode="regex": any regex matches the raw column name
    Returns first matching column or None.
    """
    cols = list(df.columns)
    if mode == "regex":
        for c in cols:
            for pat in tokens_or_patterns:
                if re.search(pat, str(c), flags=re.IGNORECASE):
                    return c
        return None
    else:
        toks = [t.lower() for t in tokens_or_patterns]
        for c in cols:
            nc = _norm(c)
            if all(t in nc for t in toks):
                return c
        return None

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

    # 1) Read
    try:
        stock_raw = pd.read_excel(stock_file, sheet_name=stock_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read STOCK sheet '{stock_sheet}': {e}"); st.stop()
    try:
        sales_raw = pd.read_excel(sales_file, sheet_name=sales_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read SALES sheet '{sales_sheet}': {e}"); st.stop()

    # 2) STOCK essentials
    stock = stock_raw.copy()

    # Detect size / variant values
    vv_col_stock = find_col(stock, ["variant","values"]) or "Variant Values" if "Variant Values" in stock.columns else None
    if vv_col_stock:
        stock["Size"] = stock[vv_col_stock].apply(extract_size_from_variant_values)
    elif "Size" in stock.columns:
        stock["Size"] = stock["Size"]
    else:
        st.error("Stock must have 'Variant Values' or a usable 'Size' column."); st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # Our Code (Color SKU)
    color_sku_col = find_col(stock, ["color","sku"]) or ("Color SKU" if "Color SKU" in stock.columns else None)
    our_code_col_stock = "Our Code" if "Our Code" in stock.columns else None
    if color_sku_col:
        stock["Our Code"] = stock[color_sku_col].apply(clean_our_code)
    elif our_code_col_stock:
        stock["Our Code"] = stock[our_code_col_stock].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' or 'Our Code'."); st.stop()

    # On Hand / Forecasted
    onhand_col = find_col(stock, ["on","hand"]) or "On Hand" if "On Hand" in stock.columns else None
    forecast_col = find_col(stock, ["forecast"]) or "Forecasted" if "Forecasted" in stock.columns else None
    if onhand_col: stock["On Hand"] = stock[onhand_col].apply(to_int_safe)
    else: stock["On Hand"] = 0
    if forecast_col: stock["Forecasted"] = stock[forecast_col].apply(to_int_safe)
    else: stock["Forecasted"] = 0

    # Optional vendor info from STOCK (fallback only)
    vendor_col_stock = find_col(stock, ["vendor"]) or find_col(stock, ["manufacturer"]) or find_col(stock, ["brand"]) 
    vendor_code_col_stock = find_col(stock, ["vendor","code"]) or find_col(stock, ["manufacturer","code"]) or find_col(stock, ["brand","code"])
    vendor_color_col_stock = find_col(stock, ["vendor","color"]) or find_col(stock, ["vendor","colour"]) or find_col(stock, ["color"]) or find_col(stock, ["colour"]) or find_col(stock, ["χρώμα"])
    # forward-fill stock vendor columns
    for c in [vendor_col_stock, vendor_code_col_stock, vendor_color_col_stock]:
        if c: stock[c] = stock[c].ffill()
    # build stock vendor map per Our Code
    stock_vendor_map = (
        stock.groupby("Our Code", as_index=False)
             .agg({
                 vendor_col_stock: "first" if vendor_col_stock else (lambda s: np.nan),
                 vendor_code_col_stock: "first" if vendor_code_col_stock else (lambda s: np.nan),
                 vendor_color_col_stock: "first" if vendor_color_col_stock else (lambda s: np.nan),
             })
             if "Our Code" in stock.columns else pd.DataFrame(columns=["Our Code"])
    )
    # rename safely
    rename_map = {}
    if vendor_col_stock: rename_map[vendor_col_stock] = "Vendor_stock"
    if vendor_code_col_stock: rename_map[vendor_code_col_stock] = "Vendor Code_stock"
    if vendor_color_col_stock: rename_map[vendor_color_col_stock] = "Vendor Color_stock"
    if not stock_vendor_map.empty:
        stock_vendor_map = stock_vendor_map.rename(columns=rename_map)

    # Variant SKU
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)

    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # 3) SALES parsing
    sales = sales_raw.copy()

    # Column holding [digits] for Variant SKU
    # Try to find any column where some cell contains [\d+]
    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d+\]").any():
                sku_col = c; break
        except Exception:
            pass
    if sku_col is None:
        # fallback: try to find a pure digits column
        for c in sales.columns:
            if sales[c].astype(str).str.fullmatch(r"\d{11}").any():
                sku_col = c; break
    if sku_col is None:
        # last fallback
        sku_col = "Unnamed: 0" if "Unnamed: 0" in sales.columns else None
    if sku_col is None:
        st.error("Could not detect a SKU column in Sales (expected something like '[12345678901]')."); st.stop()

    sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d+)\]").iloc[:,0]
    # if pattern not found, keep digits if the whole cell is digits
    mask_no_brackets = sales["Variant SKU"].isna() & sales[sku_col].astype(str).str.fullmatch(r"\d{11}")
    sales.loc[mask_no_brackets, "Variant SKU"] = sales.loc[mask_no_brackets, sku_col].astype(str)
    sales["Our Code"] = sales["Variant SKU"].str.slice(0,8)

    # Qty Ordered
    total_col = find_col(sales, ["total"]) or "Total" if "Total" in sales.columns else None
    if not total_col:
        st.error("Sales must contain a 'Total' column (ordered quantities)."); st.stop()
    sales["Qty Ordered"] = sales[total_col].apply(to_int_safe)

    # Sales by variant
    sales_by_variant = (
        sales.dropna(subset=["Variant SKU"])
             .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
    )
    # Sales by color
    sales_by_variant["ColorSKU"] = sales_by_variant["Variant SKU"].str.slice(0,8)
    sales_by_color = (
        sales_by_variant.groupby("ColorSKU", as_index=False)["Qty Ordered"].sum()
                        .rename(columns={"Qty Ordered":"Sales Color Total"})
    )

    # ----- SALES-driven vendor fields (robust column detection) -----
    col_vendor_disp = find_col(sales, ["vendors","display","name"]) or "Vendors/Display Name" if "Vendors/Display Name" in sales.columns else None
    col_vendor_code = find_col(sales, ["vendors","vendor","product","code"]) or "Vendors/Vendor Product Code" if "Vendors/Vendor Product Code" in sales.columns else None
    col_variant_values_sales = find_col(sales, ["variant","values"]) or "Variant Values" if "Variant Values" in sales.columns else None

    # build Vendor Color from sales Variant Values (Χρώμα:)
    if col_variant_values_sales:
        sales["__VendorColor_from_sales"] = sales[col_variant_values_sales].apply(extract_color_from_variant_values)
    else:
        sales["__VendorColor_from_sales"] = np.nan

    # Group per Our Code, take first non-null
    def first_non_null(s):
        s = s.dropna()
        return s.iloc[0] if not s.empty else np.nan

    sales_vendor_map_cols = ["Our Code"]
    if col_vendor_disp: sales_vendor_map_cols.append(col_vendor_disp)
    if col_vendor_code: sales_vendor_map_cols.append(col_vendor_code)
    sales_vendor_map_cols.append("__VendorColor_from_sales")

    sales_vendor_map = (
        sales.dropna(subset=["Our Code"])[sales_vendor_map_cols]
             .groupby("Our Code", as_index=False)
             .agg({(col_vendor_disp if col_vendor_disp else "__dummy"): first_non_null,
                   (col_vendor_code if col_vendor_code else "__dummy2"): first_non_null,
                   "__VendorColor_from_sales": first_non_null})
    )

    # Rename to canonical names (handle missing gracefully)
    rename_sales = {}
    if col_vendor_disp: rename_sales[col_vendor_disp] = "Vendor_sales"
    if col_vendor_code: rename_sales[col_vendor_code] = "Vendor Code_sales"
    rename_sales["__VendorColor_from_sales"] = "Vendor Color_sales"
    sales_vendor_map = sales_vendor_map.rename(columns=rename_sales)
    # ensure canonical columns exist
    for c in ["Vendor_sales","Vendor Code_sales","Vendor Color_sales"]:
        if c not in sales_vendor_map.columns:
            sales_vendor_map[c] = np.nan

    # 4) Merge
    df = stock_grp.merge(sales_by_variant, on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)
    df = df.merge(sales_by_color, left_on="Our Code", right_on="ColorSKU", how="left").drop(columns=["ColorSKU"], errors="ignore")
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # Merge vendor maps (Sales preferred, Stock fallback)
    if not sales_vendor_map.empty:
        df = df.merge(sales_vendor_map[["Our Code","Vendor_sales","Vendor Code_sales","Vendor Color_sales"]], on="Our Code", how="left")
    else:
        df["Vendor_sales"] = np.nan; df["Vendor Code_sales"] = np.nan; df["Vendor Color_sales"] = np.nan

    if not stock_vendor_map.empty:
        df = df.merge(stock_vendor_map, on="Our Code", how="left")
    else:
        df["Vendor_stock"] = np.nan; df["Vendor Code_stock"] = np.nan; df["Vendor Color_stock"] = np.nan

    # Final vendor fields with sales priority
    df["Vendor"] = df.apply(lambda r: coalesce(r.get("Vendor_sales"), r.get("Vendor_stock")), axis=1)
    df["Vendor Code"] = df.apply(lambda r: coalesce(r.get("Vendor Code_sales"), r.get("Vendor Code_stock")), axis=1)
    df["Vendor Color"] = df.apply(lambda r: coalesce(r.get("Vendor Color_sales"), r.get("Vendor Color_stock")), axis=1)

    # 5) Targets
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
        (df["Qty Ordered"] == 0) & (df["On Hand"] == 0) & (df["Forecasted"] == 0) &
        (df["Size"].isin([38,39,40])) & (df["Sales Color Total"] > 0)
    )
    df.loc[core_mask, "Adjusted Target"] = df.loc[core_mask, "Base Target"]

    # 6) Restock
    df["Restock Quantity"] = (df["Adjusted Target"] - df["Forecasted"]).clip(lower=0)

    # 7) Final file
    final_cols = [
        "Vendor", "Vendor Code", "Vendor Color",
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
        det = {
            "Detected Sales columns": {
                "SKU column": sku_col,
                "Vendor (Vendors/Display Name)": col_vendor_disp,
                "Vendor Code (Vendors/Vendor Product Code)": col_vendor_code,
                "Variant Values (Sales)": col_variant_values_sales,
            },
            "Detected Stock columns": {
                "Variant Values (Stock)": vv_col_stock,
                "Color SKU": color_sku_col,
                "Our Code (Stock)": our_code_col_stock,
                "Vendor (Stock)": vendor_col_stock,
                "Vendor Code (Stock)": vendor_code_col_stock,
                "Vendor Color (Stock)": vendor_color_col_stock,
            },
            "Non-null counts (export)": out[["Vendor","Vendor Code","Vendor Color"]].notna().sum().to_dict()
        }
        st.write(det)

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

# End
