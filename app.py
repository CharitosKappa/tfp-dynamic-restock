# app.py
# Dynamic Restock v12 – Streamlit app
# Vendor fields from Sales with robust detection + safe Vendor Color (no Color SKU)
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v1")
st.caption("Upload Stock + Sales → dynamic restock. Vendor/Vendor Code/Vendor Color mapped from Sales.")

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
    """Extract EU size 36–42 from 'Variant Values' or free text."""
    if pd.isna(text): return None
    m = re.search(r"(3[6-9]|4[0-2])\b", str(text))
    return int(m.group(1)) if m else None

def extract_color_from_text(text):
    """Extract color after 'Χρώμα:' or 'Color:' from a free-text field."""
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

def find_col(df, tokens, *, exclude_tokens=None):
    """
    Βρες στήλη που περιέχει ΟΛΑ τα tokens στο normalized name.
    Μπορείς να δώσεις exclude_tokens (κανένα να μην περιέχεται).
    """
    toks = [t.lower() for t in (tokens if isinstance(tokens, (list,tuple)) else [tokens])]
    excl = [t.lower() for t in (exclude_tokens or [])]
    for c in df.columns:
        nc = _norm(c)
        if all(t in nc for t in toks) and all(t not in nc for t in excl):
            return c
    return None

def find_any_col(df, list_of_token_sets, *, exclude_tokens=None):
    """Δοκίμασε διαδοχικά διαφορετικά token sets μέχρι να βρεθεί στήλη."""
    for tokens in list_of_token_sets:
        col = find_col(df, tokens, exclude_tokens=exclude_tokens)
        if col: return col
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

    # Size
    vv_col_stock = find_any_col(stock, [["variant","values"]]) or ("Variant Values" if "Variant Values" in stock.columns else None)
    if vv_col_stock:
        stock["Size"] = stock[vv_col_stock].apply(extract_size_from_variant_values)
    elif "Size" in stock.columns:
        stock["Size"] = stock["Size"]
    else:
        st.error("Stock must have 'Variant Values' or a usable 'Size' column."); st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # Our Code
    color_sku_col = find_col(stock, ["color","sku"]) or ("Color SKU" if "Color SKU" in stock.columns else None)
    our_code_col_stock = "Our Code" if "Our Code" in stock.columns else None
    if color_sku_col:
        stock["Our Code"] = stock[color_sku_col].apply(clean_our_code)
    elif our_code_col_stock:
        stock["Our Code"] = stock[our_code_col_stock].apply(clean_our_code)
    else:
        st.error("Stock needs 'Color SKU' or 'Our Code'."); st.stop()

    # On Hand / Forecasted
    onhand_col = find_any_col(stock, [["on","hand"]]) or ("On Hand" if "On Hand" in stock.columns else None)
    forecast_col = find_any_col(stock, [["forecast"]]) or ("Forecasted" if "Forecasted" in stock.columns else None)
    stock["On Hand"] = stock[onhand_col].apply(to_int_safe) if onhand_col else 0
    stock["Forecasted"] = stock[forecast_col].apply(to_int_safe) if forecast_col else 0

    # Stock vendor fallback (exclude Color SKU being mistaken as Vendor Color)
    vendor_col_stock = find_any_col(stock, [["vendor"],["manufacturer"],["brand"]])
    vendor_code_col_stock = find_any_col(stock, [["vendor","code"],["manufacturer","code"],["brand","code"]])
    vendor_color_col_stock = (
        find_any_col(stock, [["vendor","color"],["vendor","colour"]], exclude_tokens=["sku","code"])
        or find_any_col(stock, [["color"],["colour"],["χρώμα"]], exclude_tokens=["sku","code"])
    )
    for c in [vendor_col_stock, vendor_code_col_stock, vendor_color_col_stock]:
        if c: stock[c] = stock[c].ffill()

    # Build stock vendor map safely (only existing cols)
    agg_map = {}
    if vendor_col_stock: agg_map[vendor_col_stock] = "first"
    if vendor_code_col_stock: agg_map[vendor_code_col_stock] = "first"
    if vendor_color_col_stock: agg_map[vendor_color_col_stock] = "first"

    if agg_map:
        stock_vendor_map = stock.groupby("Our Code", as_index=False).agg(agg_map)
        rename_map = {}
        if vendor_col_stock: rename_map[vendor_col_stock] = "Vendor_stock"
        if vendor_code_col_stock: rename_map[vendor_code_col_stock] = "Vendor Code_stock"
        if vendor_color_col_stock: rename_map[vendor_color_col_stock] = "Vendor Color_stock"
        stock_vendor_map = stock_vendor_map.rename(columns=rename_map)
    else:
        stock_vendor_map = pd.DataFrame({"Our Code": stock["Our Code"].dropna().unique()})
        stock_vendor_map["Vendor_stock"] = np.nan
        stock_vendor_map["Vendor Code_stock"] = np.nan
        stock_vendor_map["Vendor Color_stock"] = np.nan

    # Variant SKU
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # 3) SALES parsing
    sales = sales_raw.copy()

    # Variant SKU detection
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
    if sku_col is None:
        sku_col = "Unnamed: 0" if "Unnamed: 0" in sales.columns else None
    if sku_col is None:
        st.error("Could not detect a SKU column in Sales (expected something like '[12345678901]')."); st.stop()

    sales["Variant SKU"] = sales[sku_col].astype(str).str.extract(r"\[(\d+)\]").iloc[:,0]
    mask_no_brackets = sales["Variant SKU"].isna() & sales[sku_col].astype(str).str.fullmatch(r"\d{11}")
    sales.loc[mask_no_brackets, "Variant SKU"] = sales.loc[mask_no_brackets, sku_col].astype(str)
    sales["Our Code"] = sales["Variant SKU"].str.slice(0,8)

    # Quantities
    total_col = find_any_col(sales, [["total"]]) or ("Total" if "Total" in sales.columns else None)
    if not total_col:
        st.error("Sales must contain a 'Total' column (ordered quantities)."); st.stop()
    sales["Qty Ordered"] = sales[total_col].apply(to_int_safe)

    sales_by_variant = (
        sales.dropna(subset=["Variant SKU"])
             .groupby("Variant SKU", as_index=False)["Qty Ordered"].sum()
    )
    sales_by_variant["ColorSKU"] = sales_by_variant["Variant SKU"].str.slice(0,8)
    sales_by_color = (
        sales_by_variant.groupby("ColorSKU", as_index=False)["Qty Ordered"].sum()
                        .rename(columns={"Qty Ordered":"Sales Color Total"})
    )

    # ----- SALES-driven vendor fields -----
    # Vendor (display name): multiple aliases
    col_vendor_disp = find_any_col(
        sales,
        [
            ["vendors","display","name"],   # Vendors/Display Name
            ["vendor","display","name"],    # Vendor/Display Name
            ["vendor","displayname"],       # Vendor Display Name
            ["vendor","name"],              # Vendor Name
            ["supplier","name"],            # Supplier Name
            ["manufacturer"],               # Manufacturer
            ["brand"],                      # Brand
        ]
    ) or ("Vendors/Display Name" if "Vendors/Display Name" in sales.columns else None)

    # Vendor Code
    col_vendor_code = find_any_col(
        sales,
        [
            ["vendors","vendor","product","code"],   # Vendors/Vendor Product Code
            ["vendor","product","code"],
            ["supplier","product","code"],
            ["manufacturer","code"],
            ["vendorcode"],
        ]
    ) or ("Vendors/Vendor Product Code" if "Vendors/Vendor Product Code" in sales.columns else None)

    # Vendor Color from Variant Values (Χρώμα:)
    col_variant_values_sales = find_any_col(sales, [["variant","values"]]) or ("Variant Values" if "Variant Values" in sales.columns else None)
    if col_variant_values_sales:
        sales["__VendorColor_from_sales"] = sales[col_variant_values_sales].apply(extract_color_from_text)
    else:
        sales["__VendorColor_from_sales"] = np.nan

    # Build sales_vendor_map ONLY with columns that actually exist
    cols_for_map = ["Our Code"]
    rename_sales = {}
    if col_vendor_disp:
        cols_for_map.append(col_vendor_disp); rename_sales[col_vendor_disp] = "Vendor_sales"
    if col_vendor_code:
        cols_for_map.append(col_vendor_code); rename_sales[col_vendor_code] = "Vendor Code_sales"
    cols_for_map.append("__VendorColor_from_sales"); rename_sales["__VendorColor_from_sales"] = "Vendor Color_sales"

    def first_non_null(s):
        s = s.dropna()
        return s.iloc[0] if not s.empty else np.nan

    sales_vendor_map = (
        sales.dropna(subset=["Our Code"])[cols_for_map]
             .groupby("Our Code", as_index=False)
             .agg(first_non_null)
             .rename(columns=rename_sales)
    )
    # Ensure columns exist
    for c in ["Vendor_sales","Vendor Code_sales","Vendor Color_sales"]:
        if c not in sales_vendor_map.columns: sales_vendor_map[c] = np.nan

    # 4) Merge + Vendor resolution
    df = stock_grp.merge(sales_by_variant, on="Variant SKU", how="left")
    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)
    df = df.merge(sales_by_color, left_on="Our Code", right_on="ColorSKU", how="left").drop(columns=["ColorSKU"], errors="ignore")
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    if not sales_vendor_map.empty:
        df = df.merge(sales_vendor_map[["Our Code","Vendor_sales","Vendor Code_sales","Vendor Color_sales"]], on="Our Code", how="left")
    else:
        df["Vendor_sales"] = np.nan; df["Vendor Code_sales"] = np.nan; df["Vendor Color_sales"] = np.nan

    if not stock_vendor_map.empty:
        df = df.merge(stock_vendor_map, on="Our Code", how="left")
    else:
        df["Vendor_stock"] = np.nan; df["Vendor Code_stock"] = np.nan; df["Vendor Color_stock"] = np.nan

    # Final vendor fields with Sales priority (fallback to Stock)
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
                "SKU": sku_col,
                "Vendor": col_vendor_disp,
                "Vendor Code": col_vendor_code,
                "Variant Values (for Vendor Color)": col_variant_values_sales,
            },
            "Detected Stock columns": {
                "Variant Values": vv_col_stock,
                "Color SKU": color_sku_col,
                "Our Code": our_code_col_stock,
                "Vendor (fallback)": vendor_col_stock,
                "Vendor Code (fallback)": vendor_code_col_stock,
                "Vendor Color (fallback)": vendor_color_col_stock,
            },
            "Non-null in export": out[["Vendor","Vendor Code","Vendor Color"]].notna().sum().to_dict()
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
