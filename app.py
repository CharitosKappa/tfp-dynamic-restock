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
    m = re.search(
        r"(?:Χρώμα|ΧΡΩΜΑ|Color)\s*[:：\-–—]?\s*"
        r"(.+?)"
        r"(?=\s*(?:Μεγ[\wΆ-ώ]+|Sizes?|Size|Taille|,|;|\||$))",
        s, flags=re.IGNORECASE
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
    vv_col_stock = "Variant Values" if "Variant Values" in stock.columns else find_any_col(stock, [["variant","values"]])
    if vv_col_stock:
        stock["Size"] = stock[vv_col_stock].apply(extract_size_from_variant_values)
    elif "Size" in stock.columns:
        stock["Size"] = stock["Size"]
    else:
        st.error("Stock must have 'Variant Values' or a usable 'Size' column."); st.stop()

    stock["Size"] = stock["Size"].apply(lambda x: int(x) if pd.notna(x) and str(x).isdigit() else x)
    stock = stock[stock["Size"].isin([36,37,38,39,40,41,42])].copy()

    # On Hand / Forecasted
    onhand_col = "On Hand" if "On Hand" in stock.columns else find_any_col(stock, [["on","hand"]])
    forecast_col = "Forecasted" if "Forecasted" in stock.columns else find_any_col(stock, [["forecast"]])
    stock["On Hand"] = stock[onhand_col].apply(to_int_safe) if onhand_col else 0
    stock["Forecasted"] = stock[forecast_col].apply(to_int_safe) if forecast_col else 0

    # ---- AUTHORITATIVE vendor fields from STOCK ----
    brand_col = "Brand" if "Brand" in stock.columns else find_col(stock, ["brand"])
    vendor_code_col_stock = (
        "Vendor Code" if "Vendor Code" in stock.columns else
        ("Vendors/Vendor Product Code" if "Vendors/Vendor Product Code" in stock.columns else
         find_any_col(stock, [["vendors","vendor","product","code"],["vendor","product","code"],["vendorcode"]]))
    )

    # --- Color από STOCK ---
    # 1) Προσπάθεια από ρητή στήλη (Color / Χρώμα / Colour), αποκλείοντας Color SKU / Code
    color_direct_col = find_any_col(stock, [["χρώμα"], ["color"], ["colour"]], exclude_tokens=["sku","code"])

    # 2) Εναλλακτικές πηγαίες στήλες κειμένου για parsing
    text_color_candidates = []
    if vv_col_stock: text_color_candidates.append(vv_col_stock)
    for cand in ["Variant Options", "Options", "Attributes", "Title"]:
        if cand in stock.columns and cand not in text_color_candidates:
            text_color_candidates.append(cand)

    # ffill για όλες τις πιθανές πηγές ώστε οι γραμμές μεγέθους να κληρονομούν τιμές
    for c in [brand_col, vendor_code_col_stock, color_direct_col] + text_color_candidates:
        if c and c in stock.columns:
            stock[c] = stock[c].ffill()

    # Υπολογισμός Color ανά γραμμή (direct -> parsed)
    # Χτίζουμε ενδιάμεσα tmp columns ώστε να μην περάσουν None keys σε agg
    tmp = pd.DataFrame({"Our Code": stock["Our Code"]})

    if brand_col:
        tmp["__VendorTmp"] = stock[brand_col]
    else:
        tmp["__VendorTmp"] = np.nan

    if vendor_code_col_stock:
        tmp["__VendorCodeTmp"] = stock[vendor_code_col_stock]
    else:
        tmp["__VendorCodeTmp"] = np.nan

    # Color ανά γραμμή
    if color_direct_col:
        color_line = stock[color_direct_col].copy()
    else:
        color_line = pd.Series([np.nan]*len(stock), index=stock.index)

    if text_color_candidates:
        def line_color(i):
            val = color_line.iat[i]
            if pd.notna(val) and str(val).strip() != "":
                return val
            for col in text_color_candidates:
                cval = stock[col].iat[i]
                c = extract_color_from_text(cval)
                if c: return c
            return np.nan
        color_line = pd.Series([line_color(i) for i in range(len(stock))], index=stock.index)

    tmp["__ColorTmp"] = color_line

    # -------- mapping ανά Our Code (mode μη-null) --------
    stock_vendor_map = (
        tmp.groupby("Our Code", as_index=False)
           .agg({"__VendorTmp": mode_non_null,
                 "__VendorCodeTmp": mode_non_null,
                 "__ColorTmp": mode_non_null})
           .rename(columns={"__VendorTmp": "Vendor_from_stock",
                            "__VendorCodeTmp": "Vendor Code_from_stock",
                            "__ColorTmp": "Color_from_stock"})
    )

    # Variant SKU
    stock["Variant SKU"] = stock.apply(lambda r: build_variant_sku(r["Our Code"], r["Size"]), axis=1)

    # Compact stock levels per variant
    stock_grp = (
        stock.groupby(["Our Code","Variant SKU","Size"], as_index=False)
             .agg({"On Hand":"max","Forecasted":"max"})
    )

    # ---------- SALES parsing (μόνο για quantities) ----------
    sales = sales_raw.copy()
    # detect Variant SKU column ([11 digits] ή καθαρό 11ψήφιο)
    sku_col = None
    for c in sales.columns:
        try:
            if sales[c].astype(str).str.contains(r"\[\d{11}\]").any():
                sku_col = c; break
        except Exception:
            pass
    if sku_col is None:
        for c in sales.columns:
            if sales[c].astype(str).str.fullmatch(r"\d{11}").any():
                sku_col = c; break

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
        sales_by_variant["Our Code"] = sales_by_variant["Variant SKU"].str.slice(0,8)
        sales_by_color = (
            sales_by_variant.groupby("Our Code", as_index=False)["Qty Ordered"].sum()
                            .rename(columns={"Qty Ordered":"Sales Color Total"})
        )
    else:
        sales_by_variant = pd.DataFrame(columns=["Variant SKU","Qty Ordered"])
        sales_by_color = pd.DataFrame(columns=["Our Code","Sales Color Total"])

    # ---------- Merge everything ----------
    df = stock_grp.copy()
    # Vendor fields από STOCK
    df = df.merge(
        stock_vendor_map[["Our Code","Vendor_from_stock","Vendor Code_from_stock","Color_from_stock"]],
        on="Our Code", how="left"
    )

    # Sales quantities
    if not sales_by_variant.empty:
        df = df.merge(sales_by_variant[["Variant SKU","Qty Ordered"]], on="Variant SKU", how="left")
    if "Qty Ordered" not in df.columns:
        df["Qty Ordered"] = 0
    if not sales_by_color.empty:
        df = df.merge(sales_by_color, on="Our Code", how="left")
    if "Sales Color Total" not in df.columns:
        df["Sales Color Total"] = 0

    df["Qty Ordered"] = df["Qty Ordered"].fillna(0).astype(int)
    df["Sales Color Total"] = df["Sales Color Total"].fillna(0).astype(int)

    # Τελικές στήλες Vendor/Code/Color
    df.rename(columns={
        "Vendor_from_stock": "Vendor",
        "Vendor Code_from_stock": "Vendor Code",
        "Color_from_stock": "Color",
    }, inplace=True)

    # ---------- Targets & Restock ----------
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

    # Restock quantity
    df["Restock Quantity"] = (df["Adjusted Target"] - df["Forecasted"]).clip(lower=0)

    # ---------- Export ----------
    final_cols = [
        "Vendor", "Vendor Code", "Color",   # A, B, C from STOCK
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
            "Detected STOCK columns": {
                "Color SKU": color_sku_col_stock,
                "Our Code (raw)": our_code_col_stock,
                "Variant Values": vv_col_stock,
                "Brand": brand_col,
                "Vendor Code": vendor_code_col_stock,
                "Color (direct)": color_direct_col,
                "Text color sources": text_color_candidates,
            },
            "Non-null (Vendor/Code/Color)": out[["Vendor","Vendor Code","Color"]].notna().sum().to_dict(),
        })
        # Δείγματα: πηγή κειμένου & εξαγόμενο Color για έλεγχο
        try:
            show_cols = [c for c in [color_direct_col] + text_color_candidates if c]
            sample = stock[["Our Code"] + show_cols].head(15).copy()
            sample["Extracted Color"] = tmp["__ColorTmp"].head(15).values
            st.write(sample)
        except Exception:
            pass

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
