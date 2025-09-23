# app.py
# Dynamic Restock v12 – Streamlit app
# Mapping:
#   Vendor (A)       <- from STOCK['Brand'] by Our Code
#   Vendor Code (B)  <- from STOCK['Vendor Code'] or STOCK['Vendors/Vendor Product Code'] by Our Code
#   Color (C)        <- from SALES line "[###########] ... (Color, Size)" by Our Code (primary),
#                       fallback to STOCK text parsing ("Χρώμα: ...") if missing
# Sales also used for quantities; targets/restock v12 rules preserved.
# Requirements: streamlit, pandas, numpy, openpyxl

import io, re, math
import numpy as np
import pandas as pd
import streamlit as st
from collections import Counter

# ---------------- UI ----------------
st.set_page_config(page_title="Dynamic Restock v12", page_icon="📦", layout="wide")
st.title("📦 Dynamic Restock v12")
st.caption("Vendor & Vendor Code από STOCK • Color από SALES (π.χ. '(Μαύρο, L/XL)')")

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

def extract_color_from_stock_text(text):
    """Color after 'Χρώμα:' ή 'Color:' από STOCK text (fallback)."""
    if pd.isna(text): return None
    s = re.sub(r"\s+", " ", str(text)).strip()
    m = re.search(
        r"(?:Χρώμα|ΧΡΩΜΑ|Color)\s*[:：\-–—]?\s*(.+?)(?=\s*(?:Μεγ[\wΆ-ώ]+|Sizes?|Size|Taille|,|;|\||$))",
        s, flags=re.IGNORECASE
    )
    color = m.group(1).strip() if m else None
    if color:
        color = re.sub(r"[\s,;|]+$", "", color).strip().strip(' "\'“”‘’')
    return color if color else None

def extract_color_from_sales_line(text):
    """
    Από SALES γραμμή τύπου:
      "[17930002013] Σλιπ ... (Μαύρο, L/XL)"
    Επιστρέφει ΠΑΝΤΑ το 1ο στοιχείο μέσα στην ΠΡΩΤΗ παρένθεση (πριν το 1ο κόμμα): 'Μαύρο'.
    """
    if pd.isna(text): return None
    s = str(text)
    m = re.search(r"\(([^)]*)\)", s)  # περιεχόμενο πρώτης παρένθεσης
    if not m:
        return None
    inside = m.group(1)                  # π.χ. "Μαύρο, L/XL"
    first_part = inside.split(",")[0]    # => "Μαύρο"
    color = first_part.strip().strip(' "\'“”‘’')
    return color if color else None

def extract_variant_sku_from_text(text):
    """Επιστρέφει 11ψήφιο Variant SKU από κείμενο τύπου '[###########]' ή σκέτο 11ψήφιο."""
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

def clip(x, lo, hi
