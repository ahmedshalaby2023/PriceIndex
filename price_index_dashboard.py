"""Streamlit app for analyzing item price indices by brand and date.

Flow
----
1. User uploads Excel workbook (Market sheet = transaction data).
2. App reads helper sheets:
   - Brands  : one column, lists the base brands we own (e.g. Atyab, Chikitita, Meatland).
   - Com     : columns [Product, Base Brand, Com] — maps each SKU to its base brand and
               the competitor brands to show alongside it.
3. User selects an Item Description from the Market sheet.
4. App resolves automatically:
   a. base_brand      → the Base Brand in the Com sheet for that SKU
                        (must be one of the Brands-sheet brands)
   b. brand_selection → [base_brand] + Com competitors for that (SKU, base_brand)
                        filtered to brands that have data in Market
5. All charts and the snapshot table are rendered for exactly those brands.
"""
from __future__ import annotations

from datetime import timedelta
from io import BytesIO
from pathlib import Path
from itertools import cycle
from typing import Any, Iterable, Optional
from difflib import get_close_matches

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Price Index Explorer", layout="wide")
st.markdown(
    """<style>.brand-snapshot-table td { padding: 0.4rem 0.6rem !important; }</style>""",
    unsafe_allow_html=True,
)

# ── Constants ─────────────────────────────────────────────────────────────────
DEFAULT_SHEET = "Market"
SUPPORTED_EXTENSIONS = {".xlsb", ".xlsx", ".xls"}
BRANDS_SHEET_NAME = "Brands"
COM_SHEET_NAME = "Com"
COM_PRODUCT_COL = "Product"
COM_BASE_BRAND_COL = "Base Brand"
COM_COMPETITOR_COL = "Com"
DEFAULT_ITEM = "Meat_Chilled Luncheon_Retail_5Kg_1_Luncheon"
ITEM_SELECT_KEY = "pie_item_select"
ITEM_INPUT_KEY = "pie_item_input"
ITEM_LAST_KEY = "pie_item_last"
SHOW_LABELS_KEY = "pie_show_labels"
BRAND_ALIASES: dict[str, tuple[str, ...]] = {
    "Atyab": ("atyab",),
    "Chikitita": ("chikitita", "chikitia"),
    "Meatland": ("meatland",),
}


# ── Pure helpers ──────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    return str(s).strip().lower()


def _match_brand(text: str, brand_list: list[str]) -> Optional[str]:
    lowered = _norm(text)
    for brand, aliases in BRAND_ALIASES.items():
        if brand in brand_list and any(a in lowered for a in aliases):
            return brand
    for brand in brand_list:
        if _norm(brand) in lowered:
            return brand
    return None


def _brand_col_map(columns: Iterable[str], required: list[str], brand_list: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    low_req = [t.lower() for t in required]
    for col in columns:
        low = _norm(col)
        if not all(t in low for t in low_req):
            continue
        b = _match_brand(col, brand_list)
        if b and b not in out:
            out[b] = col
    return out


def _find_generic_per_kg(columns: Iterable[str]) -> Optional[str]:
    for terms in [("price", "per", "kg", "promot"), ("price", "per", "kg")]:
        for col in columns:
            if all(t in _norm(col) for t in terms):
                return col
    return None


def _guess(columns: Iterable[str], keywords: list[str]) -> Optional[str]:
    for col in columns:
        if all(k in _norm(col) for k in keywords):
            return col
    return None


def _guess_item_col(columns: Iterable[str]) -> Optional[str]:
    for col in columns:
        n = _norm(col)
        if "item" in n and any(x in n for x in ("desc", "description", "name")):
            return col
    return None


def _find_price_index_cols(df: pd.DataFrame) -> list[str]:
    cols = [c for c in df.columns if "price" in _norm(c) and "index" in _norm(c)]
    return cols or df.select_dtypes(include=["number"]).columns.tolist()


def _excel_serial(v: Any) -> pd.Timestamp:
    try:
        num = float(v)
    except (TypeError, ValueError):
        return pd.NaT
    if not np.isfinite(num):
        return pd.NaT
    return pd.Timestamp("1899-12-30") + pd.to_timedelta(num, unit="D")


def _norm_pct(s: pd.Series) -> pd.Series:
    num = pd.to_numeric(s, errors="coerce")
    valid = num.dropna()
    if not valid.empty and valid.median() > 2:
        num = num / 100.0
    return num


def _fmt_kg(v: Any) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{float(v):,.2f}"


def _fmt_pct_kg(pct: Any, kg: Any) -> str:
    try:
        p = f"{float(pct):.1%}" if pct is not None and not pd.isna(pct) else "—"
    except (TypeError, ValueError):
        p = str(pct)
    return p if (kg is None or pd.isna(kg)) else f"{p} ({_fmt_kg(kg)})"


def _fmt_customers(v: Any) -> str:
    try:
        n = int(round(float(v)))
    except (TypeError, ValueError):
        return ""
    if n <= 0:
        return ""
    return "1 customer" if n == 1 else f"{n} customers"


def _border_color(v: Any, fallback: str) -> str:
    try:
        n = int(round(float(v)))
    except (TypeError, ValueError):
        return fallback
    return "#ff1744" if n == 1 else ("#1e88e5" if n >= 3 else fallback)


# ── Sheet loaders ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_sheet(upload: BytesIO, *, sheet_name: str) -> pd.DataFrame:
    suffix = Path(getattr(upload, "name", "f")).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type '{suffix}'.")
    upload.seek(0)
    df = pd.read_excel(upload, sheet_name=sheet_name, engine="pyxlsb" if suffix == ".xlsb" else None)
    upload.seek(0)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_brands(upload: BytesIO) -> list[str]:
    try:
        df = load_sheet(upload, sheet_name=BRANDS_SHEET_NAME)
    except Exception:
        return []
    vals = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    return list(dict.fromkeys(v for v in vals if v))


def load_com(upload: BytesIO) -> dict[str, dict[str, list[str]]]:
    """Returns {sku: {base_brand: [competitor, ...]}}"""
    try:
        df = load_sheet(upload, sheet_name=COM_SHEET_NAME)
    except Exception:
        return {}
    # Flexible column name matching
    rename: dict[str, str] = {}
    for col in df.columns:
        n = _norm(col)
        if "product" in n and COM_PRODUCT_COL not in rename.values():
            rename[col] = COM_PRODUCT_COL
        elif "base" in n and "brand" in n and COM_BASE_BRAND_COL not in rename.values():
            rename[col] = COM_BASE_BRAND_COL
        elif (n == "com" or n.startswith("com")) and COM_COMPETITOR_COL not in rename.values():
            rename[col] = COM_COMPETITOR_COL
    df = df.rename(columns=rename)
    if not {COM_PRODUCT_COL, COM_BASE_BRAND_COL, COM_COMPETITOR_COL}.issubset(df.columns):
        return {}
    df = df[[COM_PRODUCT_COL, COM_BASE_BRAND_COL, COM_COMPETITOR_COL]].dropna(
        subset=[COM_PRODUCT_COL, COM_BASE_BRAND_COL]
    )
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    out: dict[str, dict[str, list[str]]] = {}
    for _, row in df.iterrows():
        sku, base, comp = row[COM_PRODUCT_COL], row[COM_BASE_BRAND_COL], row[COM_COMPETITOR_COL]
        out.setdefault(sku, {}).setdefault(base, [])
        if comp and _norm(comp) not in ("nan", ""):
            out[sku][base].append(comp)
    return out


# ═══════════════════════════════════════════════════════════════════════════════
# APP
# ═══════════════════════════════════════════════════════════════════════════════

st.title("📊 Price Index Explorer")

# ── Sidebar: file + column mapping ───────────────────────────────────────────
with st.sidebar:
    st.header("1️⃣  Data source")
    uploaded_file = st.file_uploader("Upload Excel (.xlsb / .xlsx / .xls)", type=["xlsb", "xlsx", "xls"])
    sheet_input = st.text_input("Market sheet name", value=DEFAULT_SHEET)
    if uploaded_file is None:
        st.info("Upload a file to begin.")
        st.stop()
    try:
        raw_df = load_sheet(uploaded_file, sheet_name=sheet_input or DEFAULT_SHEET)
    except Exception as e:
        st.error(f"Failed to load: {e}")
        st.stop()

if raw_df.empty:
    st.warning("Worksheet is empty.")
    st.stop()

# Load Brands + Com helper sheets
brands_list = load_brands(uploaded_file)
com_map = load_com(uploaded_file)

with st.sidebar:
    if not brands_list:
        st.error("❌ 'Brands' sheet missing or empty. Please add a sheet named 'Brands' with one brand per row.")
        st.stop()

    st.header("2️⃣  Column mapping")
    all_cols = raw_df.columns.tolist()

    brand_col = st.selectbox(
        "Brand column", all_cols,
        index=all_cols.index(g) if (g := _guess(all_cols, ["brand"])) in all_cols else 0,
    )
    item_col = st.selectbox(
        "Item description column", all_cols,
        index=all_cols.index(g) if (g := _guess_item_col(all_cols)) in all_cols else 0,
    )
    channel_col = st.selectbox(
        "Sales channel (optional)", ["(none)"] + all_cols,
        index=(all_cols.index(g) + 1) if (g := _guess(all_cols, ["channel"])) in all_cols else 0,
    )
    customer_col = st.selectbox(
        "Customer column (optional)", ["(none)"] + all_cols,
        index=(all_cols.index(g) + 1) if (g := _guess(all_cols, ["customer"])) in all_cols else 0,
    )
    date_col = st.selectbox(
        "Date column (optional)", ["(none)"] + all_cols,
        index=(all_cols.index(g) + 1) if (g := _guess(all_cols, ["date"])) in all_cols else 0,
    )
    pi_candidates = _find_price_index_cols(raw_df)
    selected_pi_cols = st.multiselect(
        "Price index column(s)", all_cols,
        default=[c for c in pi_candidates if c in all_cols][: max(1, len(pi_candidates))],
    )

if not selected_pi_cols:
    st.error("Select at least one price index column.")
    st.stop()

# ── Pre-process ───────────────────────────────────────────────────────────────
wdf = raw_df.copy()

if date_col != "(none)":
    raw_d = wdf[date_col]
    dt = pd.to_datetime(raw_d, errors="coerce", infer_datetime_format=True)
    serial = raw_d.apply(_excel_serial)
    dt = dt.mask(serial.notna(), serial)
    if pd.api.types.is_datetime64tz_dtype(dt):
        dt = dt.dt.tz_convert(None)
    wdf[date_col] = dt.dt.normalize()

# Brand→price-index-col and brand→per-kg-col maps (for brands_list brands)
brand_pi_map = _brand_col_map(all_cols, ["price", "index"], brands_list)
generic_pkg = _find_generic_per_kg(all_cols)
brand_pkg_map = _build_pkg = _brand_col_map(all_cols, ["price", "per", "kg", "promoted"], brands_list)
if generic_pkg:
    for b in brands_list:
        brand_pkg_map.setdefault(b, generic_pkg)

# Normalize price-index cols
pi_to_norm = set(selected_pi_cols) | set(brand_pi_map.values())
for c in pi_to_norm:
    if c in wdf.columns:
        wdf[c] = _norm_pct(wdf[c].replace({"-": np.nan, "": np.nan}))

# Numeric per-kg cols
pkg_to_clean = set(brand_pkg_map.values())
if generic_pkg:
    pkg_to_clean.add(generic_pkg)
for c in pkg_to_clean:
    if c not in wdf.columns:
        continue
    s = wdf[c].astype(str).str.strip().replace({"-": np.nan, "": np.nan})
    s = s.str.replace(r"[^0-9,.-]", "", regex=True).str.replace(",", "", regex=False)
    wdf[c] = pd.to_numeric(s, errors="coerce")

wdf = wdf.dropna(subset=[brand_col, item_col])

# ── Data validation report ───────────────────────────────────────────────────
all_items = wdf[item_col].astype(str).sort_values().unique().tolist()
all_market_brands = set(wdf[brand_col].astype(str).unique())

if com_map:
    # Collect all issues
    issues_unmatched_sku: list[dict] = []      # Com product not in Market SKUs
    issues_base_brand: list[dict] = []          # Com base brand not in Brands sheet
    issues_com_brand: list[dict] = []           # Com competitor not in Market brands

    for sku, base_dict in com_map.items():
        # 1. Check SKU match in Market
        exact_sku_match = sku in all_items
        if not exact_sku_match:
            # Prefix match: same text before the first "_"
            prefix = sku.split("_")[0].strip()
            prefix_matches = [i for i in all_items if i.split("_")[0].strip() == prefix]
            # Only exact or prefix counts as resolved — no fuzzy guessing
            # If neither resolves, it is a true no-match → report it
            if not prefix_matches:
                issues_unmatched_sku.append({
                    "Com Product": sku,
                })

        for base_b, competitors in base_dict.items():
            # 2. Base brand vs Brands sheet
            if base_b not in brands_list:
                fuzzy_brand = get_close_matches(base_b, brands_list, n=1, cutoff=0.6)
                issues_base_brand.append({
                    "Com Base Brand": base_b,
                    "Com Product": sku,
                    "Did you mean?": fuzzy_brand[0] if fuzzy_brand else "—",
                })

            # 3. Competitor brands vs Market brands
            for comp in competitors:
                if comp not in all_market_brands:
                    fuzzy_comp = get_close_matches(comp, list(all_market_brands), n=1, cutoff=0.6)
                    issues_com_brand.append({
                        "Com Competitor": comp,
                        "Com Product": sku,
                        "Base Brand": base_b,
                        "Did you mean?": fuzzy_comp[0] if fuzzy_comp else "—",
                    })

    total_issues = len(issues_unmatched_sku) + len(issues_base_brand) + len(issues_com_brand)

    with st.expander(
        f"{'⚠️' if total_issues else '✅'} Data validation report "
        f"({'%d issue%s found' % (total_issues, 's' if total_issues != 1 else '')} — click to review)"
        if com_map else "✅ Data validation report",
        expanded=total_issues > 0,
    ):
        if total_issues == 0:
            st.success("✅ All Com sheet products, base brands, and competitors match the Market data perfectly.")
        else:
            st.caption(
                "These mismatches may cause missing data in charts or the snapshot. "
                "Fix them in the Excel file or the Com sheet to ensure accuracy."
            )

        if issues_unmatched_sku:
            st.markdown(f"#### 🔴 Com products not found in Market SKUs ({len(issues_unmatched_sku)})")
            st.caption(
                "These product names in the 'Com' sheet do not exactly match any Item Description "
                "in the Market sheet. The dashboard will try a prefix fallback but exact names are recommended."
            )
            st.dataframe(
                pd.DataFrame(issues_unmatched_sku)[["Com Product"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Com Product": st.column_config.TextColumn("Com Sheet Product — not found in Market"),
                },
            )

        if issues_base_brand:
            st.markdown(f"#### 🟠 Com base brands not in Brands sheet ({len(issues_base_brand)})")
            st.caption(
                "These base brands appear in the 'Com' sheet but are not listed in the 'Brands' sheet. "
                "They will not appear as selectable base brands in the dashboard."
            )
            st.dataframe(
                pd.DataFrame(issues_base_brand),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Com Base Brand": st.column_config.TextColumn("Base Brand in Com"),
                    "Com Product": st.column_config.TextColumn("Product"),
                    "Did you mean?": st.column_config.TextColumn("Closest in Brands sheet"),
                },
            )

        if issues_com_brand:
            st.markdown(f"#### 🟡 Competitor brands not found in Market data ({len(issues_com_brand)})")
            st.caption(
                "These competitor brands are listed in the 'Com' sheet but have no records "
                "in the Market sheet. Their prices will show as '—' in the snapshot."
            )
            st.dataframe(
                pd.DataFrame(issues_com_brand),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Com Competitor": st.column_config.TextColumn("Competitor in Com"),
                    "Com Product": st.column_config.TextColumn("Product"),
                    "Base Brand": st.column_config.TextColumn("Base Brand"),
                    "Did you mean?": st.column_config.TextColumn("Closest in Market"),
                },
            )

st.markdown("---")

# ── STEP 1: Item selection ────────────────────────────────────────────────────
if not all_items:
    st.error("No items found in data.")
    st.stop()

def_idx = all_items.index(DEFAULT_ITEM) if DEFAULT_ITEM in all_items else 0

st.markdown("### 🔍 Step 1 — Select item")
selected_item = st.selectbox("Item description", all_items, index=def_idx, key=ITEM_SELECT_KEY)

# Optional text refinement — no key, no session_state writes after widget creation
refined = st.text_input(
    "Refine item description (optional)",
    value=selected_item,
    help="Edit to narrow down. A close match will be used automatically.",
).strip()
if refined and refined != selected_item:
    if refined in all_items:
        selected_item = refined
    else:
        hits = [i for i in all_items if refined.lower() in i.lower()]
        selected_item = (
            hits[0] if hits
            else (get_close_matches(refined, all_items, n=1, cutoff=0.0) or [selected_item])[0]
        )

# ── STEP 2: Auto-resolve base brand + competitors ─────────────────────────────
#
#  Rule:
#    - base_brand comes from the "Base Brand" column of the Com sheet for this SKU
#      AND must be one of the brands in the Brands sheet.
#    - brand_selection = [base_brand] + Com competitors, filtered to brands with data.
#
st.markdown("### ⚖️ Step 2 — Base brand & competitors")

brands_with_data: set[str] = set(
    wdf.loc[wdf[item_col].astype(str) == selected_item, brand_col].astype(str).unique()
)

# Resolve Com sheet entry: try exact SKU first, then prefix before first "_"
sku_com: dict[str, list[str]] = com_map.get(selected_item, {})
com_match_key = selected_item  # track which key actually matched

if not sku_com:
    prefix = selected_item.split("_")[0].strip()
    # Find any Com sheet product that starts with the same prefix
    for com_key in com_map:
        if com_key.split("_")[0].strip() == prefix:
            sku_com = com_map[com_key]
            com_match_key = com_key
            break

# Base brand candidates: must be in both Brands sheet and Com sheet for this SKU
valid_base = [b for b in brands_list if b in sku_com and b in brands_with_data]

if not valid_base:
    # Graceful fallback: Brands-sheet brands that at least have data
    valid_base = [b for b in brands_list if b in brands_with_data]
    if not valid_base:
        st.error(
            f"No brand from the 'Brands' sheet has data for **'{selected_item}'**. "
            "Select a different item or check the data."
        )
        st.stop()
    if sku_com:
        st.warning(
            f"⚠️ Com sheet has entries for this product "
            f"({', '.join(sku_com.keys())}), but none are in the 'Brands' sheet."
        )
    else:
        st.warning(
            f"⚠️ No competition mapping found for **'{selected_item}'** "
            f"(also tried prefix '{selected_item.split('_')[0].strip()}')."
        )

# Show selectbox only if multiple base brands are possible
if len(valid_base) == 1:
    base_brand = valid_base[0]
    st.info(f"🏷️ Base brand: **{base_brand}**")
else:
    base_brand = st.selectbox(
        "Base brand (from 'Brands' sheet)",
        options=valid_base,
        help="Multiple base brands found for this product in the 'Brands' sheet.",
    )

# Competitors from Com sheet for (SKU, base_brand), filtered to what's in data
com_competitors: list[str] = sku_com.get(base_brand, [])
valid_competitors = [b for b in com_competitors if b in brands_with_data]

# Final brand list — EXACTLY these brands, nothing else
brand_selection: list[str] = list(dict.fromkeys([base_brand] + valid_competitors))
initial_visible: set[str] = set(brand_selection)

# Feedback
if valid_competitors:
    st.success(
        f"✅ **{base_brand}** vs **{', '.join(valid_competitors)}** "
        "(mapped in 'Com' sheet)"
    )
elif com_competitors:
    st.warning(
        f"⚠️ Com sheet maps {', '.join(com_competitors)} as competitors "
        f"of '{base_brand}', but they have no data here."
    )
else:
    st.warning(f"⚠️ No competitors mapped for '{base_brand}' on this product in the 'Com' sheet.")

# ── STEP 3: Optional filters (channel, customer, date) ───────────────────────
st.markdown("### 🔎 Step 3 — Filters")

# Subset: selected item + mapped brands only
subset_mask = (
    (wdf[item_col].astype(str) == selected_item)
    & (wdf[brand_col].astype(str).isin(brand_selection))
)
item_sub = wdf.loc[subset_mask].copy()

customer_sel: Optional[list[str]] = None
channel_sel: Optional[list[str]] = None
c1, c2 = st.columns(2)

with c1:
    if channel_col != "(none)" and channel_col in item_sub.columns:
        item_sub[channel_col] = (
            item_sub[channel_col].astype(str).str.strip().replace({"": np.nan, "nan": np.nan})
        )
        ch_opts = item_sub[channel_col].dropna().sort_values().unique().tolist()
        if ch_opts:
            channel_sel = st.multiselect("Sales channel", ch_opts, default=ch_opts)
            if not channel_sel:
                st.warning("Select at least one channel.")
                st.stop()
            item_sub = item_sub[item_sub[channel_col].astype(str).isin(channel_sel)]

with c2:
    if customer_col != "(none)" and customer_col in item_sub.columns:
        cu_opts = item_sub[customer_col].dropna().astype(str).sort_values().unique().tolist()
        if cu_opts:
            customer_sel = st.multiselect("Customer", cu_opts, default=cu_opts)
            if not customer_sel:
                st.warning("Select at least one customer.")
                st.stop()
            item_sub = item_sub[item_sub[customer_col].astype(str).isin(customer_sel)]

start_date = end_date = selected_date = None
with st.sidebar:
    st.header("3️⃣  Date filter")
    # Use full wdf (all items, all brands) for date options so the filter
    # applies globally across charts AND the brand snapshot
    _date_source = wdf if date_col != "(none)" and date_col in wdf.columns else None
    if _date_source is not None and _date_source[date_col].notna().any():
        date_opts = sorted({ts.date() for ts in _date_source[date_col].dropna()})
        if len(date_opts) == 1:
            start_date = end_date = date_opts[0]
            st.caption(f"Only one date available: {date_opts[0]:%Y-%m-%d}")
        else:
            start_date, end_date = st.select_slider(
                "Date range", options=date_opts,
                value=(
                    next((d for d in date_opts if d >= date_opts[-1] - timedelta(days=30)), date_opts[0]),
                    date_opts[-1],
                ),
                format_func=lambda d: d.strftime("%Y-%m-%d"),
            )
            selected_date = None
            if st.checkbox("Single date"):
                selected_date = st.date_input(
                    "Date", value=end_date, min_value=start_date, max_value=end_date, format="YYYY-MM-DD"
                )
    else:
        st.caption("No date column selected.")

# Apply date filter to item_sub so channel/customer options respect the date range
if date_col != "(none)" and date_col in item_sub.columns:
    if start_date and end_date:
        item_sub = item_sub[item_sub[date_col].dt.date.between(start_date, end_date)]
    if selected_date is not None:
        item_sub = item_sub[item_sub[date_col].dt.date == selected_date]

# ── Build filtered dataset ────────────────────────────────────────────────────
filtered = wdf[
    (wdf[item_col].astype(str) == selected_item)
    & (wdf[brand_col].astype(str).isin(brand_selection))
].copy()

if channel_col != "(none)" and channel_sel:
    filtered = filtered[filtered[channel_col].astype(str).isin(channel_sel)]
if customer_col != "(none)" and customer_sel:
    filtered = filtered[filtered[customer_col].astype(str).isin(customer_sel)]
if date_col != "(none)" and start_date and end_date:
    filtered = filtered[filtered[date_col].dt.date.between(start_date, end_date)]
if date_col != "(none)" and selected_date is not None:
    filtered = filtered[filtered[date_col].dt.date == selected_date]

has_dates = (
    date_col != "(none)"
    and date_col in filtered.columns
    and filtered[date_col].notna().any()
)
if has_dates:
    filtered = filtered.sort_values([date_col, brand_col])

if filtered.empty:
    st.warning("No records match the current filters.")
    st.stop()

# ── Resolve price-index col and per-kg cols for active brands ─────────────────
base_pi_col = brand_pi_map.get(base_brand)
if base_pi_col is None:
    for c in selected_pi_cols:
        if _norm(base_brand) in _norm(c):
            base_pi_col = c
            break
if base_pi_col is None:
    base_pi_col = selected_pi_cols[0]
active_pi = [base_pi_col]

brand_pkg_cols: dict[str, str] = {}
for b in brand_selection:
    c = brand_pkg_map.get(b) or generic_pkg
    if c and c in filtered.columns:
        brand_pkg_cols[b] = c

# Per-kg latest + by-date
latest_pkg: dict[str, float] = {}
pkg_by_date: dict[str, dict[str, float]] = {}
for b, pkg_col in brand_pkg_cols.items():
    bdf = filtered.loc[filtered[brand_col].astype(str) == b]
    s = bdf[pkg_col].dropna()
    if s.empty:
        continue
    if has_dates:
        d2 = bdf.dropna(subset=[date_col, pkg_col]).sort_values(date_col)
        if not d2.empty:
            latest_pkg[b] = float(d2.iloc[-1][pkg_col])
            pkg_by_date[b] = {
                pd.to_datetime(k).strftime("%Y-%m-%d"): float(v)
                for k, v in d2.groupby(date_col)[pkg_col].mean().items()
            }
            continue
    latest_pkg[b] = float(s.iloc[-1])


def _get_pkg(brand: str, x: Any) -> Optional[float]:
    bd = pkg_by_date.get(brand)
    if bd:
        try:
            k = pd.to_datetime(x).strftime("%Y-%m-%d")
            if k in bd:
                return bd[k]
        except Exception:
            pass
    return latest_pkg.get(brand)


# ── Colour palette ────────────────────────────────────────────────────────────
BASE_CLR = "#1b5e20"
MUTED = "#b0bec5"
_cyc = cycle(px.colors.qualitative.Safe + px.colors.qualitative.Pastel)
brand_clr: dict[str, str] = {
    b: (BASE_CLR if b == base_brand else next(_cyc)) for b in brand_selection
}


def _hl_base(col: pd.Series) -> list[str]:
    return (
        ["background-color: #fff59d; font-weight: 600"] * len(col)
        if col.name == base_brand else [""] * len(col)
    )


# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("---")
st.subheader(f"📈 {selected_item}")
st.caption(
    f"Base: **{base_brand}** | "
    f"Competitors: {', '.join(valid_competitors) or '—'} | "
    f"Brands in view: {', '.join(brand_selection)}"
)
show_labels = st.checkbox("Show price-per-kg in chart labels", key=SHOW_LABELS_KEY, value=True)

# ── Relative price-index traces ───────────────────────────────────────────────
rel_traces: dict[str, pd.DataFrame] = {}
if has_dates and base_brand in filtered[brand_col].unique():
    for pi_col in active_pi:
        pivot = (
            filtered[[date_col, brand_col, pi_col]]
            .dropna(subset=[pi_col])
            .pivot_table(index=date_col, columns=brand_col, values=pi_col, aggfunc="mean")
        )
        if pivot.empty or base_brand not in pivot.columns:
            continue
        bs = pivot[base_brand]
        valid_idx = bs.replace(0, np.nan).dropna().index
        if valid_idx.empty:
            continue
        pivot, bs = pivot.loc[valid_idx], bs.loc[valid_idx]
        rel = pivot.divide(bs, axis=0) - 1
        fv = bs.iloc[0]
        rel[base_brand] = (bs / fv - 1) if (pd.notna(fv) and not np.isclose(fv, 0.0)) else bs.pct_change().fillna(0)
        rel = (
            rel.replace([np.inf, -np.inf], np.nan).reset_index()
            .melt(id_vars=[date_col], var_name=brand_col, value_name="rel")
            .dropna(subset=["rel"])
        )
        rel["rel"] = rel["rel"].where(~np.isclose(rel["rel"], 0.0, atol=1e-9), 0.0)
        rel_traces[pi_col] = rel

# ── Line chart: price index delta ─────────────────────────────────────────────
if has_dates:
    for pi_col in active_pi:
        plot = rel_traces.get(pi_col)
        if plot is not None and not plot.empty:
            fig = px.line(
                plot, x=date_col, y="rel", color=brand_col, markers=True,
                title=f"{pi_col} — Δ vs {base_brand}",
                labels={date_col: "Date", "rel": "Δ vs base"},
                color_discrete_map=brand_clr,
            )
            fig.update_yaxes(tickformat=".0%", tickfont=dict(color="#111"))
        else:
            fig = px.line(
                filtered, x=date_col, y=pi_col, color=brand_col, markers=True,
                title=f"{pi_col} trend",
                labels={date_col: "Date", pi_col: pi_col},
                color_discrete_map=brand_clr,
            )
            fig.update_yaxes(tickformat=".0%", tickfont=dict(color="#111"))

        fig.update_layout(hovermode="x unified", font=dict(color="#111"))
        fig.update_xaxes(tickfont=dict(color="#111"))

        for trace in fig.data:
            xs = list(trace.x) if trace.x is not None else []
            ys = list(trace.y) if trace.y is not None else []
            texts, positions, seen = [], [], set()
            for i, y in enumerate(ys):
                xv = xs[i] if i < len(xs) else i
                key = (trace.name, xv)
                if key in seen:
                    texts.append(""); positions.append("top center"); continue
                seen.add(key)
                prefix = f"{trace.name}: " if i == 0 else ""
                if pd.isna(y):
                    texts.append(prefix or "")
                else:
                    vs = f"{float(y):.1%}"
                    if show_labels:
                        pkg = _get_pkg(trace.name, xv)
                        if pkg is not None:
                            vs = f"{vs} ({_fmt_kg(pkg)})"
                    texts.append(f"{prefix}{vs}")
                positions.append("top center" if i % 2 == 0 else "bottom center")
            color = brand_clr.get(trace.name, MUTED)
            trace.text = texts
            trace.textposition = positions
            trace.mode = "lines+markers+text"
            trace.textfont = dict(size=11, color=color)
            if trace.name == base_brand:
                trace.line = dict(width=2, color=color)
                trace.marker = dict(size=10, color=color)
            else:
                trace.line = dict(width=1, color=color)
                trace.marker = dict(size=6, color=color)
                trace.opacity = 0.85
            if trace.name not in initial_visible:
                trace.visible = "legendonly"

        st.plotly_chart(fig, use_container_width=True)

    # ── Line chart: price per KG ───────────────────────────────────────────────
    pkg_rows: list[pd.DataFrame] = []
    for b in brand_selection:
        pkg_col = brand_pkg_cols.get(b)
        if not pkg_col:
            continue
        bdf = filtered.loc[filtered[brand_col].astype(str) == b, [date_col, pkg_col]].dropna()
        if bdf.empty:
            continue
        avg = bdf.groupby(date_col, as_index=False)[pkg_col].mean().rename(columns={pkg_col: "pkg"})
        avg[brand_col] = b
        pkg_rows.append(avg)

    if pkg_rows:
        pkg_df = pd.concat(pkg_rows, ignore_index=True).sort_values(date_col)
        pkg_piv = (
            pkg_df.pivot_table(index=date_col, columns=brand_col, values="pkg", aggfunc="mean")
            .sort_index().replace([np.inf, -np.inf], np.nan).dropna(how="all")
        )
        if not pkg_piv.empty:
            base_pkg_s = pkg_piv.get(base_brand)
            other_max = pkg_piv[[c for c in pkg_piv.columns if c != base_brand]].max(axis=1) if len(pkg_piv.columns) > 1 else None

            pfig = go.Figure()
            for b in pkg_piv.columns:
                s = pkg_piv[b].dropna()
                if s.empty:
                    continue
                color = brand_clr.get(b, MUTED)
                labels, mk_clr, mk_brd = [], [], []
                for i, val in enumerate(s.values):
                    brd = _border_color(None, color)
                    mk_brd.append(brd)
                    if b != base_brand and base_pkg_s is not None:
                        bv = base_pkg_s.reindex(s.index).iloc[i] if i < len(s) else np.nan
                        if pd.isna(bv):
                            mk_clr.append(color); txt = _fmt_kg(val)
                        else:
                            diff = float(val) - float(bv)
                            pct = (diff / float(bv) * 100) if float(bv) != 0 else 0
                            mk_clr.append("#ff1744" if diff < 0 else color)
                            txt = f"{_fmt_kg(val)} ({pct:+.1f}%)"
                    else:
                        if other_max is not None and i < len(other_max):
                            ov = other_max.reindex(s.index).iloc[i]
                            mk_clr.append("#00e676" if (not pd.isna(ov) and float(val) < float(ov)) else color)
                        else:
                            mk_clr.append(color)
                        txt = _fmt_kg(val)
                    labels.append(txt)

                positions = ["top center" if i % 2 == 0 else "bottom center" for i in range(len(labels))]
                pfig.add_trace(go.Scatter(
                    x=s.index, y=s.values, mode="lines+markers+text", name=b,
                    text=labels, textposition=positions, textfont=dict(size=11, color=color),
                    line=dict(color=color, width=2 if b == base_brand else 1),
                    marker=dict(size=9 if b == base_brand else 7, color=mk_clr,
                                line=dict(color=mk_brd, width=2)),
                    hovertemplate="<b>%{fullData.name}</b><br>%{x|%Y-%m-%d}: %{y:,.2f}<extra></extra>",
                    visible=True if b in initial_visible else "legendonly",
                ))

            pfig.update_layout(
                title="Price per KG (promoted) — trend",
                yaxis_title="Price / KG", xaxis_title="Date",
                hovermode="x unified", font=dict(color="#111"),
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02),
            )
            pfig.update_yaxes(separatethousands=True)
            pfig.update_xaxes(tickformat="%Y-%m-%d")
            st.plotly_chart(pfig, use_container_width=True)

            # ── Response-time table ────────────────────────────────────────────
            if base_pkg_s is not None:
                resp = []
                for comp in [c for c in pkg_piv.columns if c != base_brand]:
                    aln = pd.concat(
                        [base_pkg_s.rename("base"), pkg_piv[comp].rename("comp")], axis=1
                    ).dropna().sort_index()
                    if aln.empty:
                        continue
                    diff = aln["base"] - aln["comp"]
                    last_a = ea = eb = None
                    for d, v in diff.items():
                        if v > 0:
                            last_a = d
                        elif v < 0 and last_a is not None:
                            ea, eb, last_a = last_a, d, None
                    if ea and eb:
                        resp.append({"Brand": comp, "Above date": ea, "Below date": eb,
                                     "Response time (days)": int((eb - ea).days)})
                if resp:
                    st.markdown("#### ⏱️ Response time to compete")
                    st.dataframe(
                        pd.DataFrame(resp).sort_values("Response time (days)"),
                        use_container_width=True, hide_index=True,
                        column_config={
                            "Above date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                            "Below date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                            "Response time (days)": st.column_config.NumberColumn(format="%d"),
                        },
                    )

# ── Bar chart: average Δ vs base ──────────────────────────────────────────────
rel_prim = rel_traces.get(active_pi[0]) if rel_traces else None

avg_pkg: dict[str, float] = {
    b: float(filtered.loc[filtered[brand_col].astype(str) == b, c].dropna().mean())
    for b, c in brand_pkg_cols.items()
    if not filtered.loc[filtered[brand_col].astype(str) == b, c].dropna().empty
}

if rel_prim is not None and not rel_prim.empty:
    avg_r = rel_prim.groupby(brand_col, as_index=False)["rel"].mean().sort_values("rel", ascending=False)
    avg_r["pkg"] = avg_r[brand_col].map(avg_pkg.get)
    avg_r["lbl"] = avg_r.apply(lambda r: _fmt_pct_kg(r["rel"], r["pkg"] if show_labels else None), axis=1)
    st.markdown("### Average price index (Δ vs base)")
    if start_date and end_date:
        st.caption(f"Average: {start_date:%Y-%m-%d} → {end_date:%Y-%m-%d}")
    bar = px.bar(
        avg_r, x=brand_col, y="rel", text="lbl",
        labels={brand_col: "Brand", "rel": "Δ vs base"},
        title=f"{active_pi[0]} Δ vs {base_brand} (avg)",
        color=brand_col, color_discrete_map=brand_clr,
    )
    bar.update_traces(texttemplate="%{text}", textposition="outside",
                      marker_line_color="#546e7a", marker_line_width=1,
                      textfont=dict(color="#111", size=12))
    bar.update_yaxes(tickformat=".0%", tickfont=dict(color="#111"))
    bar.update_layout(
        xaxis=dict(categoryorder="array", categoryarray=avg_r[brand_col].tolist(),
                   tickfont=dict(size=14, color="#111")),
        font=dict(color="#111"),
    )
    if base_brand in avg_r[brand_col].values:
        bv = float(avg_r.loc[avg_r[brand_col] == base_brand, "rel"].iloc[0])
        bar.add_annotation(x=base_brand, y=bv, text="Base brand",
                           showarrow=True, arrowhead=3, arrowsize=1.4, arrowwidth=2,
                           arrowcolor=BASE_CLR, ax=0, ay=-40,
                           bgcolor="rgba(253,216,83,0.25)", bordercolor=BASE_CLR,
                           font=dict(color="#000", size=12))
    st.plotly_chart(bar, use_container_width=True)

else:
    avg_r2 = (
        filtered[[brand_col] + active_pi]
        .groupby(brand_col, as_index=False)[active_pi[0]].mean()
        .sort_values(active_pi[0], ascending=False)
    )
    avg_r2["pkg"] = avg_r2[brand_col].map(avg_pkg.get)
    avg_r2["lbl"] = avg_r2.apply(lambda r: _fmt_pct_kg(r[active_pi[0]], r["pkg"] if show_labels else None), axis=1)
    st.markdown("### Average price index")
    bar = px.bar(
        avg_r2, x=brand_col, y=active_pi[0], text="lbl",
        labels={brand_col: "Brand", active_pi[0]: active_pi[0]},
        title=f"{active_pi[0]} by brand (avg)",
        color=brand_col, color_discrete_map=brand_clr,
    )
    bar.update_traces(texttemplate="%{text}", textposition="outside",
                      marker_line_color="#546e7a", marker_line_width=1,
                      textfont=dict(color="#111", size=12))
    bar.update_yaxes(tickformat=".0%", tickfont=dict(color="#111"))
    bar.update_layout(
        xaxis=dict(categoryorder="array", categoryarray=avg_r2[brand_col].tolist(),
                   tickfont=dict(size=14, color="#111")),
        font=dict(color="#111"),
    )
    if base_brand in avg_r2[brand_col].values:
        bv = float(avg_r2.loc[avg_r2[brand_col] == base_brand, active_pi[0]].iloc[0])
        bar.add_annotation(x=base_brand, y=bv, text="Base brand",
                           showarrow=True, arrowhead=3, arrowsize=1.4, arrowwidth=2,
                           arrowcolor=BASE_CLR, ax=0, ay=-40,
                           bgcolor="rgba(253,216,83,0.25)", bordercolor=BASE_CLR,
                           font=dict(color="#000", size=12))
    st.plotly_chart(bar, use_container_width=True)

if not has_dates:
    st.info("No date column selected — showing aggregated view only.")

# ── Brand snapshot ────────────────────────────────────────────────────────────
with st.expander("📋 Brand snapshot (price per KG, all items)", expanded=False):
    # Apply channel / customer / date filters to the full working dataset
    snap = wdf.copy()
    if channel_col != "(none)" and channel_sel:
        snap = snap[snap[channel_col].astype(str).isin(channel_sel)]
    if customer_col != "(none)" and customer_sel:
        snap = snap[snap[customer_col].astype(str).isin(customer_sel)]
    if date_col != "(none)" and start_date and end_date:
        snap = snap[snap[date_col].dt.date.between(start_date, end_date)]
    if date_col != "(none)" and selected_date is not None:
        snap = snap[snap[date_col].dt.date == selected_date]

    # All SKUs present in the filtered dataset
    all_skus = snap[item_col].dropna().astype(str).sort_values().unique().tolist()

    def _resolve_com(sku: str) -> tuple[str | None, list[str]]:
        """Return (base_brand, [competitors]) for a given SKU using Com map + prefix fallback."""
        entry = com_map.get(sku, {})
        if not entry:
            prefix = sku.split("_")[0].strip()
            for com_key, com_val in com_map.items():
                if com_key.split("_")[0].strip() == prefix:
                    entry = com_val
                    break
        for b in brands_list:
            if b in entry:
                return b, entry[b]
        if entry:
            first_key = next(iter(entry))
            return first_key, entry[first_key]
        return None, []

    # Build numeric price rows
    snap_rows: list[dict] = []
    for sku in all_skus:
        base_b, competitors = _resolve_com(sku)
        brands_for_sku = list(dict.fromkeys(([base_b] if base_b else []) + competitors))
        if not brands_for_sku:
            brands_for_sku = (
                snap.loc[snap[item_col].astype(str) == sku, brand_col]
                .astype(str).unique().tolist()
            )
        row: dict = {"Item": sku, "_base_brand": base_b or ""}
        for b in brands_for_sku:
            pkg_c = brand_pkg_map.get(b) or generic_pkg
            if not pkg_c or pkg_c not in snap.columns:
                row[b] = np.nan
                continue
            vals = snap.loc[
                (snap[item_col].astype(str) == sku) & (snap[brand_col].astype(str) == b),
                pkg_c,
            ].dropna()
            row[b] = float(vals.mean()) if not vals.empty else np.nan
        snap_rows.append(row)

    if not snap_rows:
        st.info("No price-per-KG data for the current selection.")
    else:
        snap_df = pd.DataFrame(snap_rows).set_index("Item")
        base_brand_col_map = snap_df.pop("_base_brand")
        snap_df = snap_df.dropna(axis=1, how="all").dropna(axis=0, how="all")

        # ── Build display dataframe with arrows + % diff for competitor cells ──
        # For each row: competitor cell = "price  ▲/▼ diff (±%)"
        # ▲ red  = competitor HIGHER than base  (base brand is cheaper → good for base)
        # ▼ green = competitor LOWER than base  (competitor is cheaper → threat)
        display_rows: list[dict] = []
        cell_styles: list[dict] = []          # parallel list of style dicts per row

        for sku, row_data in snap_df.iterrows():
            sku_base = base_brand_col_map.get(sku, "")
            base_price = row_data.get(sku_base, np.nan)
            disp_row: dict = {"Item": sku}
            style_row: dict = {"Item": ""}
            for col_name in snap_df.columns:
                val = row_data[col_name]
                if pd.isna(val):
                    disp_row[col_name] = "—"
                    style_row[col_name] = "color: #aaaaaa"
                elif col_name == sku_base:
                    # Base brand: just show price, highlighted yellow
                    disp_row[col_name] = f"{val:,.2f}"
                    style_row[col_name] = "background-color: #fff59d; font-weight: 700"
                else:
                    # Competitor: show price + arrow + diff vs base
                    if pd.notna(base_price) and base_price != 0:
                        diff = val - base_price
                        diff_pct = (diff / base_price) * 100
                        if diff > 0:
                            # Competitor more expensive → base brand is cheaper → good
                            arrow = "▲"
                            style_row[col_name] = "color: #2e7d32; font-weight: 600"  # green
                        else:
                            # Competitor cheaper → threat to base brand
                            arrow = "▼"
                            style_row[col_name] = "color: #d32f2f; font-weight: 600"  # red
                        disp_row[col_name] = f"{val:,.2f}  {arrow} {abs(diff):,.2f} ({diff_pct:+.1f}%)"
                    else:
                        disp_row[col_name] = f"{val:,.2f}"
                        style_row[col_name] = ""
            display_rows.append(disp_row)
            cell_styles.append(style_row)

        disp_df = pd.DataFrame(display_rows).set_index("Item")
        styles_df = pd.DataFrame(cell_styles).set_index("Item")

        # Streamlit styled table
        def _apply_precomputed_styles(df: pd.DataFrame, styles: pd.DataFrame) -> "pd.io.formats.style.Styler":
            styler = df.style
            def _cell_style(val, row_label, col_label):
                return styles.loc[row_label, col_label] if col_label in styles.columns else ""
            # apply cell-wise
            for col in df.columns:
                styler = styler.apply(
                    lambda s, c=col: [
                        styles.loc[idx, c] if idx in styles.index and c in styles.columns else ""
                        for idx in s.index
                    ],
                    subset=[col],
                    axis=0,
                )
            return styler

        st.caption(
            f"📊 **{len(snap_df)} SKUs** — price per KG after promotion. "
            "Base brand highlighted 🟡 | "
            "**▲ green** = competitor higher than base | **▼ red** = competitor lower than base"
        )
        styled_disp = _apply_precomputed_styles(disp_df, styles_df)
        st.dataframe(styled_disp, use_container_width=True)

        # ── Excel export with colored cells + arrow text ──────────────────────
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            wb = writer.book
            ws = wb.add_worksheet("Brand Snapshot")

            # Define formats
            header_fmt = wb.add_format({
                "bold": True, "bg_color": "#1b5e20", "font_color": "#ffffff",
                "border": 1, "align": "center", "valign": "vcenter",
            })
            base_fmt = wb.add_format({
                "bg_color": "#fff59d", "bold": True, "border": 1,
                "num_format": "#,##0.00", "align": "center",
            })
            comp_higher_fmt = wb.add_format({
                # competitor more expensive → green text (▲)
                "border": 1, "align": "center", "font_color": "#1b5e20",
                "bold": True, "text_wrap": True,
            })
            comp_lower_fmt = wb.add_format({
                # competitor cheaper → red text (▼)
                "border": 1, "align": "center", "font_color": "#c62828",
                "bold": True, "text_wrap": True,
            })
            na_fmt = wb.add_format({
                "border": 1, "align": "center", "font_color": "#aaaaaa",
            })
            item_fmt = wb.add_format({
                "bold": True, "border": 1, "text_wrap": True, "valign": "vcenter",
            })
            subheader_fmt = wb.add_format({
                "bold": True, "bg_color": "#e8f5e9", "border": 1,
                "align": "center", "italic": True, "font_color": "#555555",
            })

            all_brand_cols = snap_df.columns.tolist()
            all_export_cols = ["Item"] + all_brand_cols

            # Header row
            for ci, col_name in enumerate(all_export_cols):
                ws.write(0, ci, col_name, header_fmt)

            # Sub-header row: label each brand column as "Base" or "Competitor"
            ws.write(1, 0, "", subheader_fmt)
            for ci, col_name in enumerate(all_brand_cols, start=1):
                # Check if this brand is ever a base brand
                is_base_col = any(
                    base_brand_col_map.get(sku, "") == col_name for sku in snap_df.index
                )
                label = "◆ Base Brand" if is_base_col else "Competitor"
                ws.write(1, ci, label, subheader_fmt)

            # Data rows (start at row 2 because of sub-header)
            for ri, (sku, row_data) in enumerate(snap_df.iterrows(), start=2):
                ws.write(ri, 0, sku, item_fmt)
                ws.set_row(ri, 30)
                sku_base = base_brand_col_map.get(sku, "")
                base_price = row_data.get(sku_base, np.nan)

                for ci, col_name in enumerate(all_brand_cols, start=1):
                    val = row_data[col_name]
                    if pd.isna(val):
                        ws.write(ri, ci, "—", na_fmt)
                    elif col_name == sku_base:
                        ws.write_number(ri, ci, float(val), base_fmt)
                    else:
                        if pd.notna(base_price) and base_price != 0:
                            diff = float(val) - float(base_price)
                            diff_pct = (diff / float(base_price)) * 100
                            arrow = "▲" if diff > 0 else "▼"
                            text = f"{val:,.2f}  {arrow} {abs(diff):,.2f} ({diff_pct:+.1f}%)"
                            fmt = comp_higher_fmt if diff > 0 else comp_lower_fmt
                        else:
                            text = f"{val:,.2f}"
                            fmt = wb.add_format({"border": 1, "align": "center"})
                        ws.write(ri, ci, text, fmt)

            # Column widths + freeze
            ws.set_column(0, 0, 42)
            for ci in range(1, len(all_export_cols)):
                ws.set_column(ci, ci, 22)
            ws.freeze_panes(2, 1)

            # Legend box below data
            legend_row = len(snap_df) + 4
            legend_title_fmt = wb.add_format({"bold": True, "font_size": 11})
            ws.write(legend_row, 0, "Legend:", legend_title_fmt)
            ws.write(legend_row + 1, 0, "🟡 Yellow = Base brand price",
                     wb.add_format({"bg_color": "#fff59d", "border": 1}))
            ws.write(legend_row + 2, 0, "▲ Green = Competitor HIGHER than base (base brand is cheaper)",
                     wb.add_format({"font_color": "#1b5e20", "bold": True, "border": 1}))
            ws.write(legend_row + 3, 0, "▼ Red = Competitor LOWER than base (competitive threat)",
                     wb.add_format({"font_color": "#c62828", "bold": True, "border": 1}))

        buf.seek(0)
        st.download_button(
            "⬇️ Download full snapshot (Excel)",
            data=buf.getvalue(),
            file_name=f"brand_snapshot_all_{pd.Timestamp.now():%Y%m%d%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with st.expander("🔍 Filtered records"):
    st.dataframe(filtered, use_container_width=True, hide_index=True)

st.caption("pip install streamlit pandas numpy plotly xlsxwriter pyxlsb")