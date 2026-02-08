"""Streamlit app for analyzing item price indices by brand and date."""
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


st.set_page_config(page_title="Price Index Explorer", layout="wide")
st.markdown(
    """
    <style>
    /* Reduce default padding for Streamlit dataframes when showing styled HTML */
    .brand-snapshot-table td { padding: 0.4rem 0.6rem !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

DEFAULT_SHEET = "Market"
SUPPORTED_EXTENSIONS = {".xlsb", ".xlsx", ".xls"}
FOCUS_BRANDS = ["Atyab", "Chikitita", "Meatland"]
BRANDS_SHEET_NAME = "Brands"
DEFAULT_ITEM_DESCRIPTION = "Meat_Chilled Luncheon_Retail_5Kg_1_Luncheon"
ITEM_SELECT_KEY = "price_index_item_select"
ITEM_INPUT_KEY = "price_index_item_input"
ITEM_LAST_SELECTED_KEY = "price_index_item_last_selected"
SHOW_PRICE_LABELS_KEY = "price_index_show_price_labels"
BRAND_ALIASES: dict[str, tuple[str, ...]] = {
    "Atyab": ("atyab",),
    "Chikitita": ("chikitita", "chikitia"),
    "Meatland": ("meatland",),
}


def _normalize(name: str) -> str:
    return str(name).strip().lower()


def _match_brand_from_text(text: str) -> Optional[str]:
    lowered = _normalize(text)
    for brand, aliases in BRAND_ALIASES.items():
        if any(alias in lowered for alias in aliases):
            return brand
    return None


def _build_brand_metric_map(columns: Iterable[str], required_terms: Iterable[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    lowered_terms = [term.lower() for term in required_terms]
    for column in columns:
        lowered = _normalize(column)
        if not all(term in lowered for term in lowered_terms):
            continue
        matched_brand = _match_brand_from_text(column)
        if matched_brand and matched_brand not in mapping:
            mapping[matched_brand] = column
    return mapping


def _build_brand_price_map(columns: Iterable[str]) -> dict[str, str]:
    """Backward-compatible wrapper for previous helper name."""
    return _build_brand_metric_map(columns, ["price", "index"])


def _find_generic_price_per_kg_column(columns: Iterable[str]) -> Optional[str]:
    search_term_sequences = [
        ("price", "per", "kg", "promot"),
        ("price", "per", "kg"),
    ]
    for terms in search_term_sequences:
        for column in columns:
            lowered = _normalize(column)
            if all(term in lowered for term in terms):
                return column
    return None


def _guess_column(columns: Iterable[str], keywords: Iterable[str]) -> Optional[str]:
    for column in columns:
        lowered = _normalize(column)
        if all(keyword in lowered for keyword in keywords):
            return column
    return None


def _guess_item_column(columns: Iterable[str]) -> Optional[str]:
    for candidate in columns:
        lowered = _normalize(candidate)
        if "item" in lowered and ("desc" in lowered or "description" in lowered or "name" in lowered):
            return candidate
    return None


def _find_price_columns(df: pd.DataFrame) -> list[str]:
    matches: list[str] = []
    for column in df.columns:
        lowered = _normalize(column)
        if "price" in lowered and "index" in lowered:
            matches.append(column)
    if matches:
        return matches
    numeric_columns = df.select_dtypes(include=["number"]).columns.tolist()
    return numeric_columns


def _convert_excel_serial(value: float | int | str | None) -> pd.Timestamp | pd.NaT:
    if value is None:
        return pd.NaT
    try:
        num = float(value)
    except (TypeError, ValueError):
        return pd.NaT
    if not np.isfinite(num):
        return pd.NaT
    if num < 59:
        base = pd.Timestamp("1899-12-30")
    else:
        base = pd.Timestamp("1899-12-30")
    days = pd.to_timedelta(num, unit="D")
    return base + days


def _normalize_percent_series(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    valid = numeric.dropna()
    if valid.empty:
        return numeric
    median = valid.median()
    if median > 2:
        numeric = numeric / 100.0
    return numeric


def _format_price_per_kg_label(val: float | int | None) -> str:
    if val is None or pd.isna(val):
        return "—"
    return f"{float(val):,.2f}"


def _format_percent_with_per_kg(percent_val: float | int | None, per_kg_val: float | int | None) -> str:
    if percent_val is None or pd.isna(percent_val):
        percent_text = "—"
    else:
        try:
            percent_text = f"{float(percent_val):.1%}"
        except (TypeError, ValueError):
            percent_text = str(percent_val)
    if per_kg_val is None or pd.isna(per_kg_val):
        return percent_text
    per_kg_text = _format_price_per_kg_label(per_kg_val)
    return f"{percent_text} ({per_kg_text})"


def _highlight_base_column(col: pd.Series) -> list[str]:
    if col.name == base_brand:
        return ["background-color: #fff59d; font-weight: 600" for _ in col]
    return ["" for _ in col]


def _format_customer_count(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    try:
        count_int = int(round(float(value)))
    except (TypeError, ValueError):
        return ""
    if count_int <= 0:
        return ""
    return "1 customer" if count_int == 1 else f"{count_int} customers"


def _customer_border_color(value: Any, fallback_color: str) -> str:
    if value is None or pd.isna(value):
        return fallback_color
    try:
        count_int = int(round(float(value)))
    except (TypeError, ValueError):
        return fallback_color
    if count_int == 1:
        return "#ff1744"
    if count_int >= 3:
        return "#1e88e5"
    return fallback_color


def _default_brand_window(brands: list[str], focus: str) -> list[str]:
    if focus not in brands:
        return brands[:3] if brands else []
    idx = brands.index(focus)
    window: list[str] = []
    if idx > 0:
        window.append(brands[idx - 1])
    window.append(focus)
    if idx < len(brands) - 1:
        window.append(brands[idx + 1])
    return window


@st.cache_data(show_spinner=False)
def load_dataframe(upload: BytesIO, *, sheet_name: str) -> pd.DataFrame:
    name = getattr(upload, "name", "uploaded_file")
    suffix = Path(name).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported extension '{suffix}'. Please upload one of: {', '.join(sorted(SUPPORTED_EXTENSIONS))}."
        )
    upload.seek(0)
    if suffix == ".xlsb":
        try:
            import pyxlsb  # type: ignore  # noqa: F401
        except ImportError as exc:  # pragma: no cover - runtime feedback inside Streamlit
            raise ImportError(
                "Reading .xlsb files requires the optional 'pyxlsb' package. Install it via 'pip install pyxlsb'."
            ) from exc
        df = pd.read_excel(upload, sheet_name=sheet_name, engine="pyxlsb")
    else:
        df = pd.read_excel(upload, sheet_name=sheet_name)
    upload.seek(0)
    return df


st.title("📊 Price Index Explorer")
st.caption("Upload market tracker data to benchmark Atyab, Chikitita, and Meatland across time.")

with st.sidebar:
    st.header("1️⃣ Data source")
    uploaded_file = st.file_uploader("Upload Excel (.xlsb/.xlsx/.xls)", type=["xlsb", "xlsx", "xls"], accept_multiple_files=False)
    sheet_name = st.text_input("Sheet name", value=DEFAULT_SHEET)

    if uploaded_file is None:
        st.info("Upload a file to begin.")
        st.stop()

    try:
        raw_df = load_dataframe(uploaded_file, sheet_name=sheet_name or DEFAULT_SHEET)
    except Exception as err:
        st.error(f"Failed to load workbook: {err}")
        st.stop()

if raw_df.empty:
    st.warning("The selected worksheet is empty.")
    st.stop()

raw_df.columns = [str(col).strip() for col in raw_df.columns]

priority_brand_list: list[str] = []
try:
    brands_df = load_dataframe(uploaded_file, sheet_name=BRANDS_SHEET_NAME)
except ValueError:
    brands_df = pd.DataFrame()
except Exception as exc:
    st.warning(f"Unable to read '{BRANDS_SHEET_NAME}' sheet: {exc}")
    brands_df = pd.DataFrame()

if not brands_df.empty:
    first_column = brands_df.columns[0]
    series = brands_df[first_column].dropna().astype(str).str.strip()
    priority_brand_list = [brand for brand in series.tolist() if brand]
    # Preserve ordering but remove duplicates
    priority_brand_list = list(dict.fromkeys(priority_brand_list))

columns = raw_df.columns.tolist()
brand_price_column_map = _build_brand_metric_map(columns, ["price", "index"])
brand_price_per_kg_map = _build_brand_metric_map(columns, ["price", "per", "kg", "promoted"])
generic_price_per_kg_col = _find_generic_price_per_kg_column(columns)

if generic_price_per_kg_col:
    for focus_brand in FOCUS_BRANDS:
        brand_price_per_kg_map.setdefault(focus_brand, generic_price_per_kg_col)

with st.sidebar:
    st.header("2️⃣ Column mapping")
    brand_col_guess = _guess_column(columns, ["brand"])
    item_col_guess = _guess_item_column(columns)
    customer_col_guess = _guess_column(columns, ["customer"])
    channel_col_guess = _guess_column(columns, ["channel"])
    date_col_guess = _guess_column(columns, ["date"])
    price_candidates = _find_price_columns(raw_df)

    brand_col = st.selectbox("Brand column", options=columns, index=columns.index(brand_col_guess) if brand_col_guess in columns else 0)
    item_col = st.selectbox("Item description column", options=columns, index=columns.index(item_col_guess) if item_col_guess in columns else 0)
    channel_col = st.selectbox(
        "Sales channel column",
        options=["(none)"] + columns,
        index=(columns.index(channel_col_guess) + 1) if channel_col_guess in columns else 0,
        help="Choose the column that stores sales channel names (optional).",
    )
    customer_col = st.selectbox(
        "Customer name column",
        options=["(none)"] + columns,
        index=(columns.index(customer_col_guess) + 1) if customer_col_guess in columns else 0,
        help="Choose the column that stores customer names (optional).",
    )
    date_col = st.selectbox(
        "Date column",
        options=["(none)"] + columns,
        index=(columns.index(date_col_guess) + 1) if date_col_guess in columns else 0,
        help="Choose the column that stores transaction or observation dates.",
    )
    selected_price_cols = st.multiselect(
        "Price index column(s)",
        options=columns,
        default=[col for col in price_candidates if col in columns][: max(1, len(price_candidates))],
        help="Select the numeric column(s) that contain price index values.",
    )

if not selected_price_cols:
    st.error("Select at least one price index column to continue.")
    st.stop()

working_df = raw_df.copy()

if date_col != "(none)":
    raw_dates = working_df[date_col]
    dt_series = pd.to_datetime(raw_dates, errors="coerce", infer_datetime_format=True)
    excel_candidates = raw_dates.apply(_convert_excel_serial)
    excel_mask = excel_candidates.notna()
    dt_series = dt_series.mask(excel_mask, excel_candidates)
    if pd.api.types.is_datetime64tz_dtype(dt_series):
        dt_series = dt_series.dt.tz_convert(None)
    dt_series = dt_series.dt.normalize()
    working_df[date_col] = dt_series

columns_to_normalize: set[str] = set(selected_price_cols)
columns_to_normalize.update(brand_price_column_map.values())
for col in columns_to_normalize:
    if col not in working_df.columns:
        continue
    working_df[col] = _normalize_percent_series(
        working_df[col].replace({"-": np.nan, "": np.nan})
    )

per_kg_columns_to_numeric = set(brand_price_per_kg_map.values())
if generic_price_per_kg_col:
    per_kg_columns_to_numeric.add(generic_price_per_kg_col)

for col in per_kg_columns_to_numeric:
    if col not in working_df.columns:
        continue
    cleaned_series = (
        working_df[col]
        .astype(str)
        .str.strip()
        .replace({"-": np.nan, "": np.nan})
    )
    cleaned_series = cleaned_series.str.replace(r"[^0-9,.-]", "", regex=True)
    cleaned_series = cleaned_series.str.replace(",", "", regex=False)
    working_df[col] = pd.to_numeric(cleaned_series, errors="coerce")

required_subset = [brand_col, item_col]
if customer_col != "(none)":
    required_subset.append(customer_col)
working_df = working_df.dropna(subset=required_subset)

available_items = working_df[item_col].astype(str).sort_values().unique().tolist()

if not available_items:
    st.error("No item descriptions found after cleaning. Adjust the column mapping.")
    st.stop()

default_item_index = (
    available_items.index(DEFAULT_ITEM_DESCRIPTION)
    if DEFAULT_ITEM_DESCRIPTION in available_items
    else 0
)

st.markdown("### 🔍 Item selection")
selected_item = st.selectbox(
    "Item description",
    options=available_items,
    index=default_item_index,
    key=ITEM_SELECT_KEY,
    help="Choose a starting item. You can tweak the text below to jump to a similar description without retyping it from scratch.",
)

if st.session_state.get(ITEM_LAST_SELECTED_KEY) != selected_item:
    st.session_state[ITEM_INPUT_KEY] = selected_item
    st.session_state[ITEM_LAST_SELECTED_KEY] = selected_item

editable_item = st.text_input(
    "Refine item description",
    key=ITEM_INPUT_KEY,
    help="Edit the selected description. A close match from the data will be used automatically.",
).strip()

resolved_item = selected_item
if editable_item:
    if editable_item in available_items:
        resolved_item = editable_item
    else:
        partial_matches = [item for item in available_items if editable_item.lower() in item.lower()]
        if partial_matches:
            resolved_item = partial_matches[0]
        else:
            close_match = get_close_matches(editable_item, available_items, n=1, cutoff=0.0)
            if close_match:
                resolved_item = close_match[0]
            else:
                st.warning("No item matches the text you entered. Using the original selection.")

if resolved_item != selected_item:
    st.session_state[ITEM_SELECT_KEY] = resolved_item
    selected_item = resolved_item
    st.session_state[ITEM_INPUT_KEY] = resolved_item
    st.session_state[ITEM_LAST_SELECTED_KEY] = resolved_item
else:
    selected_item = resolved_item

item_mask = working_df[item_col].astype(str) == selected_item
item_subset = working_df.loc[item_mask].copy()

customer_selection: Optional[list[str]] = None
channel_selection: Optional[list[str]] = None

filters_container = st.sidebar.container()

with filters_container:
    st.header("3️⃣ Additional filters")
    st.caption("Refine analysis with customer, brand, and date filters.")

    if channel_col != "(none)" and channel_col in item_subset.columns:
        item_subset.loc[:, channel_col] = (
            item_subset[channel_col].astype(str).str.strip().replace({"": np.nan, "nan": np.nan})
        )
        channel_options = (
            item_subset[channel_col].dropna().astype(str).sort_values().unique().tolist()
        )
        if channel_options:
            channel_selection = st.multiselect(
                "Sales channel",
                options=channel_options,
                default=channel_options,
                help="Filter the view to specific sales channels.",
            )
            if not channel_selection:
                st.warning("Select at least one sales channel to continue.")
                st.stop()
            channel_mask = item_subset[channel_col].astype(str).isin(channel_selection)
            item_subset = item_subset.loc[channel_mask]
        else:
            st.info("No sales channels available for the selected item.")

    if customer_col != "(none)":
        customer_options = (
            item_subset[customer_col]
            .dropna()
            .astype(str)
            .sort_values()
            .unique()
            .tolist()
        )
        if customer_options:
            customer_selection = st.multiselect(
                "Customer name",
                options=customer_options,
                default=customer_options,
                help="Filter results to specific customers.",
            )
            if not customer_selection:
                st.warning("Select at least one customer to continue.")
                st.stop()
            customer_mask = item_subset[customer_col].astype(str).isin(customer_selection)
            item_subset = item_subset.loc[customer_mask]
        else:
            st.info("No customer names available for the selected item.")

    available_brands = (
        item_subset[brand_col]
        .astype(str)
        .sort_values()
        .unique()
        .tolist()
    )
    if not available_brands:
        st.warning("No brands found for the selected item and customer filters. Adjust the selections.")
        st.stop()

preferred_brands = priority_brand_list or FOCUS_BRANDS
base_brand_options = [brand for brand in preferred_brands if brand in available_brands]

if not base_brand_options:
    base_brand_options = available_brands

if not base_brand_options:
    st.warning("Atyab, Chikitita, and Meatland are not available for the current selections. Adjust filters to include one of them.")
    st.stop()

default_base_brand = base_brand_options[0]

st.markdown("### ⚖️ Base brand selection")
base_brand = st.selectbox(
    "Base brand for price comparison",
    options=base_brand_options,
    index=base_brand_options.index(default_base_brand) if default_base_brand in base_brand_options else 0,
    help="All brands for the chosen item are shown. The base brand acts as the reference for price index deltas.",
)

base_price_col = brand_price_column_map.get(base_brand)
if base_price_col is None:
    st.warning(
        "Could not find a price index column for the selected base brand. Ensure the sheet includes a column whose name contains the brand."
    )
    st.stop()

base_price_per_kg_col = brand_price_per_kg_map.get(base_brand)
if base_price_per_kg_col is None and generic_price_per_kg_col:
    base_price_per_kg_col = generic_price_per_kg_col

selected_price_cols = [base_price_col]
caption_parts = [f"Using price index column: {base_price_col}"]
if base_price_per_kg_col:
    caption_parts.append(f"Price per KG promoted column: {base_price_per_kg_col}")
st.caption(" | ".join(caption_parts))

with filters_container:
    default_brand_selection = _default_brand_window(available_brands, base_brand)
    brand_selection = st.multiselect(
        "Brands to display",
        options=available_brands,
        default=available_brands,
        help="All brands are selected by default. Adjust to focus analysis if needed.",
    )
    if base_brand not in brand_selection:
        brand_selection.append(base_brand)
    ordered_unique: list[str] = []
    for brand in available_brands:
        if brand in brand_selection and brand not in ordered_unique:
            ordered_unique.append(brand)
    brand_selection = ordered_unique or [base_brand]
    initial_visible_brands: set[str] = set(default_brand_selection or [base_brand])
    initial_visible_brands.add(base_brand)

    if date_col != "(none)" and working_df[date_col].notna().any():
        date_series = (
            working_df.loc[working_df[item_col].astype(str) == selected_item, date_col]
            .dropna()
            .sort_values()
        )
        if not date_series.empty:
            date_options = sorted({ts.date() for ts in date_series})
            if len(date_options) == 1:
                only_date = date_options[0]
                st.caption(f"Only one date available: {only_date:%Y-%m-%d}")
                start_date = end_date = only_date
                selected_date = only_date
            else:
                start_date, end_date = st.select_slider(
                    "Date range",
                    options=date_options,
                    value=(
                        next(
                            (
                                dt
                                for dt in date_options
                                if dt >= (date_options[-1] - timedelta(days=30))
                            ),
                            date_options[0],
                        ),
                        date_options[-1],
                    ),
                    format_func=lambda dt: dt.strftime("%Y-%m-%d"),
                )
                enable_single_date = st.checkbox(
                    "Filter to a single date",
                    value=False,
                    help="Enable to focus charts and records on one date within the selected range.",
                )
                selected_date = None
                if enable_single_date:
                    default_focus = min(max(end_date, start_date), end_date)
                    selected_date = st.date_input(
                        "Select date",
                        value=default_focus,
                        min_value=start_date,
                        max_value=end_date,
                        format="YYYY-MM-DD",
                    )
        else:
            start_date = end_date = selected_date = None
    else:
        start_date = end_date = selected_date = None


if not brand_selection:
    st.warning("Select at least one brand.")
    st.stop()

filtered = working_df[working_df[item_col].astype(str) == selected_item]
filtered = filtered[filtered[brand_col].astype(str).isin(brand_selection)]

if customer_col != "(none)" and customer_selection:
    filtered = filtered[filtered[customer_col].astype(str).isin(customer_selection)]

if date_col != "(none)" and start_date and end_date:
    mask = filtered[date_col].dt.date.between(start_date, end_date)
    filtered = filtered[mask]

if date_col != "(none)" and selected_date is not None:
    filtered = filtered[filtered[date_col].dt.date == selected_date]

if filtered.empty:
    st.warning("No records match the current filters.")
    st.stop()

brand_per_kg_columns: dict[str, str] = {}
for brand in brand_selection:
    per_kg_col = brand_price_per_kg_map.get(brand)
    if per_kg_col is None and generic_price_per_kg_col:
        per_kg_col = generic_price_per_kg_col
    if per_kg_col and per_kg_col in filtered.columns:
        brand_per_kg_columns[brand] = per_kg_col

brand_latest_per_kg: dict[str, float] = {}
brand_per_kg_by_date: dict[str, dict[str, float]] = {}
brand_str_series = filtered[brand_col].astype(str)
has_valid_dates = (
    date_col != "(none)"
    and date_col in filtered.columns
    and filtered[date_col].notna().any()
)

for brand, per_kg_col in brand_per_kg_columns.items():
    brand_mask = brand_str_series == brand
    brand_subset = filtered.loc[brand_mask].copy()
    if brand_subset.empty:
        continue
    per_series = brand_subset[per_kg_col].dropna()
    if per_series.empty:
        continue
    if has_valid_dates:
        dated_subset = brand_subset.dropna(subset=[date_col, per_kg_col])
        if not dated_subset.empty:
            dated_subset = dated_subset.sort_values(by=date_col)
            brand_latest_per_kg[brand] = float(dated_subset.iloc[-1][per_kg_col])
            grouped = dated_subset.groupby(dated_subset[date_col])[per_kg_col].mean()
            brand_per_kg_by_date[brand] = {
                pd.to_datetime(idx).strftime("%Y-%m-%d"): float(val)
                for idx, val in grouped.items()
            }
            continue
    brand_latest_per_kg[brand] = float(per_series.iloc[-1])


def _lookup_per_kg_value(brand: str, x_value: Any) -> float | None:
    per_kg_map = brand_per_kg_by_date.get(brand)
    if per_kg_map:
        try:
            dt_value = pd.to_datetime(x_value)
        except Exception:  # pragma: no cover - defensive against unexpected axis data types
            dt_value = None
        if dt_value is not None and not pd.isna(dt_value):
            key = pd.to_datetime(dt_value).strftime("%Y-%m-%d")
            mapped = per_kg_map.get(key)
            if mapped is not None:
                return mapped
    return brand_latest_per_kg.get(brand)


base_brand_data = filtered[filtered[brand_col].astype(str) == base_brand]
if date_col != "(none)" and date_col in filtered.columns and filtered[date_col].notna().any():
    base_brand_data = base_brand_data.sort_values(by=date_col)

st.subheader(f"Item: {selected_item}")

controls_col, _ = st.columns([1, 6])
with controls_col:
    if SHOW_PRICE_LABELS_KEY not in st.session_state:
        st.session_state[SHOW_PRICE_LABELS_KEY] = True
    st.checkbox(
        "Show price values in chart labels",
        key=SHOW_PRICE_LABELS_KEY,
        help="Toggle per-kg price information inside chart annotations.",
    )
    show_price_labels = st.session_state[SHOW_PRICE_LABELS_KEY]

st.caption(
    "Adjust selections in the sidebar to refine the analysis. Values are aggregated by brand and date where applicable."
)

if date_col != "(none)" and filtered[date_col].notna().any():
    filtered = filtered.sort_values(by=[date_col, brand_col])

snapshot_df = working_df.copy()
if channel_col != "(none)" and channel_selection:
    snapshot_df = snapshot_df[snapshot_df[channel_col].astype(str).isin(channel_selection)]
if customer_col != "(none)" and customer_selection:
    snapshot_df = snapshot_df[snapshot_df[customer_col].astype(str).isin(customer_selection)]
if brand_selection:
    snapshot_df = snapshot_df[snapshot_df[brand_col].astype(str).isin(brand_selection)]
if date_col != "(none)" and start_date and end_date:
    snapshot_df = snapshot_df[snapshot_df[date_col].dt.date.between(start_date, end_date)]
if date_col != "(none)" and selected_date is not None:
    snapshot_df = snapshot_df[snapshot_df[date_col].dt.date == selected_date]

items_index = (
    snapshot_df[item_col]
    .dropna()
    .astype(str)
    .sort_values()
    .unique()
    .tolist()
)

pivot_summary = pd.DataFrame(index=items_index)
for brand_name in brand_selection:
    price_col = brand_price_per_kg_map.get(brand_name)
    if price_col is None and generic_price_per_kg_col:
        price_col = generic_price_per_kg_col
    if price_col is None or price_col not in working_df.columns:
        continue
    brand_mask = snapshot_df[brand_col].astype(str) == brand_name
    brand_subset = snapshot_df.loc[brand_mask, [item_col, price_col]].copy()
    if brand_subset.empty:
        continue
    brand_subset[item_col] = brand_subset[item_col].astype(str)
    avg_by_item = brand_subset.groupby(item_col)[price_col].mean()
    pivot_summary[brand_name] = avg_by_item

pivot_summary.index.name = "Item description"
pivot_summary = pivot_summary.dropna(how="all")
present_columns = [brand for brand in brand_selection if brand in pivot_summary.columns]
other_columns = [col for col in pivot_summary.columns if col not in present_columns]
pivot_summary = pivot_summary.loc[:, present_columns + other_columns]

formatters = {col: _format_price_per_kg_label for col in pivot_summary.columns}

if not pivot_summary.empty:
    if base_brand in pivot_summary.columns:
        pivot_summary = pivot_summary.sort_values(by=base_brand, ascending=False)

styled_summary = pivot_summary.style.format(formatters).apply(_highlight_base_column, axis=0)

highlight_color = "#0d47a1"
muted_color = "#b0bec5"
base_bar_color = "#1b5e20"

palette = px.colors.qualitative.Safe + px.colors.qualitative.Pastel
color_cycle = cycle(palette)
brand_color_map: dict[str, str] = {}
for brand in brand_selection:
    if brand == base_brand:
        brand_color_map[brand] = base_bar_color
    else:
        brand_color_map[brand] = next(color_cycle)

highlight_color = brand_color_map.get(base_brand, base_bar_color)

relative_traces: dict[str, pd.DataFrame] = {}
if date_col != "(none)" and filtered[date_col].notna().any() and base_brand in filtered[brand_col].unique():
    for price_col in selected_price_cols:
        pivot = (
            filtered[[date_col, brand_col, price_col]]
            .dropna(subset=[price_col])
            .pivot_table(index=date_col, columns=brand_col, values=price_col, aggfunc="mean")
        )
        if pivot.empty or base_brand not in pivot.columns:
            continue
        base_series = pivot[base_brand]
        valid_index = base_series.replace(0, np.nan).dropna().index
        if len(valid_index) == 0:
            continue
        pivot = pivot.loc[valid_index]
        base_series = base_series.loc[valid_index]
        relative = pivot.divide(base_series, axis=0) - 1
        first_base_value = base_series.iloc[0] if not base_series.empty else np.nan
        if pd.notna(first_base_value) and not np.isclose(first_base_value, 0.0):
            base_trend = (base_series / first_base_value) - 1
        else:
            base_trend = base_series.pct_change().fillna(0)
        relative[base_brand] = base_trend
        relative = relative.replace([np.inf, -np.inf], np.nan).reset_index()
        relative = relative.melt(id_vars=[date_col], var_name=brand_col, value_name="relative_value")
        relative = relative.dropna(subset=["relative_value"])
        relative["relative_value"] = relative["relative_value"].where(~np.isclose(relative["relative_value"], 0.0, atol=1e-9), 0.0)
        relative_traces[price_col] = relative

if date_col != "(none)" and filtered[date_col].notna().any():
    for price_col in selected_price_cols:
        plot_df = relative_traces.get(price_col)
        if plot_df is not None and not plot_df.empty:
            fig = px.line(
                plot_df,
                x=date_col,
                y="relative_value",
                color=brand_col,
                markers=True,
                title=f"{price_col} Δ vs {base_brand}",
                labels={date_col: "Date", "relative_value": "Δ vs base"},
                color_discrete_map=brand_color_map,
            )
            fig.update_yaxes(tickformat=".0%", tickfont=dict(color="#111111"))
        else:
            fig = px.line(
                filtered,
                x=date_col,
                y=price_col,
                color=brand_col,
                markers=True,
                title=f"{price_col} trend",
                labels={date_col: "Date", price_col: price_col, brand_col: "Brand"},
                color_discrete_map=brand_color_map,
            )
            fig.update_yaxes(tickformat=".0%", tickfont=dict(color="#111111"))
        fig.update_traces(
            texttemplate="%{text}",
            textposition="middle center",
            textfont=dict(size=11, color="#111111"),
            mode="lines+markers+text",
        )
        fig.update_layout(hovermode="x unified", font=dict(color="#111111"))
        fig.update_xaxes(tickfont=dict(color="#111111"))
        for trace_index, trace in enumerate(fig.data):
            y_values = trace.y if trace.y is not None else []
            x_values = trace.x if trace.x is not None else []
            formatted_text: list[str] = []
            text_positions: list[str] = []
            seen_points: set[tuple[Any, Any]] = set()
            for idx, y in enumerate(y_values):
                label_prefix = f"{trace.name}: " if idx == 0 else ""
                x_val = x_values[idx] if idx < len(x_values) else idx
                point_key = (trace.name, x_val)
                if point_key in seen_points:
                    formatted_text.append("")
                    text_positions.append("top center" if (idx % 2 == 0) else "bottom center")
                    continue
                seen_points.add(point_key)
                if pd.isna(y):
                    formatted_text.append(f"{trace.name}" if idx == 0 else "")
                    text_positions.append("top center" if (idx % 2 == 0) else "bottom center")
                else:
                    try:
                        value_text = f"{float(y):.1%}"
                    except (TypeError, ValueError):
                        value_text = str(y)
                    per_kg_value = None
                    if idx < len(x_values):
                        per_kg_value = _lookup_per_kg_value(trace.name, x_values[idx])
                    if show_price_labels and per_kg_value is not None and not pd.isna(per_kg_value):
                        per_kg_label = _format_price_per_kg_label(per_kg_value)
                        value_text = f"{value_text} ({per_kg_label})"
                    formatted_text.append(f"{label_prefix}{value_text}")
                    text_positions.append("top center" if (idx % 2 == 0) else "bottom center")
            if formatted_text:
                trace.text = formatted_text
            else:
                x_values = trace.x if trace.x is not None else []
                trace.text = [str(trace.name)] + [""] * (len(x_values) - 1)
                text_positions = ["top center" if (idx % 2 == 0) else "bottom center" for idx in range(len(trace.text))]
            color = brand_color_map.get(trace.name, muted_color)
            if trace.name == base_brand:
                trace.update(line=dict(width=2, color=color), marker=dict(size=10, color=color))
            else:
                trace.update(line=dict(width=1, color=color), marker=dict(size=6, color=color), opacity=0.75)
            if trace.name not in initial_visible_brands:
                trace.visible = "legendonly"
            if len(text_positions) < len(trace.text):
                text_positions.extend([
                    "top center" if (idx % 2 == 0) else "bottom center" for idx in range(len(text_positions), len(trace.text))
                ])
            trace.update(textposition=text_positions, textfont=dict(size=11, color=color))
        st.plotly_chart(fig, use_container_width=True)

    per_kg_trend_rows: list[pd.DataFrame] = []
    if brand_per_kg_columns and date_col != "(none)" and filtered[date_col].notna().any():
        for brand in brand_selection:
            per_kg_col = brand_per_kg_columns.get(brand)
            if per_kg_col is None:
                continue
            brand_slice = filtered.loc[
                filtered[brand_col].astype(str) == brand, [date_col, per_kg_col]
            ].dropna()
            if brand_slice.empty:
                continue
            daily_avg = (
                brand_slice.groupby(date_col, as_index=False)[per_kg_col]
                .mean()
                .rename(columns={per_kg_col: "price_per_kg"})
            )
            daily_avg[brand_col] = brand
            per_kg_trend_rows.append(daily_avg)

    customer_count_pivot = None
    if per_kg_trend_rows:
        per_kg_trend_df = pd.concat(per_kg_trend_rows, ignore_index=True)
        per_kg_trend_df = per_kg_trend_df.sort_values(date_col)
        per_kg_pivot = (
            per_kg_trend_df.pivot_table(
                index=date_col,
                columns=brand_col,
                values="price_per_kg",
                aggfunc="mean",
            )
            .sort_index()
            .replace([np.inf, -np.inf], np.nan)
            .dropna(how="all")
        )
        if "customer_count" in per_kg_trend_df.columns:
            customer_count_pivot = (
                per_kg_trend_df.pivot_table(
                    index=date_col,
                    columns=brand_col,
                    values="customer_count",
                    aggfunc="first",
                )
                .reindex(per_kg_pivot.index)
                .sort_index()
            )
        if not per_kg_pivot.empty:
            price_fig = go.Figure()
            base_series_per_date = per_kg_pivot[base_brand] if base_brand in per_kg_pivot.columns else None
            other_max_per_date = None
            if base_series_per_date is not None:
                other_cols = [col for col in per_kg_pivot.columns if col != base_brand]
                if other_cols:
                    other_max_per_date = per_kg_pivot[other_cols].max(axis=1)
            for brand in per_kg_pivot.columns:
                series = per_kg_pivot[brand].dropna()
                if series.empty:
                    continue
                color = brand_color_map.get(brand, muted_color)
                marker_colors: list[str] | str = color
                count_values: np.ndarray | list[Any]
                if customer_count_pivot is not None and brand in customer_count_pivot.columns:
                    aligned_counts = customer_count_pivot[brand].reindex(series.index)
                    count_values = aligned_counts.values
                else:
                    count_values = [None] * len(series)
                marker_border_colors: list[str] = []
                if base_series_per_date is not None and brand != base_brand:
                    aligned_base = base_series_per_date.reindex(series.index)
                    label_text = []
                    marker_colors = []
                    for val, base_val, count_val in zip(series.values, aligned_base.values, count_values):
                        count_text = _format_customer_count(count_val)
                        border_color = _customer_border_color(count_val, color)
                        if pd.isna(base_val):
                            text_value = f"{val:,.2f}"
                            marker_colors.append(color)
                        else:
                            diff_val = float(val) - float(base_val)
                            if float(base_val) == 0:
                                text_value = f"{val:,.2f} (n/a)"
                            else:
                                diff_pct = (diff_val / float(base_val)) * 100.0
                                text_value = f"{val:,.2f} ({diff_pct:+.1f}%)"
                            marker_colors.append("#ff1744" if diff_val < 0 else color)
                        marker_border_colors.append(border_color)
                        if count_text:
                            text_value = f"{text_value}<br>{count_text}"
                        label_text.append(text_value)
                else:
                    label_text = []
                    if brand == base_brand and other_max_per_date is not None:
                        aligned_other_max = other_max_per_date.reindex(series.index)
                        marker_colors = []
                        for val, other_val, count_val in zip(series.values, aligned_other_max.values, count_values):
                            if pd.isna(other_val):
                                marker_colors.append(color)
                            else:
                                marker_colors.append("#00e676" if float(val) < float(other_val) else color)
                            count_text = _format_customer_count(count_val)
                            marker_border_colors.append(_customer_border_color(count_val, color))
                            text_value = f"{val:,.2f}"
                            if count_text:
                                text_value = f"{text_value}<br>{count_text}"
                            label_text.append(text_value)
                    else:
                        marker_colors = [color] * len(series)
                        for val, count_val in zip(series.values, count_values):
                            count_text = _format_customer_count(count_val)
                            marker_border_colors.append(_customer_border_color(count_val, color))
                            text_value = f"{val:,.2f}"
                            if count_text:
                                text_value = f"{text_value}<br>{count_text}"
                            label_text.append(text_value)
                    if marker_border_colors and len(marker_border_colors) < len(label_text):
                        marker_border_colors.extend([
                            _customer_border_color(None, color)
                            for _ in range(len(label_text) - len(marker_border_colors))
                        ])
                    elif not marker_border_colors:
                        marker_border_colors = [
                            _customer_border_color(count_val, color) for count_val in count_values
                        ]
                label_positions = [
                    "top center" if (idx % 2 == 0) else "bottom center" for idx in range(len(label_text))
                ]
                price_fig.add_trace(
                    go.Scatter(
                        x=series.index,
                        y=series.values,
                        mode="lines+markers+text",
                        name=brand,
                        visible=True if brand == base_brand else "legendonly",
                        text=label_text,
                        textposition=label_positions,
                        textfont=dict(size=11, color=color),
                        line=dict(
                            color=color,
                            width=2 if brand == base_brand else 1,
                        ),
                        marker=dict(
                            size=9 if brand == base_brand else 7,
                            color=marker_colors,
                            line=dict(
                                color=marker_border_colors,
                                width=2,
                            ),
                        ),
                        hovertemplate="<b>%{fullData.name}</b><br>Date: %{x|%Y-%m-%d}<br>Price per KG: %{y:,.2f}<extra></extra>",
                    )
                )
            price_fig.update_layout(
                title="Price per KG promoted trend",
                yaxis_title="Price per KG",
                xaxis_title="Date",
                hovermode="x unified",
                font=dict(color="#111111"),
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02),
            )
            price_fig.update_yaxes(tickprefix="", separatethousands=True)
            price_fig.update_xaxes(tickformat="%Y-%m-%d")
            st.plotly_chart(price_fig, use_container_width=True)

            if base_series_per_date is not None:
                response_rows: list[dict[str, object]] = []
                for competitor in [col for col in per_kg_pivot.columns if col != base_brand]:
                    aligned = pd.concat(
                        [
                            base_series_per_date.rename("base"),
                            per_kg_pivot[competitor].rename("competitor"),
                        ],
                        axis=1,
                    ).dropna()
                    if aligned.empty:
                        continue
                    aligned = aligned.sort_index()
                    diff = aligned["base"] - aligned["competitor"]

                    last_above_date = None
                    event_above_date = None
                    event_below_date = None
                    for current_date, current_diff in diff.items():
                        if current_diff > 0:
                            last_above_date = current_date
                        elif current_diff < 0 and last_above_date is not None:
                            event_above_date = last_above_date
                            event_below_date = current_date
                            last_above_date = None

                    if event_above_date is not None and event_below_date is not None:
                        response_days = (event_below_date - event_above_date).days
                        response_rows.append(
                            {
                                "Brand": competitor,
                                "Above date": event_above_date,
                                "Below date": event_below_date,
                                "Response time (days)": int(response_days),
                            }
                        )

                if response_rows:
                    response_df = pd.DataFrame(response_rows).sort_values(
                        "Response time (days)", ascending=True
                    )
                    st.markdown("#### Response time to compete")
                    st.dataframe(
                        response_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Above date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                            "Below date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                            "Response time (days)": st.column_config.NumberColumn(format="%d"),
                        },
                    )

    relative_primary = relative_traces.get(selected_price_cols[0])
    if relative_primary is not None and not relative_primary.empty:
        avg_relative = (
            relative_primary.groupby(brand_col, as_index=False)["relative_value"].mean()
            .sort_values("relative_value", ascending=False)
        )
        brand_order = avg_relative[brand_col].tolist()
        brand_avg_per_kg: dict[str, float] = {}
        for brand, per_kg_col in brand_per_kg_columns.items():
            if per_kg_col in filtered.columns:
                per_values = filtered.loc[filtered[brand_col].astype(str) == brand, per_kg_col].dropna()
                if not per_values.empty:
                    brand_avg_per_kg[brand] = float(per_values.mean())
        avg_relative["price_per_kg_value"] = avg_relative[brand_col].map(brand_avg_per_kg.get)
        avg_relative["display_text"] = avg_relative.apply(
            lambda row: (
                _format_percent_with_per_kg(row["relative_value"], row["price_per_kg_value"])
                if show_price_labels
                else _format_percent_with_per_kg(row["relative_value"], None)
            ),
            axis=1,
        )
        st.markdown("### Average price index (Δ vs base)")
        if start_date and end_date:
            st.caption(f"Average across selected date range: {start_date:%Y-%m-%d} to {end_date:%Y-%m-%d}")
        bar_fig = px.bar(
            avg_relative,
            x=brand_col,
            y="relative_value",
            text="display_text",
            labels={brand_col: "Brand", "relative_value": "Δ vs base"},
            title=f"{selected_price_cols[0]} Δ vs {base_brand} (avg)",
            color=brand_col,
            color_discrete_map=brand_color_map,
        )
        bar_fig.update_traces(
            texttemplate="%{text}",
            textposition="outside",
            marker_line_color="#546e7a",
            marker_line_width=1,
            textfont=dict(color="#111111", size=12),
        )
        bar_fig.update_yaxes(tickformat=".0%", tickfont=dict(color="#111111"))
        for trace in bar_fig.data:
            if trace.name not in initial_visible_brands:
                trace.visible = "legendonly"
        bar_fig.update_layout(
            xaxis=dict(
                categoryorder="array",
                categoryarray=brand_order,
                tickfont=dict(size=14, color="#111111"),
            ),
            font=dict(color="#111111"),
        )
        if base_brand in avg_relative[brand_col].values:
            base_row = avg_relative.loc[avg_relative[brand_col] == base_brand]
            if not base_row.empty:
                base_value = float(base_row["relative_value"].iloc[0])
                bar_fig.add_annotation(
                    x=base_brand,
                    y=base_value,
                    text="Base brand",
                    showarrow=True,
                    arrowhead=3,
                    arrowsize=1.4,
                    arrowwidth=2,
                    arrowcolor=base_bar_color,
                    ax=0,
                    ay=-40,
                    bgcolor="rgba(253,216,83,0.25)",
                    bordercolor=base_bar_color,
                    font=dict(color="#000000", size=12, family="Segoe UI Semibold"),
                )
        st.plotly_chart(bar_fig, use_container_width=True)
    else:
        avg_dates = (
            filtered[[brand_col] + selected_price_cols]
            .groupby(brand_col, as_index=False)[selected_price_cols[0]]
            .mean()
            .sort_values(selected_price_cols[0], ascending=False)
        )
        brand_order = avg_dates[brand_col].tolist()
        brand_avg_per_kg: dict[str, float] = {}
        for brand, per_kg_col in brand_per_kg_columns.items():
            if per_kg_col in filtered.columns:
                per_values = filtered.loc[filtered[brand_col].astype(str) == brand, per_kg_col].dropna()
                if not per_values.empty:
                    brand_avg_per_kg[brand] = float(per_values.mean())
        avg_dates["price_per_kg_value"] = avg_dates[brand_col].map(brand_avg_per_kg.get)
        avg_dates["display_text"] = avg_dates.apply(
            lambda row: (
                _format_percent_with_per_kg(row[selected_price_cols[0]], row["price_per_kg_value"])
                if show_price_labels
                else _format_percent_with_per_kg(row[selected_price_cols[0]], None)
            ),
            axis=1,
        )
        st.markdown("### Average price index")
        if start_date and end_date:
            st.caption(f"Average across selected date range: {start_date:%Y-%m-%d} to {end_date:%Y-%m-%d}")
        bar_fig = px.bar(
            avg_dates,
            x=brand_col,
            y=selected_price_cols[0],
            text="display_text",
            labels={brand_col: "Brand", selected_price_cols[0]: selected_price_cols[0]},
            title=f"{selected_price_cols[0]} by brand (avg)",
            color=brand_col,
            color_discrete_map=brand_color_map,
        )
        bar_fig.update_traces(
            texttemplate="%{text}",
            textposition="outside",
            marker_line_color="#546e7a",
            marker_line_width=1,
            textfont=dict(color="#111111", size=12),
        )
        bar_fig.update_yaxes(tickformat=".0%", tickfont=dict(color="#111111"))
        for trace in bar_fig.data:
            if trace.name not in initial_visible_brands:
                trace.visible = "legendonly"
        bar_fig.update_layout(
            xaxis=dict(
                categoryorder="array",
                categoryarray=brand_order,
                tickfont=dict(size=14, color="#111111"),
            ),
            font=dict(color="#111111"),
        )
        if base_brand in avg_dates[brand_col].values:
            base_row = avg_dates.loc[avg_dates[brand_col] == base_brand]
            if not base_row.empty:
                base_value = float(base_row[selected_price_cols[0]].iloc[0])
                bar_fig.add_annotation(
                    x=base_brand,
                    y=base_value,
                    text="Base brand",
                    showarrow=True,
                    arrowhead=3,
                    arrowsize=1.4,
                    arrowwidth=2,
                    arrowcolor=base_bar_color,
                    ax=0,
                    ay=-40,
                    bgcolor="rgba(253,216,83,0.25)",
                    bordercolor=base_bar_color,
                    font=dict(color="#000000", size=12, family="Segoe UI Semibold"),
                )
        st.plotly_chart(bar_fig, use_container_width=True)

else:
    st.info("Date column not provided or contains no valid timestamps. Showing aggregated view only.")

with st.expander("Brand snapshot", expanded=False):
    if pivot_summary.empty:
        st.info("No price per KG values are available to summarize for the selected period.")
    else:
        export_table = pivot_summary.copy()
        export_reset = export_table.reset_index()
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            export_table.to_excel(writer, sheet_name="Price per KG")
            worksheet = writer.sheets["Price per KG"]
            for col_idx, column in enumerate(export_reset.columns):
                max_len = max(export_reset[column].astype(str).map(len).max(), len(str(column)))
                worksheet.set_column(col_idx, col_idx, max_len + 2)
        buffer.seek(0)
        st.download_button(
            "⬇️ Download prices (Excel)",
            data=buffer.getvalue(),
            file_name=f"brand_snapshot_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.write(
            styled_summary.to_html(classes="brand-snapshot-table", escape=False),
            unsafe_allow_html=True,
        )

with st.expander("Show filtered records"):
    st.dataframe(filtered, use_container_width=True, hide_index=True)

st.caption(
    "Tip: Install required packages with `pip install streamlit pandas numpy plotly pyxlsb` (pyxlsb is only needed for .xlsb files)."
)
