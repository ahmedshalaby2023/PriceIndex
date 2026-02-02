"""Streamlit app for analyzing item price indices by brand and date."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
from itertools import cycle
from typing import Any, Iterable, Optional
from difflib import get_close_matches

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="Price Index Explorer", layout="wide")

DEFAULT_SHEET = "Market"
SUPPORTED_EXTENSIONS = {".xlsb", ".xlsx", ".xls"}
FOCUS_BRANDS = ["Atyab", "Chikitita", "Meatland"]
BRANDS_SHEET_NAME = "Brands"
DEFAULT_ITEM_DESCRIPTION = "Meat_Chilled Luncheon_Retail_5Kg_1_Luncheon"
ITEM_SELECT_KEY = "price_index_item_select"
ITEM_INPUT_KEY = "price_index_item_input"
ITEM_LAST_SELECTED_KEY = "price_index_item_last_selected"
SHOW_PRICE_LABELS_KEY = "price_index_show_price_labels"
BRAND_SELECTION_KEY = "price_index_brand_selection"
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
    stored_brands = st.session_state.get(BRAND_SELECTION_KEY)
    if stored_brands is None:
        stored_brands = list(available_brands)
    else:
        stored_brands = [brand for brand in stored_brands if brand in available_brands]
        if not stored_brands:
            stored_brands = list(available_brands)
    if base_brand not in stored_brands:
        stored_brands = list(dict.fromkeys(list(stored_brands) + [base_brand]))
    st.session_state[BRAND_SELECTION_KEY] = stored_brands
    brand_selection = st.multiselect(
        "Brands to display",
        options=available_brands,
        default=stored_brands,
        help="All brands are selected by default. Adjust to focus analysis if needed.",
        key=BRAND_SELECTION_KEY,
    )
    brand_selection = st.session_state.get(BRAND_SELECTION_KEY, available_brands)
    if base_brand not in brand_selection:
        brand_selection = list(dict.fromkeys(list(brand_selection) + [base_brand]))
    ordered_unique: list[str] = []
    for brand in available_brands:
        if brand in brand_selection and brand not in ordered_unique:
            ordered_unique.append(brand)
    brand_selection = ordered_unique or [base_brand]
    initial_visible_brands = set(brand_selection)

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
                    value=(date_options[0], date_options[-1]),
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

price_summary_rows = []
for brand_name, brand_df in filtered.groupby(brand_col):
    summary = {"Brand": brand_name}
    for price_col in selected_price_cols:
        col_values = brand_df[price_col].dropna()
        summary[f"Avg {price_col}"] = col_values.mean() if not col_values.empty else np.nan
        summary[f"Latest {price_col}"] = col_values.iloc[-1] if not col_values.empty else np.nan
    price_summary_rows.append(summary)

summary_df = pd.DataFrame(price_summary_rows)
if base_brand and base_brand in summary_df["Brand"].values:
    ref_rows = summary_df[summary_df["Brand"] == base_brand]
    for price_col in selected_price_cols:
        ref_value = ref_rows[f"Latest {price_col}"].iloc[0]
        relative_col = f"Δ vs {base_brand} ({price_col})"
        if pd.notna(ref_value) and ref_value != 0:
            summary_df[relative_col] = summary_df[f"Latest {price_col}"] / ref_value - 1
        else:
            summary_df[relative_col] = np.nan

display_summary = summary_df.set_index("Brand")
if selected_price_cols:
    primary_metric = f"Latest {selected_price_cols[0]}"
    if primary_metric in display_summary.columns:
        display_summary = display_summary.sort_values(primary_metric, ascending=False)

def _format_percent_cell(val: float | int | None) -> str:
    if pd.isna(val):
        return "—"
    return f"{val:.0%}"


formatters = {col: _format_percent_cell for col in display_summary.columns}

def _highlight_base(row: pd.Series) -> list[str]:
    if row.name == base_brand:
        return ["background-color: #fff59d; font-weight: 600" for _ in row]
    return ["" for _ in row]

styled_summary = display_summary.style.format(formatters).apply(_highlight_base, axis=1)

highlight_color = "#0d47a1"
muted_color = "#b0bec5"
base_bar_color = "#fdd835"

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
            textposition="top center",
            textfont=dict(size=11, color="#111111"),
            mode="lines+markers+text",
        )
        fig.update_layout(hovermode="x unified", font=dict(color="#111111"))
        fig.update_xaxes(tickfont=dict(color="#111111"))
        for trace in fig.data:
            y_values = trace.y if trace.y is not None else []
            x_values = trace.x if trace.x is not None else []
            formatted_text: list[str] = []
            for idx, y in enumerate(y_values):
                label_prefix = f"{trace.name}: " if idx == 0 else ""
                if pd.isna(y):
                    formatted_text.append(f"{trace.name}" if idx == 0 else "")
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
            if formatted_text:
                trace.text = formatted_text
            else:
                x_values = trace.x if trace.x is not None else []
                trace.text = [str(trace.name)] + [""] * (len(x_values) - 1)
            color = brand_color_map.get(trace.name, muted_color)
            if trace.name == base_brand:
                trace.update(line=dict(width=4, color=color), marker=dict(size=10, color=color))
            else:
                trace.update(line=dict(width=2, color=color), marker=dict(size=6, color=color), opacity=0.75)
            if trace.name not in initial_visible_brands:
                trace.visible = "legendonly"
        st.plotly_chart(fig, use_container_width=True)

    relative_primary = relative_traces.get(selected_price_cols[0])
    if relative_primary is not None and not relative_primary.empty:
        latest_relative = (
            relative_primary.sort_values(date_col)
            .groupby(brand_col, as_index=False)
            .tail(1)
        )
        latest_relative = latest_relative.sort_values("relative_value", ascending=False)
        brand_order = latest_relative[brand_col].tolist()
        latest_relative["price_per_kg_value"] = latest_relative[brand_col].map(brand_latest_per_kg.get)
        latest_relative["display_text"] = latest_relative.apply(
            lambda row: (
                _format_percent_with_per_kg(row["relative_value"], row["price_per_kg_value"])
                if show_price_labels
                else _format_percent_with_per_kg(row["relative_value"], None)
            ),
            axis=1,
        )
        st.markdown("### Latest recorded price index (Δ vs base)")
        latest_date_value = relative_primary[date_col].max()
        if pd.notna(latest_date_value):
            st.caption(f"Latest observation date across selected brands: {latest_date_value:%Y-%m-%d}")
        bar_fig = px.bar(
            latest_relative,
            x=brand_col,
            y="relative_value",
            text="display_text",
            labels={brand_col: "Brand", "relative_value": "Δ vs base"},
            title=f"{selected_price_cols[0]} Δ vs {base_brand} (latest)",
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
        if base_brand in latest_relative[brand_col].values:
            base_row = latest_relative.loc[latest_relative[brand_col] == base_brand]
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
        latest_dates = (
            filtered[[date_col, brand_col] + selected_price_cols]
            .sort_values(date_col)
            .groupby(brand_col, as_index=False)
            .tail(1)
        )
        latest_dates = latest_dates.sort_values(selected_price_cols[0], ascending=False)
        brand_order = latest_dates[brand_col].tolist()
        latest_dates["price_per_kg_value"] = latest_dates[brand_col].map(brand_latest_per_kg.get)
        latest_dates["display_text"] = latest_dates.apply(
            lambda row: (
                _format_percent_with_per_kg(row[selected_price_cols[0]], row["price_per_kg_value"])
                if show_price_labels
                else _format_percent_with_per_kg(row[selected_price_cols[0]], None)
            ),
            axis=1,
        )
        st.markdown("### Latest recorded price index")
        latest_date_value = latest_dates[date_col].max()
        if pd.notna(latest_date_value):
            st.caption(f"Latest observation date across selected brands: {latest_date_value:%Y-%m-%d}")
        bar_fig = px.bar(
            latest_dates,
            x=brand_col,
            y=selected_price_cols[0],
            text="display_text",
            labels={brand_col: "Brand", selected_price_cols[0]: selected_price_cols[0]},
            title=f"{selected_price_cols[0]} by brand (latest)",
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
        if base_brand in latest_dates[brand_col].values:
            base_row = latest_dates.loc[latest_dates[brand_col] == base_brand]
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
    st.dataframe(styled_summary, use_container_width=True)

with st.expander("Show filtered records"):
    st.dataframe(filtered, use_container_width=True, hide_index=True)

st.caption(
    "Tip: Install required packages with `pip install streamlit pandas numpy plotly pyxlsb` (pyxlsb is only needed for .xlsb files)."
)
