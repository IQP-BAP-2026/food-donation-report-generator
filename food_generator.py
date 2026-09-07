import asyncio
import base64
import mimetypes
import re
from pathlib import Path
from typing import Any

import pandas as pd
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright


# ============================================================
# PROJECT CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "MASTER-SHEET.xlsx"
TEMPLATE_FILE = "food-traceability-template.html"
TEMPLATE_PATH = BASE_DIR / TEMPLATE_FILE

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Food-report assets are expected here.
ASSET_DIR = BASE_DIR / "assets" / "food"


# ============================================================
# GENERAL HELPERS
# ============================================================

def norm(value: Any) -> str:
    """Return a clean string for an Excel value."""
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return re.sub(r"\s+", " ", str(value)).strip()


def slug(value: Any) -> str:
    """
    Normalize column names so matching is tolerant of:
    accents, punctuation, spaces, underscores, etc.
    """
    text = norm(value).lower()

    replacements = str.maketrans(
        "áéíóúüñÁÉÍÓÚÜÑ",
        "aeiouunAEIOUUN",
    )

    text = text.translate(replacements)
    return re.sub(r"[^a-z0-9]+", "", text)


def safe_float(value: Any) -> float:
    """Convert common Excel number formats into a float."""
    if value is None:
        return 0.0

    try:
        if pd.isna(value):
            return 0.0
    except (TypeError, ValueError):
        pass

    if isinstance(value, (int, float)):
        return float(value)

    text = norm(value)

    if not text:
        return 0.0

    # Handle common formats:
    # 3,933
    # 3.933
    # 3,933.50
    # B/. 1,500
    text = text.replace("B/.", "")
    text = text.replace("B/", "")
    text = text.replace("$", "")
    text = text.strip()

    # If comma is decimal separator and there is no period,
    # preserve it as decimal. Otherwise remove grouping commas.
    if "," in text and "." not in text:
        parts = text.split(",")

        if len(parts) == 2 and len(parts[1]) <= 2:
            text = ".".join(parts)
        else:
            text = "".join(parts)
    else:
        text = text.replace(",", "")

    text = re.sub(r"[^0-9.\-]", "", text)

    try:
        return float(text)
    except ValueError:
        return 0.0


def fmt_num(value: Any, decimals: int = 0) -> str:
    """Format a number with thousands separators."""
    number = safe_float(value)

    if decimals == 0:
        return f"{number:,.0f}"

    return f"{number:,.{decimals}f}"


def get_image_base64(relative_path: str) -> str:
    """
    Optional helper for donor-specific photos.
    The template can use these fields if desired.
    """
    image_path = BASE_DIR / relative_path

    if not image_path.exists():
        return ""

    mime_type, _ = mimetypes.guess_type(str(image_path))
    mime_type = mime_type or "application/octet-stream"

    encoded = base64.b64encode(
        image_path.read_bytes()
    ).decode("ascii")

    return f"data:{mime_type};base64,{encoded}"


# ============================================================
# DATAFRAME HELPERS
# ============================================================

def dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make duplicate Excel column names unique.
    """
    df = df.copy()

    used = {}
    new_columns = []

    for column in df.columns:
        name = str(column)

        if name not in used:
            used[name] = 0
            new_columns.append(name)
        else:
            used[name] += 1
            new_columns.append(
                f"{name}_{used[name]}"
            )

    df.columns = new_columns
    return df


def find_col(
    df: pd.DataFrame,
    keywords: list[str],
    exclude: list[str] | None = None,
) -> str | None:
    """
    Find the first column whose normalized name contains
    at least one requested keyword and none of the excludes.
    """
    exclude = exclude or []

    normalized = [
        (column, slug(column))
        for column in df.columns
    ]

    for column, normalized_name in normalized:
        if not any(
            slug(keyword) in normalized_name
            for keyword in keywords
        ):
            continue

        if any(
            slug(item) in normalized_name
            for item in exclude
        ):
            continue

        return column

    return None


def matching_columns(
    df: pd.DataFrame,
    keywords: list[str],
    exclude: list[str] | None = None,
) -> list[str]:
    """Return every matching column, not just the first."""
    exclude = exclude or []

    matches = []

    for column in df.columns:
        normalized_name = slug(column)

        if not any(
            slug(keyword) in normalized_name
            for keyword in keywords
        ):
            continue

        if any(
            slug(item) in normalized_name
            for item in exclude
        ):
            continue

        matches.append(column)

    return matches


# ============================================================
# DONOR MATCHING
# ============================================================

def get_donor_column(df: pd.DataFrame) -> str | None:
    return find_col(
        df,
        [
            "DONOR",
            "DONANTE",
            "EMPRESA",
            "COMPANY",
        ],
    )


def get_donor_matches(
    df: pd.DataFrame,
    donor_query: str,
) -> pd.DataFrame:
    donor_col = get_donor_column(df)

    if donor_col is None:
        raise KeyError(
            "Could not find a donor column. "
            "Expected something like DONOR, DONANTE, EMPRESA, or COMPANY."
        )

    mask = (
        df[donor_col]
        .astype(str)
        .str.contains(
            donor_query,
            case=False,
            na=False,
            regex=False,
        )
    )

    return df[mask].copy()


# ============================================================
# MONTH DETECTION
# ============================================================

MONTHS_ES = [
    ("Ene", ["ene", "enero"]),
    ("Feb", ["feb", "febrero"]),
    ("Mar", ["mar", "marzo"]),
    ("Abr", ["abr", "abril"]),
    ("May", ["may", "mayo"]),
    ("Jun", ["jun", "junio"]),
    ("Jul", ["jul", "julio"]),
    ("Ago", ["ago", "agosto"]),
    ("Sep", ["sep", "sept", "septiembre"]),
    ("Oct", ["oct", "octubre"]),
    ("Nov", ["nov", "noviembre"]),
    ("Dic", ["dic", "diciembre"]),
]


def detect_month_in_text(text: Any) -> int | None:
    """Return month number 1-12 without matching month fragments inside words."""
    raw = norm(text).lower()

    for index, (_, aliases) in enumerate(MONTHS_ES, start=1):
        for alias in aliases:
            pattern = rf"(?<![a-záéíóúüñ]){re.escape(alias.lower())}(?![a-záéíóúüñ])"
            if re.search(pattern, raw):
                return index

    return None


def detect_year_in_text(text: Any) -> int | None:
    match = re.search(r"(20\d{2})", norm(text))

    if match:
        return int(match.group(1))

    return None


# ============================================================
# MONTHLY DATA EXTRACTION
# ============================================================

def empty_month_series() -> list[float]:
    return [0.0] * 12


def choose_month_column_columns(
    df: pd.DataFrame,
    category_keywords: list[str],
) -> dict[int, str]:
    """
    Look for columns such as:
        KG Ene
        Enero Kg
        Beneficiarios Marzo
        Platos_Abr
    and map them to month numbers.

    Returns:
        {month_number: column_name}
    """
    result = {}

    for column in df.columns:
        column_slug = slug(column)

        if not any(
            slug(keyword) in column_slug
            for keyword in category_keywords
        ):
            continue

        month = detect_month_in_text(column)

        if month is None:
            continue

        result.setdefault(month, column)

    return result


def monthly_from_wide_columns(
    donor_rows: pd.DataFrame,
    category_keywords: list[str],
) -> list[float]:
    """
    Extract a monthly series from a spreadsheet arranged horizontally.
    """
    month_columns = choose_month_column_columns(
        donor_rows,
        category_keywords,
    )

    values = empty_month_series()

    for month, column in month_columns.items():
        values[month - 1] = donor_rows[column].apply(
            safe_float
        ).sum()

    return values


def monthly_from_date_column(
    donor_rows: pd.DataFrame,
    date_col: str,
    value_col: str,
) -> list[float]:
    """
    Extract monthly values when the workbook has normalized rows:
        DATE | KG
    """
    values = empty_month_series()

    dates = pd.to_datetime(
        donor_rows[date_col],
        errors="coerce",
        dayfirst=True,
    )

    for index, date_value in dates.items():
        if pd.isna(date_value):
            continue

        month = int(date_value.month)

        values[month - 1] += safe_float(
            donor_rows.loc[index, value_col]
        )

    return values


def get_monthly_series(
    donor_rows: pd.DataFrame,
    category_keywords: list[str],
) -> list[float]:
    """
    Try both common spreadsheet structures:
      1. one row per donor with monthly columns
      2. many rows with a date/month column
    """

    # --------------------------------------------------------
    # Wide format: columns named with months
    # --------------------------------------------------------

    wide = monthly_from_wide_columns(
        donor_rows,
        category_keywords,
    )

    if any(value > 0 for value in wide):
        return wide

    # --------------------------------------------------------
    # Normalized format: DATE/MONTH + VALUE
    # --------------------------------------------------------

    date_col = find_col(
        donor_rows,
        [
            "fecha",
            "date",
            "mes",
            "month",
            "periodo",
            "period",
        ],
    )

    value_col = find_col(
        donor_rows,
        category_keywords,
    )

    if date_col and value_col:
        normalized = monthly_from_date_column(
            donor_rows,
            date_col,
            value_col,
        )

        if any(value > 0 for value in normalized):
            return normalized

    return []


def format_month_series(
    values: list[float],
) -> list[str]:
    return [
        fmt_num(value) if value > 0 else "0"
        for value in values
    ]


def total_from_series(values: list[float]) -> str:
    total = sum(values)

    return fmt_num(total) if total > 0 else ""


def average_from_series(values: list[float]) -> str:
    """Average only over months that contain a positive reported value."""
    valid = [value for value in values if value > 0]

    if not valid:
        return ""

    average = sum(valid) / len(valid)
    return fmt_num(average, 1)


def infer_month_from_monthly_value(
    target_value: float,
    monthly_values: list[float],
) -> int | None:
    """Infer the reporting month when the sheet has no reliable month field.

    This is useful for workbooks where a row contains both a monthly KG field
    and separate Ene/Feb/Mar/... columns but no explicit report-month cell.
    """
    if target_value <= 0 or not monthly_values:
        return None

    exact_matches = [
        index + 1
        for index, value in enumerate(monthly_values)
        if value > 0 and abs(value - target_value) < 0.01
    ]

    if exact_matches:
        return exact_matches[-1]

    return None


def find_reliable_reporting_month(
    df: pd.DataFrame,
    row: pd.Series,
    rescued_value: float,
    monthly_kilos: list[float],
) -> int | None:
    """Find an explicit report month first, then infer it from monthly KG data."""
    # Look only at columns that are genuinely month/date fields.
    month_candidates = []
    for column in df.columns:
        name = slug(column)
        if any(token in name for token in [
            "mes", "fecha", "date", "periodo", "period", "reportmonth", "month",
        ]) and not any(token in name for token in [
            "donated", "donacion", "donaciones", "kilos", "kilo", "kg",
        ]):
            month_candidates.append(column)

    for column in month_candidates:
        raw = norm(row.get(column))
        if not raw:
            continue

        month_number = detect_month_in_text(raw)
        if month_number:
            return month_number

        try:
            date_value = pd.to_datetime(raw, errors="coerce", dayfirst=True)
            if not pd.isna(date_value):
                return int(date_value.month)
        except Exception:
            pass

    return infer_month_from_monthly_value(
        rescued_value,
        monthly_kilos,
    )


# ============================================================
# FOOD METRIC EXTRACTION
# ============================================================

def get_first_positive_value(
    row: pd.Series,
    df: pd.DataFrame,
    keywords: list[str],
    exclude: list[str] | None = None,
) -> tuple[str | None, float]:
    """
    Find a likely metric column and return its positive value.
    """
    column = find_col(
        df,
        keywords,
        exclude=exclude,
    )

    if column is None:
        return None, 0.0

    return column, safe_float(row.get(column))


def extract_food_metrics(
    donor_rows: pd.DataFrame,
    first_row: pd.Series,
) -> dict:
    """Extract the five required top metrics from the master sheet.

    The master workbook has explicit fields for the monthly donation, food
    formula percentages, plate conversion, and beneficiary organizations.
    Those are preferred over broad keyword matching so summary columns such as
    Kilos_Total cannot accidentally be used for the monthly hero figure.
    """

    # 1) KG rescatados en el mes
    rescued_col = None
    rescued_value = 0.0
    exact = ["{KG DONADOS EN EL MES}", "{KG DONATED IN THE MONTH}", "KG DONADOS EN EL MES", "KG DONATED IN THE MONTH"]
    for column in donor_rows.columns:
        if slug(column) in {slug(x) for x in exact}:
            rescued_col = column
            rescued_value = safe_float(first_row.get(column))
            break
    if rescued_col is None:
        rescued_col, rescued_value = get_first_positive_value(
            first_row, donor_rows,
            ["kg donated in the month", "kg donados en el mes", "kg rescatados en el mes"],
        )

    # 2) Quality percentages are authoritative in the sheet.
    used_pct_col = None
    used_pct = 0.0
    merma_pct_col = None
    merma_pct = 0.0
    for column in donor_rows.columns:
        n = slug(column)
        if n in {
            slug("{FÓRMULA DE ALIMENTO % UTILIZADO}"),
            slug("FÓRMULA DE ALIMENTO % UTILIZADO"),
            slug("{FEED FORMULA % USED}"),
            slug("FEED FORMULA % USED"),
        }:
            used_pct_col = column
            used_pct = safe_float(first_row.get(column))
        if n in {
            slug("{FÓRMULA DE ALIMENTO % MERMA}"),
            slug("FÓRMULA DE ALIMENTO % MERMA"),
            slug("{FOOD FORMULA % MERMA}"),
            slug("FOOD FORMULA % MERMA"),
        }:
            merma_pct_col = column
            merma_pct = safe_float(first_row.get(column))

    # Explicit usable/waste kg columns, when present, still take precedence.
    usable_col, usable_value = get_first_positive_value(
        first_row, donor_rows,
        ["{KG APROVECHABES EN EL MES}", "KG APROVECHABES EN EL MES", "kilos aprovechables", "kg aprovechables", "kilos aptos", "kg aptos", "usable kg"],
    )
    waste_col, waste_value = get_first_positive_value(
        first_row, donor_rows,
        ["{KG MERMA EN EL MES}", "KG MERMA EN EL MES", "kilos merma", "kg merma", "merma", "waste kg", "kg waste", "no aptos"],
    )

    if rescued_value > 0 and usable_value <= 0:
        usable_value = rescued_value * used_pct / 100.0
    if rescued_value > 0 and waste_value <= 0:
        waste_value = rescued_value * merma_pct / 100.0

    # 3) Food plates
    plates_col, plates_value = get_first_positive_value(
        first_row, donor_rows,
        ["{FÓRMULA CONVERSIÓN A PLATOS DE COMIDA}", "FÓRMULA CONVERSIÓN A PLATOS DE COMIDA", "formula conversion to food plates", "platos de comida", "platos comida", "food plates"],
    )

    # 4) Organizations - prefer the workbook's Total Obs/PRO field.
    orgs_col = None
    orgs_value = 0.0
    for column in donor_rows.columns:
        if slug(column) in {slug("Total Obs/PRO"), slug("Total Obs / PRO")}:
            orgs_col = column
            orgs_value = safe_float(first_row.get(column))
            break
    if orgs_col is None:
        orgs_col, orgs_value = get_first_positive_value(
            first_row, donor_rows,
            ["organizaciones beneficiarias", "organizaciones", "organization", "org"],
            exclude=["beneficiarios", "beneficiarias"],
        )

    population_value = 0.0
    population_col = None
    for column in donor_rows.columns:
        if slug(column) in {
            slug("Total BENEFICIARIOS"),
            slug("Benefeciarios en Total"),
        }:
            population_col = column
            population_value = safe_float(first_row.get(column))
            break

    return {
        "rescued_value": rescued_value,
        "rescued_col": rescued_col,
        "usable_value": usable_value,
        "usable_col": usable_col,
        "waste_value": waste_value,
        "waste_col": waste_col,
        "plates_value": plates_value,
        "plates_col": plates_col,
        "orgs_value": orgs_value,
        "orgs_col": orgs_col,
        "population_value": population_value,
        "population_col": population_col,
        "used_pct": round(used_pct, 1),
        "merma_pct": round(merma_pct, 1),
        "usable_pct": round((usable_value / rescued_value * 100.0) if rescued_value > 0 else 0.0, 1),
        "waste_pct": round((waste_value / rescued_value * 100.0) if rescued_value > 0 else 0.0, 1),
        "rescued_pct": round(100.0 if rescued_value > 0 else 0.0, 1),
    }


# ============================================================
# REGIONAL EXTRACTION
# ============================================================

REGIONS = [
    ("Panamá Centro", ["P CENTRO", "PANAMA CENTRO"]),
    ("Panamá Este", ["ESTE", "PANAMA ESTE"]),
    ("Panamá Norte", ["NORTE", "PANAMA NORTE"]),
    ("San Miguelito", ["SAN MIGUELITO"]),
    ("Panamá Oeste", ["OESTE", "PANAMA OESTE"]),
    ("Coclé", ["COCLE"]),
    ("Colón", ["COLON"]),
    ("Darién", ["DARIEN"]),
    ("Herrera", ["HERRERA"]),
    ("Los Santos", ["LOS SANTOS"]),
    ("Veraguas", ["VERAGUAS"]),
    ("Chiriquí", ["CHIQUIRI"]),
    ("Bocas del Toro", ["BOCAS"]),
    ("Comarca Ngäbe Buglé", ["COMARCA NGABE BUGLE"]),
]


def extract_regional_data(
    row: pd.Series,
) -> list[dict]:
    """
    Find regional kg columns and beneficiary columns in the donor row.
    """

    results = []

    for display_name, aliases in REGIONS:

        kg = 0.0
        beneficiaries = 0.0

        for column in row.index:

            column_slug = slug(column)

            if not any(
                slug(alias) in column_slug
                for alias in aliases
            ):
                continue

            if "kg" in column_slug and (
                "asign" in column_slug
                or "donat" in column_slug
                or "kilo" in column_slug
            ):
                kg = max(
                    kg,
                    safe_float(row.get(column)),
                )

            if "benef" in column_slug:
                beneficiaries = max(
                    beneficiaries,
                    safe_float(row.get(column)),
                )

        if kg > 0 or beneficiaries > 0:

            results.append(
                {
                    "region": display_name,
                    "kg_raw": kg,
                    "ben_raw": beneficiaries,
                    "kg": fmt_num(kg),
                    "ben": fmt_num(beneficiaries),
                }
            )

    total_kg = sum(
        item["kg_raw"]
        for item in results
    )

    for item in results:
        item["pct"] = (
            round(
                item["kg_raw"] / total_kg * 100,
                1,
            )
            if total_kg > 0
            else 0
        )

    return results


# ============================================================
# MONTHLY REPORT DATA
# ============================================================

def extract_monthly_data(
    donor_rows: pd.DataFrame,
) -> dict:
    """
    Build monthly series for:
      - kilos
      - organizations
      - population

    The raw numeric series are retained for reliable month inference and
    averages; underscore-prefixed keys are not displayed directly in the HTML.
    """
    kilos_raw = []
    kilo_columns = {}
    for column in donor_rows.columns:
        name = slug(column)
        if name.startswith("kilos"):
            month = detect_month_in_text(column)
            if month:
                kilo_columns[month] = column
    if kilo_columns:
        kilos_raw = empty_month_series()
        for month, column in kilo_columns.items():
            kilos_raw[month - 1] = donor_rows[column].apply(safe_float).sum()
    else:
        kilos_raw = get_monthly_series(
            donor_rows,
            ["kg", "kilo", "kilos", "kilogram"],
        ) or []

    organizations_raw = get_monthly_series(
        donor_rows,
        [
            "organizacion",
            "organization",
            "org",
        ],
    ) or []

    # The master workbook uses OB_Ene ... OB_Dic for beneficiary organizations.
    # These exact fields are preferred over generic matching.
    ob_columns = {}
    for column in donor_rows.columns:
        name = slug(column)
        if name.startswith("ob"):
            month = detect_month_in_text(column)
            if month:
                ob_columns[month] = column
    if ob_columns:
        organizations_raw = empty_month_series()
        for month, column in ob_columns.items():
            organizations_raw[month - 1] = donor_rows[column].apply(safe_float).sum()

    population_raw = []
    population_columns = {}
    for column in donor_rows.columns:
        name = slug(column)
        if name.startswith("poblacion"):
            month = detect_month_in_text(column)
            if month:
                population_columns[month] = column
    if population_columns:
        population_raw = empty_month_series()
        for month, column in population_columns.items():
            population_raw[month - 1] = donor_rows[column].apply(safe_float).sum()
    else:
        population_raw = get_monthly_series(
            donor_rows,
            ["beneficiario", "beneficiarios", "poblacion", "población", "personas"],
        ) or []

    result = {
        "_kilos_raw": kilos_raw,
        "_organizations_raw": organizations_raw,
        "_population_raw": population_raw,
    }

    if kilos_raw:
        result["kilos"] = format_month_series(kilos_raw)
        result["kilos_total"] = total_from_series(kilos_raw)

    if organizations_raw:
        result["organizations"] = format_month_series(organizations_raw)
        result["organizations_total"] = total_from_series(organizations_raw)
        result["organizations_average"] = average_from_series(organizations_raw)

    if population_raw:
        result["population"] = format_month_series(population_raw)
        result["population_total"] = total_from_series(population_raw)
        result["population_average"] = average_from_series(population_raw)

    return result


# ============================================================
# MAP POINTS
# ============================================================

# Approximate positions on the Panama map image.
# These positions are visual anchors only; the actual values
# still come from the spreadsheet.
MAP_POSITIONS = {
    # Positions are percentages relative to the transparent Panama map asset.
    # Labels are intentionally spread inside their province while remaining on-map.
    "Bocas del Toro": (8, 28),
    "Chiriquí": (24, 56),
    "Veraguas": (36, 66),
    "Coclé": (49, 56),
    "Herrera": (54, 68),
    "Los Santos": (62, 74),
    "Panamá Oeste": (61, 45),
    "Panamá Centro": (70, 47),
    "Panamá Este": (84, 51),
    "Darién": (94, 66),
    "Colón": (63, 27),
    "Panamá Norte": (72, 37),
    "San Miguelito": (73, 43),
    "Comarca Ngäbe Buglé": (18, 44),
}


def build_map_points(
    regional_table: list[dict],
) -> list[dict]:
    points = []

    for item in regional_table:

        position = MAP_POSITIONS.get(
            item["region"]
        )

        if not position:
            continue

        x, y = position

        points.append(
            {
                "x": x,
                "y": y,
                "value": item["kg"],
                "region": item["region"],
            }
        )

    return points


# ============================================================
# BUILD JINJA DATA
# ============================================================

def build_report_data(
    df: pd.DataFrame,
    donor_query: str,
) -> dict | None:

    donor_col = get_donor_column(df)

    if donor_col is None:
        print(
            "\n❌ Could not find the donor column."
        )
        return None

    matches = get_donor_matches(
        df,
        donor_query,
    )

    if matches.empty:
        print(
            f"\n❌ No donor matching "
            f"'{donor_query}' was found."
        )
        return None

    row = matches.iloc[0]

    company_name = norm(
        row.get(donor_col)
    )

    print(
        f"\n✅ Processing Food Traceability: "
        f"{company_name}"
    )

    # --------------------------------------------------------
    # Core metrics
    # --------------------------------------------------------

    food = extract_food_metrics(
        matches,
        row,
    )

    rescued = food["rescued_value"]
    usable = food["usable_value"]
    waste = food["waste_value"]
    plates = food["plates_value"]
    orgs = food["orgs_value"]
    population = food["population_value"]

    # Some versions of the master sheet store food quality as percentages
    # instead of explicit usable/merma kilogram columns. Derive the kg values
    # from the rescued amount without inventing any unsupported numbers.
    if rescued > 0 and usable <= 0:
        used_pct_col = find_col(
            df,
            [
                "feed formula % used",
                "food formula % used",
                "formula % used",
                "% used",
            ],
        )
        if used_pct_col:
            used_pct = safe_float(row.get(used_pct_col))
            if used_pct > 0:
                usable = rescued * used_pct / 100.0

    if rescued > 0 and waste <= 0:
        waste_pct_col = find_col(
            df,
            [
                "food formula % merma",
                "formula % merma",
                "% merma",
                "merma %",
            ],
        )
        if waste_pct_col:
            merma_pct = safe_float(row.get(waste_pct_col))
            if merma_pct > 0:
                waste = rescued * merma_pct / 100.0

    # Recalculate displayed percentages after any derived kg values are added.
    usable_pct = (usable / rescued * 100.0) if rescued > 0 and usable > 0 else 0.0
    waste_pct = (waste / rescued * 100.0) if rescued > 0 and waste > 0 else 0.0
    food["usable_pct"] = round(usable_pct, 1)
    food["waste_pct"] = round(waste_pct, 1)

    # If organization count is not directly in the donor row,
    # try counting actual organization records.
    org_count_fallback = find_col(
        matches,
        [
            "organizacion",
            "organization",
            "aliada",
            "beneficiaria",
        ],
        exclude=[
            "beneficiarios",
            "beneficiarias",
        ],
    )

    if orgs <= 0 and org_count_fallback:
        unique_orgs = (
            matches[org_count_fallback]
            .dropna()
            .astype(str)
            .str.strip()
        )

        unique_orgs = unique_orgs[
            unique_orgs != ""
        ]

        unique_org_count = (
            unique_orgs.nunique()
        )

        if unique_org_count > 0:
            orgs = float(
                unique_org_count
            )

    # --------------------------------------------------------
    # Regional allocation
    # --------------------------------------------------------

    regional_table = extract_regional_data(
        row
    )

    # --------------------------------------------------------
    # Monthly
    # --------------------------------------------------------

    monthly_data = extract_monthly_data(
        matches
    )

    # --------------------------------------------------------
    # Reporting month
    # --------------------------------------------------------

    month_number = find_reliable_reporting_month(
        df,
        row,
        rescued,
        monthly_data.get("_kilos_raw", []),
    )

    # If there is no explicit month/date cell, the monthly KG series is the
    # source of truth. Pick the latest month with reported KG, which matches
    # the current monthly donation represented by {KG DONATED IN THE MONTH}.
    if not month_number:
        kilos_series = monthly_data.get("_kilos_raw", [])
        nonzero = [i + 1 for i, value in enumerate(kilos_series) if value > 0]
        month_number = nonzero[-1] if nonzero else None

    month_name = (
        MONTHS_ES[month_number - 1][0]
        if month_number
        else ""
    )

    # The top organization and population figures are for the same report
    # month, not annual totals.
    if month_number:
        idx = month_number - 1
        org_series = monthly_data.get("_organizations_raw", [])
        population_series = monthly_data.get("_population_raw", [])
        if idx < len(org_series):
            orgs = org_series[idx]
        if idx < len(population_series):
            population = population_series[idx]
    # --------------------------------------------------------
    # Year
    # --------------------------------------------------------

    year = "2026"

    # Search a year field first.
    year_col = find_col(
        df,
        [
            "año",
            "ano",
            "year",
        ],
    )

    if year_col:
        year_number = detect_year_in_text(
            row.get(year_col)
        )

        if year_number:
            year = str(year_number)

    # Search column names for a year if needed.
    if year == "2026":

        for column in df.columns:

            detected_year = detect_year_in_text(
                column
            )

            if detected_year:
                year = str(
                    detected_year
                )
                break

    # --------------------------------------------------------
    # National delivered kg
    # --------------------------------------------------------

    delivered_col = find_col(
        df,
        [
            "kilos totales entregados",
            "kilos entregados",
            "kg entregados",
            "national delivered",
            "delivered kg",
        ],
    )

    national_delivered = (
        safe_float(row.get(delivered_col))
        if delivered_col
        else 0.0
    )

    # --------------------------------------------------------
    # Regional narrative
    # --------------------------------------------------------

    top_region_summary = ""

    if regional_table:

        top_region = max(
            regional_table,
            key=lambda item: item["kg_raw"],
        )

        if top_region["kg_raw"] > 0:

            top_region_summary = (
                f"La mayor asignación regional "
                f"se concentra en "
                f"{top_region['region']}, "
                f"con {top_region['kg']} kg."
            )

    # --------------------------------------------------------
    # Optional donor photo
    # --------------------------------------------------------

    distribution_photo = ""

    candidate_photos = [
        "photos/community_photo1.jpg",
        "photos/community_photo2.jpg",
    ]

    for candidate in candidate_photos:
        encoded = get_image_base64(
            candidate
        )

        if encoded:
            distribution_photo = encoded
            break

    # --------------------------------------------------------
    # Return template data
    # --------------------------------------------------------

    return {

        "company_name": company_name,

        "year": year,

        "month_name": month_name,

        # Core food figures
        "kilos_donated": fmt_num(rescued),

        "kilos_usable": fmt_num(usable),

        "kilos_waste": fmt_num(waste),

        "meals_served": fmt_num(plates),

        "orgs_helped": fmt_num(orgs),

        "beneficiaries": fmt_num(population),

        # Percentages
        "kilos_share": food["rescued_pct"],

        "usable_pct": food["usable_pct"],

        "waste_pct": food["waste_pct"],

        # Monthly
        "monthly_data": monthly_data,

        "month_labels": [
            month[0]
            for month in MONTHS_ES
        ],

        "monthly_total_label": True,

        # Regional
        "regional_table": regional_table,

        # Stable left/right callout columns keep every regional figure
        # separated from the map and from one another.
        "regional_left": regional_table[0::2],
        "regional_right": regional_table[1::2],

        "map_points": build_map_points(
            regional_table
        ),

        "national_delivered_kg": (
            fmt_num(national_delivered)
            if national_delivered > 0
            else fmt_num(rescued)
        ),

        "top_region_summary": (
            top_region_summary
        ),

        # Optional image
        "distribution_photo": (
            distribution_photo
        ),
    }


# ============================================================
# PDF RENDERING
# ============================================================

async def generate_pdf(
    data: dict,
) -> None:

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Missing template:\n{TEMPLATE_PATH}"
        )

    env = Environment(
        loader=FileSystemLoader(
            str(BASE_DIR)
        ),
        autoescape=True,
    )

    template = env.get_template(
        TEMPLATE_FILE
    )

    rendered_html = template.render(
        **data
    )

    safe_name = re.sub(
        r"[^A-Za-z0-9 _-]+",
        "",
        data["company_name"],
    ).strip().replace(
        " ",
        "_",
    )

    month_part = (
        data["month_name"]
        if data["month_name"]
        else "Annual"
    )

    rendered_path = (
        BASE_DIR
        / f"._rendered_food_{safe_name}.html"
    )

    output_path = (
        OUTPUT_DIR
        / (
            f"Food_Traceability_"
            f"{safe_name}_"
            f"{month_part}_"
            f"{data['year']}.pdf"
        )
    )

    rendered_path.write_text(
        rendered_html,
        encoding="utf-8",
    )

    async with async_playwright() as p:

        browser = await p.chromium.launch()

        page = await browser.new_page(
            viewport={
                "width": 900,
                "height": 1800,
            },
            device_scale_factor=1,
        )

        # Opening the local HTML file makes relative
        # CSS and asset references resolve correctly.
        await page.goto(
            rendered_path.as_uri(),
            wait_until="networkidle",
        )

        await page.evaluate(
            """
            document.fonts
              ? document.fonts.ready
              : Promise.resolve()
            """
        )

        await page.wait_for_function(
            """
            Array.from(document.images)
              .every(img => img.complete)
            """
        )

        await page.pdf(
            path=str(output_path),
            print_background=True,
            prefer_css_page_size=True,
            margin={
                "top": "0mm",
                "right": "0mm",
                "bottom": "0mm",
                "left": "0mm",
            },
        )

        await browser.close()

    try:
        rendered_path.unlink()
    except OSError:
        pass

    print(
        "\n🎉 Food Traceability PDF generated:"
    )
    print(
        f"   {output_path}"
    )


# ============================================================
# MAIN
# ============================================================

def main() -> None:

    if not EXCEL_FILE.exists():

        print(
            f"\n❌ Excel file not found:\n"
            f"   {EXCEL_FILE}"
        )

        print(
            "\nPlace MASTER-SHEET.xlsx next to "
            "food_generator.py."
        )

        return

    if not TEMPLATE_PATH.exists():

        print(
            f"\n❌ Food template not found:\n"
            f"   {TEMPLATE_PATH}"
        )

        return

    print(
        "\nLoading Excel database..."
    )

    df = pd.read_excel(
        EXCEL_FILE
    )

    df = dedupe_columns(
        df
    )

    donor_col = get_donor_column(
        df
    )

    if donor_col is None:

        print(
            "\n❌ Could not identify a "
            "DONOR column."
        )

        print(
            "Expected something like:"
            "\n  DONOR"
            "\n  DONANTE"
            "\n  EMPRESA"
            "\n  COMPANY"
        )

        return

    print(
        f"\nDonor column detected: "
        f"{donor_col}"
    )

    user_input = input(
        "\nEnter company/donor name "
        "(or type 'ALL'): "
    ).strip()

    if not user_input:
        print(
            "\n❌ Please enter a donor."
        )
        return

    if user_input.upper() == "ALL":

        donors = (
            df[donor_col]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        print(
            f"\n📄 Generating food reports "
            f"for {len(donors)} donors..."
        )

        async def generate_all():

            for donor in donors:

                data = build_report_data(
                    df,
                    donor,
                )

                if data:
                    await generate_pdf(
                        data
                    )

        asyncio.run(
            generate_all()
        )

        return

    data = build_report_data(
        df,
        user_input,
    )

    if data is None:
        return

    asyncio.run(
        generate_pdf(
            data
        )
    )


if __name__ == "__main__":
    main()