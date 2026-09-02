import asyncio
import re
import base64
import mimetypes
from pathlib import Path
from typing import Any

import pandas as pd
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "master-excel.xlsx"
TEMPLATE_FILE = "report-template.html"
TEMPLATE_PATH = BASE_DIR / TEMPLATE_FILE
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def norm(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def slug(v: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", norm(v).lower())


def safe_float(v: Any) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)) or norm(v) == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = norm(v).replace(",", "")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return 0.0


def fmt_num(v: Any) -> str:
    return f"{safe_float(v):,.0f}"


def fmt_money(v: Any, currency="B/.") -> str:
    return f"{currency} {safe_float(v):,.2f}"


def get_image_base64(relative_path: str) -> str:
    p = BASE_DIR / relative_path
    if not p.exists():
        return ""
    mime, _ = mimetypes.guess_type(str(p))
    mime = mime or "application/octet-stream"
    return f"data:{mime};base64,{base64.b64encode(p.read_bytes()).decode('ascii')}"


def find_col(df: pd.DataFrame, keywords: list[str], *, exclude: list[str] | None = None) -> str | None:
    exclude = exclude or []
    candidates = []
    for c in df.columns:
        sc = slug(c)
        if any(slug(k) in sc for k in keywords) and not any(slug(x) in sc for x in exclude):
            candidates.append(c)
    return candidates[0] if candidates else None


def all_matching_cols(df: pd.DataFrame, keywords: list[str], exclude: list[str] | None = None) -> list[str]:
    exclude = exclude or []
    return [c for c in df.columns if any(slug(k) in slug(c) for k in keywords) and not any(slug(x) in slug(c) for x in exclude)]


def dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        idxs = cols[cols == dup].index.tolist()
        cols.loc[idxs] = [dup if i == 0 else f"{dup}_{i}" for i in range(len(idxs))]
    df = df.copy()
    df.columns = cols
    return df


def extract_regional_data(row: pd.Series) -> tuple[list[dict], list[str], list[float]]:
    region_mapping = [
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
    rows = []
    for display, prefixes in region_mapping:
        kg = 0.0
        ben = 0.0
        for c in row.index:
            s = slug(c)
            if any(slug(p) in s for p in prefixes) and "kg" in s and ("asign" in s or "donat" in s or "kilo" in s):
                kg = max(kg, safe_float(row.get(c)))
            if any(slug(p) in s for p in prefixes) and "benef" in s:
                ben = max(ben, safe_float(row.get(c)))
        if kg > 0 or ben > 0:
            rows.append({"region": display, "kg_raw": kg, "ben_raw": ben})
    total_kg = sum(x["kg_raw"] for x in rows)
    for x in rows:
        x["pct"] = round((x["kg_raw"] / total_kg * 100) if total_kg else 0, 1)
        x["kg"] = fmt_num(x["kg_raw"])
        x["ben"] = fmt_num(x["ben_raw"])
    return rows, [x["region"] for x in rows], [x["kg_raw"] for x in rows]


def extract_org_rows(matches: pd.DataFrame) -> list[dict]:
    """Best-effort individual organization extraction for normalized or semi-normalized sheets."""
    org_col = find_col(matches, ["organizacion", "organization", "aliada", "beneficiaria", "beneficiario"])
    if not org_col:
        return []
    region_col = find_col(matches, ["region", "provincia", "area", "zona"])
    program_col = find_col(matches, ["programa", "program"])
    kg_col = find_col(matches, ["kg", "kilos", "kilogram"], exclude=["total", "mes"])
    ben_col = find_col(matches, ["beneficiario", "beneficiarios", "personas", "beneficiadas"])
    plates_col = find_col(matches, ["plato", "plates", "food plates", "equivalencia"])
    money_col = find_col(matches, ["monto", "aporte", "valor", "donacion", "donación", "amount", "spent", "gasto", "inversion"])

    data = []
    for _, r in matches.iterrows():
        name = norm(r.get(org_col))
        if not name or name.lower() in {"nan", "none", "0"}:
            continue
        item = {
            "name": name,
            "region": norm(r.get(region_col)) if region_col else "",
            "program": norm(r.get(program_col)) if program_col else "",
            "kg": fmt_num(r.get(kg_col)) if kg_col and safe_float(r.get(kg_col)) else None,
            "ben": fmt_num(r.get(ben_col)) if ben_col and safe_float(r.get(ben_col)) else None,
            "plates": fmt_num(r.get(plates_col)) if plates_col and safe_float(r.get(plates_col)) else None,
            "amount": fmt_money(r.get(money_col)) if money_col and safe_float(r.get(money_col)) else None,
        }
        data.append(item)
    # Aggregate duplicates by organization, retaining max/total sensible measures.
    merged: dict[str, dict] = {}
    for x in data:
        key = x["name"].casefold() 
        if key not in merged:
            merged[key] = x.copy()
        else:
            for fld in ("kg", "ben", "plates"):
                if x[fld] is not None:
                    merged[key][fld] = fmt_num(safe_float(merged[key][fld]) + safe_float(x[fld]))
            if not merged[key].get("region") and x.get("region"):
                merged[key]["region"] = x["region"]
            if not merged[key].get("program") and x.get("program"):
                merged[key]["program"] = x["program"]
    return list(merged.values())[:12]


def build_data(df: pd.DataFrame, donor_query: str) -> dict | None:
    donor_col = find_col(df, ["DONOR", "DONANTE", "EMPRESA", "COMPANY"])
    if not donor_col:
        raise KeyError("No pude encontrar una columna DONOR/DONANTE/EMPRESA/COMPANY en el Excel.")

    matches = df[df[donor_col].astype(str).str.contains(donor_query, case=False, na=False, regex=False)].copy()
    if matches.empty:
        print(f"\n❌ No se encontró un donante que coincida con '{donor_query}'.")
        return None

    row = matches.iloc[0]
    company_name = norm(row.get(donor_col))
    print(f"\n✅ Processing Donor: {company_name}")

    kilos_col = find_col(df, ["Kilos_Total", "kg donated in the month", "kilos", "kilogram"], exclude=["asignados", "beneficiarios"])
    meals_col = find_col(df, ["formula conversion to food plates", "food plates", "platos", "plates"])
    beneficiaries_col = find_col(df, ["Benefeciarios en Total", "Total BENEFICIARIOS", "beneficiarios en total", "beneficiarios"])
    orgs_col = find_col(df, ["Total Obs/PRO", "organizaciones", "organization", "org"])
    money_col = find_col(df, ["monto", "aporte", "valor donacion", "valor donación", "amount", "spent", "gasto", "inversion", "inversión"], exclude=["total beneficiarios"])
    contact_col = find_col(df, ["contacto principal", "contact"])
    email_col = find_col(df, ["correo electronico", "correo", "email"])
    phone_col = find_col(df, ["telefono", "teléfono", "phone"])

    regional_table, _, _ = extract_regional_data(row)
    if not regional_table:
        regional_table = [{"region": "Sin área asignada", "kg_raw": 0, "ben_raw": 0, "kg": "0", "ben": "0", "pct": 0}]

    top = max(regional_table, key=lambda x: x["kg_raw"])
    if top["kg_raw"] > 0:
        top_region_summary = f"La mayor concentración del aporte aparece en {top['region']}, con {top['kg']} kg ({top['pct']}% del volumen regional identificado)."
    else:
        top_region_summary = "No se encontró un desglose regional positivo en esta fila del maestro."

    age_specs = [
        ("Niños 0–5 años", find_col(df, ["Niños 0-5", "ninos 0-5"])) ,
        ("Niños 6–10 años", find_col(df, ["Niños 6-10", "ninos 6-10"])),
        ("Jóvenes 11–18 años", find_col(df, ["Niños 11-18", "ninos 11-18", "jovenes 11-18"])),
        ("Adultos", find_col(df, ["Total ADULTOS", "adultos"])),
        ("Sin Edad Especificada", find_col(df, ["Sin Edad", "sin edad"])) ,
    ]
    age_table = []
    for label, col in age_specs:
        if col:
            n = safe_float(row.get(col))
            if n > 0:
                age_table.append({"group": label, "count": fmt_num(n), "raw": n})
    age_total = sum(x["raw"] for x in age_table)
    for x in age_table:
        x["share"] = round(x["raw"] / age_total * 100, 1) if age_total else 0
    children = sum(x["raw"] for x in age_table if "Niños" in x["group"] or "Jóvenes" in x["group"])
    adults = next((x["raw"] for x in age_table if x["group"] == "Adultos"), 0)
    others = max(age_total - children - adults, 0)
    denom = max(children + adults + others, 1)

    monetary_total = safe_float(row.get(money_col)) if money_col else 0
    data = {
        "company_name": company_name,
        "year": str(2026),
        "contact_person": norm(row.get(contact_col)) if contact_col else "No reportado",
        "contact_email": norm(row.get(email_col)) if email_col else "No reportado",
        "contact_phone": norm(row.get(phone_col)) if phone_col else "No reportado",
        "kilos_donated": fmt_num(row.get(kilos_col)) if kilos_col else "0",
        "meals_served": fmt_num(row.get(meals_col)) if meals_col else "0",
        "beneficiaries": fmt_num(row.get(beneficiaries_col)) if beneficiaries_col else "0",
        "orgs_helped": fmt_num(row.get(orgs_col)) if orgs_col else fmt_num(len(extract_org_rows(matches))),
        "children_count": fmt_num(children),
        "adults_count": fmt_num(adults),
        "other_age_count": fmt_num(others),
        "children_pct": round(children / denom * 100, 1),
        "adults_pct": round(adults / denom * 100, 1),
        "other_age_pct": round(others / denom * 100, 1),
        "demographics_total": fmt_num(age_total),
        "regional_table": regional_table,
        "region_count": len([x for x in regional_table if x["kg_raw"] > 0 or x["ben_raw"] > 0]),
        "age_table": age_table or [{"group": "Sin datos demográficos", "count": "0", "share": 0}],
        "organizations": extract_org_rows(matches),
        "monetary_total": monetary_total,
        "monetary_total_display": fmt_money(monetary_total) if monetary_total else "",
        "top_region_summary": top_region_summary,
        "image_path_1": get_image_base64("photos/community_photo1.jpg"),
        "image_path_2": get_image_base64("photos/community_photo2.jpg"),
    }
    return data


async def generate_pdf(data: dict):
    env = Environment(loader=FileSystemLoader(str(BASE_DIR)), autoescape=True)
    template = env.get_template(TEMPLATE_FILE)
    rendered_html = template.render(**data)

    safe_name = re.sub(r"[^A-Za-z0-9 _-]+", "", data["company_name"]).strip().replace(" ", "_")
    rendered_path = BASE_DIR / f"._rendered_{safe_name}.html"
    output_path = OUTPUT_DIR / f"Traceability_{safe_name}.pdf"
    rendered_path.write_text(rendered_html, encoding="utf-8")

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={"width": 794, "height": 1123}, device_scale_factor=1)
        await page.goto(rendered_path.as_uri(), wait_until="networkidle")
        await page.evaluate("document.fonts ? document.fonts.ready : Promise.resolve()")
        await page.wait_for_function("Array.from(document.images).every(img => img.complete)")
        await page.pdf(
            path=str(output_path),
            format="A4",
            print_background=True,
            prefer_css_page_size=True,
            margin={"top": "0mm", "right": "0mm", "bottom": "0mm", "left": "0mm"},
        )
        await browser.close()
    try:
        rendered_path.unlink()
    except OSError:
        pass
    print(f"🎉 PDF generated: {output_path}")


def main():
    if not EXCEL_FILE.exists():
        print(f"❌ Missing Excel: {EXCEL_FILE}")
        return
    if not TEMPLATE_PATH.exists():
        print(f"❌ Missing template: {TEMPLATE_PATH}")
        return

    print("Loading Excel database...")
    df = dedupe_columns(pd.read_excel(EXCEL_FILE))
    donor_col = find_col(df, ["DONOR", "DONANTE", "EMPRESA", "COMPANY"])
    if not donor_col:
        print("❌ Could not identify donor column.")
        return

    user_input = input("\nEnter company/donor name (or type 'ALL'): ").strip()

    if user_input.upper() == "ALL":
        async def run_all():
            for donor in df[donor_col].dropna().astype(str).unique():
                data = build_data(df, donor)
                if data:
                    await generate_pdf(data)
        asyncio.run(run_all())
    else:
        data = build_data(df, user_input)
        if data:
            asyncio.run(generate_pdf(data))


if __name__ == "__main__":
    main()
