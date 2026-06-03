from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = BASE_DIR / "data" / "normalized_inputs"
OUTPUT_DIR.mkdir(exist_ok=True)

DOWNLOADS = Path.home() / "Downloads"


@dataclass(frozen=True)
class SourceFiles:
    invoices: Path = DOWNLOADS / "invoices file.xlsx"
    control: Path = DOWNLOADS / "CONTROL ARRI ADMON.xlsx"
    contracts: Path = DOWNLOADS / "03. Contratos con condiciones Marzo 2026 1.xlsx"
    prorateo: Path = DOWNLOADS / "PRORATEO.xlsx"
    distribution: Path = DOWNLOADS / "Cuadro de distribucion 2023.xls"
    history: Path = DOWNLOADS / "FACTURAS CONTABILIZADAS ARRIENDOS.xlsx"


SOURCES = SourceFiles()


def normalize_text(value) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).replace("\xa0", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text or None


def normalize_col(value) -> str:
    text = normalize_text(value) or ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-zA-Z0-9]+", "_", text).strip("_").lower()
    return text or "unnamed"


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [normalize_col(col) for col in cleaned.columns]
    for idx, col in enumerate(cleaned.columns):
        series = cleaned.iloc[:, idx]
        if series.dtype == object:
            cleaned.iloc[:, idx] = series.map(normalize_text)
    return cleaned


def open_sheet_by_alias(path: Path, aliases: list[str], header: int = 0) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    alias_keys = {normalize_col(alias) for alias in aliases}
    for sheet_name in xl.sheet_names:
        if normalize_col(sheet_name) in alias_keys:
            return pd.read_excel(path, sheet_name=sheet_name, header=header)
    raise ValueError(f"None of the sheet aliases {aliases} were found in {path.name}")


def extract_primary_code(value) -> str | None:
    text = normalize_text(value)
    if not text:
        return None
    match = re.search(r"\d+", text)
    return match.group(0) if match else None


def to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def build_invoices_normalized(path: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    raw = clean_dataframe(pd.read_excel(path))
    normalized = pd.DataFrame(
        {
            "invoice_id": raw.get("numero_de_factura"),
            "vendor_code": raw.get("vendor").map(extract_primary_code),
            "vendor_name": raw.get("proveedor"),
            "store": raw.get("tienda"),
            "ceco": raw.get("ceco"),
            "total_invoice": to_numeric(raw.get("valor_facturado")),
            "invoice_date": pd.NaT,
            "subtotal": pd.NA,
            "vat_total": pd.NA,
            "nit": raw.get("nit"),
            "source_file": path.name,
        }
    )
    sheets = {
        "invoices_master": normalized,
        "invoices_raw_standardized": raw,
    }
    return sheets, summary_sheet(path.name, normalized, "Invoice intake normalized from email Excel")


def build_control_normalized(path: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    arriendo = clean_dataframe(pd.read_excel(path, sheet_name="ARRIENDO 2022", header=1))
    admon = clean_dataframe(pd.read_excel(path, sheet_name="ADMON 2022", header=1))
    control_adm = clean_dataframe(pd.read_excel(path, sheet_name="CONTROL ADM", header=4))
    az_admon = clean_dataframe(pd.read_excel(path, sheet_name="a-z admon", header=0))

    arriendo_master = pd.DataFrame(
        {
            "store": arriendo.get("tienda"),
            "store_name": arriendo.get("centro"),
            "city": arriendo.get("ciudad"),
            "ceco": arriendo.get("c_c"),
            "vendor_code": arriendo.get("acreedor_arriendo").map(extract_primary_code),
            "vendor_name": arriendo.get("proveedor_arriendo"),
            "nit": arriendo.get("nit"),
            "monthly_canon_reference": to_numeric(arriendo.get("canon_mes_2")),
            "source_sheet": "ARRIENDO 2022",
        }
    )
    admon_master = pd.DataFrame(
        {
            "store": admon.get("tienda"),
            "store_name": admon.get("descripcion_administracion"),
            "ceco": admon.get("c_c"),
            "vendor_code": admon.get("acreedor_admon").map(extract_primary_code),
            "vendor_name": admon.get("proveedor_administracion"),
            "nit": admon.get("nit"),
            "monthly_fee_reference": to_numeric(admon.get("canon_mes")),
            "source_sheet": "ADMON 2022",
        }
    )
    control_adm_master = pd.DataFrame(
        {
            "store": control_adm.get("tienda"),
            "store_name": control_adm.get("centro"),
            "ceco": control_adm.get("c_c"),
            "vendor_code_rent": control_adm.get("acreedor_arriendo").map(extract_primary_code),
            "vendor_name_rent": control_adm.get("proveedor_arriendo"),
            "source_sheet": "CONTROL ADM",
        }
    )

    sheets = {
        "control_arriendo_master": arriendo_master,
        "control_admon_master": admon_master,
        "control_reference_master": control_adm_master,
        "vendor_dictionary_admon": az_admon,
    }
    return sheets, summary_sheet(path.name, control_adm_master, "Control workbook reorganized into machine-readable tables")


def build_contracts_normalized(path: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    xl = pd.ExcelFile(path)
    frames = []
    for sheet_name in xl.sheet_names:
        if sheet_name.strip().upper() in {"CONDICIONES", "CONTABILIZACION", "CUENTAS"}:
            continue
        raw = clean_dataframe(xl.parse(sheet_name, header=4))
        if "ceco" not in raw.columns or "end_of_term" not in raw.columns:
            continue
        frame = pd.DataFrame(
            {
                "country_sheet": sheet_name.strip(),
                "ceco": raw.get("ceco"),
                "contract": raw.get("contract"),
                "contract_name": raw.get("contract_name"),
                "contract_start_date": raw.get("contract_start_date"),
                "end_of_term": raw.get("end_of_term"),
                "end_of_term_virtual": raw.get("end_of_term_en_virtual_contract"),
                "status_virtual": raw.get("status_en_virtual_contract"),
                "status_rem": raw.get("status_en_rem"),
                "rent_min": to_numeric(raw.get("rent_min_rent")),
                "sell_media": to_numeric(raw.get("sell_media")),
                "sell_cam": to_numeric(raw.get("sell_cam")),
            }
        )
        frames.append(frame)

    contracts_all = pd.concat(frames, ignore_index=True)
    contracts_latest = (
        contracts_all.sort_values(["ceco", "end_of_term"], ascending=[True, False])
        .drop_duplicates(subset=["ceco"], keep="first")
        .reset_index(drop=True)
    )
    sheets = {
        "contracts_all": contracts_all,
        "contracts_latest_by_ceco": contracts_latest,
    }
    return sheets, summary_sheet(path.name, contracts_latest, "Contracts workbook normalized for CECO-based matching")


def build_prorateo_normalized(path: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    raw = clean_dataframe(pd.read_excel(path, sheet_name="CONSOLIDADO ARRIENDOS PROY IVA"))
    master = pd.DataFrame(
        {
            "vendor_code": raw.get("acreedor_arriendo").map(extract_primary_code),
            "vendor_name": raw.get("proveedor_arriendo"),
            "store": raw.get("tienda"),
            "store_name": raw.get("nombre_tienda"),
            "city": raw.get("ciudad"),
            "ceco": raw.get("ceco"),
            "nit": raw.get("nit"),
            "current_canon_value": to_numeric(raw.get("valor_actual_canon")),
            "vat_canon": to_numeric(raw.get("iva_canon")),
            "vat_vw_amount": to_numeric(raw.get("iva_deducible_vw")),
            "vat_vq_amount": to_numeric(raw.get("iva_prorrateo_vq")),
            "vat_vw_percent": to_numeric(raw.get("porcentaje_iva_vw")),
            "vat_vq_percent": to_numeric(raw.get("porcentaje_iva_vq")),
        }
    )
    sheets = {
        "prorateo_master": master,
    }
    return sheets, summary_sheet(path.name, master, "VAT proration workbook normalized for vendor/store/CECO lookup")


def build_distribution_normalized(path: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    raw_distribution = clean_dataframe(open_sheet_by_alias(path, ["Distrubución", "Distrubucion"], header=1))
    raw_unique = clean_dataframe(open_sheet_by_alias(path, ["Grupo Unico"], header=0))
    raw_exito = clean_dataframe(open_sheet_by_alias(path, ["Grupo Éxito", "Grupo Exito"], header=2))

    rules = []
    current_source_vendor = None
    current_contract = None
    for _, row in raw_distribution.iterrows():
        contract_name = row.iloc[0] if len(row) > 0 else None
        label = row.iloc[2] if len(row) > 2 else None
        target_vendor = row.iloc[3] if len(row) > 3 else None
        source_vendor = row.iloc[5] if len(row) > 5 else None
        if normalize_text(contract_name):
            current_contract = contract_name
            current_source_vendor = source_vendor
        pct = extract_percent(label)
        if pct is None or pd.isna(target_vendor):
            continue
        rules.append(
            {
                "contract_name": current_contract,
                "source_vendor_code": extract_primary_code(current_source_vendor),
                "target_vendor_code": extract_primary_code(target_vendor),
                "distribution_percent": pct,
                "distribution_label": normalize_text(label),
            }
        )

    rules_df = pd.DataFrame(rules)
    sheets = {
        "distribution_rules": rules_df,
        "group_unico_dictionary": raw_unique,
        "group_exito_dictionary": raw_exito,
    }
    return sheets, summary_sheet(path.name, rules_df, "Distribution workbook normalized into explicit split rules")


def build_history_normalized(path: Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    raw = clean_dataframe(pd.read_excel(path, sheet_name="Hoja1"))
    history = pd.DataFrame(
        {
            "invoice_id": raw.get("unnamed_0"),
            "vendor_name": raw.get("facturas_contabilizadas_arriendos"),
            "sap_document": raw.get("unnamed_2"),
            "source_sheet": "Hoja1",
        }
    )
    accounts = clean_dataframe(pd.read_excel(path, sheet_name="CONSOLIDADO DE CUENTAS ARRI", header=0))
    sheets = {
        "historical_invoices": history,
        "accounts_reference": accounts,
    }
    return sheets, summary_sheet(path.name, history, "Historical accounting workbook normalized for auditing/reference")


def extract_percent(value) -> float | None:
    text = normalize_text(value)
    if not text:
        return None
    match = re.search(r"(\d+(?:[.,]\d+)?)\s*%", text)
    if not match:
        return None
    return float(match.group(1).replace(",", ".")) / 100.0


def summary_sheet(source_name: str, df: pd.DataFrame, purpose: str) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"source_file": source_name, "rows": len(df), "columns": len(df.columns), "purpose": purpose},
        ]
    )


def write_workbook(path: Path, sheets: dict[str, pd.DataFrame], summary: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)


def main() -> None:
    workbook_builders = [
        ("invoices_input_normalized.xlsx", build_invoices_normalized, SOURCES.invoices),
        ("control_arri_admon_normalized.xlsx", build_control_normalized, SOURCES.control),
        ("contracts_conditions_normalized.xlsx", build_contracts_normalized, SOURCES.contracts),
        ("prorateo_normalized.xlsx", build_prorateo_normalized, SOURCES.prorateo),
        ("distribution_normalized.xlsx", build_distribution_normalized, SOURCES.distribution),
        ("historical_facturas_normalized.xlsx", build_history_normalized, SOURCES.history),
    ]

    manifest_rows = []
    for output_name, builder, source_path in workbook_builders:
        sheets, summary = builder(source_path)
        output_path = OUTPUT_DIR / output_name
        write_workbook(output_path, sheets, summary)
        manifest_rows.append(
            {
                "output_file": output_name,
                "source_file": source_path.name,
                "sheets_created": len(sheets) + 1,
                "description": summary.loc[0, "purpose"],
            }
        )

    manifest = pd.DataFrame(manifest_rows)
    manifest.to_excel(OUTPUT_DIR / "normalized_manifest.xlsx", index=False)
    print(f"Normalized workbooks created in {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
