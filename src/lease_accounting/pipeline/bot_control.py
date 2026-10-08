import json
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path


def write_bot_control(
    output_dir: Path,
    postings: list[dict],
    processing_issues: list[dict],
    json_name: str,
    bundle_name: str,
    preparation_errors: dict[str, list[str]],
) -> list[str]:
    json_path = output_dir / json_name
    if json_path.exists():
        raise ValueError("El control del bot ya existe; genera una nueva ejecucion para conservar su avance.")

    processed_at = datetime.now(timezone.utc).isoformat()
    invoices: dict[str, dict] = {}
    copied_supports: dict[str, str] = {}
    bundle_paths = set()
    warnings = []

    for posting in postings:
        invoice_id = str(posting["invoice_id"])
        invoice = invoices.setdefault(invoice_id, {
            "folio": invoice_id,
            "vendor": None,
            "vendors": [],
            "ruta": None,
            "estado": "preparada",
            "lista_para_bot": False,
            "documento_sap": None,
            "fecha_procesamiento": processed_at,
            "fecha_actualizacion": processed_at,
            "error": None,
            "pasos_confirmados": ["preparacion_completada"],
            "partes": [],
        })
        raw_vendor = posting.get("vendor")
        vendor = str(raw_vendor).strip() if raw_vendor is not None else None
        if vendor and vendor not in invoice["vendors"]:
            invoice["vendors"].append(vendor)
        if posting.get("invoice_total") is None:
            raise ValueError(f"La factura {invoice_id} no tiene un importe confirmado.")
        for relative_path in posting["archivos_csv"].values():
            path = output_dir / relative_path
            if not path.is_file():
                raise ValueError(f"No existe el CSV vinculado: {relative_path}")
        bundle_paths.add(posting["archivos_csv"]["allocated_costs"])

        source_path = posting.get("support_source_path")
        support_path = None
        if source_path and Path(source_path).suffix.lower() == ".xml":
            if source_path not in copied_supports:
                source = Path(source_path)
                if not source.is_file():
                    raise ValueError(f"No existe el soporte vinculado: {source}")
                support_dir = output_dir / "soportes_facturas"
                support_dir.mkdir(parents=True, exist_ok=True)
                target = support_dir / f"{len(copied_supports) + 1}_{source.name}"
                shutil.copy2(source, target)
                copied_supports[source_path] = str(target.relative_to(output_dir))
            support_path = copied_supports[source_path]

        part = {
            "parte": posting["posting_index"],
            "vendor": vendor or None,
            "tienda": posting.get("store"),
            "ceco": posting.get("ceco"),
            "profit_center": posting.get("profit_center"),
            "importe": posting.get("invoice_total"),
            "base": posting.get("amount"),
            "iva": posting.get("vat_total"),
            "moneda": posting.get("currency"),
            "tipo_documento": posting.get("document_type"),
            "fecha_factura": posting.get("invoice_date"),
            "identificador_lucy": posting.get("lucy_id"),
            "barcode": posting.get("barcode"),
            "archivo_factura": support_path,
            "archivo_origen": source_path,
            "csv": posting["archivos_csv"],
            "allocated_costs": posting["allocated_costs"],
            "estado": "preparada" if support_path else "pendiente_archivo",
            "documento_sap": None,
            "fecha_actualizacion": processed_at,
            "error": None if support_path else "No se dispone del soporte de la factura.",
            "pasos_confirmados": ["preparacion_completada"],
            "copia_lucy": None,
            "compensacion_sap": None,
        }
        invoice["partes"].append(part)

    for invoice in invoices.values():
        errors = list(dict.fromkeys(preparation_errors.get(invoice["folio"], [])))
        if errors:
            invoice["estado"] = "pendiente_revision"
        if len(invoice["vendors"]) == 1:
            invoice["vendor"] = invoice["vendors"][0]
            invoice["ruta"] = "single_vendor"
        else:
            invoice["estado"] = "pendiente_ruta"
            errors.append("La ruta single_vendor requiere exactamente un vendor; multivendor no esta habilitada.")
        if any(part["archivo_factura"] is None for part in invoice["partes"]):
            if invoice["estado"] == "preparada":
                invoice["estado"] = "pendiente_archivo"
            errors.append("Falta el soporte vinculado de una o mas partes.")
        invoice["importe"] = sum(part["importe"] for part in invoice["partes"])
        invoice["error"] = " ".join(errors) or None
        invoice["lista_para_bot"] = not errors
        invoice["validaciones"] = {
            "preparacion": "validada" if not errors else "pendiente_revision",
            "duplicados": "no_ejecutada",
        }
        for part in invoice["partes"]:
            if errors and part["estado"] == "preparada":
                part["estado"] = "pendiente_revision"
                part["error"] = invoice["error"]
            if not errors:
                part["pasos_confirmados"].append("validacion_preparacion_completada")
        if not errors:
            invoice["pasos_confirmados"].append("validacion_preparacion_completada")
        if errors:
            warnings.append(f"Factura {invoice['folio']}: {invoice['error']}")

    payload = {
        "schema_version": 2,
        "ejecucion": output_dir.name,
        "fecha_procesamiento": processed_at,
        "rutas_habilitadas": ["single_vendor"],
        "base_rutas": "directorio_del_json",
        "verificacion_duplicados": "no_ejecutada",
        "alcance_lista_para_bot": "datos_y_archivos_preparados; no confirma ausencia de duplicados",
        "resumen_validacion": {
            "total_facturas": len(invoices),
            "preparacion_validada": sum(invoice["lista_para_bot"] for invoice in invoices.values()),
            "pendientes": sum(not invoice["lista_para_bot"] for invoice in invoices.values()),
        },
        "facturas": list(invoices.values()),
        "incidencias_procesamiento": processing_issues,
    }
    serialized = json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False)
    temporary_path = json_path.with_suffix(".json.tmp")
    try:
        temporary_path.write_text(serialized, encoding="utf-8")
        temporary_path.replace(json_path)
    finally:
        temporary_path.unlink(missing_ok=True)

    with zipfile.ZipFile(output_dir / bundle_name, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
        bundle.write(json_path, arcname=json_name)
        for relative_path in sorted(bundle_paths):
            bundle.write(output_dir / relative_path, arcname=relative_path)
    return warnings
