from pathlib import Path
from uuid import uuid4
import json
import re
import shutil
from datetime import datetime

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

from .pipeline.processor import LeaseAccountingPipeline, PipelineError, normalize_learning_phrase


BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "output"
INPUT_DIR = DATA_DIR / "input"
RAW_DIR = DATA_DIR / "raw"
LEARNING_DICTIONARY_PATH = BASE_DIR / "Diccionario_Conceptos_Simple.xlsx"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv", ".xml"}
SUPPORT_EXTENSIONS = {".pdf", ".xml"}


app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
app.config["SECRET_KEY"] = "lease-accounting-local-dev"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
INPUT_DIR.mkdir(exist_ok=True)
RAW_DIR.mkdir(exist_ok=True)


# Standard filenames for reference databases
MASTER_FILENAMES = {
    "control_file": "CONTROL_ARRI_ADMON.xlsx",
    "contracts_file": "Contratos_con_condiciones.xlsx",
    "prorateo_file": "PRORATEO.xlsx",
    "distribution_file": "Cuadro_de_distribucion.xls",
    "history_file": "FACTURAS_CONTABILIZADAS.xlsx",
    "macro_workbook_file": "Arriendos_Macro.xlsm"
}


def _allowed(filename: str, allowed_extensions: set[str] | None = None) -> bool:
    target_extensions = allowed_extensions or ALLOWED_EXTENSIONS
    return Path(filename).suffix.lower() in target_extensions


def _save_upload(file_storage, batch_dir: Path, allowed_extensions: set[str] | None = None) -> Path:
    filename = secure_filename(file_storage.filename or "")
    if not filename:
        raise PipelineError("One of the uploaded files does not have a valid filename.")
    if not _allowed(filename, allowed_extensions=allowed_extensions):
        allowed_hint = ", ".join(sorted(ext.lstrip(".") for ext in (allowed_extensions or ALLOWED_EXTENSIONS)))
        raise PipelineError(f"Unsupported file type for {filename}. Use {allowed_hint}.")

    target = batch_dir / filename
    file_storage.save(target)
    return target


def _resolve_and_save_master(key: str, file_storage, batch_dir: Path) -> Path | None:
    standard_name = MASTER_FILENAMES.get(key)
    if not standard_name:
        return None

    # If the user uploaded a new version of the file:
    if file_storage and file_storage.filename:
        temp_path = _save_upload(file_storage, batch_dir)
        target_master_path = INPUT_DIR / standard_name
        
        # If the file already exists in data/input, back it up to data/raw/
        if target_master_path.exists():
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem = target_master_path.stem
            suffix = target_master_path.suffix
            backup_name = f"{stem}_replaced_{timestamp}{suffix}"
            
            backup_path = RAW_DIR / backup_name
            shutil.move(str(target_master_path), str(backup_path))
        
        # Copy the new file to data/input
        shutil.copy2(str(temp_path), str(target_master_path))
        return target_master_path

    # If the user did NOT upload a file, check if it exists in data/input/
    target_master_path = INPUT_DIR / standard_name
    if target_master_path.exists():
        return target_master_path

    # If it is a required master file but doesn't exist anywhere
    if key in ["control_file", "contracts_file", "prorateo_file", "distribution_file"]:
        raise PipelineError(f"El archivo maestro obligatorio '{standard_name}' no se encuentra en el servidor. Por favor, súbelo al menos una vez.")
        
    return None


def _job_path(batch_id: str) -> Path:
    return UPLOAD_DIR / batch_id / "concept_review_job.json"


def _save_review_job(batch_id: str, job: dict) -> None:
    path = _job_path(batch_id)
    path.write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_review_job(batch_id: str) -> dict:
    path = _job_path(batch_id)
    if not path.exists():
        raise PipelineError("No se encontro la sesion de revision de conceptos. Vuelve a procesar el archivo.")
    return json.loads(path.read_text(encoding="utf-8"))


def _run_pipeline_from_job(job: dict) -> dict:
    pipeline = LeaseAccountingPipeline(output_dir=OUTPUT_DIR, learning_dictionary_path=LEARNING_DICTIONARY_PATH)
    return pipeline.run(
        invoices_path=Path(job["invoices_path"]),
        control_path=Path(job["control_path"]),
        contracts_path=Path(job["contracts_path"]),
        prorateo_path=Path(job["prorateo_path"]),
        distribution_path=Path(job["distribution_path"]),
        history_path=Path(job["history_path"]) if job.get("history_path") else None,
        support_paths=[Path(path) for path in job.get("support_paths", [])],
        support_split_factors=job.get("support_split_factors") or None,
        macro_template_path=Path(job["macro_template_path"]) if job.get("macro_template_path") else None,
        period=job.get("period") or None,
    )


def _append_learning_dictionary_phrases(assignments: list[dict]) -> None:
    if not LEARNING_DICTIONARY_PATH.exists():
        raise PipelineError(f"No existe el diccionario de conceptos: {LEARNING_DICTIONARY_PATH.name}")
    try:
        workbook = load_workbook(LEARNING_DICTIONARY_PATH)
    except PermissionError as exc:
        raise PipelineError("Cierra el archivo Diccionario_Conceptos_Simple.xlsx antes de guardar nuevas reglas.") from exc
    if "Diccionario" not in workbook.sheetnames:
        raise PipelineError("El diccionario debe tener una hoja llamada 'Diccionario'.")

    sheet = workbook["Diccionario"]
    headers = {str(cell.value or "").strip().lower(): idx for idx, cell in enumerate(sheet[1], start=1)}
    concept_col = headers.get("concepto")
    phrases_col = headers.get("frases/palabras") or headers.get("frases") or headers.get("palabras")
    if not concept_col or not phrases_col:
        raise PipelineError("El diccionario debe tener columnas 'Concepto' y 'Frases/palabras'.")

    concept_rows = {}
    for row_idx in range(2, sheet.max_row + 1):
        concept = str(sheet.cell(row_idx, concept_col).value or "").strip()
        if concept:
            concept_rows[concept.upper()] = row_idx

    for assignment in assignments:
        concept = (assignment.get("concept") or "").strip()
        phrase_text = (assignment.get("phrase") or "").strip()
        if not concept or not phrase_text:
            continue
        row_idx = concept_rows.get(concept.upper())
        if not row_idx:
            row_idx = sheet.max_row + 1
            sheet.cell(row_idx, concept_col).value = concept
            concept_rows[concept.upper()] = row_idx
        current = str(sheet.cell(row_idx, phrases_col).value or "").strip()
        existing = [part.strip() for part in current.split(",") if part.strip()]
        existing_keys = {part.upper() for part in existing}
        for phrase in [part.strip() for part in phrase_text.split(",") if part.strip()]:
            phrase = normalize_learning_phrase(phrase)
            if phrase and phrase.upper() not in existing_keys:
                existing.append(phrase)
                existing_keys.add(phrase.upper())
        sheet.cell(row_idx, phrases_col).value = ", ".join(existing)

    try:
        workbook.save(LEARNING_DICTIONARY_PATH)
    except PermissionError as exc:
        raise PipelineError("No pude guardar el diccionario. Cierra Diccionario_Conceptos_Simple.xlsx e intenta de nuevo.") from exc


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


def _dictionary_sheet(workbook):
    if "Diccionario" not in workbook.sheetnames:
        raise PipelineError("El diccionario debe tener una hoja llamada 'Diccionario'.")
    sheet = workbook["Diccionario"]
    headers = {str(cell.value or "").strip().lower(): idx for idx, cell in enumerate(sheet[1], start=1)}
    concept_col = headers.get("concepto")
    account_col = headers.get("gl account") or headers.get("gl_account") or headers.get("cuenta")
    phrases_col = headers.get("frases/palabras") or headers.get("frases") or headers.get("palabras")
    if not concept_col or not account_col or not phrases_col:
        raise PipelineError("El diccionario debe tener columnas 'Concepto', 'GL Account' y 'Frases/palabras'.")
    return sheet, concept_col, account_col, phrases_col


def _load_dictionary_workbook():
    if not LEARNING_DICTIONARY_PATH.exists():
        raise PipelineError(f"No existe el diccionario de conceptos: {LEARNING_DICTIONARY_PATH.name}")
    try:
        return load_workbook(LEARNING_DICTIONARY_PATH)
    except PermissionError as exc:
        raise PipelineError("Cierra Diccionario_Conceptos_Simple.xlsx para modificar la configuracion.") from exc


def _validate_concept_account(concept: str, account: str) -> tuple[str, str]:
    concept = re.sub(r"\s+", " ", concept).strip().upper()
    account = re.sub(r"\s+", "_", account).strip().upper()
    if not concept or not account:
        raise PipelineError("El concepto y el GL Account son obligatorios.")
    if account != "SEGUN_CONTRATO" and not re.fullmatch(r"\d{6,15}", account):
        raise PipelineError("El GL Account debe contener entre 6 y 15 digitos, o usar SEGUN_CONTRATO para RENTA.")
    return concept, account


@app.route("/settings", methods=["GET", "POST"])
def settings():
    try:
        workbook = _load_dictionary_workbook()
        sheet, concept_col, account_col, phrases_col = _dictionary_sheet(workbook)

        if request.method == "POST":
            action = request.form.get("action", "").strip().lower()
            rows_by_concept = {
                str(sheet.cell(row_idx, concept_col).value or "").strip().upper(): row_idx
                for row_idx in range(2, sheet.max_row + 1)
                if str(sheet.cell(row_idx, concept_col).value or "").strip()
            }

            if action == "delete":
                original = request.form.get("original_concept", "").strip().upper()
                row_idx = rows_by_concept.get(original)
                if not row_idx:
                    raise PipelineError("El concepto que intentas eliminar ya no existe.")
                sheet.delete_rows(row_idx, 1)
                success_message = f"Concepto {original} eliminado correctamente."
            else:
                concept, account = _validate_concept_account(
                    request.form.get("concept", ""),
                    request.form.get("account", ""),
                )

            if action == "add":
                if concept in rows_by_concept:
                    raise PipelineError(f"El concepto {concept} ya existe. Puedes editarlo en la tabla.")
                row_idx = sheet.max_row + 1
                sheet.cell(row_idx, concept_col).value = concept
                sheet.cell(row_idx, account_col).value = account
                sheet.cell(row_idx, phrases_col).value = ""
                success_message = f"Concepto {concept} agregado correctamente."
            elif action == "update":
                original = request.form.get("original_concept", "").strip().upper()
                row_idx = rows_by_concept.get(original)
                if not row_idx:
                    raise PipelineError("El concepto que intentas editar ya no existe.")
                duplicate_row = rows_by_concept.get(concept)
                if duplicate_row and duplicate_row != row_idx:
                    raise PipelineError(f"Ya existe otro concepto llamado {concept}.")
                sheet.cell(row_idx, concept_col).value = concept
                sheet.cell(row_idx, account_col).value = account
                success_message = f"Concepto {concept} actualizado correctamente."
            elif action != "delete":
                raise PipelineError("Accion de configuracion no valida.")

            try:
                workbook.save(LEARNING_DICTIONARY_PATH)
            except PermissionError as exc:
                raise PipelineError("Cierra Diccionario_Conceptos_Simple.xlsx antes de guardar los cambios.") from exc
            flash(success_message, "success")
            return redirect(url_for("settings"))

        concepts = []
        for row_idx in range(2, sheet.max_row + 1):
            concept = str(sheet.cell(row_idx, concept_col).value or "").strip()
            if not concept:
                continue
            phrases = [part.strip() for part in str(sheet.cell(row_idx, phrases_col).value or "").split(",") if part.strip()]
            concepts.append(
                {
                    "concept": concept,
                    "account": str(sheet.cell(row_idx, account_col).value or "").strip(),
                    "phrase_count": len(phrases),
                }
            )
        return render_template("settings.html", concepts=concepts)
    except PipelineError as exc:
        flash(str(exc), "error")
        return redirect(url_for("settings")) if request.method == "POST" else redirect(url_for("index"))


@app.route("/process", methods=["POST"])
def process():
    invoices_file = request.files.get("invoices_file")
    if not invoices_file or not invoices_file.filename:
        flash("Por favor, sube el archivo de facturas consolidado (Invoices file).")
        return redirect(url_for("index"))

    batch_id = uuid4().hex
    batch_dir = UPLOAD_DIR / batch_id
    batch_dir.mkdir(parents=True, exist_ok=True)

    try:
        invoices_path = _save_upload(invoices_file, batch_dir)
        
        control_path = _resolve_and_save_master("control_file", request.files.get("control_file"), batch_dir)
        contracts_path = _resolve_and_save_master("contracts_file", request.files.get("contracts_file"), batch_dir)
        prorateo_path = _resolve_and_save_master("prorateo_file", request.files.get("prorateo_file"), batch_dir)
        distribution_path = _resolve_and_save_master("distribution_file", request.files.get("distribution_file"), batch_dir)
        
        history_path = _resolve_and_save_master("history_file", request.files.get("history_file"), batch_dir)
        macro_workbook_path = _resolve_and_save_master("macro_workbook_file", request.files.get("macro_workbook_file"), batch_dir)

        support_paths = []
        support_split_factors = {}
        for support_file in request.files.getlist("invoice_support_files"):
            if support_file and support_file.filename:
                support_paths.append(_save_upload(support_file, batch_dir, allowed_extensions=SUPPORT_EXTENSIONS))
        for field_name, split_factor in [("invoice_support_files_split_2", 2), ("invoice_support_files_split_3", 3)]:
            for support_file in request.files.getlist(field_name):
                if support_file and support_file.filename:
                    support_path = _save_upload(support_file, batch_dir, allowed_extensions=SUPPORT_EXTENSIONS)
                    support_paths.append(support_path)
                    support_split_factors[str(support_path)] = split_factor

        period = request.form.get("period", "").strip()
        job = {
            "invoices_path": str(invoices_path),
            "control_path": str(control_path),
            "contracts_path": str(contracts_path),
            "prorateo_path": str(prorateo_path),
            "distribution_path": str(distribution_path),
            "history_path": str(history_path) if history_path else None,
            "support_paths": [str(path) for path in support_paths],
            "support_split_factors": support_split_factors,
            "macro_template_path": str(macro_workbook_path) if macro_workbook_path else None,
            "period": period or None,
        }

        pipeline = LeaseAccountingPipeline(output_dir=OUTPUT_DIR, learning_dictionary_path=LEARNING_DICTIONARY_PATH)
        data, output_df, validation_df = pipeline.prepare(
            invoices_path=invoices_path,
            control_path=control_path,
            contracts_path=contracts_path,
            prorateo_path=prorateo_path,
            distribution_path=distribution_path,
            history_path=history_path,
            support_paths=support_paths,
            support_split_factors=support_split_factors or None,
            macro_template_path=macro_workbook_path,
            period=period or None,
        )
        review_items = pipeline.build_concept_review_items(output_df)
        if review_items:
            _save_review_job(batch_id, job)
            concepts = pipeline.learning_concepts()
            return render_template("concept_review.html", batch_id=batch_id, items=review_items, concepts=concepts)

        result = _run_pipeline_from_job(job)
    except PipelineError as exc:
        flash(str(exc))
        return redirect(url_for("index"))
    except Exception as exc:  # pragma: no cover
        flash(f"Unexpected error: {exc}")
        return redirect(url_for("index"))

    return render_template("result.html", result=result)


@app.route("/concept-review/<batch_id>", methods=["POST"])
def concept_review(batch_id: str):
    try:
        job = _load_review_job(batch_id)
        item_count = int(request.form.get("item_count", "0"))
        assignments = []
        for idx in range(item_count):
            concept = request.form.get(f"concept_{idx}", "").strip()
            phrase = request.form.get(f"phrase_{idx}", "").strip()
            if not concept or not phrase:
                raise PipelineError("Todos los items pendientes necesitan concepto y frase/palabra.")
            assignments.append({"concept": concept, "phrase": phrase})
        _append_learning_dictionary_phrases(assignments)
        result = _run_pipeline_from_job(job)
    except PipelineError as exc:
        flash(str(exc))
        return redirect(url_for("index"))
    except Exception as exc:  # pragma: no cover
        flash(f"Unexpected error: {exc}")
        return redirect(url_for("index"))

    return render_template("result.html", result=result)


@app.route("/download/<path:filename>", methods=["GET"])
def download(filename: str):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        flash("Requested file does not exist.")
        return redirect(url_for("index"))
    return send_file(file_path, as_attachment=True)
