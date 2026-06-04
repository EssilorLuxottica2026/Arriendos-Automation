from pathlib import Path
from uuid import uuid4
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

from .pipeline.processor import LeaseAccountingPipeline, PipelineError


BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "output"
INPUT_DIR = DATA_DIR / "input"
RAW_DIR = DATA_DIR / "raw"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}
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


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


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
        for support_file in request.files.getlist("invoice_support_files"):
            if support_file and support_file.filename:
                support_paths.append(_save_upload(support_file, batch_dir, allowed_extensions=SUPPORT_EXTENSIONS))

        period = request.form.get("period", "").strip()
        pipeline = LeaseAccountingPipeline(output_dir=OUTPUT_DIR)
        result = pipeline.run(
            invoices_path=invoices_path,
            control_path=control_path,
            contracts_path=contracts_path,
            prorateo_path=prorateo_path,
            distribution_path=distribution_path,
            history_path=history_path,
            support_paths=support_paths,
            macro_template_path=macro_workbook_path,
            period=period or None,
        )
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
