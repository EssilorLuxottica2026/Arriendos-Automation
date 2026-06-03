from pathlib import Path
from uuid import uuid4

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

from pipeline.processor import LeaseAccountingPipeline, PipelineError


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "output"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}
SUPPORT_EXTENSIONS = {".pdf", ".xml"}


app = Flask(__name__)
app.config["SECRET_KEY"] = "lease-accounting-local-dev"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


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


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    required_keys = [
        "control_file",
        "contracts_file",
        "prorateo_file",
        "distribution_file",
        "invoices_file",
    ]
    missing = [key for key in required_keys if key not in request.files or not request.files[key].filename]
    if missing:
        flash("Please upload all required files before processing.")
        return redirect(url_for("index"))

    batch_id = uuid4().hex
    batch_dir = UPLOAD_DIR / batch_id
    batch_dir.mkdir(parents=True, exist_ok=True)

    try:
        files = {key: _save_upload(request.files[key], batch_dir) for key in required_keys}
        optional_history = request.files.get("history_file")
        history_path = _save_upload(optional_history, batch_dir) if optional_history and optional_history.filename else None
        optional_macro = request.files.get("macro_workbook_file")
        macro_workbook_path = _save_upload(optional_macro, batch_dir) if optional_macro and optional_macro.filename else None
        support_paths = []
        for support_file in request.files.getlist("invoice_support_files"):
            if support_file and support_file.filename:
                support_paths.append(_save_upload(support_file, batch_dir, allowed_extensions=SUPPORT_EXTENSIONS))

        period = request.form.get("period", "").strip()
        pipeline = LeaseAccountingPipeline(output_dir=OUTPUT_DIR)
        result = pipeline.run(
            invoices_path=files["invoices_file"],
            control_path=files["control_file"],
            contracts_path=files["contracts_file"],
            prorateo_path=files["prorateo_file"],
            distribution_path=files["distribution_file"],
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


if __name__ == "__main__":
    app.run(debug=True)
