"""Local web app: bulk logger PDFs -> Temperature / Humidity Excel."""

from __future__ import annotations

import math
import os
import shutil
import threading
import traceback
import uuid
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from excel_export import write_workbook
from parser import (
    DL_NAME_RE,
    logger_column_name,
    logger_sort_key,
    merge_columns,
    parse_one_worker,
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR") or APP_DIR)
OUTPUT_DIR = DATA_DIR / "output"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# On Render there is no local DL folder to point at, so the folder-path input
# is hidden and everything arrives through the browser upload.
CLOUD_MODE = bool(os.environ.get("RENDER") or os.environ.get("CLOUD_MODE") == "1")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or uuid.uuid4().hex
app.config["MAX_CONTENT_LENGTH"] = 2000 * 1024 * 1024
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

DEFAULT_FOLDER = r"C:\Users\Juby John\Downloads\Aramax WH\Aramax WH"
# Render Free is 512 MB. Many parse processes will get the service killed.
WORKERS = 2 if CLOUD_MODE else min(12, max(4, os.cpu_count() or 4))

JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


def _safe_upload_name(filename: str) -> str | None:
    name = Path(str(filename).replace("\\", "/")).name
    if Path(name).suffix.lower() not in {".pdf", ".xls", ".xlsx"}:
        return None
    if not DL_NAME_RE.search(Path(name).stem):
        return None
    return name


def _clean(value: float) -> float | None:
    """JSON has no NaN, so blank readings go out as null."""
    return None if math.isnan(value) else round(value, 1)


def _update_job(job_id: str, **fields) -> None:
    with JOBS_LOCK:
        JOBS.setdefault(job_id, {}).update(fields)


def _collect_source_files(folder: Path) -> list[Path]:
    files = []
    for path in folder.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in {".pdf", ".xls", ".xlsx"}:
            continue
        if path.name.startswith("~$") or path.name.lower().startswith("temperature"):
            continue
        if not DL_NAME_RE.search(path.stem):
            continue
        files.append(path)
    return sorted(files, key=lambda p: p.name.lower())


def _dedupe_prefer_pdf(files: list[Path]) -> list[Path]:
    """One source per logger. Render prefers Excel; desktop prefers PDF."""
    by_logger: dict[str, Path] = {}
    for path in files:
        key = logger_column_name(path)
        existing = by_logger.get(key)
        if existing is None:
            by_logger[key] = path
            continue
        path_excel = path.suffix.lower() in {".xls", ".xlsx"}
        existing_excel = existing.suffix.lower() in {".xls", ".xlsx"}
        if CLOUD_MODE:
            if path_excel and not existing_excel:
                by_logger[key] = path
        elif path.suffix.lower() == ".pdf" and existing.suffix.lower() != ".pdf":
            by_logger[key] = path
    return sorted(by_logger.values(), key=lambda p: logger_sort_key(logger_column_name(p)))


def run_job(job_id: str, files: list[Path], batch_dir: Path | None = None) -> None:
    try:
        files = [path for path in files if DL_NAME_RE.search(path.stem)]
        files = _dedupe_prefer_pdf(files)
        if not files:
            _update_job(job_id, status="error", error="No PDF or Excel logger files found. Upload files named like DL 01, DL 02.")
            return

        _update_job(job_id, status="running", total=len(files), current=0, message="Reading logger files...")
        file_columns: dict[str, tuple] = {}
        warnings: list[str] = []

        done = 0

        def _take_result(logger, columns, error, path):
            nonlocal done
            done += 1
            if error:
                warnings.append(error)
            else:
                file_columns[logger] = columns
            _update_job(job_id, current=done, message=f"Parsed {done}/{len(files)}: {path.name}")
            if done == 1 or done % 10 == 0 or done == len(files):
                print(f"Parsed {done}/{len(files)}: {path.name}", flush=True)

        if CLOUD_MODE:
            for path in files:
                logger, columns, error = parse_one_worker(str(path.resolve()))
                _take_result(logger, columns, error, path)
        else:
            try:
                ctx = multiprocessing.get_context("spawn")
                with ProcessPoolExecutor(max_workers=WORKERS, mp_context=ctx) as pool:
                    futures = {pool.submit(parse_one_worker, str(path.resolve())): path for path in files}
                    for future in as_completed(futures):
                        logger, columns, error = future.result()
                        _take_result(logger, columns, error, futures[future])
            except Exception as exc:
                print(f"process pool unavailable ({exc}); reading files here instead", flush=True)
                for path in files:
                    logger, columns, error = parse_one_worker(str(path.resolve()))
                    _take_result(logger, columns, error, path)

        if not file_columns:
            detail = warnings[0] if warnings else "The files had no DATE / Time / temperature / humidity rows."
            extra = f" ({len(warnings)} files)" if len(warnings) > 1 else ""
            _update_job(
                job_id,
                status="error",
                error=f"Could not extract readings.{extra} {detail}",
                warnings=warnings,
            )
            return

        _update_job(job_id, message="Aligning timestamps and splitting Temperature / Humidity...")
        loggers, timestamps, temps, hums = merge_columns(file_columns)
        if not timestamps:
            _update_job(job_id, status="error", error="No timestamps found after merge.")
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"Temperature_Humidity_{stamp}.xlsx"
        out_path = OUTPUT_DIR / out_name
        _update_job(job_id, message="Writing Excel workbook...")
        write_workbook(
            out_path,
            loggers,
            timestamps,
            temps,
            hums,
            on_progress=lambda message: _update_job(job_id, message=message),
        )

        preview_rows = 40
        preview_loggers = loggers[:18]
        preview = []
        for i, ts in enumerate(timestamps[:preview_rows]):
            row = {
                "DATE": ts.strftime("%d/%m/%Y"),
                "Time": ts.strftime("%H:%M"),
                "temp": {name: _clean(temps[name][i]) for name in preview_loggers},
                "rh": {name: _clean(hums[name][i]) for name in preview_loggers},
            }
            preview.append(row)

        _update_job(
            job_id,
            status="done",
            message="Finished",
            download=out_name,
            warnings=warnings,
            stats={
                "loggers": len(loggers),
                "rows": len(timestamps),
                "files": len(files),
                "start": timestamps[0].strftime("%d/%m/%Y %H:%M"),
                "end": timestamps[-1].strftime("%d/%m/%Y %H:%M"),
            },
            preview_loggers=preview_loggers,
            preview=preview,
        )
    except Exception:
        _update_job(job_id, status="error", error=traceback.format_exc())
    finally:
        # Uploaded sources are only needed while the job runs; the cloud disk is small.
        if batch_dir is not None and batch_dir.is_dir():
            shutil.rmtree(batch_dir, ignore_errors=True)


@app.before_request
def require_password():
    """One shared password, only when APP_PASSWORD is configured."""
    if not APP_PASSWORD or session.get("ok"):
        return None
    if request.endpoint in {"login", "static", "healthz"}:
        return None
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "error": "Session expired. Reload and sign in again."}), 401
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    error = ""
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["ok"] = True
            session.permanent = True
            return redirect(url_for("index"))
        error = "Wrong password."
    return render_template("login.html", error=error)


@app.route("/")
def index():
    default_folder = "" if CLOUD_MODE else (DEFAULT_FOLDER if Path(DEFAULT_FOLDER).exists() else "")
    return render_template("index.html", default_folder=default_folder, cloud_mode=CLOUD_MODE)


@app.post("/api/batch")
def api_batch_new():
    """Start an upload batch; the browser then posts files a few at a time."""
    batch_id = uuid.uuid4().hex
    (UPLOAD_DIR / batch_id).mkdir(parents=True, exist_ok=True)
    return jsonify({"ok": True, "batch_id": batch_id})


@app.post("/api/batch/<batch_id>")
def api_batch_add(batch_id: str):
    batch = UPLOAD_DIR / Path(batch_id).name
    if not batch.is_dir():
        return jsonify({"ok": False, "error": "Upload session expired. Start again."}), 400

    added = 0
    for item in request.files.getlist("files"):
        if not item.filename:
            continue
        name = _safe_upload_name(item.filename)
        if not name:
            continue
        item.save(batch / name)
        added += 1

    total = sum(1 for _ in batch.iterdir())
    return jsonify({"ok": True, "added": added, "total": total})


@app.post("/api/scan")
def api_scan():
    if CLOUD_MODE:
        return jsonify({"ok": False, "error": "Folder scanning is only available in the desktop version."}), 400
    folder = (request.form.get("folder") or "").strip().strip('"')
    if not folder:
        return jsonify({"ok": False, "error": "Enter a folder path."}), 400
    folder_path = Path(folder)
    if not folder_path.exists():
        return jsonify({"ok": False, "error": f"Folder not found: {folder}"}), 400
    files = _collect_source_files(folder_path)
    loggers = _dedupe_prefer_pdf(files)
    return jsonify(
        {
            "ok": True,
            "files": len(files),
            "loggers": len(loggers),
            "sample": [logger_column_name(p) for p in loggers[:6]],
            "last": logger_column_name(loggers[-1]) if loggers else None,
        }
    )


@app.post("/api/process")
def api_process():
    folder = (request.form.get("folder") or "").strip().strip('"')
    batch_id = (request.form.get("batch_id") or "").strip()
    saved: list[Path] = []
    batch_dir: Path | None = None

    if batch_id:
        batch_dir = UPLOAD_DIR / Path(batch_id).name
        if not batch_dir.is_dir():
            return jsonify({"ok": False, "error": "Upload session expired. Start again."}), 400
        saved.extend(p for p in batch_dir.iterdir() if p.is_file())

    for item in request.files.getlist("files"):
        if not item.filename:
            continue
        name = _safe_upload_name(item.filename)
        if not name:
            continue
        if batch_dir is None:
            batch_dir = UPLOAD_DIR / uuid.uuid4().hex
            batch_dir.mkdir(parents=True, exist_ok=True)
        dest = batch_dir / name
        item.save(dest)
        saved.append(dest)

    if folder:
        if CLOUD_MODE:
            return jsonify({"ok": False, "error": "This server has no access to your PC folders. Please upload the files."}), 400
        folder_path = Path(folder)
        if not folder_path.exists():
            return jsonify({"ok": False, "error": f"Folder not found: {folder}"}), 400
        saved.extend(_collect_source_files(folder_path))

    if not saved:
        message = "Choose your logger files first." if CLOUD_MODE else "Upload PDF files or enter a folder path."
        return jsonify({"ok": False, "error": message}), 400

    job_id = uuid.uuid4().hex
    _update_job(job_id, status="queued", current=0, total=len(saved), message="Queued")
    thread = threading.Thread(target=run_job, args=(job_id, saved, batch_dir), daemon=False)
    thread.start()
    return jsonify({"ok": True, "job_id": job_id})


@app.get("/api/status/<job_id>")
def api_status(job_id: str):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Unknown job"}), 404
    return jsonify({"ok": True, **job})


@app.get("/healthz")
def healthz():
    return "ok v7", 200


@app.get("/download/<name>")
def download(name: str):
    path = OUTPUT_DIR / Path(name).name
    if not path.exists():
        return "File not found", 404
    return send_file(path, as_attachment=True, download_name=path.name)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0" if CLOUD_MODE else "127.0.0.1"
    print(f"Open http://127.0.0.1:{port}")
    app.run(host=host, port=port, debug=False, threaded=True)
