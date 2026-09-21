"""Process an entire DL folder into Temperature / Humidity Excel."""

from __future__ import annotations

import sys
from pathlib import Path

from app import DEFAULT_FOLDER, JOBS, JOBS_LOCK, OUTPUT_DIR, _collect_source_files, run_job


def main():
    folder = Path(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FOLDER)
    if not folder.exists():
        raise SystemExit(f"Folder not found: {folder}")
    files = _collect_source_files(folder)
    print(f"Found {len(files)} logger files in {folder}")
    job_id = "cli"
    run_job(job_id, files)
    with JOBS_LOCK:
        job = JOBS.get(job_id, {})
    if job.get("status") != "done":
        print(job.get("error") or job)
        raise SystemExit(1)
    print("Wrote", OUTPUT_DIR / job["download"])
    print("Loggers", job["stats"]["loggers"], "rows", job["stats"]["rows"])
    print("Range", job["stats"]["start"], "->", job["stats"]["end"])
    if job.get("warnings"):
        print("Warnings", len(job["warnings"]))
        for w in job["warnings"][:20]:
            print(" ", w)


if __name__ == "__main__":
    main()
