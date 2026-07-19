"""In-memory job registry and temp-file lifecycle.

Jobs are held in a process-global dict keyed by job id. Each finished job keeps
its output workbook on disk plus its log in memory; heavy resources (the loaded
openpyxl workbook and RAW DataFrames) are released as soon as the job reaches a
terminal state, and the oldest terminal jobs are evicted past MAX_JOBS.
"""
from __future__ import annotations

import os
import tempfile
from typing import Any

from .config import MAX_JOBS

# job_id -> job dict. A job dict carries public keys (status, progress, log,
# download_filename, error) and private ones (_ctx, _raw_path, _rekap_path,
# pending) used only server-side.
jobs: dict[str, dict[str, Any]] = {}

# Per-process temp dir for uploaded inputs and generated outputs.
UPLOAD_FOLDER: str = tempfile.mkdtemp()


def remove_file(path: str | None) -> None:
    """Delete a file if it exists, ignoring OS errors."""
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def free_job_memory(job: dict[str, Any]) -> None:
    """Release a terminal job's heavy resources: the loaded openpyxl workbook +
    raw DataFrames (held in _ctx) and the input files on disk. The output file
    and the log survive — they're still needed for /download and /status."""
    job.pop('_ctx', None)  # drops the workbook + all_raw/all_proktor
    remove_file(job.pop('_raw_path', None))
    remove_file(job.pop('_rekap_path', None))


def evict_old_jobs() -> None:
    """Cap memory/disk by dropping the oldest terminal (done/error) jobs, and
    their output files, once we exceed MAX_JOBS. Active and awaiting_review jobs
    are never evicted. `jobs` preserves insertion order, so the terminal list
    below is oldest-first."""
    if len(jobs) <= MAX_JOBS: return
    terminal = [jid for jid, j in jobs.items() if j.get('status') in ('done', 'error')]
    for jid in terminal:
        if len(jobs) <= MAX_JOBS: break
        j = jobs.pop(jid, None)
        if j:
            remove_file(os.path.join(UPLOAD_FOLDER, j['download_filename'])
                        if j.get('download_filename') else None)
