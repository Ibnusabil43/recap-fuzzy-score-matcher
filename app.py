"""Psikotes Score Automation — API entrypoint.

Thin Flask layer over the `recap` package: it only wires HTTP routes to the
pipeline in `recap.engine` and the job registry in `recap.jobstore`. All matching,
scoring, and workbook logic lives in the package.

No HTML UI is served here (except an opt-in local dev console) — the Ordinat
Dashboard (Next.js) is the only frontend, and it proxies every request with a
shared bearer token. Never expose this service directly to the public internet.
"""
from __future__ import annotations

import json
import os
import threading
import uuid

from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from recap.auth import require_service_token
from recap.config import MAX_CONTENT_LENGTH, TEST_CONSOLE_PATH
from recap.engine import finalize_job, process_job
from recap.jobstore import UPLOAD_FOLDER, evict_old_jobs, jobs

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH


# ── Routes ──
# Every route below (except /healthz and the opt-in dev console) requires the
# shared service token.

@app.route('/healthz')
def healthz():
    """Unauthenticated on purpose — platform health checks (Render etc.) hit
    this before any token exists. Reveals nothing beyond "process is alive"."""
    return jsonify({'status': 'ok'}), 200


@app.route('/console')
def dev_console():
    """Serve the local test console (test_console.html) same-origin so its API
    calls avoid CORS. Off by default to preserve the "no public UI" posture —
    opt in with ENABLE_TEST_CONSOLE=1 for local dev only. The page itself still
    carries the bearer token on every API call; this route only serves static
    HTML and exposes nothing."""
    if os.environ.get('ENABLE_TEST_CONSOLE', '').lower() not in ('1', 'true', 'yes'):
        return "Not found", 404
    if not TEST_CONSOLE_PATH.exists():
        return "console not found", 404
    return send_file(str(TEST_CONSOLE_PATH))


@app.route('/process', methods=['POST'])
@require_service_token
def start_process():
    if 'raw_file' not in request.files or 'rekap_file' not in request.files:
        return jsonify({'error': 'Upload kedua file'}), 400
    rf, rkf = request.files['raw_file'], request.files['rekap_file']
    # Prefix stored files with the job id so concurrent uploads that happen to
    # share a filename (e.g. two admins both uploading "RAW.xlsx") never
    # overwrite each other on disk while a background thread is still reading.
    jid = str(uuid.uuid4())[:8]
    rp = os.path.join(UPLOAD_FOLDER, f"{jid}_raw_{secure_filename(rf.filename)}")
    rkp = os.path.join(UPLOAD_FOLDER, f"{jid}_rekap_{secure_filename(rkf.filename)}")
    rf.save(rp); rkf.save(rkp)
    ov = {}
    try:
        ov = json.loads(request.form.get('manual_overrides', '{}'))
    except (ValueError, TypeError):
        pass
    th = float(request.form.get('threshold', 0.78))
    tgl_p = request.form.get('tgl_pemeriksaan', '').strip() or None
    pend = request.form.get('pendidikan', '').strip() or None
    jobs[jid] = {'status': 'queued', 'progress': 0, 'log': None, 'download_filename': None,
                 'error': None, '_raw_path': rp, '_rekap_path': rkp}
    evict_old_jobs()
    threading.Thread(target=process_job, args=(jid, rp, rkp, ov, th, tgl_p, pend),
                     daemon=True).start()
    return jsonify({'job_id': jid})


@app.route('/status/<job_id>')
@require_service_token
def job_status(job_id):
    j = jobs.get(job_id)
    if not j: return jsonify({'error': 'Job not found'}), 404
    resp = {'status': j['status'], 'progress': j['progress'], 'log': j.get('log'),
            'download_filename': j.get('download_filename'), 'error': j.get('error')}
    if j.get('status') == 'awaiting_review':
        resp['pending'] = [{k: v for k, v in p.items() if not k.startswith('_')}
                           for p in j.get('pending', [])]
    return jsonify(resp)


@app.route('/review/<job_id>', methods=['POST'])
@require_service_token
def submit_review(job_id):
    j = jobs.get(job_id)
    if not j: return jsonify({'error': 'Job not found'}), 404
    if j.get('status') != 'awaiting_review':
        return jsonify({'error': 'Job tidak sedang menunggu review'}), 400
    data = request.get_json(silent=True) or {}
    decisions = data.get('decisions', {})
    ctx = j['_ctx']; wb = ctx['wb']; log = ctx['log']
    headers = ctx.get('headers', {}); answer_sections = ctx.get('answer_sections') or {}
    ge_qcols = ctx.get('ge_qcols', []); raw_path = ctx.get('raw_path')
    from recap.engine import _apply_leftover_confirm
    for item in j.get('pending', []):
        dec = decisions.get(str(item['id']))
        if not dec: continue
        kind = item.get('kind')
        if kind == 'leftover':
            # Unlike value/answer_choice (boolean per item), a leftover item's
            # decision is the LIST of subtest codes the admin checked in the
            # review UI's per-subtest dropdown — not confirmed wholesale, since
            # a candidate name plausible enough to suggest isn't automatically
            # plausible enough to blindly apply to every one of a student's
            # subtests. Only 'confirmable' items (see process_job) have a
            # resolved candidate row to write into at all.
            if not item.get('confirmable') or not isinstance(dec, list) or not dec:
                continue
            if _apply_leftover_confirm(wb, raw_path, item, dec, headers, answer_sections, ge_qcols):
                log['total_scores'] += 1
                log['cross_kelas'] += 1  # always lands on a different roster row than the one RAW claimed
            continue
        ws = wb[f"Gugus {item['gugus']}"]
        if kind == 'answer_choice':
            for col, val in zip(item['_targets'], item['_answers']):
                if col and val is not None:
                    ws.cell(row=item['_row_idx'], column=col, value=val)
        else:
            ws.cell(row=item['_row_idx'], column=item['_col_idx'], value=item['_value'])
        log['total_scores'] += 1
        if item['source'] != 'own': log['cross_kelas'] += 1
    j['pending'] = []
    j['status'] = 'Melanjutkan proses...'; j['progress'] = 80
    threading.Thread(target=finalize_job, args=(job_id,), daemon=True).start()
    return jsonify({'ok': True})


@app.route('/download/<filename>')
@require_service_token
def download(filename):
    p = os.path.join(UPLOAD_FOLDER, secure_filename(filename))
    if not os.path.exists(p): return "File not found", 404
    return send_file(p, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    # debug=True exposes Werkzeug's interactive debugger (an RCE surface on any
    # reachable instance) — never on by default. Opt in explicitly for local dev.
    debug = os.environ.get('FLASK_DEBUG', '').lower() in ('1', 'true', 'yes')
    app.run(debug=debug, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
