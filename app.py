"""
Psikotes Score Automation — API (v5)
- Improved matching: initials, acronyms, bidirectional word match
- Date normalization: DD/MM/YYYY
- Uniform fields: TGL PEMERIKSAAN and PENDIDIKAN
- Yellow detail list per gugus
- No HTML UI — this is now an internal API consumed only by the Ordinat
  Dashboard (Next.js), which proxies every request with a shared bearer
  token. Never expose this service directly to the public internet.
"""
import os, io, json, uuid, tempfile, threading, re, hmac
from functools import wraps
from datetime import datetime
from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from difflib import SequenceMatcher
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
UPLOAD_FOLDER = tempfile.mkdtemp()
jobs = {}
# Each finished job retains a loaded workbook + its output file. Without a cap
# this grows unbounded on a long-running server, so we evict the oldest
# terminal jobs past this limit (see _evict_old_jobs).
MAX_JOBS = 30


def require_service_token(fn):
    """Gate a route behind Authorization: Bearer <RECAP_SERVICE_TOKEN>.
    Fails closed — an unset token never authorizes anything. This is the same
    shared secret the dashboard uses for its Flask -> Next.js webhook, reused
    here for the Next.js -> Flask direction."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        token = os.environ.get('RECAP_SERVICE_TOKEN')
        auth = request.headers.get('Authorization', '')
        if not token or not auth.startswith('Bearer ') or not hmac.compare_digest(auth[7:], token):
            return jsonify({'error': 'Unauthorized'}), 401
        return fn(*args, **kwargs)
    return wrapper

# Skor kecocokan nama di bawah ini butuh konfirmasi manual sebelum ditulis ke Rekap.
AUTO_CONFIDENCE = 0.90


# ── Name Matching ──

def normalize_name(name):
    if pd.isna(name): return ""
    n = str(name).strip().upper()
    for ch in ["\u2018","\u2019","'",'"']: n = n.replace(ch,"'")
    n = n.replace("."," ").replace("-"," ")
    while "  " in n: n = n.replace("  "," ")
    return n.strip()

def _ordered_word_match(short_w, long_w):
    """Match words in order. A short token (<=4 chars) that's an exact prefix
    of a long-side word counts as a partial (initial/abbreviation) match —
    covers both single-letter initials ("M" -> MUHAMMAD) and the common
    Indonesian 3-4 letter prefix abbreviations ("Moh"/"Muh"/"Moch" -> MUHAMMAD)."""
    li, matched = 0, 0
    for sw in short_w:
        for j in range(li, len(long_w)):
            lw = long_w[j]
            if sw == lw: matched += 1; li = j+1; break
            elif len(sw) <= 4 and lw.startswith(sw): matched += 0.8; li = j+1; break
    return matched

def calc_match(a, b):
    if not a or not b: return 0
    if a == b: return 1.0
    s = SequenceMatcher(None, a, b).ratio()
    # Substring check (min 4 chars to avoid short false positives)
    if len(a)>=4 and a in b: s = max(s, len(a)/len(b), 0.82)
    if len(b)>=4 and b in a: s = max(s, len(b)/len(a), 0.82)
    wa, wb = a.split(), b.split()
    if len(wa) >= 2 and len(wb) >= 2:
        # 1) Ordered word match with initials/prefix abbreviations
        #    (M->MUHAMMAD, MOH->MUHAMMAD, MUH->MUHAMMAD, P->PUTRA)
        for sw, lw in [(wa,wb),(wb,wa)]:
            m = _ordered_word_match(sw, lw)
            if len(sw) > 0:
                ratio = m / len(sw)
                # Scaled to word count, not a fixed m>=2 — a fixed floor was
                # unreachable for 2-word names where one word is a single
                # initial (e.g. "M. IRFAN" vs "MUKHAMAD IRFAN": max possible
                # m = 0.8 (initial) + 1.0 (exact) = 1.8, never >= 2), so this
                # very common Indonesian abbreviation pattern (Muhammad/
                # Mohammad/Muhamad -> "M.") was silently unmatchable no
                # matter how confident the rest of the name was.
                if ratio >= 0.6 and m >= len(sw) - 0.5: s = max(s, ratio * 0.92)
        # 2) Acronym detection — strict rule:
        #    A token (2-6 chars) is an acronym ONLY if each character is the
        #    FIRST LETTER of a CONSECUTIVE word on the other side.
        #    e.g. "SPS" = S(haista) P(utri) S(etiawan)
        #    Remaining non-acronym words must also match.
        for all_w, check_w in [(wa,wb),(wb,wa)]:
            for token in all_w:
                if len(token) < 2 or len(token) > 6: continue
                if token in check_w: continue  # already a real word, skip
                for start in range(len(check_w)):
                    needed = len(token)
                    if start + needed > len(check_w): break
                    candidate = check_w[start:start+needed]
                    initials = ''.join(w[0] for w in candidate)
                    if initials == token:
                        # Acronym match! Score based on remaining word matches
                        rest_tokens = [w for w in all_w if w != token]
                        rest_targets = [w for w in check_w if w not in candidate]
                        if not rest_tokens:
                            s = max(s, 0.85)
                        else:
                            rm = sum(1 for rt in rest_tokens if rt in rest_targets)
                            tr = (rm + 1) / (len(rest_tokens) + 1)
                            if tr > 0.5: s = max(s, tr * 0.90)
                        break
    return s


# ── Date Normalization ──

def normalize_date(val):
    """Convert any date value to DD/MM/YYYY string."""
    if pd.isna(val) or val is None: return None
    if isinstance(val, (datetime, pd.Timestamp)):
        y = val.year
        if y < 100: y += 2000
        return f"{val.day:02d}/{val.month:02d}/{y:04d}" if 1900 <= y <= 2100 else None
    s = str(val).strip()
    if not s: return None
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m:
        a, b, ys = int(m.group(1)), int(m.group(2)), m.group(3)
        y = int(ys)
        if y < 1900: y = 2000 + int(ys[-2:])  # "0209"->2009, "0008"->2008
        if a > 12 and b <= 12: a, b = b, a
        if 1<=a<=12 and 1<=b<=31: return f"{b:02d}/{a:02d}/{y:04d}"
        if 1<=b<=12 and 1<=a<=31: return f"{a:02d}/{b:02d}/{y:04d}"
        return s
    m2 = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{2,4})$', s)
    if m2:
        a, b, y = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if y < 100: y += 2000
        if a > 12 and b <= 12: a, b = b, a
        return f"{b:02d}/{a:02d}/{y:04d}" if 1<=a<=12 else s
    try:
        dt = pd.to_datetime(val)
        y = dt.year
        if y < 100: y += 2000
        return f"{dt.day:02d}/{dt.month:02d}/{y:04d}" if 1900<=y<=2100 else s
    except: return s


# ── Sheet Detection ──

def fast_detect_sheets(raw_path):
    wb = load_workbook(raw_path, read_only=True, data_only=True)
    subtes, other = [], []
    for sn in wb.sheetnames:
        ws = wb[sn]
        row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        hdrs = [str(h).strip() if h else "" for h in row1]
        if 'Score' in hdrs and 'NAMA LENGKAP' in hdrs and 'KELAS' in hdrs: subtes.append(sn)
        else: other.append(sn)
    wb.close()
    return subtes, other

def read_raw_minimal(raw_path, sheet_name):
    needed = ['Timestamp','Score','NAMA LENGKAP','KELAS',
              'JENIS KELAMIN','TEMPAT LAHIR','TANGGAL LAHIR','TANGGAL TES','PENDIDIKAN ']
    return pd.read_excel(raw_path, sheet_name=sheet_name, usecols=lambda c: c in needed)

def detect_kelas_fmt(vals):
    for v in vals:
        s = str(v)
        if s.startswith('KELAS X'): return 'KELAS X'
        if 'XI.' in s: return 'XI.'
    return 'X'

def extract_kelas_num(kstr, fmt):
    try:
        s = str(kstr).strip()
        if fmt == 'KELAS X': return int(s.replace('KELAS X','').strip())
        if fmt == 'XI.': return int(s.replace('XI.','').strip())
        return int(s.replace('X','').strip())
    except: return 0

def find_id_sheet(raw_path, subtes_sheets):
    wb = load_workbook(raw_path, read_only=True, data_only=True)
    res = None
    for sn in subtes_sheets:
        ws = wb[sn]
        row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        hdrs = [str(h).upper().strip() if h else "" for h in row1]
        if any('JENIS KELAMIN' in h for h in hdrs): res = sn; break
    wb.close()
    return res or ('SE' if 'SE' in subtes_sheets else subtes_sheets[0] if subtes_sheets else None)


# ── Background Job ──

def _remove_file(path):
    if path and os.path.exists(path):
        try: os.remove(path)
        except OSError: pass

def _free_job_memory(job):
    """Release a terminal job's heavy resources: the loaded openpyxl workbook +
    raw DataFrames (held in _ctx) and the input files on disk. The output file
    and the log survive — they're still needed for /download and /status."""
    job.pop('_ctx', None)  # drops the workbook + all_raw/all_proktor
    _remove_file(job.pop('_raw_path', None))
    _remove_file(job.pop('_rekap_path', None))

def _evict_old_jobs():
    """Cap memory/disk by dropping the oldest terminal (done/error) jobs, and
    their output files, once we exceed MAX_JOBS. Active and awaiting_review
    jobs are never evicted. `jobs` preserves insertion order, so the terminal
    list below is oldest-first."""
    if len(jobs) <= MAX_JOBS: return
    terminal = [jid for jid, j in jobs.items() if j.get('status') in ('done', 'error')]
    for jid in terminal:
        if len(jobs) <= MAX_JOBS: break
        j = jobs.pop(jid, None)
        if j:
            _remove_file(os.path.join(UPLOAD_FOLDER, j['download_filename']) if j.get('download_filename') else None)

def queue_or_write(ws, row_idx, col_idx, sv, best_s, gnum, subtes, gn_display, matched_name, source, log, pending):
    """Write score directly if high-confidence match; otherwise queue it for manual review.
    Returns True if the row is considered handled (found), either way."""
    if pd.isna(sv): return False
    # Scores come from Google Forms responses — a stray non-numeric cell must
    # skip just this row (it'll surface as missing/yellow), never crash the job.
    try:
        score_val = int(float(sv))
    except (ValueError, TypeError):
        return False
    if best_s >= AUTO_CONFIDENCE:
        ws.cell(row=row_idx, column=col_idx, value=score_val)
        log['total_scores'] += 1
        if source != 'own': log['cross_kelas'] += 1
    else:
        pending.append({
            'id': len(pending), 'gugus': gnum, 'subtes': subtes,
            'nama_rekap': gn_display, 'nama_raw': str(matched_name),
            'score': round(best_s, 3), 'source': source,
            '_row_idx': row_idx, '_col_idx': col_idx, '_value': score_val,
        })
    return True

def process_job(job_id, raw_path, rekap_path, overrides, threshold, tgl_pemeriksaan, pendidikan):
    job = jobs[job_id]
    try:
        job['status'] = 'Mendeteksi sheet...'; job['progress'] = 3
        norm_ov = {normalize_name(k): (normalize_name(v) if v else None) for k,v in overrides.items()}
        subtes_sheets, _ = fast_detect_sheets(raw_path)
        if not subtes_sheets:
            job['status'] = 'error'; job['error'] = 'Tidak ada sheet subtes di RAW'; return

        job['progress'] = 8
        all_raw = {}; all_proktor = {}; kelas_fmt = None
        for i, sn in enumerate(subtes_sheets):
            job['status'] = f'Membaca {sn}...'
            raw = read_raw_minimal(raw_path, sn)
            raw['_norm'] = raw['NAMA LENGKAP'].apply(normalize_name)
            # Split: normal kelas vs Proktor entries
            is_proktor = raw['KELAS'].astype(str).str.contains('Proktor', na=False)
            proktor_df = raw[is_proktor].drop_duplicates(subset=['_norm'], keep='first')
            normal_df = raw[~is_proktor].drop_duplicates(subset=['_norm','KELAS'], keep='first')
            all_raw[sn] = normal_df
            all_proktor[sn] = proktor_df
            if kelas_fmt is None:
                kv = [v for v in normal_df['KELAS'].dropna().unique()]
                kelas_fmt = detect_kelas_fmt(kv)
            job['progress'] = 8 + int(22*(i+1)/len(subtes_sheets))

        job['status'] = 'Membaca Rekap...'; job['progress'] = 32
        wb = load_workbook(rekap_path)
        gsheets = sorted([s for s in wb.sheetnames if s.lower().startswith('gugus')],
                          key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))
        if not gsheets:
            job['status'] = 'error'; job['error'] = 'Tidak ada sheet Gugus'; return

        ws0 = wb[gsheets[0]]
        headers = {}
        for col in range(1, ws0.max_column+1):
            v = ws0.cell(row=1, column=col).value
            if v: headers[str(v).strip()] = col
        name_col = headers.get('NAMA', 2)
        rekap_subtes = [s for s in subtes_sheets if s in headers]

        gugus_rows, gugus_all_names, gugus_nums = {}, {}, []
        for sn in gsheets:
            try: gnum = int(''.join(filter(str.isdigit, sn)))
            except: continue
            gugus_nums.append(gnum)
            ws = wb[sn]
            gugus_rows[gnum] = {}
            for row in range(2, ws.max_row+1):
                cv = ws.cell(row=row, column=name_col).value
                if cv: gugus_rows[gnum][row] = normalize_name(cv)
            gugus_all_names[gnum] = list(gugus_rows[gnum].values())

        job['progress'] = 38; job['status'] = 'Mencocokkan nama...'
        log = {'total_scores':0,'cross_kelas':0,'identity_filled':0,'yellow_count':0,
               'per_gugus':{},'unmatched_summary':[],'yellow_detail':[],
               'subtes_detected':subtes_sheets,'rekap_subtes':rekap_subtes}

        total_steps = len(rekap_subtes) * len(gugus_nums); step = 0
        pending = []
        for subtes in rekap_subtes:
            raw = all_raw[subtes]; proktor = all_proktor[subtes]; col_idx = headers[subtes]
            for gnum in gugus_nums:
                if gnum not in gugus_rows: step+=1; continue
                kstr = f'{kelas_fmt} {gnum}' if kelas_fmt != 'XI.' else f'XI.{gnum}'
                ws = wb[f'Gugus {gnum}']
                raw_own = raw[raw['KELAS']==kstr]
                for row_idx, gn in gugus_rows[gnum].items():
                    found = False; mt = norm_ov.get(gn)
                    # Step 1: own kelas
                    best_s, best_i = 0, None
                    for ri, r in raw_own.iterrows():
                        s = calc_match(gn, r['_norm'])
                        if mt and calc_match(mt, r['_norm'])>0.9: s=0.95
                        if s>best_s: best_s,best_i=s,ri
                    if best_s>=threshold and best_i is not None:
                        sv = raw.loc[best_i,'Score']
                        nm = raw.loc[best_i,'NAMA LENGKAP']
                        found = queue_or_write(ws, row_idx, col_idx, sv, best_s, gnum, subtes, gn, nm, 'own', log, pending)
                    # Step 2: cross-kelas
                    if not found:
                        raw_other = raw[raw['KELAS']!=kstr]; best_s,best_i=0,None
                        for ri, r in raw_other.iterrows():
                            s = calc_match(gn, r['_norm'])
                            if mt and calc_match(mt, r['_norm'])>0.9: s=0.95
                            if s>best_s: best_s,best_i=s,ri
                        if best_s>=threshold and best_i is not None:
                            rr = raw.loc[best_i]; rg = extract_kelas_num(rr['KELAS'], kelas_fmt)
                            claimed = rg in gugus_all_names and any(calc_match(og, rr['_norm'])>=threshold for og in gugus_all_names[rg])
                            if not claimed:
                                sv = rr['Score']
                                found = queue_or_write(ws, row_idx, col_idx, sv, best_s, gnum, subtes, gn, rr['NAMA LENGKAP'], 'cross', log, pending)
                    # Step 3: search Proktor entries (students who picked wrong kelas option)
                    if not found and len(proktor) > 0:
                        best_s, best_i = 0, None
                        for ri, r in proktor.iterrows():
                            s = calc_match(gn, r['_norm'])
                            if mt and calc_match(mt, r['_norm'])>0.9: s=0.95
                            if s>best_s: best_s,best_i=s,ri
                        if best_s>=threshold and best_i is not None:
                            sv = proktor.loc[best_i,'Score']
                            nm = proktor.loc[best_i,'NAMA LENGKAP']
                            found = queue_or_write(ws, row_idx, col_idx, sv, best_s, gnum, subtes, gn, nm, 'proktor', log, pending)
                step += 1
                job['progress'] = min(38 + int(40*step/max(total_steps,1)), 78)
            job['status'] = f'Mencocokkan {subtes}...'

        ctx = {'wb':wb,'headers':headers,'name_col':name_col,'gugus_rows':gugus_rows,
               'gugus_nums':gugus_nums,'rekap_subtes':rekap_subtes,'raw_path':raw_path,
               'subtes_sheets':subtes_sheets,'all_raw':all_raw,'all_proktor':all_proktor,
               'kelas_fmt':kelas_fmt,'tgl_pemeriksaan':tgl_pemeriksaan,'pendidikan':pendidikan,
               'log':log,'threshold':threshold,'norm_ov':norm_ov}
        if pending:
            job['_ctx'] = ctx; job['pending'] = pending
            job['status'] = 'awaiting_review'; job['progress'] = 78
            return
        job['_ctx'] = ctx
        finalize_job(job_id)
    except Exception as e:
        job['status'] = 'error'; job['error'] = f'{type(e).__name__}: {str(e)}'
        _free_job_memory(job)


def finalize_job(job_id):
    job = jobs[job_id]
    ctx = job['_ctx']
    wb = ctx['wb']; headers = ctx['headers']; name_col = ctx['name_col']
    gugus_rows = ctx['gugus_rows']; gugus_nums = ctx['gugus_nums']; rekap_subtes = ctx['rekap_subtes']
    raw_path = ctx['raw_path']; subtes_sheets = ctx['subtes_sheets']
    all_raw = ctx['all_raw']; all_proktor = ctx['all_proktor']
    kelas_fmt = ctx['kelas_fmt']; tgl_pemeriksaan = ctx['tgl_pemeriksaan']; pendidikan = ctx['pendidikan']
    log = ctx['log']; threshold = ctx['threshold']; norm_ov = ctx['norm_ov']
    try:
        # ── Identity ──
        job['status'] = 'Mengisi identitas...'; job['progress'] = 80
        id_sheet = find_id_sheet(raw_path, subtes_sheets)
        if id_sheet and id_sheet in all_raw:
            # Combine normal + proktor entries for identity search
            rid = pd.concat([all_raw[id_sheet], all_proktor.get(id_sheet, pd.DataFrame())], ignore_index=True)
            rid = rid.drop_duplicates(subset=['_norm'], keep='first')
            id_map = {}
            for c in rid.columns:
                u = str(c).upper().strip()
                if 'JENIS KELAMIN' in u and 'JK' in headers: id_map[c]='JK'
                elif 'TEMPAT LAHIR' in u and 'TEMPAT LAHIR' in headers: id_map[c]='TEMPAT LAHIR'
                elif 'TANGGAL LAHIR' in u and 'TGL LAHIR' in headers: id_map[c]='TGL LAHIR'
            for gnum in gugus_nums:
                if gnum not in gugus_rows: continue
                ws = wb[f'Gugus {gnum}']
                for row_idx, gn in gugus_rows[gnum].items():
                    mt = norm_ov.get(gn); best_s,best_i=0,None
                    for ri, r in rid.iterrows():
                        s = calc_match(gn, r['_norm'])
                        if mt and calc_match(mt, r['_norm'])>0.9: s=0.95
                        if s>best_s: best_s,best_i=s,ri
                    if best_s<threshold or best_i is None: continue
                    r = rid.loc[best_i]
                    for rc, rkc in id_map.items():
                        if rc in r.index and rkc in headers:
                            v = r[rc]
                            if pd.notna(v):
                                cur = ws.cell(row=row_idx, column=headers[rkc]).value
                                if cur is None or str(cur).strip()=='':
                                    # Normalize dates
                                    if rkc == 'TGL LAHIR':
                                        v = normalize_date(v) or v
                                    ws.cell(row=row_idx, column=headers[rkc], value=v)
                                    log['identity_filled'] += 1

        # ── Fix existing dates + fill uniform fields ──
        job['status'] = 'Memperbaiki format...'; job['progress'] = 88
        tgl_lahir_col = headers.get('TGL LAHIR')
        tgl_periksa_col = headers.get('TGL PEMERIKSAAN')
        pendidikan_col = headers.get('PENDIDIKAN')

        for gnum in gugus_nums:
            ws = wb[f'Gugus {gnum}']
            for row in range(2, ws.max_row+1):
                nm = ws.cell(row=row, column=name_col).value
                if not nm: continue
                # Fix TGL LAHIR format
                if tgl_lahir_col:
                    v = ws.cell(row=row, column=tgl_lahir_col).value
                    if v is not None:
                        fixed = normalize_date(v)
                        if fixed: ws.cell(row=row, column=tgl_lahir_col, value=fixed)
                # Fill TGL PEMERIKSAAN (uniform)
                if tgl_periksa_col and tgl_pemeriksaan:
                    ws.cell(row=row, column=tgl_periksa_col, value=tgl_pemeriksaan)
                # Fill PENDIDIKAN (uniform)
                if pendidikan_col and pendidikan:
                    ws.cell(row=row, column=pendidikan_col, value=pendidikan)

        # ── Yellow highlight + stats ──
        job['status'] = 'Finalisasi...'; job['progress'] = 93
        subtes_cols = [headers[s] for s in rekap_subtes if s in headers]
        yf = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        for gnum in gugus_nums:
            ws = wb[f'Gugus {gnum}']; ts,fc,yc=0,0,0
            yellow_names = []
            for row in range(2, ws.max_row+1):
                nm = ws.cell(row=row, column=name_col).value
                if not nm: continue
                ts += 1
                rf = sum(1 for c in subtes_cols if ws.cell(row=row,column=c).value is not None)
                fc += rf
                if rf == 0:
                    for col in range(1, ws.max_column+1): ws.cell(row=row, column=col).fill = yf
                    yc += 1; log['yellow_count'] += 1
                    yellow_names.append(str(nm))
                elif rf < len(subtes_cols):
                    gn_name = normalize_name(nm)
                    missing = [rekap_subtes[j] for j,c in enumerate(subtes_cols) if ws.cell(row=row,column=c).value is None]
                    log['unmatched_summary'].append({'gugus':gnum,'nama':gn_name,'missing':missing,'all_missing':False})
            tc = ts * len(subtes_cols)
            log['per_gugus'][f'Gugus {gnum}'] = {'siswa':ts,'filled':fc,'total':tc,
                'pct':fc*100//tc if tc else 0,'yellow':yc}
            if yellow_names:
                log['yellow_detail'].append({'gugus':gnum, 'names':yellow_names})

        # job_id in the name so two jobs finishing in the same second don't collide.
        out_name = f"Rekap_Otomatis_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{job_id}.xlsx"
        out_path = os.path.join(UPLOAD_FOLDER, out_name)
        wb.save(out_path)
        job['progress'] = 100; job['status'] = 'done'; job['log'] = log; job['download_filename'] = out_name
        _free_job_memory(job)
    except Exception as e:
        job['status'] = 'error'; job['error'] = f'{type(e).__name__}: {str(e)}'
        _free_job_memory(job)


# ── Routes ──
# No HTML UI is served here anymore — the Ordinat Dashboard is the only
# frontend. Every route below requires the shared service token.

@app.route('/process', methods=['POST'])
@require_service_token
def start_process():
    if 'raw_file' not in request.files or 'rekap_file' not in request.files:
        return jsonify({'error':'Upload kedua file'}), 400
    rf, rkf = request.files['raw_file'], request.files['rekap_file']
    # Prefix stored files with the job id so concurrent uploads that happen to
    # share a filename (e.g. two admins both uploading "RAW.xlsx") never
    # overwrite each other on disk while a background thread is still reading.
    jid = str(uuid.uuid4())[:8]
    rp = os.path.join(UPLOAD_FOLDER, f"{jid}_raw_{secure_filename(rf.filename)}")
    rkp = os.path.join(UPLOAD_FOLDER, f"{jid}_rekap_{secure_filename(rkf.filename)}")
    rf.save(rp); rkf.save(rkp)
    ov = {}
    try: ov = json.loads(request.form.get('manual_overrides','{}'))
    except: pass
    th = float(request.form.get('threshold', 0.78))
    tgl_p = request.form.get('tgl_pemeriksaan', '').strip() or None
    pend = request.form.get('pendidikan', '').strip() or None
    jobs[jid] = {'status':'queued','progress':0,'log':None,'download_filename':None,
                 'error':None,'_raw_path':rp,'_rekap_path':rkp}
    _evict_old_jobs()
    threading.Thread(target=process_job, args=(jid, rp, rkp, ov, th, tgl_p, pend), daemon=True).start()
    return jsonify({'job_id': jid})

@app.route('/status/<job_id>')
@require_service_token
def job_status(job_id):
    j = jobs.get(job_id)
    if not j: return jsonify({'error':'Job not found'}), 404
    resp = {'status':j['status'],'progress':j['progress'],'log':j.get('log'),
            'download_filename':j.get('download_filename'),'error':j.get('error')}
    if j.get('status') == 'awaiting_review':
        resp['pending'] = [{k:v for k,v in p.items() if not k.startswith('_')} for p in j.get('pending', [])]
    return jsonify(resp)

@app.route('/review/<job_id>', methods=['POST'])
@require_service_token
def submit_review(job_id):
    j = jobs.get(job_id)
    if not j: return jsonify({'error':'Job not found'}), 404
    if j.get('status') != 'awaiting_review':
        return jsonify({'error':'Job tidak sedang menunggu review'}), 400
    data = request.get_json(silent=True) or {}
    decisions = data.get('decisions', {})
    ctx = j['_ctx']; wb = ctx['wb']; log = ctx['log']
    for item in j.get('pending', []):
        if decisions.get(str(item['id'])):
            ws = wb[f"Gugus {item['gugus']}"]
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
