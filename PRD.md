# PRD — Psikotes Score Automation (Recap Fuzzy Score Matcher)

> Reverse-engineered from the current codebase (`app.py` v5) as of 2026-07-12.
> This service is a backend-only API. It has no UI of its own — it is consumed
> exclusively by the Ordinat Dashboard (Next.js), which proxies every request
> with a shared bearer token.

## 1. Problem

Psychotest ("psikotes") results are collected per subtest via Google Forms
into a multi-sheet RAW Excel export. Staff must then hand-copy each student's
score from RAW into a master REKAP workbook (one sheet per "Gugus" / cohort
group), matching students by name across sheets that were entered inconsistently
(typos, initials, nicknames, wrong class selected). This is slow and error-prone
at the volumes involved (hundreds of students × multiple subtests).

## 2. Goal

Given a RAW workbook and a REKAP template workbook, automatically:
- Match each REKAP student row to their RAW score entries across all subtests, tolerating name variation.
- Fill in scores, with low-confidence matches routed to a human for approval instead of silently guessing.
- Backfill missing identity fields (sex, birthplace, birth date) from RAW.
- Normalize inconsistent date formats.
- Stamp uniform fields (exam date, education level) across all rows.
- Flag rows with zero matched scores so staff can spot them immediately (yellow highlight).
- Return a downloadable, ready-to-use REKAP workbook plus a completion report.

## 3. Non-goals

- No end-user UI — the Ordinat Dashboard owns all UX; this service is API-only.
- No persistent storage/database — jobs and files live in-memory/tmp-disk for the life of the process only.
- No multi-tenant auth — a single shared bearer token gates all non-health routes.
- No public exposure — this service is designed to sit behind a private network boundary, reachable only by the dashboard's server.

## 4. Users

- **Psikotes admin staff** (via Ordinat Dashboard UI): upload files, review borderline matches, download the finished recap.
- **Ordinat Dashboard backend** (the only direct API consumer): holds the shared `RECAP_SERVICE_TOKEN`, proxies staff actions to this service.

## 5. Current architecture (as built)

- **Stack**: Flask + pandas + openpyxl, single-process, in-memory job store (`jobs` dict), background work via Python `threading`. Deployed as a Docker image (Gunicorn, 2 workers/4 threads) to Render, gated by `/healthz`.
- **Auth**: every route except `/healthz` requires `Authorization: Bearer <RECAP_SERVICE_TOKEN>`, checked with constant-time comparison, fails closed if unset.
- **Job lifecycle**: `queued` → matching → (`awaiting_review` if any low-confidence matches) → finalizing → `done` | `error`. Terminal jobs beyond `MAX_JOBS = 30` are evicted oldest-first, freeing workbook memory and temp files.
- **Matching engine** (`calc_match`): combines `difflib.SequenceMatcher` ratio, substring containment, ordered-word/initials matching (e.g. "M." → "MUHAMMAD"), and strict consecutive-initial acronym detection (e.g. "SPS" → "Shaista Putri Setiawan"). Matches run in three passes per REKAP row: own-kelas RAW rows → cross-kelas RAW rows (guarding against a name already claimed by its rightful class) → "Proktor" catch-all entries (students who picked the wrong class option in the form).
- **Confidence routing**: matches ≥ `AUTO_CONFIDENCE` (0.90) auto-write; matches between the caller-supplied `threshold` (default 0.78) and 0.90 queue for manual review via `/review`.

## 6. Shipped features (checklist)

### Core matching & scoring
- [x] Fuzzy name normalization (case, punctuation, whitespace, apostrophe variants)
- [x] Sequence-ratio + substring similarity scoring
- [x] Ordered-word / initials matching ("M IRFAN" ↔ "MUKHAMAD IRFAN")
- [x] Strict acronym detection from consecutive-word initials
- [x] Manual name-override map (`manual_overrides`) to force a specific match
- [x] Configurable match threshold per job
- [x] Three-tier search: own-kelas → cross-kelas (with claim-guard) → Proktor fallback
- [x] Auto-write above `AUTO_CONFIDENCE` (0.90); queue borderline matches for human review

### Data normalization & enrichment
- [x] Date normalization to `DD/MM/YYYY` across multiple input formats (slash, dash, Excel datetime, 2-digit year)
- [x] Identity backfill: JENIS KELAMIN, TEMPAT LAHIR, TANGGAL LAHIR pulled from RAW into REKAP, never overwriting existing values
- [x] Uniform field stamping: TGL PEMERIKSAAN and PENDIDIKAN applied to every filled row
- [x] Kelas-format auto-detection (`KELAS X`, `XI.`, bare `X`) to match RAW class labels against Gugus numbers

### Reporting & QA
- [x] Per-Gugus completion stats (siswa count, filled cells, % complete, yellow count)
- [x] Yellow highlight for rows with zero matched scores across all subtests
- [x] Unmatched-subtest summary (rows with partial but not full coverage)
- [x] Yellow-detail listing (names per Gugus needing manual attention)

### Job orchestration API
- [x] `POST /process` — upload RAW + REKAP, kick off async job, returns `job_id`
- [x] `GET /status/<job_id>` — poll progress %, status text, log, pending review items, error
- [x] `POST /review/<job_id>` — submit accept/reject decisions for borderline matches, resumes finalization
- [x] `GET /download/<filename>` — fetch the finished `.xlsx`
- [x] `GET /healthz` — unauthenticated liveness check for platform health probes

### Operational hardening
- [x] Shared-secret bearer auth, fail-closed, constant-time compare
- [x] 50 MB request body cap
- [x] Job eviction (`MAX_JOBS = 30`) with memory + temp-file cleanup for terminal jobs
- [x] Per-job temp file isolation (job-id-prefixed filenames) to avoid collisions on concurrent uploads
- [x] Exception isolation per job (`status: error`, message captured, resources freed) — a bad job can't crash the process
- [x] `debug` mode opt-in only via `FLASK_DEBUG` env var (RCE surface closed by default)
- [x] Dockerized (Gunicorn, timeout 300s, 2 workers × 4 threads)
- [x] Render Blueprint (`render.yaml`) with `/healthz` health check path

## 7. Known gaps / risks (not yet addressed)

- [ ] **No persistence** — all job state and files live in process memory / local tmp disk. A restart or redeploy (e.g. Render free-tier sleep/wake, or any crash) loses every in-flight and completed-but-undownloaded job.
- [ ] **No automated tests** — matching logic, date normalization, and job lifecycle have zero test coverage; regressions rely on manual QA.
- [ ] **Single shared token, no per-admin identity** — the API cannot distinguish which staff member triggered a job or reviewed a match; all audit trail lives only in the dashboard, not here.
- [ ] **No job cancellation endpoint** — a stuck/long-running job can only be left to finish, error out, or age out via eviction.
- [ ] **In-memory job store is not horizontally scalable** — running >1 instance would split jobs across processes with no shared state; free-tier/Gunicorn single-instance only.
- [ ] **No rate limiting** — a single valid token can issue unbounded `/process` calls, bounded only by `MAX_JOBS` eviction and the 50 MB body cap.
- [ ] **No structured logging/observability** — job status strings are user-facing Indonesian text, not machine-parseable logs; no error tracking (e.g. Sentry) integration.
- [ ] **No malicious-file hardening documented** — `.xlsx` parsing (openpyxl/pandas) trusts uploaded files; no explicit size/row-count guard beyond the 50 MB total body cap.

## 8. Roadmap

### Phase 0 — Foundation (shipped)
Everything in §6. The service is functional and deployed.

### Phase 1 — Reliability hardening
- [ ] Add automated test suite for `calc_match`, `normalize_date`, `normalize_name` (pure functions, easy to unit test first)
- [ ] Add integration test for the full `/process` → `/review` → `/download` job lifecycle using fixture workbooks
- [ ] Persist job state + output files to durable storage (e.g. S3-compatible bucket or a small DB) so a redeploy doesn't drop in-flight/completed jobs
- [ ] Add `/cancel/<job_id>` to abort a running job early

### Phase 2 — Observability & ops
- [ ] Structured JSON logging (job id, phase, duration) separate from the human-facing status string
- [ ] Error tracking integration (e.g. Sentry) for uncaught exceptions in background threads
- [ ] Basic request rate limiting per token
- [ ] Metrics/alerting on job failure rate and average processing time

### Phase 3 — Matching quality
- [ ] Track match-acceptance/rejection outcomes from `/review` to tune `AUTO_CONFIDENCE`/`threshold` defaults over time
- [ ] Support additional identity fields beyond JK/TEMPAT LAHIR/TGL LAHIR if new subtest sheets require them
- [ ] Configurable acronym/initial matching rules per school (current rules are tuned for Indonesian naming conventions specifically)

### Phase 4 — Scale
- [ ] Move job queue off in-process `threading` to a real task queue (e.g. RQ/Celery + Redis) to support multi-instance deployment
- [ ] Multi-instance deploy behind a load balancer once job store is externalized (depends on Phase 1 persistence)

## 9. User stories

**As a psikotes admin**, I want to upload the RAW export and REKAP template and get a finished, filled-in REKAP file without manually cross-referencing hundreds of names, so that recap production takes minutes instead of hours.
- [x] Supported via `POST /process` + `GET /status` + `GET /download`.

**As a psikotes admin**, I want the system to handle nicknames, initials, and typos in student names, so that I don't have to manually fix every RAW entry before matching.
- [x] Supported via `calc_match`'s multi-strategy scoring (sequence ratio, substring, ordered-word/initials, acronym).

**As a psikotes admin**, I want to review and approve/reject only the uncertain matches, not every match, so that I retain control without re-doing the computer's work.
- [x] Supported via the `awaiting_review` status and `POST /review`.

**As a psikotes admin**, I want rows with completely missing scores to stand out visually, so that I can immediately see which students still need manual data entry.
- [x] Supported via yellow `PatternFill` highlighting and `yellow_detail` in the log.

**As a psikotes admin**, I want a specific name mapping I already know about to be forced, so that ambiguous or unusual name pairs I've verified myself aren't re-litigated by the fuzzy matcher.
- [x] Supported via `manual_overrides`.

**As the Ordinat Dashboard**, I want a single shared-secret auth mechanism, so that only my server can drive this internal service.
- [x] Supported via `require_service_token`.

**As a platform operator**, I want an unauthenticated liveness endpoint, so that Render's health checks don't need credentials to keep the service alive.
- [x] Supported via `/healthz`.

**As a psikotes admin**, I want my in-progress job to survive a server restart, so that a Render free-tier sleep/wake cycle doesn't force me to re-upload and re-review everything.
- [ ] Not yet supported — see Phase 1 (no persistence).

**As a developer maintaining this service**, I want confidence that a change to the matching algorithm doesn't silently break existing behavior, so that I can iterate safely.
- [ ] Not yet supported — see Phase 1 (no automated tests).

---

# Version 2.0 — Planning (new subtest types)

> ⚠️ **Superseded by Version 2.1** (below). Added 2026-07-12 as a pre-fixture
> draft; the checkboxes here are **not** the live task list — use §2.1.5.
> Fixtures (`… (RAW).xlsx`, `Template Rekap.xlsx`, `RUMUS GE.xlsx`) corrected
> several assumptions this draft made, so it is kept only as a pointer to avoid
> duplicate/contradictory tasks. What changed:
>
> - **"Holland"** → the real RAW sheet is **RIASEC** (a Holland-code test) and
>   is **answer-type** (`Ya`/`Tidak` per question), not value-transfer.
> - **EPPS** has **no `Score` column** — it is `answer_choice` (`A.`/`B.` per
>   question), not value-transfer as assumed here.
> - **Gaya Belajar (GB)** is **`answer_choice`** (`a.`/`b.`/`c.` per question,
>   30 columns) — *not* an A/B/C count tally into 3 columns.
> - **GE rubric** is already extracted → `data/ge_rubric.json` (that task is
>   done); GE is confirmed `computed_ge`.
> - V2.0's open questions on GE answer source and GB layout are now answered by
>   the fixtures; only the GE raw→scaled *norma* question remains (§2.1.7).
>
> _The original V2.0 planning body (goal, feature breakdown, roadmap A–D, user
> stories) was removed to prevent doubled tasks — see git history for the
> pre-fixture draft. Its one still-relevant task, "extend the completion `log`
> to report computed subtests (GE distribution, answer-block coverage per
> Gugus)", is carried into §2.1.5 Phase D._

---

# Version 2.1 — Confirmed Implementation Plan

> Status: **ready to build**. Scope locked against real fixtures:
> `REKAP JAWABAN SMAN 1 GEGESIK (RAW).xlsx`, `Template Rekap.xlsx`,
> `RUMUS GE.xlsx`. Added 2026-07-18.
>
> Guiding constraint: **do not change the matching algorithm** (`calc_match`,
> `normalize_name`, `normalize_date`, and the own→cross→proktor match order).
> Everything below is additive plumbing + new per-subtest write strategies +
> an anomaly-review layer. The Score-copy path for existing subtests keeps its
> exact behavior; only its row/column offsets shift for the new template.

## 2.1.1 New template format (confirmed)

`Template Rekap.xlsx` uses a **two-row header**, data from **row 3**:

- **Row 1** — merged section labels: `BIODATA` (A1:Q1), `EPPS` (S1:II1),
  `RIASEC` (IK1:MN1), `GAYA BELAJAR` (MP1:NS1).
- **Row 2** — field headers: biodata `NOMOR/NAMA/JK/TEMPAT LAHIR/TGL LAHIR/
  TGL PEMERIKSAAN/PENDIDIKAN/TUJUAN PEMERIKSAAN` (A–H); Score subtests
  `SE,WA,AN,GE,ME,RA,ZR,FA,WU` (I–Q); then question numbers `1..225` (EPPS,
  S–II), `1..108` (RIASEC, IK–MN), `1..30` (GAYA BELAJAR, MP–NS).
- **Row 3+** — student roster / data.

Consequence: header read moves row 1→2, data loops row 2→3. Once headers read
from row 2, existing Score subtests map unchanged (they live at I–Q). Numeric
question-number header cells must be skipped when building the `headers` dict.

## 2.1.2 Subtest taxonomy (confirmed)

| Subtest | RAW shape | Kind | Value written | Missing/unmatched → |
|---|---|---|---|---|
| SE, WA, AN, ME, RA, ZR, FA, WU | `Score` col | `value` *(existing)* | RAW `Score` copied | **`0`** (matched students only) |
| **GE** | `Score` + 16 answer cols (Q61–76) | **`computed_ge`** | **rubric sum 0–32; RAW `Score` ignored** | **`0`** |
| EPPS | 225 answer cols | `answer_choice` | `A.`/`B.` per question | **blank cell** |
| RIASEC | 108 answer cols | `answer_choice` | `Ya`/`Tidak` per question | **blank cell** |
| GB (→ section `GAYA BELAJAR`) | 30 answer cols | `answer_choice` | `a.`/`b.`/`c.` per question | **blank cell** |

Detection: EPPS/RIASEC/GB lack a `Score` column, so the existing
`fast_detect_sheets` already routes them to "other" and never touches them — a
new detector picks them up by (`NAMA LENGKAP` + `KELAS` + numbered question
columns, no `Score`). GE keeps its `Score` column but is intercepted and
recomputed.

`extract_choice(text)`: `re.match(r'^([A-Za-z]+)\.', s)` → returns the token
**with its dot, case preserved** (`A.`, `b.`); otherwise returns the trimmed
value (`Ya`/`Tidak`); blank/NaN → `None` (empty cell). Answers map to REKAP
question columns **positionally 1:1** (RAW question order = template `1..N`);
a count mismatch is an anomaly (see 2.1.4, trigger S3).

## 2.1.3 GE rubric engine (confirmed)

- Bundle `data/ge_rubric.json` in the repo — the 16-question keyword→weight map
  extracted from `RUMUS GE.xlsx` (Q61–76, columns D–S). **Never** read the
  broken-`#REF!` workbook at runtime.
- `score_ge(answers)`: for each of the 16 questions, normalize the answer
  (lowercase + trim + collapse internal whitespace, replicating Excel's
  case-insensitive `=`), first-match lookup against the ordered keyword list,
  sum weights → integer **0–32**. Exact keyword match only (no fuzzy/typo
  tolerance — matches the original formula, so e.g. `indera` ≠ `indra` scores 0).
- The GE RAW reader must retain the 16 answer columns (`read_raw_minimal`
  currently drops them).
- Write path: name-match via the **same** 3-tier `calc_match`, then write
  `score_ge(...)` into the GE column (L). A large gap between computed and RAW
  `Score` is surfaced as an anomaly (trigger V2), not silently resolved.

## 2.1.4 Confirmation triggers — "data janggal" review queue

Anything uncertain **pauses for human confirmation** instead of being written
blind. Each trigger produces a `pending` review item carrying a `reason`, a
`severity` (`block` = must decide before write / `warn` = written but
flagged), and enough context to decide. Triggers are grouped; thresholds are
tunable constants.

### A. Name / identity
| id | Trigger | Severity | Suggested action |
|---|---|---|---|
| N1 | Borderline match: `threshold ≤ score < AUTO_CONFIDENCE` *(existing)* | block | accept / reject |
| N2 | **Ambiguous** — ≥2 RAW rows match the same REKAP student above threshold | block | pick the correct RAW candidate |
| N3 | **Collision** — one RAW row is the best match for ≥2 REKAP students | block | assign to one, reject others |
| N4 | Cross-kelas match (student found in another gugus) *(currently silent)* | warn | confirm identity |
| N5 | Gibberish / tester name in RAW (no vowels, `Proktor/Tester…` pattern) matched a roster row | block | reject / confirm |
| N6 | Duplicate names within one gugus roster | warn | confirm which row gets the score |
| N7 | RAW student unmatched to any REKAP row (leftover) | warn | confirm not a missed student |

### B. Answer content (per matched student)
| id | Trigger | Severity | Suggested action |
|---|---|---|---|
| C1 | Straight-lining: one choice ≥ 90% of answered questions | block | confirm / discard as ngawur |
| C2 | Sparse: < 50% of questions answered | warn | confirm partial is acceptable |
| C3 | Fully blank answer sheet for a matched student | warn | leave blank / investigate |
| C4 | Unrecognized choice token (not `A./B./…`, not `Ya/Tidak`) | warn | confirm raw value written |
| C5 | RIASEC value outside `{Ya, Tidak}` | warn | confirm / normalize |

### C. Scores (value + GE)
| id | Trigger | Severity | Suggested action |
|---|---|---|---|
| V1 | RAW `Score` non-numeric / empty for a matched student | warn | write `0` / investigate |
| V2 | GE: `|computed − RAW Score|` exceeds gap threshold | warn | confirm computed is trusted |
| V3 | GE zero-signal: score `0` with ≥ N non-blank (gibberish) answers | block | confirm / discard as ngawur |
| V4 | Score out of plausible range (negative / above subtest max) | warn | confirm / clamp |
| V5 | Duplicate RAW rows, same student, conflicting scores | block | pick which score |

### D. Date / biodata
| id | Trigger | Severity | Suggested action |
|---|---|---|---|
| D1 | Birth date unparseable | warn | leave blank / enter manually |
| D2 | Ambiguous date (DD/MM vs MM/DD both valid, e.g. `05/07`) | warn | confirm interpretation |
| D3 | Implausible age (birth year → student age out of range) | warn | confirm |
| D4 | Conflicting biodata across RAW sheets (JK / birthplace differ) | warn | pick source of truth |

### E. Structure / sheet
| id | Trigger | Severity | Suggested action |
|---|---|---|---|
| S1 | Expected RAW sheet missing (EPPS/RIASEC/GB/GE not found) | block | proceed without / abort |
| S2 | Template section block missing (no EPPS columns in REKAP) | block | skip that subtest / abort |
| S3 | RAW answer-column count ≠ template block size → positional map unsafe | block | abort mapping / confirm alignment |
| S4 | Kelas format undetected or mixed across sheets | warn | confirm detected format |

### Extended `pending` item schema
```
{
  id, reason,            # e.g. "N2_ambiguous", "C1_straightlining"
  severity,              # "block" | "warn"
  gugus, subtes, kind,   # context
  nama_rekap, nama_raw,
  score,                 # name-match score (or computed value where relevant)
  detail,                # human-readable explanation of the anomaly
  candidates?,           # for N2/V5: [{nama_raw, score, _row_ref}] to choose from
  _write?                # internal: how to apply on accept (col/value or block)
}
```
`GET /status` exposes everything except `_`-prefixed keys (unchanged rule).
`POST /review` accepts `{decisions: {id: bool}}` for warn/accept-reject, plus
`{choices: {id: candidate_index}}` for N2/N3/V5 pick-one triggers.

## 2.1.5 Implementation phases

### Phase A — New-template adapter *(foundation)* — ✅ DONE
- Header read row 1→2 (skip numeric question-number cells); data loops row 2→3
  (`gugus_rows`, finalize date/uniform, finalize yellow).
- Section-block mapper: from row-1 merged labels build
  `section → {question_num → column}`; alias RAW `GB` → `GAYA BELAJAR`.
- **Exit:** existing SE/WA/…/WU output is byte-identical to pre-change on a
  fixture roster; no answer logic yet.

### Phase B — `answer_choice` kind (EPPS / RIASEC / GB) — ✅ DONE
- New answer-sheet detector + reader (keeps ordered question columns).
- `extract_choice`; block writer (N cells, positional map); reuse 3-tier match.
- **Exit:** EPPS/RIASEC/GB blocks fill with parsed tokens; blanks stay blank.

### Phase C — `computed_ge` kind
- Bundle `data/ge_rubric.json`; `score_ge`; GE reader keeps Q61–76; intercept
  GE to write computed sum instead of RAW `Score`.
- **Exit:** GE column shows rubric sum; `score_ge` unit-tested against rubric.

### Phase D — Missing-value normalization
- Score columns (incl. GE) → `0` for matched students; answer blocks → blank.
- Reconcile with yellow: yellow now means **unmatched or ngawur**, not "empty
  score cell". 0-fill applies to matched rows only.
- Extend the completion `log` to report computed/answer subtests sensibly
  (GE score distribution, EPPS/RIASEC/GB answer-block coverage per Gugus)
  *(carried over from the superseded V2.0 draft)*.
- **Exit:** matched students never show empty score cells; blanks preserved in
  answer blocks; yellow only on unmatched/ngawur; log reports all kinds.

### Phase E — Anomaly detection & confirmation review
- Implement triggers 2.1.4 (start with N1–N5, C1–C3, V1–V3, S1–S3; rest as
  follow-ups). Extend `pending` schema + `/review` to handle pick-one choices.
- `assess_quality(kind, answers)` for the ngawur/straight-line/sparse checks.
- **Exit:** anomalous data routes to review with actionable detail; nothing
  janggal is written unconfirmed at `block` severity.

### Phase F — Test console + dev ergonomics
- Ship `test_console.html` (single-file harness, see 2.1.6).
- Optional dev-only route `GET /console` (serves the file, same-origin, no
  CORS) — gated behind an env flag; keeps the "no public UI" posture.
- **Exit:** full `/process → /status → review → /download` loop drivable from
  a browser against a local instance.

## 2.1.6 Test console (frontend)

`test_console.html` — one self-contained page (no build, no external deps):
- Config: base URL + bearer token (persisted in `localStorage`).
- Start a job: `raw_file`, `rekap_file`, `threshold`, `tgl_pemeriksaan`,
  `pendidikan`, `manual_overrides` (JSON).
- Live poll of `/status`: progress bar, status text, streaming log.
- **Review queue** grouped by `reason`/`severity`: accept/reject toggles for
  warn/binary items and candidate-pickers for N2/N3/V5; submits `/review`.
- Authenticated download of the finished `.xlsx` (fetch → blob).

To avoid CORS, it is meant to be served same-origin via the optional
`GET /console` route (Phase F). Standalone use requires CORS enabled or a
browser started with web-security disabled — documented in the page itself.

## 2.1.7 Remaining open questions
- [x] **GE norma** — **CONFIRMED 2026-07-18: the rubric sum (0–32) is the
      final GE value.** No raw→scaled conversion. Phase C writes the sum directly.
- [x] **0-fill scope** — **CONFIRMED: only *matched* students get score→`0`;**
      unmatched/ngawur stay blank + yellow.
- [x] **Ngawur thresholds / severity** — **CONFIRMED: use the §2.1.4 defaults
      as specified** (straight-line ≥ 90%, sparse < 50%, block/warn split as
      tabled). Tunable later.
- [ ] **Ambiguous/collision (N2/N3)** — auto-pick highest score and only
      surface ties, or always ask on ≥2 above-threshold candidates? *(defer to
      Phase E build)*

## 2.1.8 Build progress
- [x] **Phase A** — new-template adapter (header 1↔2 / data 2↔3 auto-detect,
      section-block mapper). Verified on real template + legacy fallback.
- [x] **Phase B** — `answer_choice` kind (EPPS/RIASEC/GB). Verified: 6 students
      × 225/108/30 cells filled with correct tokens; borderline → review block.
- [ ] **Phase C** — `computed_ge` (rubric sum is final; `data/ge_rubric.json`
      ready). *Next up.*
- [ ] **Phase D** — missing-value 0-fill (matched students only) + yellow
      reconcile + log for answer/computed kinds.
- [ ] **Phase E** — anomaly/confirmation review (§2.1.4 defaults).
- [ ] **Phase F** — test console route (`GET /console`) + docs.
