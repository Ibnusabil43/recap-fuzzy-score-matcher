"""Recap Fuzzy Score Matcher — psikotes score automation.

The Flask entrypoint (`app.py` at the repo root) is intentionally thin; all
logic lives here, split by concern:

- `config`     — constants, thresholds, paths
- `matching`   — fuzzy name matching (`calc_match`, memoized)
- `dates`      — date normalization to DD/MM/YYYY
- `raw_reader` — RAW workbook detection + column readers
- `layout`     — REKAP two-row-header parsing + answer-section mapping
- `scoring`    — GE rubric scoring + answer-choice token parsing
- `quality`    — "ngawur"/anomaly quality gates
- `jobstore`   — in-memory job registry + temp-file lifecycle
- `engine`     — the matching / review / finalize pipeline
- `auth`       — shared-token route guard

Behavior is identical to the original single-file implementation; this package
only reorganizes it and adds result-preserving performance optimizations
(memoized matching, pool pre-extraction, early-exit leftover scan).
"""
