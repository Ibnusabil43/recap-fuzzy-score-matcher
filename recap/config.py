"""Shared configuration constants and filesystem paths."""
from __future__ import annotations

from pathlib import Path

# Repo root. This file lives in recap/, so parent.parent is the project root
# where data/ and test_console.html sit.
BASE_DIR: Path = Path(__file__).resolve().parent.parent
GE_RUBRIC_PATH: Path = BASE_DIR / "data" / "ge_rubric.json"
TEST_CONSOLE_PATH: Path = BASE_DIR / "test_console.html"

# Upload size cap for incoming multipart requests (RAW + REKAP workbooks).
MAX_CONTENT_LENGTH: int = 50 * 1024 * 1024

# Each finished job retains a loaded workbook + its output file. Without a cap
# this grows unbounded on a long-running server, so we evict the oldest terminal
# jobs past this limit (see jobstore.evict_old_jobs).
MAX_JOBS: int = 30

# Name-match score at/above which a score is written without manual review.
# Skor kecocokan nama di bawah ini butuh konfirmasi manual sebelum ditulis ke Rekap.
AUTO_CONFIDENCE: float = 0.90

# Floor below which a leftover student's "closest roster name" is just noise
# (coincidental partial overlap, not a plausible same-person typo) — e.g. a junk
# RAW row scoring 15-50% against an unrelated name. Below this, the
# closest_candidate is still shown for transparency but not offered as a
# one-click-confirmable match in the review UI.
LEFTOVER_CANDIDATE_MIN: float = 0.55

# Quality-gate thresholds ("ngawur" detection) — tunable.
STRAIGHTLINE_FRAC: float = 0.90   # one choice dominating this fraction = suspect
SPARSE_MIN_FRAC: float = 0.50     # answered below this fraction = suspect
GE_ZERO_MIN_ANSWERS: int = 5      # GE = 0 despite this many filled answers = suspect

# Upper bound on distinct (name, name) pairs cached by matching.calc_match. Large
# enough that a single school's whole run never evicts (one school ≈ a few
# hundred-thousand unique pairs), while still capping long-running-server memory.
MATCH_CACHE_SIZE: int = 1_048_576

# Answer-type REKAP sections: a numbered question block, not a single score cell.
# Keys are the row-1 section labels as they appear in the template.
ANSWER_SECTION_LABELS: frozenset[str] = frozenset({"EPPS", "RIASEC", "GAYA BELAJAR"})

# RAW sheet name -> REKAP section label (row-1 merged label).
RAW_TO_SECTION: dict[str, str] = {"EPPS": "EPPS", "RIASEC": "RIASEC", "GB": "GAYA BELAJAR"}
