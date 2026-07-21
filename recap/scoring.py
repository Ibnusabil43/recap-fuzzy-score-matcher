"""GE computed scoring and answer-choice token parsing.

GE answers are free text, not a score. Each question has an ordered
keyword->weight rubric (from RUMUS GE.xlsx, extracted once into
data/ge_rubric.json, keyed by question number "61".."76"). Per question the
first exact keyword match wins (best synonym=2, looser term=1, else 0); the GE
value is the sum (0-32). The RAW "Score" column is ignored — recomputed here.

Answer-choice subtests (EPPS/RIASEC/GB) instead carry one answer string per
question; `extract_choice` pulls the leading choice token to write verbatim.
"""
from __future__ import annotations

import json
import re

import pandas as pd

from .config import GE_RUBRIC_PATH


def _load_rubric() -> dict[str, list[dict[str, object]]]:
    """Load the GE rubric once at import; empty dict if the file is missing so
    GE simply falls back to writing nothing computable."""
    try:
        with open(GE_RUBRIC_PATH, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


GE_RUBRIC: dict[str, list[dict[str, object]]] = _load_rubric()


def extract_choice(val: object) -> str | None:
    """Take the leading choice token from an answer, dot included, case
    preserved: "A. Saya suka..." -> "A.", "b. Mendengarkan..." -> "b.".
    PAPI answers lead with a numbered choice instead: "1). Saya seorang
    pekerja..." -> "1)." (paren + dot kept). Values without either prefix
    (RIASEC's "Ya"/"Tidak") pass through trimmed. Blank/NaN -> None (write
    nothing / empty cell)."""
    if pd.isna(val): return None
    s = str(val).strip()
    if not s: return None
    m = re.match(r'^([A-Za-z]+)\.', s)
    if m: return m.group(0)
    m = re.match(r'^(\d+\)\.?)', s)
    if m: return m.group(1)
    return s


def _norm_ge_answer(val: object) -> str:
    """Lowercase + trim + collapse whitespace, replicating Excel's `=` compare."""
    if pd.isna(val): return ""
    return re.sub(r'\s+', ' ', str(val).strip().lower())


def score_ge(row_series: pd.Series, qcols: list[object]) -> int:
    """Sum GE rubric weights across the answer columns; first exact keyword match
    per question wins, unmatched = 0. Returns int 0-32."""
    total = 0
    for col in qcols:
        m = re.match(r'^(\d+)\.', str(col).strip())
        if not m: continue
        rubric = GE_RUBRIC.get(m.group(1))
        if not rubric: continue
        ans = _norm_ge_answer(row_series.get(col))
        if not ans: continue
        for entry in rubric:  # ordered -> first match wins
            if ans == entry['kw']:
                total += entry['w']; break
    return total


def value_for(subtes: object, row_series: pd.Series, ge_qcols: list[object]) -> object:
    """Value to write for a matched RAW row: computed sum for GE, copied Score
    otherwise. Keeps the existing value-copy behavior for every other subtest."""
    if str(subtes).strip().upper() == 'GE':
        return score_ge(row_series, ge_qcols)
    return row_series.get('Score')
