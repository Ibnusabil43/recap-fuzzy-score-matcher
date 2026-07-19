"""Quality gates ("ngawur" / careless-response detection).

Even a confident name match must be held for manual confirmation if the content
looks careless — a straight-lined answer sheet, a mostly-blank one, or a GE
score of 0 despite the student clearly having answered. Each check returns a
(reason, severity, detail) triple to queue for review, or None if the response
looks genuine.
"""
from __future__ import annotations

from collections import Counter

import pandas as pd

from .config import GE_ZERO_MIN_ANSWERS, SPARSE_MIN_FRAC, STRAIGHTLINE_FRAC

Flag = tuple[str, str, str]


def assess_answer_quality(answers: list[str | None]) -> Flag | None:
    """Flag careless/nonsense ('ngawur') answer sheets. `answers` is the list of
    parsed tokens for one student. Returns (reason, severity, detail) or None if
    the response looks genuine."""
    n_total = len(answers)
    if n_total == 0: return None
    vals = [a for a in answers if a not in (None, '')]
    n_ans = len(vals)
    if n_ans == 0:
        return ('C3_blank', 'warn', 'Lembar jawaban kosong')
    top, cnt = Counter(vals).most_common(1)[0]
    if n_ans >= 5 and cnt / n_ans >= STRAIGHTLINE_FRAC:
        return ('C1_straightlining', 'block',
                f'{int(cnt / n_ans * 100)}% jawaban sama ("{top}") — kemungkinan ngawur')
    if n_ans / n_total < SPARSE_MIN_FRAC:
        return ('C2_sparse', 'warn', f'hanya {n_ans}/{n_total} soal terisi')
    return None


def assess_ge(score_val: object, row_series: pd.Series, qcols: list[object]) -> Flag | None:
    """Flag a GE score of 0 that came from a student who *did* answer — usually
    means gibberish/typo'd free-text (nothing matched the rubric), worth a look.
    Returns (reason, severity, detail) or None."""
    if score_val != 0: return None
    n_ans = sum(1 for c in qcols
                if not pd.isna(row_series.get(c)) and str(row_series.get(c)).strip())
    if n_ans >= GE_ZERO_MIN_ANSWERS:
        return ('V3_ge_zero', 'block',
                f'GE = 0 padahal {n_ans} soal terisi — kemungkinan ngawur/typo')
    return None
