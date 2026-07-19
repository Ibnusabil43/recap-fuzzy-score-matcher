"""Fuzzy name matching for reconciling RAW answers against the REKAP roster.

`calc_match` is memoized: it is a pure function of two normalized strings, and
the same name pairs recur heavily across every subtest, the identity pass, and
the leftover check. Caching removes the redundant SequenceMatcher work — the
single biggest speedup here — without changing any score it returns.
"""
from __future__ import annotations

from difflib import SequenceMatcher
from functools import lru_cache

import pandas as pd

from .config import MATCH_CACHE_SIZE


def normalize_name(name: object) -> str:
    """Upper-case, strip, and flatten punctuation so name variants compare equal."""
    if pd.isna(name):
        return ""
    n = str(name).strip().upper()
    for ch in ["‘", "’", "'", '"']:
        n = n.replace(ch, "'")
    n = n.replace(".", " ").replace("-", " ")
    while "  " in n:
        n = n.replace("  ", " ")
    return n.strip()


def _ordered_word_match(short_w: list[str], long_w: list[str]) -> float:
    """Match words in order. A short token (<=4 chars) that's an exact prefix
    of a long-side word counts as a partial (initial/abbreviation) match —
    covers both single-letter initials ("M" -> MUHAMMAD) and the common
    Indonesian 3-4 letter prefix abbreviations ("Moh"/"Muh"/"Moch" -> MUHAMMAD).
    A token that exactly equals two CONSECUTIVE long-side words concatenated
    also counts as a full match — covers a compact-vs-split spelling of the
    same name ("AZZAHRA" == "AZ"+"ZAHRA"), which the prefix rule above can't
    reach since it only ever compares one short token to one long word."""
    li, matched = 0, 0.0
    for sw in short_w:
        for j in range(li, len(long_w)):
            lw = long_w[j]
            if sw == lw:
                matched += 1; li = j + 1; break
            elif len(sw) <= 4 and lw.startswith(sw):
                matched += 0.8; li = j + 1; break
            elif j + 1 < len(long_w) and sw == lw + long_w[j + 1]:
                matched += 1; li = j + 2; break
    return matched


@lru_cache(maxsize=MATCH_CACHE_SIZE)
def calc_match(a: str, b: str) -> float:
    """Similarity in [0, 1] between two normalized names.

    Combines SequenceMatcher ratio, substring containment, ordered word match
    with initial/prefix abbreviations, and strict acronym detection. Pure and
    memoized — identical inputs always return the identical score."""
    if not a or not b:
        return 0
    if a == b:
        return 1.0
    s = SequenceMatcher(None, a, b).ratio()
    # Substring check (min 4 chars to avoid short false positives)
    if len(a) >= 4 and a in b: s = max(s, len(a) / len(b), 0.82)
    if len(b) >= 4 and b in a: s = max(s, len(b) / len(a), 0.82)
    wa, wb = a.split(), b.split()
    if len(wa) >= 2 and len(wb) >= 2:
        # 1) Ordered word match with initials/prefix abbreviations
        #    (M->MUHAMMAD, MOH->MUHAMMAD, MUH->MUHAMMAD, P->PUTRA)
        for sw, lw in [(wa, wb), (wb, wa)]:
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
        for all_w, check_w in [(wa, wb), (wb, wa)]:
            for token in all_w:
                if len(token) < 2 or len(token) > 6: continue
                if token in check_w: continue  # already a real word, skip
                for start in range(len(check_w)):
                    needed = len(token)
                    if start + needed > len(check_w): break
                    candidate = check_w[start:start + needed]
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


# A pool of RAW candidates, pre-extracted once as (row_index, normalized_name)
# tuples so the per-student matching loops never pay pandas .iterrows() overhead.
NamePairs = list[tuple[object, str]]


def as_pairs(df: pd.DataFrame) -> NamePairs:
    """Freeze a RAW DataFrame into (index, _norm) tuples for repeated scanning.

    Iterates in the DataFrame's row order, exactly matching what .iterrows()
    yielded before, so the winner and its tie-breaking are unchanged."""
    return list(zip(df.index, df["_norm"]))


def best_name_match(gn: str, mt: str | None, pairs: NamePairs) -> tuple[float, object]:
    """Highest calc_match over a pre-extracted pool; a confident manual override
    wins. Returns (best_score, best_index) with best_index None if pool empty.

    Equivalent to scanning df.iterrows() and keeping the first strict-max row —
    same order, same `> best_s` tie-break — but without rebuilding a Series per
    row."""
    best_s: float = 0
    best_i: object = None
    for idx, norm in pairs:
        s = calc_match(gn, norm)
        if mt and calc_match(mt, norm) > 0.9: s = 0.95
        if s > best_s: best_s, best_i = s, idx
    return best_s, best_i


def closest_roster_candidate(
    k: str,
    gugus_rows: dict[int, dict[int, str]],
    threshold: float,
) -> tuple[bool, float, str | None, int | None, int | None]:
    """Leftover (N7) membership test for one RAW student's normalized name `k`.

    Returns (represented, best_score, best_candidate, cand_gugus, cand_row).

    If any roster name scores >= threshold the student IS represented; we can
    stop early and return represented=True (the caller skips them, and the
    candidate fields are irrelevant). Otherwise we scan the whole roster to find
    the single closest below-threshold name — the full max, identical to the
    original exhaustive loop — which the review UI offers as a confirmable
    candidate."""
    best: float = 0
    best_candidate: str | None = None
    cand_gugus: int | None = None
    cand_row: int | None = None
    for gn2, rows in gugus_rows.items():
        for ri, rn in rows.items():
            sc = calc_match(k, rn)
            if sc >= threshold:
                return True, sc, rn, gn2, ri
            if sc > best:
                best, best_candidate, cand_gugus, cand_row = sc, rn, gn2, ri
    return False, best, best_candidate, cand_gugus, cand_row
