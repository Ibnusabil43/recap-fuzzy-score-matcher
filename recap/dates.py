"""Date normalization to a uniform DD/MM/YYYY string."""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd


def normalize_date(val: object) -> str | None:
    """Convert any date value to a DD/MM/YYYY string, or None if unparseable."""
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
        if 1 <= a <= 12 and 1 <= b <= 31: return f"{b:02d}/{a:02d}/{y:04d}"
        if 1 <= b <= 12 and 1 <= a <= 31: return f"{a:02d}/{b:02d}/{y:04d}"
        return s
    m2 = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{2,4})$', s)
    if m2:
        a, b, y = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if y < 100: y += 2000
        if a > 12 and b <= 12: a, b = b, a
        return f"{b:02d}/{a:02d}/{y:04d}" if 1 <= a <= 12 else s
    try:
        dt = pd.to_datetime(val)
        y = dt.year
        if y < 100: y += 2000
        return f"{dt.day:02d}/{dt.month:02d}/{y:04d}" if 1900 <= y <= 2100 else s
    except Exception:
        return s
