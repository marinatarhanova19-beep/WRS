"""Shared helpers from desktop WRS app."""
from __future__ import annotations

import re
from typing import Any, Optional

import pandas as pd


def clamp01(x: float) -> float:
    return max(0.0, min(1.0, x))


def fmt_sl(sl: float) -> str:
    return f"{int(sl * 1000) / 10:.1f}%"


def safe_int(x: Any, d: int = 0) -> int:
    try:
        return int(x)
    except Exception:
        return d


def int_mask(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str).str.lower().str.replace("ё", "е", regex=False)
    return (
        s.str.contains(r"\bв магазине\b", na=False)
        | s.str.contains("внутрен", na=False)
        | s.str.contains("помещен", na=False)
    )


def adj_mask(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.lower()
        .str.replace("ё", "е", regex=False)
        .str.contains("прилега", na=False)
    )


def parse_hours(value: Any) -> Optional[float]:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        v = float(value)
        return v * 24 if 0 <= v < 10 else v
    t = str(value).strip()
    if not t:
        return None
    try:
        return pd.to_timedelta(t).total_seconds() / 3600
    except Exception:
        pass
    m = re.fullmatch(r"(\d{1,3}):(\d{2})(?::(\d{2}))?", t)
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60 + int(m.group(3) or 0) / 3600
    try:
        v = float(t.replace(",", "."))
        return v * 24 if 0 <= v < 10 else v
    except Exception:
        return None


def fmt_hours(h: float) -> str:
    m = round(float(h) * 60)
    hh = m // 60
    mm = m % 60
    if hh == 0 and mm == 0:
        return "0 ч"
    if mm == 0:
        return f"{hh} ч"
    return f"{hh} ч {mm} мин"


def fmt_money(rub: float) -> str:
    return f"{rub:,.0f} ₽".replace(",", " ")
