"""Address normalization utilities (from desktop WRS app)."""
from __future__ import annotations

import re
from typing import Any

import pandas as pd


def normalize_address(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().lower().replace("ё", "е")
    for ch in [",", ".", ";", '"', "'"]:
        s = s.replace(ch, " ")
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip()


def shorten_address(address: str) -> str:
    if not address:
        return ""
    parts = [p.strip() for p in str(address).split(",") if p.strip()]
    return ", ".join(parts[3:]).strip() if len(parts) >= 4 else str(address).strip()


def prettify_address(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    parts = [p.strip() for p in text.split(",") if p.strip()]
    cleaned, seen = [], set()
    for part in parts:
        n = normalize_address(part)
        if re.fullmatch(r"\d{5,6}", part):
            continue
        if any(t in n for t in [" обл", " область", " край", " респ"]):
            continue
        if n in seen:
            continue
        cleaned.append(part)
        seen.add(n)
    return ", ".join(cleaned) if cleaned else text


def is_moscow(value: Any) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    t = normalize_address(value)
    if not t:
        return False
    return "москва" in t


def street_key(addr: str, level: int = 1) -> str:
    if not addr or pd.isna(addr):
        return ""
    s = str(addr).lower().replace("ё", "е")
    s = re.sub(r"\b\d{6}\b", "", s)
    for ch in [",", ".", ";", "'", '"']:
        s = s.replace(ch, " ")
    s = re.sub(
        r"\b(г|пгт|с|сл|обл|р-н|ул|пр-кт|проспект|пр|пл|пер|пр-д|бул|ш|"
        r"дом|д|стр|строение|корп|зд|влад|помещение|пом|п)\b\.?",
        "",
        s,
    )
    s = re.sub(r"\b(орловская|курская|липецкая|орловский|курский|липецкий)\b", "", s)
    while "  " in s:
        s = s.replace("  ", " ")
    tokens = [t.strip() for t in s.split() if t.strip() and len(t.strip()) > 1]
    if not tokens:
        return ""
    if level == 1:
        return " ".join(tokens[-3:])
    if level == 2:
        return " ".join(tokens[-2:])
    if level == 3:
        return " ".join(tokens[-4:])
    return " ".join(tokens)


def safe_sap(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    v = str(x).strip()
    if not v or v.lower() == "nan":
        return ""
    try:
        return str(int(float(v)))
    except Exception:
        return v.lower()
