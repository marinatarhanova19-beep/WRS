"""Авторизация дворников (прилегающая территория) по графику."""
from __future__ import annotations

import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd

from address import is_moscow, normalize_address, prettify_address, shorten_address


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


def addr_key(x: Any) -> str:
    """Ключ адреса: улица + номер дома (цифры сохраняем)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).lower().replace("ё", "е")
    s = re.sub(r"\b\d{6}\b", "", s)
    for ch in [",", ".", ";", '"', "'"]:
        s = s.replace(ch, " ")
    s = re.sub(
        r"\b(г|пгт|с|сл|обл|р-н|ул|пр-кт|проспект|пр|пл|пер|пр-д|проезд|бул|ш|"
        r"дом|д|стр|строение|корп|к|зд|здание|влад|помещение|пом|п)\b",
        " ",
        s,
    )
    while "  " in s:
        s = s.replace("  ", " ")
    tokens = [t.strip() for t in s.split() if t.strip()]
    tokens = [t for t in tokens if t.isdigit() or len(t) > 1]
    return " ".join(tokens[-5:])


def should_work(schedule_str: str, weekday: int) -> bool:
    """Должен ли магазин иметь уборку прилегающей в этот день недели (0=пн)."""
    if not schedule_str or (isinstance(schedule_str, float) and pd.isna(schedule_str)):
        return False
    s = str(schedule_str).lower().replace("ё", "е")
    if "не требуется" in s:
        return False
    if "ежедневно" in s:
        return True
    if "6/1" in s:
        return weekday <= 5
    if "5/2" in s:
        return weekday <= 4
    day_map = {
        "пн": 0,
        "вт": 1,
        "ср": 2,
        "чт": 3,
        "пт": 4,
        "сб": 5,
        "вс": 6,
        "вск": 6,
    }
    req = [v for k, v in day_map.items() if k in s]
    if req:
        return weekday in req
    if "1 раз" in s:
        return weekday == 0
    return False


def parse_schedule_excel(df: pd.DataFrame) -> List[dict]:
    """Разобрать Excel графика дворников → список {address, schedule, sap, _key}."""
    cols_l = {str(c).lower(): c for c in df.columns}

    def pick(*needles: str):
        for n in needles:
            for low, orig in cols_l.items():
                if n in low:
                    return orig
        return None

    addr_col = pick("адрес")
    sched_col = pick("график")
    sap_col = pick("sap") or pick("№ тт", "номер тт") or pick("тт")

    if not addr_col or not sched_col:
        raise ValueError(
            "Не найдены колонки «Адрес» и «График».\n"
            f"Колонки в файле: {list(df.columns)}"
        )

    rows: List[dict] = []
    for _, r in df.iterrows():
        addr = str(r.get(addr_col, "") or "").strip()
        sched = str(r.get(sched_col, "") or "").strip()
        sap = safe_sap(r.get(sap_col, "")) if sap_col else ""
        if addr and sched:
            rows.append(
                {
                    "address": addr,
                    "schedule": sched,
                    "sap": sap,
                    "_key": addr_key(addr),
                }
            )
    return rows


def detect_weekday_from_shifts(
    shifts_df: pd.DataFrame, start_col: Optional[str] = None
) -> Tuple[int, str]:
    """Определить день недели из файла смен. Возвращает (weekday 0=пн, label)."""
    today = datetime.now()
    weekday = today.weekday()
    label = today.strftime("%d.%m.%Y")

    if shifts_df is None or shifts_df.empty:
        return weekday, label

    col = start_col
    if not col or col not in shifts_df.columns:
        for c in ["Начало работы", "Дата начала", "Начало смены", "Дата"]:
            if c in shifts_df.columns:
                col = c
                break

    if not col or col not in shifts_df.columns:
        return weekday, label

    try:
        series = shifts_df[col].dropna()
        if series.empty:
            return weekday, label
        # try parse as datetime
        parsed = pd.to_datetime(series, dayfirst=True, errors="coerce").dropna()
        if not parsed.empty:
            mode = parsed.mode()
            d = (mode.iloc[0] if not mode.empty else parsed.iloc[0]).to_pydatetime()
            return d.weekday(), d.strftime("%d.%m.%Y (%a)")
        # fallback: string dd.mm.yyyy
        texts = series.astype(str).str.strip().str[:10]
        most = texts.value_counts().idxmax()
        d = datetime.strptime(most, "%d.%m.%Y")
        return d.weekday(), d.strftime("%d.%m.%Y")
    except Exception:
        return weekday, label


def build_adj_keys_from_shifts(shift_rows: List[dict]) -> Tuple[Set[str], Set[str]]:
    """
    Из строк смен прилегающей собрать множества ключей адреса и SAP.
    shift_rows: [{address, sap, type}] или сырые dict с колонками.
    """
    keys: Set[str] = set()
    saps: Set[str] = set()
    for r in shift_rows or []:
        addr = r.get("address") or r.get("Адрес магазина") or ""
        sap = safe_sap(r.get("sap") or r.get("SAP магазина") or "")
        if addr:
            k = addr_key(addr)
            if k:
                keys.add(k)
            # also normalized full
            n = normalize_address(addr)
            if n:
                keys.add(n)
        if sap:
            saps.add(sap)
    return keys, saps


def extract_adj_from_dataframe(df: pd.DataFrame) -> List[dict]:
    """Вытащить строки «уборка прилегающих» из файла смен."""
    if df is None or df.empty:
        return []
    if "Тип работы" not in df.columns or "Адрес магазина" not in df.columns:
        return []
    wt = df["Тип работы"].astype(str).str.strip().str.lower()
    adj = df[wt.str.contains("прилега", na=False)].copy()
    if adj.empty:
        return []
    rows = []
    for _, r in adj.iterrows():
        addr = str(r.get("Адрес магазина", "") or "").strip()
        if not addr or is_moscow(addr):
            continue
        rows.append(
            {
                "address": addr,
                "sap": safe_sap(r.get("SAP магазина", "")) if "SAP магазина" in adj.columns else "",
                "address_short": shorten_address(addr),
            }
        )
    return rows


def evaluate_dvornik(
    schedule_rows: List[dict],
    adj_rows: List[dict],
    weekday: int,
    ab_df: pd.DataFrame,
    regions: Optional[dict] = None,
) -> dict:
    """
    Сравнить график с фактическими сменами прилегающей.

    Возвращает:
      nonauth, mgr_should, mgr_worked, mgr_nonauth, n_total, n_worked,
      region_stats, day_label (caller sets), report_text pieces
    """
    regions = regions or {}
    adj_keys, adj_saps = build_adj_keys_from_shifts(adj_rows)

    # manager maps from address book
    key_to_mgr: Dict[str, str] = {}
    if ab_df is not None and not ab_df.empty:
        for _, r in ab_df.iterrows():
            addr = str(r.get("Адрес магазина", "") or "")
            mgr = str(r.get("Менеджер", "") or "").strip()
            if not addr or not mgr:
                continue
            k = addr_key(addr)
            n = normalize_address(addr)
            if k:
                key_to_mgr.setdefault(k, mgr)
            if n:
                key_to_mgr.setdefault(n, mgr)

    def resolve_manager(addr: str, sap: str) -> str:
        if sap and sap in key_to_mgr:
            return key_to_mgr[sap]
        k = addr_key(addr)
        n = normalize_address(addr)
        return key_to_mgr.get(k) or key_to_mgr.get(n) or "Менеджер не определён"

    def is_worked(addr: str, sap: str) -> bool:
        if sap and sap in adj_saps:
            return True
        k = addr_key(addr)
        n = normalize_address(addr)
        if k and k in adj_keys:
            return True
        if n and n in adj_keys:
            return True
        # soft: substring token overlap for last tokens
        if k:
            for ak in adj_keys:
                if k == ak:
                    return True
                # house number match on same street stem
                if len(k) > 4 and (k in ak or ak in k):
                    return True
        return False

    nonauth: List[dict] = []
    mgr_should: Dict[str, int] = {}
    mgr_worked: Dict[str, int] = {}
    mgr_nonauth: Dict[str, int] = {}
    n_total = 0
    n_worked = 0

    for row in schedule_rows or []:
        sched = row.get("schedule", "")
        if not should_work(sched, weekday):
            continue
        addr = row.get("address", "")
        if is_moscow(addr):
            continue
        sap = safe_sap(row.get("sap", ""))
        mgr = resolve_manager(addr, sap)
        n_total += 1
        mgr_should[mgr] = mgr_should.get(mgr, 0) + 1
        worked = is_worked(addr, sap)
        if worked:
            n_worked += 1
            mgr_worked[mgr] = mgr_worked.get(mgr, 0) + 1
        else:
            mgr_nonauth[mgr] = mgr_nonauth.get(mgr, 0) + 1
            nonauth.append(
                {
                    "manager": mgr,
                    "address": prettify_address(addr) or addr,
                    "address_short": shorten_address(addr) or addr,
                    "schedule": sched,
                    "sap": sap,
                }
            )

    # region aggregation
    def region_for(manager: str) -> str:
        for region, managers in regions.items():
            if manager in managers:
                return region
        return "Без региона"

    region_stats: Dict[str, dict] = {}
    for mgr, sh in mgr_should.items():
        reg = region_for(mgr)
        rs = region_stats.setdefault(reg, {"should": 0, "worked": 0, "nonauth": 0})
        rs["should"] += sh
        rs["worked"] += mgr_worked.get(mgr, 0)
        rs["nonauth"] += mgr_nonauth.get(mgr, 0)

    mgr_rows = []
    for mgr in sorted(mgr_should.keys(), key=lambda m: (m == "Менеджер не определён", m)):
        sh = mgr_should.get(mgr, 0)
        wr = mgr_worked.get(mgr, 0)
        na = mgr_nonauth.get(mgr, 0)
        sl = (wr / sh) if sh else 0.0
        mgr_rows.append(
            {
                "manager": mgr,
                "region": region_for(mgr),
                "should": sh,
                "worked": wr,
                "nonauth": na,
                "sl": sl,
                "sl_percent": f"{sl * 100:.1f}%",
            }
        )

    region_rows = []
    for reg, rs in sorted(region_stats.items()):
        sh = rs["should"]
        wr = rs["worked"]
        sl = (wr / sh) if sh else 0.0
        region_rows.append(
            {
                "region": reg,
                "should": sh,
                "worked": wr,
                "nonauth": rs["nonauth"],
                "sl": sl,
                "sl_percent": f"{sl * 100:.1f}%",
            }
        )

    nonauth_sorted = sorted(
        nonauth,
        key=lambda r: (
            r["manager"] == "Менеджер не определён",
            r["manager"],
            r["address"],
        ),
    )

    return {
        "nonauth": nonauth_sorted,
        "mgr_rows": mgr_rows,
        "region_rows": region_rows,
        "n_total": n_total,
        "n_worked": n_worked,
        "n_nonauth": len(nonauth_sorted),
    }


def build_dvornik_report(
    result: dict,
    day_label: str,
    with_sl: bool = True,
) -> str:
    sep = "─" * 32
    n_total = result.get("n_total", 0)
    n_worked = result.get("n_worked", 0)
    nonauth = result.get("nonauth") or []

    lines = [
        "🌿 АВТОРИЗАЦИЯ ПРИЛЕГАЮЩЕЙ ТЕРРИТОРИИ",
        f"    {day_label}  |  Должны работать: {n_total}  |  Открыли: {n_worked}",
        sep,
    ]

    if nonauth:
        by_mgr: Dict[str, List] = {}
        for r in nonauth:
            by_mgr.setdefault(r["manager"], []).append(r)
        for mgr in sorted(by_mgr.keys(), key=lambda m: (m == "Менеджер не определён", m)):
            items = by_mgr[mgr]
            lines.append(f"{mgr} ({len(items)}):")
            for r in items:
                lines.append(f"  🧹 {r.get('address_short') or r.get('address')}")
            lines.append(sep)
    else:
        lines.append("✅ Все магазины авторизованы.")
        lines.append(sep)

    if with_sl:
        reg_rows = result.get("region_rows") or []
        if reg_rows:
            lines.append("SL по регионам:")
            for r in reg_rows:
                lines.append(
                    f"  {r['region']}: SL = {r['sl_percent']}  "
                    f"(открыли {r['worked']} из {r['should']}, не авт. {r['nonauth']})"
                )
            lines.append(sep)

    return "\n".join(lines)
