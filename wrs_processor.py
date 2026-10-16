"""WRS report processing (core logic from desktop app)."""
from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

from address import (
    is_moscow,
    normalize_address,
    prettify_address,
    shorten_address,
)
from helpers import adj_mask, clamp01, fmt_sl, int_mask, safe_int


def is_under_reconstruction(
    addr_norm: str, recon_rows: list, on_date: Optional[date] = None
) -> bool:
    if on_date is None:
        on_date = date.today()
    for r in recon_rows or []:
        if normalize_address(r.get("address", "")) != addr_norm:
            continue
        try:
            d_from = datetime.strptime(r.get("date_from", ""), "%Y-%m-%d").date()
            d_to = datetime.strptime(r.get("date_to", ""), "%Y-%m-%d").date()
        except Exception:
            continue
        if d_from <= on_date <= d_to:
            return True
    return False


def region_for(manager: str, regions: dict) -> str:
    for region, managers in (regions or {}).items():
        if manager in managers:
            return region
    return "Без региона"


def sl_manager(
    manager: str,
    total: dict,
    non_auth: dict,
    adj_only: dict,
    corrections: dict,
) -> float:
    """
    SL = (Всего − (Н/к уборка + Не авт + Прилег) + (Реконстр + Тех)) / Всего
    """
    t = int(total.get(manager, 0) or 0)
    if t <= 0:
        return 1.0
    corr = corrections.get(manager) or {}
    q = safe_int(corr.get("q", 0))
    nr = safe_int(corr.get("nr", 0))
    tech = safe_int(corr.get("tech", 0))
    na = len(non_auth.get(manager, []) or [])
    adj = len(adj_only.get(manager, []) or [])
    score = (t - (q + na + adj) + (nr + tech)) / t
    return clamp01(score)


def sl_region(
    managers: list,
    total: dict,
    non_auth: dict,
    adj_only: dict,
    corrections: dict,
) -> float:
    num = den = 0.0
    for m in managers:
        t = int(total.get(m, 0) or 0)
        if t <= 0:
            continue
        den += t
        num += t * sl_manager(m, total, non_auth, adj_only, corrections)
    if den <= 0:
        return 1.0
    return clamp01(num / den)


def process_wrs(
    shops: pd.DataFrame,
    ab_df: pd.DataFrame,
    recon: list,
    regions: dict,
    tracker: dict,
    corrections: Optional[dict] = None,
    cancel: Optional[Callable[[], bool]] = None,
) -> dict:
    """Process WRS Excel and return structured result."""
    corrections = corrections or {}
    req = {"Адрес объекта", "Статус", "Тип работ"}
    if not req.issubset(set(shops.columns)):
        missing = req - set(shops.columns)
        raise ValueError("В WRS-файле нет колонок: " + ", ".join(missing))

    shops = shops.copy()
    ab_df = ab_df.copy()
    shops["_n"] = shops["Адрес объекта"].fillna("").astype(str).apply(normalize_address)
    ab_df["_n"] = ab_df["Адрес магазина"].apply(normalize_address)
    merged = shops.merge(ab_df[["_n", "Менеджер"]], on="_n", how="left")

    non_auth: Dict[str, List[str]] = {}
    adj_only: Dict[str, List[str]] = {}
    internal_no_adj: List[dict] = []
    nonauth_items: List[dict] = []
    unmatched: List[dict] = []
    seen_u = set()

    total = (
        ab_df.drop_duplicates(subset=["_n"])
        .groupby("Менеджер")["_n"]
        .nunique()
        .to_dict()
    )

    for an, grp in merged.groupby("_n"):
        if cancel and cancel():
            raise RuntimeError("Отменено")
        addr = grp["Адрес объекта"].iloc[0]
        if is_moscow(addr):
            continue
        if is_under_reconstruction(str(an), recon):
            continue
        mgr = grp["Менеджер"].iloc[0]
        if pd.isna(mgr) or not str(mgr).strip():
            if an not in seen_u:
                seen_u.add(an)
                unmatched.append(
                    {
                        "address_full": str(addr),
                        "address_short": shorten_address(str(addr)),
                    }
                )
            continue
        mgr = str(mgr).strip()
        stat = grp["Статус"].fillna("").astype(str)
        na_mask = stat.str.contains("Не авторизован", case=False, na=False)
        wt = grp["Тип работ"]
        im, am = int_mask(wt), adj_mask(wt)
        ia_cnt = int((im & ~na_mask).sum())
        aa_cnt = int((am & ~na_mask).sum())
        sh = shorten_address(str(addr))
        if na_mask.any():
            non_auth.setdefault(mgr, []).append(sh)
            nonauth_items.append(
                {
                    "addr_norm": str(an),
                    "manager": mgr,
                    "address_full": str(addr),
                    "address_short": sh,
                }
            )
        if am.any() and not im.any():
            adj_only.setdefault(mgr, []).append(sh)
        if ia_cnt > 0 and aa_cnt == 0:
            dlines = [
                f"{str(r.get('Тип работ', '')).strip()} | {str(r.get('Статус', '')).strip()}"
                for _, r in grp.iterrows()
            ]
            reason = (
                "Нет строки по прилегающей"
                if not am.any()
                else "Прилегающая не авторизована"
            )
            internal_no_adj.append(
                {
                    "manager": mgr,
                    "region": region_for(mgr, regions),
                    "address_short": sh,
                    "address_full": str(addr),
                    "reason": reason,
                    "internal_auth_count": ia_cnt,
                    "adjacent_auth_count": aa_cnt,
                    "detail_lines": dlines,
                }
            )

    today = date.today().strftime("%Y-%m-%d")
    for a in list(tracker.keys()):
        if isinstance(tracker[a], dict):
            tracker[a]["active_nonauth"] = False
    for it in nonauth_items:
        a = it["addr_norm"]
        rec = tracker.get(a) or {}
        if not rec.get("first_seen"):
            rec["first_seen"] = today
        rec["last_manager"] = it["manager"]
        rec["last_address_short"] = it["address_short"]
        rec["last_address_full"] = it["address_full"]
        ls = str(rec.get("last_seen_nonauth", "") or "")
        if ls != today:
            yd = (
                datetime.strptime(today, "%Y-%m-%d").date() - timedelta(days=1)
            ).strftime("%Y-%m-%d")
            rec["streak_nonauth"] = (
                (int(rec.get("streak_nonauth", 0) or 0) + 1) if ls == yd else 1
            )
            rec["last_seen_nonauth"] = today
        rec["active_nonauth"] = True
        tracker[a] = rec

    repeat = [
        {
            "streak": int(r.get("streak_nonauth", 0)),
            "manager": r.get("last_manager", ""),
            "address_short": r.get("last_address_short", ""),
            "address_full": r.get("last_address_full", ""),
            "last_seen": r.get("last_seen_nonauth", ""),
        }
        for r in tracker.values()
        if isinstance(r, dict)
        and int(r.get("streak_nonauth", 0)) >= 2
        and r.get("active_nonauth")
        and r.get("last_seen_nonauth", "") == today
    ]
    repeat.sort(key=lambda x: x["streak"], reverse=True)

    managers_order = list(
        dict.fromkeys(
            list((regions or {}).get(r, []) for r in (regions or {}))
            and sum((list(ms) for ms in (regions or {}).values()), [])
            or list(total.keys())
        )
    )
    # prefer settings order if available via regions values flattened
    flat = []
    for ms in (regions or {}).values():
        for m in ms:
            if m not in flat:
                flat.append(m)
    for m in total:
        if m not in flat:
            flat.append(m)
    managers_order = flat

    mgr_rows: List[dict] = []
    for mgr in managers_order:
        t = int(total.get(mgr, 0))
        if t <= 0:
            continue
        corr = corrections.get(mgr) or {}
        q = safe_int(corr.get("q", 0))
        nr = safe_int(corr.get("nr", 0))
        tech = safe_int(corr.get("tech", 0))
        na = len(non_auth.get(mgr, []) or [])
        adj = len(adj_only.get(mgr, []) or [])
        sl = sl_manager(mgr, total, non_auth, adj_only, corrections)
        mgr_rows.append(
            {
                "manager": mgr,
                "region": region_for(mgr, regions),
                "total": t,
                "non_auth": na,
                "adj_only": adj,
                "q": q,
                "nr": nr,
                "tech": tech,
                "sl": sl,
                "sl_percent": fmt_sl(sl),
                "non_auth_shops": non_auth.get(mgr, []),
                "adj_only_shops": adj_only.get(mgr, []),
            }
        )

    reg_rows: List[dict] = []
    for region, mgrs in (regions or {}).items():
        slr = sl_region(mgrs, total, non_auth, adj_only, corrections)
        reg_rows.append(
            {"region": region, "sl": slr, "sl_percent": fmt_sl(slr)}
        )

    return {
        "total": total,
        "non_auth": non_auth,
        "adj_only": adj_only,
        "unmatched": unmatched,
        "repeat": repeat,
        "internal": internal_no_adj,
        "nonauth_items": nonauth_items,
        "tracker": tracker,
        "mgr_rows": mgr_rows,
        "region_rows": reg_rows,
        "today": today,
    }



def build_basic_report(result: dict, now_str: str, with_sl: bool = False) -> str:
    """Отчёт по авторизации в магазине (формат как в десктопе)."""
    sep = "─" * 32
    blocks: List[str] = []
    for r in result.get("mgr_rows") or []:
        na = r.get("non_auth", 0)
        adj = r.get("adj_only", 0)
        q, nr, tech = r.get("q", 0), r.get("nr", 0), r.get("tech", 0)
        if not (na > 0 or adj > 0 or q > 0 or nr > 0 or tech > 0):
            continue
        h = f"{r['manager']} ({r['total']} магазинов"
        if na:
            h += f", {na} не авториз."
        if q:
            h += f", {q} Н/к уборка"
        if nr:
            h += f", {nr} не рек."
        if tech:
            h += f", {tech} тех."
        if adj:
            h += f", {adj} только прилег."
        if with_sl and r.get("sl_percent"):
            h += f", SL {r['sl_percent']}"
        h += ")"
        lines = [h]
        if na:
            lines.append("❌ Не авторизован:")
            for a in r.get("non_auth_shops") or []:
                lines.append(f"• {a}")
        if adj:
            lines.append("🌿 Только прилегающая:")
            for a in r.get("adj_only_shops") or []:
                lines.append(f"• {a}")
        blocks.append("\n".join(lines))
    if not blocks:
        body = (
            f"🧹 АВТОРИЗАЦИЯ В МАГАЗИНЕ\nСитуация на {now_str}:\n{sep}\n✅ Проблем нет."
        )
    else:
        body = (
            f"🧹 АВТОРИЗАЦИЯ В МАГАЗИНЕ\nСитуация на {now_str}:\n{sep}\n"
            + f"\n{sep}\n".join(blocks)
        )
    if with_sl:
        reg_lines = []
        for r in result.get("region_rows") or []:
            reg_lines.append(f"{r['region']}: SL = {r['sl_percent']}")
        if reg_lines:
            body += f"\n{sep}\nSL по регионам:\n" + "\n".join(reg_lines)
    return body


def build_short_shifts_block(short_rows: list) -> str:
    """Блок коротких смен для общего отчёта."""
    if not short_rows:
        return ""
    by_mgr: Dict[str, List] = {}
    for r in short_rows:
        mgr = r.get("manager") or "Менеджер не найден"
        by_mgr.setdefault(mgr, []).append(r)
    lines = ["⏱ КОРОТКИЕ СМЕНЫ"]
    for mgr in sorted(by_mgr.keys()):
        items = by_mgr[mgr]
        lines.append(f"{mgr} ({len(items)}):")
        for r in items:
            addr = r.get("address") or ""
            dur = r.get("duration") or ""
            lines.append(f"• {addr} — {dur}")
    return "\n".join(lines)


def build_combined_store_report(result: dict, now_str: str, short_rows: list, with_sl: bool = False) -> str:
    """Авторизация в магазине + короткие смены одним текстом."""
    base = build_basic_report(result, now_str, with_sl=with_sl)
    short = build_short_shifts_block(short_rows or [])
    if short:
        sep = "─" * 32
        return base + "\n" + sep + "\n" + short
    return base


