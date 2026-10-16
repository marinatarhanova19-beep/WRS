"""Short shifts analysis (core logic from desktop)."""
from __future__ import annotations

from datetime import datetime
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from address import (
    is_moscow,
    normalize_address,
    prettify_address,
    shorten_address,
)
from helpers import fmt_hours, parse_hours
from wrs_processor import is_under_reconstruction


def process_shifts(
    df: pd.DataFrame,
    ab_df: pd.DataFrame,
    recon: list,
    cancel: Optional[Callable[[], bool]] = None,
) -> dict:
    req = ["Адрес магазина", "Тип работы", "Длительность смены"]
    miss = [c for c in req if c not in df.columns]
    if miss:
        raise ValueError("Нет колонок: " + ", ".join(miss))

    def pick(cands):
        return next((c for c in cands if c in df.columns), None)

    sc = pick(["Начало работы", "Дата начала", "Начало смены"])
    ec = pick(["Конец работы", "Дата окончания", "Конец смены"])

    mgr_df = ab_df.copy()
    mgr_df["_n"] = mgr_df["Адрес магазина"].apply(normalize_address)
    mgr_map = (
        mgr_df.dropna(subset=["Менеджер"])
        .drop_duplicates(subset=["_n"], keep="first")
        .set_index("_n")["Менеджер"]
        .to_dict()
    )

    wt = df["Тип работы"].astype(str).str.strip().str.lower()
    ca = df[wt == "уборка в магазине"].copy()
    ca["_h"] = ca["Длительность смены"].apply(parse_hours)
    ca["_n"] = ca["Адрес магазина"].apply(normalize_address)
    ca = ca[~ca["Адрес магазина"].apply(is_moscow)].copy()
    ca = ca[~ca["_n"].apply(lambda n: is_under_reconstruction(n, recon))].copy()

    adj = df[wt == "уборка прилегающих территорий"].copy()
    if not adj.empty:
        adj["_n"] = adj["Адрес магазина"].apply(normalize_address)
        adj = adj[~adj["Адрес магазина"].apply(is_moscow)].copy()
        if not adj.empty:
            adj = adj[
                ~adj["_n"].apply(lambda n: is_under_reconstruction(n, recon))
            ].copy()
        if not adj.empty:
            adj["_short"] = adj["Адрес магазина"].apply(shorten_address)
            adj["Менеджер"] = adj["_n"].map(mgr_map).fillna("Менеджер не найден")

    shift_report_date = ""
    if sc and sc in df.columns:
        try:
            dates = df[sc].dropna().apply(lambda x: str(x).strip()[:10])
            most_common = dates.value_counts().idxmax()
            shift_report_date = datetime.strptime(most_common, "%d.%m.%Y").strftime(
                "%d.%m"
            )
        except Exception:
            pass

    def bkey(row):
        sap = str(row.get("SAP магазина", "")).strip()
        num = str(row.get("Номер магазина", "")).strip()
        if sap and sap.lower() != "nan":
            return "sap::" + sap.lower()
        if num and num.lower() != "nan":
            return "num::" + num.lower()
        return "addr::" + row.get("_n", "")

    ca["_key"] = ca.apply(bkey, axis=1)
    ca["_open"] = ca["_h"].isna()
    cl = ca[ca["_h"].notna()].copy()

    agg: Dict[str, Any] = {
        "_h": "sum",
        "Адрес магазина": "first",
        "_n": "first",
    }
    if "SAP магазина" in cl.columns:
        agg["SAP магазина"] = "first"
    if "Номер магазина" in cl.columns:
        agg["Номер магазина"] = "first"

    if not cl.empty:
        st = cl.groupby("_key", as_index=False).agg(agg)
        cnt = cl.groupby("_key").size().rename("_cnt").reset_index()
        st = st.merge(cnt, on="_key", how="left")
        of = ca.groupby("_key")["_open"].max().rename("_open2").reset_index()
        st = st.merge(of, on="_key", how="left")
        st["_open2"] = st["_open2"].fillna(False)
        st["Менеджер"] = st["_n"].map(mgr_map).fillna("Менеджер не найден")
        st["_addr"] = st["Адрес магазина"].apply(prettify_address)
        st["_txt"] = st["_h"].apply(fmt_hours)

        def mst(row):
            if row["_open2"]:
                return "Не в отчете: открытая смена"
            if float(row["_h"]) < 3:
                return "В отчете"
            return "Не в отчете: >= 3 ч"

        st["_status"] = st.apply(mst, axis=1)
        rep = st[st["_status"] == "В отчете"].copy()
    else:
        st = pd.DataFrame()
        rep = st

    short_items: List[dict] = []
    if not rep.empty:
        for _, row in rep.iterrows():
            short_items.append(
                {
                    "address_norm": str(row["_n"]),
                    "manager": row["Менеджер"],
                    "address_short": shorten_address(str(row["Адрес магазина"])),
                    "address_full": str(row["Адрес магазина"]),
                    "duration_text": row["_txt"],
                    "hours": float(row["_h"]),
                }
            )

    report_rows = []
    if not rep.empty:
        for _, row in rep.iterrows():
            report_rows.append(
                {
                    "manager": row["Менеджер"],
                    "address": row["_addr"],
                    "duration": row["_txt"],
                    "hours": float(row["_h"]),
                    "shifts_count": int(row.get("_cnt", 0) or 0),
                    "key": str(row["_key"]),
                }
            )

    return {
        "shift_stores": st,
        "shift_report": rep,
        "shift_details": ca,
        "shift_adj": adj if isinstance(adj, pd.DataFrame) else pd.DataFrame(),
        "shift_ctx": {"sc": sc, "ec": ec, "report_date": shift_report_date},
        "short_items": short_items,
        "report_rows": report_rows,
    }
