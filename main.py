"""WRS Mobile — single-page FastAPI app."""
from __future__ import annotations

import gc
import io
import os
import pickle
import threading
import traceback
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from address import normalize_address
from config import (
    ADDRESS_FILE,
    APP_TITLE,
    APP_VER,
    CORRECTIONS_FILE,
    DATA_DIR,
    DVORNIK_FILE,
    HISTORY_FILE,
    RECON_FILE,
    SETTINGS_FILE,
    TRACKER_FILE,
    UPLOAD_DIR,
)
from dvornik import (
    build_dvornik_report,
    detect_weekday_from_shifts,
    evaluate_dvornik,
    extract_adj_from_dataframe,
    parse_schedule_excel,
)
from excel_export import build_shifts_excel, msk_now
from helpers import fmt_sl, safe_int
from shifts_processor import process_shifts
from storage import default_settings, load_json, now_str, save_json
from wrs_processor import (
    build_basic_report,
    build_combined_store_report,
    process_wrs,
    region_for,
    sl_manager,
    sl_region,
)

app = FastAPI(title=APP_TITLE, version=APP_VER)

BASE = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=BASE)
try:
    app.mount("/static", StaticFiles(directory=BASE), name="static")
except Exception:
    pass

_LAST: Dict[str, Any] = {}
_JOB: Dict[str, Any] = {"running": False, "error": "", "started": ""}
_JOB_LOCK = threading.Lock()
SHIFTS_CACHE = os.path.join(DATA_DIR, "last_shifts.pkl")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


@app.exception_handler(Exception)
async def _unhandled(request: Request, exc: Exception):
    traceback.print_exc()
    return HTMLResponse(
        f"<html><body style='font-family:sans-serif;padding:24px'>"
        f"<h2>Ошибка</h2><pre style='white-space:pre-wrap'>{type(exc).__name__}: {exc}</pre>"
        f"<p><a href='/'>На главную</a></p></body></html>",
        status_code=500,
    )


def get_settings() -> dict:
    data = load_json(SETTINGS_FILE, default_settings())
    if not isinstance(data, dict):
        data = default_settings()
    data.setdefault("managers", default_settings()["managers"])
    data.setdefault("regions", default_settings()["regions"])
    return data


def get_ab() -> dict:
    raw = load_json(ADDRESS_FILE, {"rows": [], "source_file": "", "updated_at": ""})
    if isinstance(raw, list):
        raw = {"rows": raw, "source_file": "", "updated_at": ""}
    if not isinstance(raw, dict):
        raw = {"rows": [], "source_file": "", "updated_at": ""}
    rows = []
    for item in raw.get("rows") or []:
        if not isinstance(item, dict):
            continue
        addr = str(item.get("Адрес магазина") or item.get("address") or "").strip()
        mgr = str(item.get("Менеджер") or item.get("manager") or "").strip()
        if addr:
            rows.append({"Адрес магазина": addr, "Менеджер": mgr})
    return {
        "rows": rows,
        "source_file": raw.get("source_file") or "",
        "updated_at": raw.get("updated_at") or "",
    }


def save_ab(store: dict) -> None:
    save_json(ADDRESS_FILE, store)


def ab_to_df(ab: dict) -> pd.DataFrame:
    rows = ab.get("rows") or []
    if not rows:
        return pd.DataFrame(columns=["Адрес магазина", "Менеджер"])
    return pd.DataFrame(rows)


def get_corrections() -> dict:
    data = load_json(CORRECTIONS_FILE, {})
    return data if isinstance(data, dict) else {}


def save_corrections(data: dict) -> None:
    save_json(CORRECTIONS_FILE, data)


def get_recon() -> list:
    data = load_json(RECON_FILE, [])
    return data if isinstance(data, list) else []


def get_dvornik() -> list:
    data = load_json(DVORNIK_FILE, [])
    return data if isinstance(data, list) else []


def save_dvornik(rows: list) -> None:
    save_json(DVORNIK_FILE, rows)


def append_history(entry: dict) -> None:
    hist = load_json(HISTORY_FILE, [])
    if not isinstance(hist, list):
        hist = []
    hist.insert(0, entry)
    save_json(HISTORY_FILE, hist[:50])


def _save_shifts_cache(payload: dict) -> None:
    try:
        # Serialize DataFrames to records for durable storage
        details = payload.get("shift_details")
        adj = payload.get("shift_adj")
        disk = {
            "shift_ctx": payload.get("shift_ctx") or {},
            "details_records": details.to_dict(orient="records") if isinstance(details, pd.DataFrame) and not details.empty else [],
            "details_columns": list(details.columns) if isinstance(details, pd.DataFrame) else [],
            "adj_records": adj.to_dict(orient="records") if isinstance(adj, pd.DataFrame) and not adj.empty else [],
            "adj_columns": list(adj.columns) if isinstance(adj, pd.DataFrame) else [],
        }
        with open(SHIFTS_CACHE, "wb") as f:
            pickle.dump(disk, f)
        # also keep raw pickle of original for same process
        with open(SHIFTS_CACHE + ".raw", "wb") as f:
            pickle.dump(payload, f)
    except Exception:
        traceback.print_exc()


def _load_shifts_cache() -> Optional[dict]:
    # prefer in-memory
    full = _LAST.get("shifts_full")
    if full and full.get("shift_details") is not None:
        return full
    # raw pickle
    try:
        if os.path.exists(SHIFTS_CACHE + ".raw"):
            with open(SHIFTS_CACHE + ".raw", "rb") as f:
                return pickle.load(f)
    except Exception:
        pass
    # records pickle
    try:
        if os.path.exists(SHIFTS_CACHE):
            with open(SHIFTS_CACHE, "rb") as f:
                disk = pickle.load(f)
            details = pd.DataFrame(disk.get("details_records") or [], columns=disk.get("details_columns") or None)
            adj = pd.DataFrame(disk.get("adj_records") or [], columns=disk.get("adj_columns") or None)
            return {
                "shift_details": details,
                "shift_adj": adj,
                "shift_ctx": disk.get("shift_ctx") or {},
            }
    except Exception:
        traceback.print_exc()
    return None


def page_ctx(request: Request, error: str = "") -> dict:
    settings = get_settings()
    ab = get_ab()
    managers = settings.get("managers") or []
    regions = settings.get("regions") or {}
    corr = get_corrections()
    wrs = _LAST.get("wrs") or {}
    shifts = _LAST.get("shifts") or {}
    dv = _LAST.get("dvornik") or {}
    schedule = get_dvornik()

    corr_items = []
    total = wrs.get("total") or {}
    non_auth = wrs.get("non_auth") or {}
    adj_only = wrs.get("adj_only") or {}
    for m in managers:
        c = corr.get(m) or {}
        row = next((r for r in (wrs.get("mgr_rows") or []) if r.get("manager") == m), None)
        corr_items.append(
            {
                "manager": m,
                "region": region_for(m, regions),
                "total": (row or {}).get("total") or total.get(m) or 0,
                "q": int(c.get("q") or 0),
                "nr": int(c.get("nr") or 0),
                "tech": int(c.get("tech") or 0),
                "sl_percent": (row or {}).get("sl_percent") or "—",
            }
        )

    has_results = bool(wrs or shifts or dv or _LAST.get("basic_report"))
    job_running = bool(_JOB.get("running"))
    return {
        "request": request,
        "title": APP_TITLE,
        "ver": APP_VER,
        "error": error,
        "has_results": has_results,
        "job_running": job_running,
        "job_error": _JOB.get("error") or "",
        "has_excel": bool(_load_shifts_cache()),
        "report_date": (shifts.get("shift_ctx") or {}).get("report_date") or "",
        "generated_at": _LAST.get("generated_at") or "",
        "kpi": _LAST.get("kpi") or {},
        "ab_count": len(ab.get("rows") or []),
        "ab_updated": ab.get("updated_at") or "",
        "schedule_count": len(schedule),
        "wrs_report": _LAST.get("basic_report") or "",
        "wrs_report_sl": _LAST.get("basic_report_sl") or "",
        "wrs_mgr_rows": wrs.get("mgr_rows") or [],
        "wrs_region_rows": wrs.get("region_rows") or [],
        "dv_day": dv.get("day_label") or shifts.get("day_label") or "",
        "dv_n_total": dv.get("n_total"),
        "dv_n_worked": dv.get("n_worked"),
        "dv_n_nonauth": dv.get("n_nonauth"),
        "dv_nonauth": dv.get("nonauth") or [],
        "dv_mgr_rows": dv.get("mgr_rows") or [],
        "dv_region_rows": dv.get("region_rows") or [],
        "dv_report": dv.get("report_text") or "",
        "dv_report_sl": dv.get("report_text_sl") or "",
        "short_rows": shifts.get("report_rows") or [],
        "short_ctx": shifts.get("shift_ctx") or {},
        "corr_items": corr_items,
        "managers": managers,
        "regions": regions,
    }


def _read_upload(upload: Optional[UploadFile]) -> bytes:
    """Read upload bytes; treat empty/missing as b''."""
    if upload is None:
        return b""
    name = getattr(upload, "filename", None) or ""
    if not str(name).strip():
        return b""
    # sync/async: in FastAPI UploadFile.read is async
    return b""  # placeholder — not used


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", page_ctx(request))




def _save_upload(name: str, data: bytes) -> str:
    path = os.path.join(UPLOAD_DIR, name)
    with open(path, "wb") as f:
        f.write(data)
    return path


def _read_excel(path: str, usecols=None) -> pd.DataFrame:
    kwargs: dict = {"engine": "openpyxl"}
    if usecols:
        kwargs["usecols"] = usecols
    return pd.read_excel(path, **kwargs)


def _run_pipeline(ab_bytes: bytes, wrs_bytes: bytes, sh_bytes: bytes, dv_bytes: bytes,
                  ab_name: str, wrs_name: str, sh_name: str) -> None:
    """Обработка с минимальной памятью: файлы на диск, DF по одному, gc."""
    global _JOB
    errors: List[str] = []
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        ab_path = wrs_path = sh_path = dv_path = ""

        if ab_bytes:
            ab_path = _save_upload("ab.xlsx", ab_bytes)
        if wrs_bytes:
            wrs_path = _save_upload("wrs.xlsx", wrs_bytes)
        if sh_bytes:
            sh_path = _save_upload("shifts.xlsx", sh_bytes)
        if dv_bytes:
            dv_path = _save_upload("dvornik.xlsx", dv_bytes)

        # освобождаем RAM от сырых байтов
        del ab_bytes, wrs_bytes, sh_bytes, dv_bytes
        gc.collect()

        ab_store = get_ab()
        if ab_path:
            ab_df_raw = _read_excel(ab_path)
            # fallback if usecols lambda fails
            if "Адрес магазина" not in ab_df_raw.columns or "Менеджер" not in ab_df_raw.columns:
                ab_df_raw = _read_excel(ab_path)
            req = {"Адрес магазина", "Менеджер"}
            if not req.issubset(set(ab_df_raw.columns)):
                raise ValueError("Адресная книга: нужны колонки «Адрес магазина», «Менеджер»")
            clean = ab_df_raw.loc[:, ["Адрес магазина", "Менеджер"]].copy()
            del ab_df_raw
            clean["Адрес магазина"] = clean["Адрес магазина"].fillna("").astype(str).str.strip()
            clean["Менеджер"] = clean["Менеджер"].fillna("").astype(str).str.strip()
            clean = clean.loc[clean["Адрес магазина"] != ""].copy()
            clean["_n"] = clean["Адрес магазина"].map(normalize_address)
            rows = (
                clean.drop_duplicates(subset=["_n"], keep="last")[["Адрес магазина", "Менеджер"]]
                .to_dict("records")
            )
            del clean
            ab_store = {"rows": rows, "source_file": ab_name, "updated_at": now_str()}
            save_ab(ab_store)
            gc.collect()

        if not (ab_store.get("rows") or []):
            raise ValueError("Нет адресной книги. Загрузите её один раз.")

        ab_df = ab_to_df(ab_store)
        settings = get_settings()
        corr = get_corrections()
        recon = get_recon()
        regions = settings.get("regions") or {}
        _LAST.clear()

        if dv_path:
            try:
                dv_df = _read_excel(dv_path)
                save_dvornik(parse_schedule_excel(dv_df))
                del dv_df
                gc.collect()
            except Exception as e:
                errors.append(f"График дворников: {e}")
        schedule = get_dvornik()

        wrs_result: dict = {}
        try:
            wrs_df = _read_excel(wrs_path)
            wrs_result = process_wrs(wrs_df, ab_df, recon, regions, {}, corrections=corr)
            _LAST["wrs"] = wrs_result
            total_na = sum(len(v) for v in (wrs_result.get("non_auth") or {}).values())
            total_adj = sum(len(v) for v in (wrs_result.get("adj_only") or {}).values())
            shops_count = int(ab_df["Адрес магазина"].map(normalize_address).nunique())
            _LAST["kpi"] = {"shops": shops_count, "nonauth": total_na, "adj": total_adj}
            del wrs_df
            gc.collect()
        except Exception as e:
            errors.append(f"WRS: {e}")
            traceback.print_exc()
            _LAST["wrs"] = {}

        shifts_light: dict = {}
        try:
            sh_df = _read_excel(sh_path)
            sh_result = process_shifts(sh_df, ab_df, recon)
            sc = (sh_result.get("shift_ctx") or {}).get("sc")
            try:
                weekday, day_from_shifts = detect_weekday_from_shifts(sh_df, sc)
            except Exception:
                weekday, day_from_shifts = datetime.now().weekday(), now_str()[:10]
            day_label = (sh_result.get("shift_ctx") or {}).get("report_date") or day_from_shifts
            adj_src = sh_result.get("shift_adj")
            if adj_src is None or (hasattr(adj_src, "empty") and adj_src.empty):
                adj_rows = extract_adj_from_dataframe(sh_df)
            else:
                adj_rows = extract_adj_from_dataframe(adj_src)
            shifts_light = {
                "report_rows": sh_result.get("report_rows") or [],
                "short_items": sh_result.get("short_items") or [],
                "shift_ctx": sh_result.get("shift_ctx") or {},
                "adj_rows": adj_rows,
                "weekday": weekday,
                "day_label": day_label,
            }
            _LAST["shifts"] = shifts_light
            full = {
                "shift_details": sh_result.get("shift_details"),
                "shift_adj": sh_result.get("shift_adj"),
                "shift_ctx": sh_result.get("shift_ctx") or {},
            }
            _LAST["shifts_full"] = full
            _LAST["generated_at"] = now_str()
            _save_shifts_cache(full)
            del sh_df, sh_result
            gc.collect()
        except Exception as e:
            errors.append(f"Смены: {e}")
            traceback.print_exc()
            _LAST["shifts"] = {}

        short_rows = (_LAST.get("shifts") or {}).get("report_rows") or []
        if wrs_result:
            try:
                _LAST["basic_report"] = build_combined_store_report(
                    wrs_result, now_str(), short_rows, with_sl=False
                )
                _LAST["basic_report_sl"] = build_combined_store_report(
                    wrs_result, now_str(), short_rows, with_sl=True
                )
            except Exception as e:
                errors.append(f"Отчёт: {e}")
                _LAST["basic_report"] = build_basic_report(wrs_result, now_str(), False)
                _LAST["basic_report_sl"] = build_basic_report(wrs_result, now_str(), True)

        if schedule and shifts_light:
            try:
                weekday = shifts_light.get("weekday")
                if weekday is None:
                    weekday = datetime.now().weekday()
                day_label = shifts_light.get("day_label") or now_str()[:10]
                dv_res = evaluate_dvornik(
                    schedule,
                    shifts_light.get("adj_rows") or [],
                    int(weekday),
                    ab_df,
                    regions=regions,
                )
                dv_res["day_label"] = day_label
                dv_res["report_text"] = build_dvornik_report(dv_res, day_label, with_sl=False)
                dv_res["report_text_sl"] = build_dvornik_report(dv_res, day_label, with_sl=True)
                _LAST["dvornik"] = dv_res
            except Exception as e:
                errors.append(f"Дворники: {e}")
                traceback.print_exc()
                _LAST["dvornik"] = {}
        else:
            _LAST["dvornik"] = {}

        append_history(
            {
                "at": now_str(),
                "file_ab": ab_name,
                "file_wrs": wrs_name,
                "file_shifts": sh_name,
                "errors": errors,
            }
        )
        with _JOB_LOCK:
            _JOB["running"] = False
            _JOB["error"] = "; ".join(errors) if errors else ""
    except Exception as e:
        traceback.print_exc()
        with _JOB_LOCK:
            _JOB["running"] = False
            _JOB["error"] = str(e)
    finally:
        gc.collect()


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", page_ctx(request))



def _run_pipeline(ab_bytes: bytes, wrs_bytes: bytes, sh_bytes: bytes, dv_bytes: bytes,
                  ab_name: str, wrs_name: str, sh_name: str) -> None:
    """Тяжёлая обработка в фоне — чтобы Render не отдавал 502 по таймауту."""
    global _JOB
    errors: List[str] = []
    try:
        ab_store = get_ab()
        if ab_bytes:
            ab_df_raw = pd.read_excel(io.BytesIO(ab_bytes), engine="openpyxl")
            req = {"Адрес магазина", "Менеджер"}
            if not req.issubset(set(ab_df_raw.columns)):
                raise ValueError("Адресная книга: нужны колонки «Адрес магазина», «Менеджер»")
            clean = ab_df_raw[["Адрес магазина", "Менеджер"]].copy()
            clean["Адрес магазина"] = clean["Адрес магазина"].fillna("").astype(str).str.strip()
            clean["Менеджер"] = clean["Менеджер"].fillna("").astype(str).str.strip()
            clean = clean[clean["Адрес магазина"] != ""].copy()
            clean.loc[:, "_n"] = clean["Адрес магазина"].map(normalize_address)
            rows = (
                clean.drop_duplicates(subset=["_n"], keep="last")[["Адрес магазина", "Менеджер"]]
                .to_dict("records")
            )
            ab_store = {
                "rows": rows,
                "source_file": ab_name,
                "updated_at": now_str(),
            }
            save_ab(ab_store)
            del ab_df_raw, clean

        if not (ab_store.get("rows") or []):
            raise ValueError("Нет адресной книги. Загрузите её один раз.")

        ab_df = ab_to_df(ab_store)
        settings = get_settings()
        corr = get_corrections()
        recon = get_recon()
        regions = settings.get("regions") or {}

        _LAST.clear()

        if dv_bytes:
            try:
                dv_df = pd.read_excel(io.BytesIO(dv_bytes), engine="openpyxl")
                save_dvornik(parse_schedule_excel(dv_df))
                del dv_df
            except Exception as e:
                errors.append(f"График дворников: {e}")
        schedule = get_dvornik()

        wrs_result: dict = {}
        try:
            wrs_df = pd.read_excel(io.BytesIO(wrs_bytes), engine="openpyxl")
            wrs_result = process_wrs(wrs_df, ab_df, recon, regions, {}, corrections=corr)
            _LAST["wrs"] = wrs_result
            total_na = sum(len(v) for v in (wrs_result.get("non_auth") or {}).values())
            total_adj = sum(len(v) for v in (wrs_result.get("adj_only") or {}).values())
            shops_count = int(ab_df["Адрес магазина"].map(normalize_address).nunique())
            _LAST["kpi"] = {"shops": shops_count, "nonauth": total_na, "adj": total_adj}
            del wrs_df
        except Exception as e:
            errors.append(f"WRS: {e}")
            traceback.print_exc()
            _LAST["wrs"] = {}

        shifts_light: dict = {}
        try:
            sh_df = pd.read_excel(io.BytesIO(sh_bytes), engine="openpyxl")
            sh_result = process_shifts(sh_df, ab_df, recon)
            sc = (sh_result.get("shift_ctx") or {}).get("sc")
            try:
                weekday, day_from_shifts = detect_weekday_from_shifts(sh_df, sc)
            except Exception:
                weekday, day_from_shifts = datetime.now().weekday(), now_str()[:10]
            day_label = (sh_result.get("shift_ctx") or {}).get("report_date") or day_from_shifts
            adj_src = sh_result.get("shift_adj")
            if adj_src is None or (hasattr(adj_src, "empty") and adj_src.empty):
                adj_src = sh_df
            adj_rows = extract_adj_from_dataframe(adj_src)
            shifts_light = {
                "report_rows": sh_result.get("report_rows") or [],
                "short_items": sh_result.get("short_items") or [],
                "shift_ctx": sh_result.get("shift_ctx") or {},
                "adj_rows": adj_rows,
                "weekday": weekday,
                "day_label": day_label,
            }
            _LAST["shifts"] = shifts_light
            full = {
                "shift_details": sh_result.get("shift_details"),
                "shift_adj": sh_result.get("shift_adj"),
                "shift_ctx": sh_result.get("shift_ctx") or {},
            }
            _LAST["shifts_full"] = full
            _LAST["generated_at"] = now_str()
            _save_shifts_cache(full)
            del sh_df
        except Exception as e:
            errors.append(f"Смены: {e}")
            traceback.print_exc()
            _LAST["shifts"] = {}

        short_rows = (_LAST.get("shifts") or {}).get("report_rows") or []
        if wrs_result:
            try:
                _LAST["basic_report"] = build_combined_store_report(
                    wrs_result, now_str(), short_rows, with_sl=False
                )
                _LAST["basic_report_sl"] = build_combined_store_report(
                    wrs_result, now_str(), short_rows, with_sl=True
                )
            except Exception as e:
                errors.append(f"Отчёт: {e}")
                _LAST["basic_report"] = build_basic_report(wrs_result, now_str(), False)
                _LAST["basic_report_sl"] = build_basic_report(wrs_result, now_str(), True)

        if schedule and shifts_light:
            try:
                weekday = shifts_light.get("weekday")
                if weekday is None:
                    weekday = datetime.now().weekday()
                day_label = shifts_light.get("day_label") or now_str()[:10]
                dv_res = evaluate_dvornik(
                    schedule,
                    shifts_light.get("adj_rows") or [],
                    int(weekday),
                    ab_df,
                    regions=regions,
                )
                dv_res["day_label"] = day_label
                dv_res["report_text"] = build_dvornik_report(dv_res, day_label, with_sl=False)
                dv_res["report_text_sl"] = build_dvornik_report(dv_res, day_label, with_sl=True)
                _LAST["dvornik"] = dv_res
            except Exception as e:
                errors.append(f"Дворники: {e}")
                traceback.print_exc()
                _LAST["dvornik"] = {}
        else:
            _LAST["dvornik"] = {}

        append_history(
            {
                "at": now_str(),
                "file_ab": ab_name,
                "file_wrs": wrs_name,
                "file_shifts": sh_name,
                "errors": errors,
            }
        )
        with _JOB_LOCK:
            _JOB["running"] = False
            _JOB["error"] = "; ".join(errors) if errors else ""
    except Exception as e:
        traceback.print_exc()
        with _JOB_LOCK:
            _JOB["running"] = False
            _JOB["error"] = str(e)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", page_ctx(request))


@app.head("/")
async def index_head():
    return Response(status_code=200)


@app.post("/process")
async def process_all(
    request: Request,
    file_ab: Optional[UploadFile] = File(None),
    file_wrs: Optional[UploadFile] = File(None),
    file_shifts: Optional[UploadFile] = File(None),
    file_dvornik: Optional[UploadFile] = File(None),
):
    """Принимает файлы быстро и считает отчёты в фоне (защита от 502 на Render)."""

    async def read_file(uf: Optional[UploadFile]) -> bytes:
        if uf is None:
            return b""
        fn = (getattr(uf, "filename", None) or "").strip()
        if not fn:
            return b""
        try:
            return (await uf.read()) or b""
        except Exception:
            return b""

    with _JOB_LOCK:
        if _JOB.get("running"):
            return templates.TemplateResponse(
                "index.html",
                page_ctx(request, error="Уже идёт обработка — подождите несколько секунд и обновите страницу."),
            )

    ab_bytes = await read_file(file_ab)
    wrs_bytes = await read_file(file_wrs)
    sh_bytes = await read_file(file_shifts)
    dv_bytes = await read_file(file_dvornik)

    if not wrs_bytes:
        return templates.TemplateResponse(
            "index.html", page_ctx(request, error="Загрузите файл WRS."), status_code=400
        )
    if not sh_bytes:
        return templates.TemplateResponse(
            "index.html", page_ctx(request, error="Загрузите файл смен."), status_code=400
        )
    if not ab_bytes and not (get_ab().get("rows") or []):
        return templates.TemplateResponse(
            "index.html",
            page_ctx(request, error="Загрузите адресную книгу (в первый раз обязательно)."),
            status_code=400,
        )

    ab_name = (file_ab.filename if file_ab and file_ab.filename else "") or ""
    wrs_name = (file_wrs.filename if file_wrs and file_wrs.filename else "") or ""
    sh_name = (file_shifts.filename if file_shifts and file_shifts.filename else "") or ""

    with _JOB_LOCK:
        _JOB["running"] = True
        _JOB["error"] = ""
        _JOB["started"] = now_str()

    th = threading.Thread(
        target=_run_pipeline,
        args=(ab_bytes, wrs_bytes, sh_bytes, dv_bytes, ab_name, wrs_name, sh_name),
        daemon=True,
    )
    th.start()

    # сразу отдаём страницу «считается…» — соединение не висит 30+ секунд
    return templates.TemplateResponse("index.html", page_ctx(request))


@app.post("/corrections/save")
async def corrections_save(
    request: Request,
    manager: str = Form(...),
    q: int = Form(0),
    nr: int = Form(0),
    tech: int = Form(0),
):
    data = get_corrections()
    data[manager] = {"q": max(0, int(q)), "nr": max(0, int(nr)), "tech": max(0, int(tech))}
    save_corrections(data)

    last = _LAST.get("wrs")
    if last:
        settings = get_settings()
        regions = settings.get("regions") or {}
        total = last.get("total") or {}
        non_auth = last.get("non_auth") or {}
        adj_only = last.get("adj_only") or {}
        managers = settings.get("managers") or list(total.keys())
        mgr_rows = []
        for m in managers:
            t = int(total.get(m, 0) or 0)
            if t <= 0:
                continue
            sl = sl_manager(m, total, non_auth, adj_only, data)
            mgr_rows.append(
                {
                    "manager": m,
                    "region": region_for(m, regions),
                    "total": t,
                    "non_auth": len(non_auth.get(m, []) or []),
                    "adj_only": len(adj_only.get(m, []) or []),
                    "q": safe_int((data.get(m) or {}).get("q", 0)),
                    "nr": safe_int((data.get(m) or {}).get("nr", 0)),
                    "tech": safe_int((data.get(m) or {}).get("tech", 0)),
                    "sl": sl,
                    "sl_percent": fmt_sl(sl),
                    "non_auth_shops": non_auth.get(m, []),
                    "adj_only_shops": adj_only.get(m, []),
                }
            )
        reg_rows = []
        for region, mgrs in regions.items():
            slr = sl_region(mgrs, total, non_auth, adj_only, data)
            reg_rows.append({"region": region, "sl": slr, "sl_percent": fmt_sl(slr)})
        last["mgr_rows"] = mgr_rows
        last["region_rows"] = reg_rows
        last["corrections"] = data
        _LAST["wrs"] = last
        short_rows = (_LAST.get("shifts") or {}).get("report_rows") or []
        _LAST["basic_report"] = build_combined_store_report(last, now_str(), short_rows, False)
        _LAST["basic_report_sl"] = build_combined_store_report(last, now_str(), short_rows, True)

    return RedirectResponse("/", status_code=303)


@app.get("/export/shifts")
async def export_shifts():
    full = _load_shifts_cache()
    if not full or full.get("shift_details") is None:
        return HTMLResponse(
            "<h3>Нет данных для Excel</h3>"
            "<p>Сначала загрузите WRS + смены и нажмите «Сформировать отчёты».</p>"
            "<p><a href='/'>На главную</a></p>",
            status_code=400,
        )
    details = full.get("shift_details")
    try:
        empty = getattr(details, "empty", True)
    except Exception:
        empty = True
    if empty:
        return HTMLResponse(
            "<h3>Пустые смены</h3><p>В файле смен нет строк.</p><p><a href='/'>На главную</a></p>",
            status_code=400,
        )
    ab_df = ab_to_df(get_ab())
    settings = get_settings()
    try:
        data = build_shifts_excel(
            details,
            full.get("shift_adj"),
            ab_df,
            full.get("shift_ctx") or {},
            manager_order=settings.get("managers") or [],
            regions=settings.get("regions") or {},
        )
    except Exception as e:
        traceback.print_exc()
        return HTMLResponse(
            f"<h3>Ошибка Excel</h3><pre>{e}</pre><p><a href='/'>На главную</a></p>",
            status_code=500,
        )
    rd = (full.get("shift_ctx") or {}).get("report_date") or msk_now().strftime("%d.%m.%Y")
    fname = f"smeny_{str(rd).replace('.', '-')}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{fname}"',
            "Cache-Control": "no-store",
        },
    )


@app.get("/health")
async def health():
    return {"status": "ok", "version": APP_VER, "data_dir": DATA_DIR}


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
