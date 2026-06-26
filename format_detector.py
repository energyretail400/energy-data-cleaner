"""
format_detector.py - Auto-detect energy meter file format from column headers.

Returns a cfg dict compatible with clean_data.process_format().
Returns None when the format cannot be identified.
"""

import csv
import io
import re
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _file_type(filename: str) -> str:
    return {".xlsx": "xlsx", ".xls": "xlsx", ".xlsb": "xlsb", ".tsv": "tsv"}.get(
        Path(filename).suffix.lower(), "csv"
    )


def _read_sample(file_bytes: bytes, filename: str, nrows: int = 6, header_row: int = 0):
    """Return (df, raw_rows).  df may be None on read error.  raw_rows is CSV only."""
    suffix = Path(filename).suffix.lower()
    raw_rows = None
    try:
        if suffix in (".xlsx",):
            df = pd.read_excel(
                io.BytesIO(file_bytes), dtype=object,
                nrows=nrows, engine="openpyxl", header=header_row,
            )
            return df, raw_rows
        if suffix in (".xls",):
            df = pd.read_excel(
                io.BytesIO(file_bytes), dtype=object,
                nrows=nrows, engine="xlrd", header=header_row,
            )
            return df, raw_rows
        if suffix in (".xlsb", ".tsv"):
            df = pd.read_csv(
                io.BytesIO(file_bytes), sep="\t", dtype=object,
                nrows=nrows, comment="#", encoding="utf-8", errors="replace",
                header=header_row,
            )
            return df, raw_rows
        # CSV / unknown: also capture raw rows for NEM12 marker check
        text = file_bytes.decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(text))
        raw_rows = [row for i, row in enumerate(reader) if i < nrows + 2]
        df = pd.read_csv(
            io.StringIO(text), dtype=object, nrows=nrows, header=header_row,
        )
        return df, raw_rows
    except Exception:
        return None, raw_rows


def _norm_map(df: pd.DataFrame) -> dict:
    """Return {normalised_lower_stripped: original_col_name}."""
    result = {}
    for c in df.columns:
        key = str(c).lower().strip().lstrip("﻿")
        result[key] = str(c)
    return result


def _get(nm: dict, *candidates: str) -> Optional[str]:
    """Return the original column name for the first matching candidate (case-insensitive)."""
    for c in candidates:
        found = nm.get(c.lower().strip())
        if found is not None:
            return found
    return None


def _contains(nm: dict, fragment: str) -> Optional[str]:
    """Return the first original column name whose normalised form contains fragment."""
    frag = fragment.lower()
    for key, orig in nm.items():
        if frag in key:
            return orig
    return None


def _extract_nmi(filename: str) -> str:
    stem = Path(filename).stem
    m = re.search(r"[Nn]\d{9}|\d{10}", stem)
    return m.group(0)[:10].upper() if m else ""


def _is_ampm(df: pd.DataFrame, col: str) -> bool:
    """True if the datetime column values look like AM/PM (12-hour clock)."""
    if col not in df.columns:
        return False
    for val in df[col].dropna().head(5):
        if re.search(r"\b(am|pm)\b", str(val), re.IGNORECASE):
            return True
    return False


def _is_nem12_xlsx(file_bytes: bytes) -> bool:
    """Peek at an xlsx to see if its first two cells are '100' and 'NEM12'."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True, max_row=2))
        wb.close()
        if rows and rows[0]:
            row0 = [str(v).strip() for v in rows[0] if v is not None]
            if len(row0) >= 2 and row0[0] == "100" and row0[1] == "NEM12":
                return True
        return False
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def detect_format(file_bytes: bytes, filename: str) -> Optional[dict]:
    """
    Detect energy meter format from file bytes + filename.
    Returns a cfg dict for clean_data.process_format(), or None if undetected.
    """
    ft = _file_type(filename)

    # ── NEM12 xlsx ──────────────────────────────────────────────────────────
    if ft == "xlsx" and _is_nem12_xlsx(file_bytes):
        return {
            "file_type": "xlsx", "layout": "nem12", "interval_minutes": 30,
            "datetime_col": "_nem12_", "ce_col": "_nem12_",
        }

    df, raw_rows = _read_sample(file_bytes, filename)

    # ── NEM12 CSV ────────────────────────────────────────────────────────────
    if raw_rows:
        for row in raw_rows[:3]:
            if (len(row) >= 2
                    and str(row[0]).strip() == "100"
                    and str(row[1]).strip() == "NEM12"):
                return {
                    "file_type": ft, "layout": "nem12", "interval_minutes": 30,
                    "datetime_col": "_nem12_", "ce_col": "_nem12_",
                }

    # ── Wide daily CSV: NMI/metadata rows + Date/Time row + HH:MM interval cols
    if raw_rows and len(raw_rows) > 3:
        r0, r3 = raw_rows[0], raw_rows[3]
        if (r0 and str(r0[0]).strip() == "NMI"
                and r3 and str(r3[0]).strip() == "Date/Time"
                and len(r3) > 2
                and re.match(r"^\d{1,2}:\d{2}$", str(r3[1]).strip())):
            nmi = str(r0[1]).strip()[:10] if len(r0) > 1 else ""
            return {
                "file_type": ft,
                "layout": "wide_daily",
                "const_nmi": nmi,
                "datetime_col": "Date/Time",
                "datetime_kind": "string_start",
                "ce_col": "_wide_daily_",
                "header_row": 3,
                "interval_minutes": 30,
            }

    if df is None or df.empty:
        return None

    nm = _norm_map(df)
    cols = set(nm.keys())

    # ── 1. Gentrack standard (nmi + endtime + kwh) ───────────────────────────
    if {"nmi", "endtime", "kwh"}.issubset(cols):
        cfg = {
            "file_type": ft,
            "nmi_col":      _get(nm, "nmi"),
            "datetime_col": _get(nm, "endtime"),
            "datetime_kind": "string_end",
            "ce_col":       _get(nm, "kwh"),
            "interval_minutes": 30,
        }
        soe = _get(nm, "generated_kwh")
        if soe:
            cfg["soe_col"] = soe
        return cfg

    # ── 2. Gentrack no-NMI (endtime + kwh, NMI from meter_serial or filename) ─
    if {"endtime", "kwh"}.issubset(cols) and "nmi" not in cols:
        cfg = {
            "file_type": ft,
            "datetime_col": _get(nm, "endtime"),
            "datetime_kind": "string_end",
            "ce_col":       _get(nm, "kwh"),
            "interval_minutes": 30,
        }
        meter_nmi = _get(nm, "meter_serial", "meternmi", "meter nmi")
        if meter_nmi:
            cfg["nmi_col"] = meter_nmi
        else:
            cfg["const_nmi"] = _extract_nmi(filename)
        soe = _get(nm, "generated_kwh")
        if soe:
            cfg["soe_col"] = soe
        return cfg

    # ── 3. Origin / Retailer  ReadingDateTime + E (Usage kWh) ───────────────
    rdt_col  = _contains(nm, "readingdatetime")
    e_use_col = _contains(nm, "e (usage")
    b_gen_col = _contains(nm, "b (generation")
    if rdt_col and e_use_col:
        interval = 5 if _is_ampm(df, rdt_col) else 30
        cfg = {
            "file_type": ft,
            "datetime_col":  rdt_col,
            "datetime_kind": "string_start",
            "ce_col":        e_use_col,
            "interval_minutes": interval,
        }
        nmi_col = _get(nm, "nmi")
        if nmi_col:
            cfg["nmi_col"] = nmi_col
        else:
            cfg["const_nmi"] = _extract_nmi(filename)
        if b_gen_col:
            cfg["soe_col"] = b_gen_col
        return cfg

    # ── 4. Citipower  NMI + Read Datetime + E + B ───────────────────────────
    read_dt = _get(nm, "read datetime", "read_datetime")
    if read_dt and _get(nm, "nmi") and _get(nm, "e") and _get(nm, "b"):
        return {
            "file_type": ft,
            "nmi_col":      _get(nm, "nmi"),
            "datetime_col": read_dt,
            "datetime_kind": "string_start",
            "ce_col":       _get(nm, "e"),
            "soe_col":      _get(nm, "b"),
            "interval_minutes": 30,
        }

    # ── 5. Citipower  NMI + Date (or Datetime) + E + B ──────────────────────
    if _get(nm, "nmi") and _get(nm, "e") and _get(nm, "b"):
        dt_col = _get(nm, "datetime", "date", "starttime", "start time", "period")
        if dt_col:
            return {
                "file_type": ft,
                "nmi_col":      _get(nm, "nmi"),
                "datetime_col": dt_col,
                "datetime_kind": "string_start",
                "ce_col":       _get(nm, "e"),
                "soe_col":      _get(nm, "b"),
                "interval_minutes": 30,
            }

    # ── 6. Long: Nmi / Datetime / Meter Type / Active ────────────────────────
    active_col = _get(nm, "active")
    if _get(nm, "nmi") and _get(nm, "datetime") and _get(nm, "meter type") and active_col:
        return {
            "file_type": ft,
            "layout":        "long",
            "nmi_col":       _get(nm, "nmi"),
            "datetime_col":  _get(nm, "datetime"),
            "datetime_kind": "string_start",
            "type_col":      _get(nm, "meter type"),
            "ce_type_value":  "E1",
            "soe_type_value": "B1",
            "ce_col":        active_col,
            "interval_minutes": 5,
        }

    # ── 7. Long: Nmi / Date / Time / Meter Type ──────────────────────────────
    if _get(nm, "nmi") and _get(nm, "date") and _get(nm, "time") and _get(nm, "meter type"):
        act = _get(nm, "active") or _contains(nm, "active")
        return {
            "file_type": ft,
            "layout":        "long",
            "nmi_col":       _get(nm, "nmi"),
            "datetime_col":  _get(nm, "date"),
            "time_col":      _get(nm, "time"),
            "datetime_kind": "string_start",
            "type_col":      _get(nm, "meter type"),
            "ce_type_value":  "E1",
            "soe_type_value": "B1",
            "ce_col":        act,
            "interval_minutes": 5,
        }

    # ── Usage from Grid / Feed In (multi-section xlsx, real header at row 1) ──
    if _contains(nm, "usage from the grid") and _contains(nm, "feed in"):
        df1, _ = _read_sample(file_bytes, filename, nrows=8, header_row=1)
        if df1 is not None and not df1.empty:
            return {
                "file_type": ft,
                "header_row": 1,
                "data_skip_rows": 4,
                "const_nmi": _extract_nmi(filename) or Path(filename).stem[:10],
                "datetime_col": str(df1.columns[0]),
                "datetime_kind": "string_start",
                "ce_col": "Active Amt (kWh)",
                "soe_col": "Active Amt (kWh).1",
                "interval_minutes": 30,
            }

    # ── 8. Alba Thermal / Usage Data CSV  (Site + From Date + Direction + Unit + Value) ──
    from_dt = _get(nm, "from date", "from_date")
    if _get(nm, "site") and from_dt and _get(nm, "direction") and _get(nm, "unit") and _get(nm, "value"):
        cfg = {
            "file_type": ft,
            "layout":        "long",
            "filter_col":    _get(nm, "unit"),
            "filter_value":  "kWh",
            "nmi_col":       _get(nm, "site"),
            "datetime_col":  from_dt,
            "datetime_kind": "string_start",
            "type_col":      _get(nm, "direction"),
            "ce_type_value":  "I",
            "soe_type_value": "E",
            "ce_col":        _get(nm, "value"),
            "interval_minutes": 30,
        }
        from_tm = _get(nm, "from time", "from_time")
        if from_tm:
            cfg["time_col"] = from_tm
        return cfg

    # ── Direction + Interval Date/Time + Amount + UOM (long, IMPORT/EXPORT) ──
    int_dt_col = _contains(nm, "interval date")
    if int_dt_col and _get(nm, "direction") and _get(nm, "amount") and _get(nm, "uom"):
        return {
            "file_type": ft,
            "layout": "long",
            "filter_col": _get(nm, "uom"),
            "filter_value": "KWH",
            "const_nmi": _extract_nmi(filename) or Path(filename).stem[:10],
            "datetime_col": int_dt_col,
            "datetime_kind": "string_end",
            "type_col": _get(nm, "direction"),
            "ce_type_value": "IMPORT",
            "soe_type_value": "EXPORT",
            "ce_col": _get(nm, "amount"),
            "interval_minutes": 30,
        }

    # ── 9. EndDate + EndTime + Consumption ───────────────────────────────────
    if _get(nm, "enddate") and _get(nm, "endtime") and _get(nm, "consumption"):
        cfg = {
            "file_type": ft,
            "datetime_col":  _get(nm, "enddate"),
            "time_col":      _get(nm, "endtime"),
            "datetime_kind": "string_end",
            "ce_col":        _get(nm, "consumption"),
            "interval_minutes": 30,
        }
        nmi_col = _get(nm, "nmi")
        if nmi_col:
            cfg["nmi_col"] = nmi_col
        else:
            cfg["const_nmi"] = _extract_nmi(filename)
        uom_col = _get(nm, "consumptionuom") or _contains(nm, "uom")
        if uom_col:
            cfg["filter_col"]   = uom_col
            cfg["filter_value"] = "kWh"
        qual_col = _get(nm, "consumptionquality") or _contains(nm, "quality")
        if qual_col:
            cfg["quality_col"] = qual_col
        return cfg

    # ── 10. nmi + Date + Time + intervalvalue_k*h ────────────────────────────
    iv_col = _contains(nm, "intervalvalue_k")
    if _get(nm, "nmi") and _get(nm, "date") and _get(nm, "time") and iv_col:
        return {
            "file_type": ft,
            "nmi_col":       _get(nm, "nmi"),
            "datetime_col":  _get(nm, "date"),
            "time_col":      _get(nm, "time"),
            "datetime_kind": "string_end",
            "datetime_format": "%d/%m/%Y %I:%M:%S %p",
            "ce_col":        iv_col,
            "interval_minutes": 30,
        }

    # ── 11. Momentum / Service Point + Import kWh ────────────────────────────
    svc_col    = _get(nm, "service point", "service_point")
    import_col = _contains(nm, "import kwh") or _contains(nm, "import_kwh")
    if svc_col and import_col:
        dt_col     = _get(nm, "period start", "date", "datetime", "read datetime")
        export_col = _contains(nm, "export kwh") or _contains(nm, "export_kwh")
        cfg = {
            "file_type": ft,
            "nmi_col":       svc_col,
            "datetime_col":  dt_col,
            "datetime_kind": "string_start",
            "ce_col":        import_col,
            "interval_minutes": 30,
        }
        if export_col:
            cfg["soe_col"] = export_col
        return cfg

    # ── 12. Read datetime + Export kWh / Import kWh ──────────────────────────
    rd_col  = _get(nm, "read datetime")
    imp_col = _contains(nm, "import kwh")
    exp_col = _contains(nm, "export kwh")
    if rd_col and imp_col:
        cfg = {
            "file_type": ft,
            "datetime_col":  rd_col,
            "datetime_kind": "string_end",
            "ce_col":        imp_col,
            "interval_minutes": 30,
        }
        nmi_col = _get(nm, "nmi") or _get(nm, "site") or svc_col
        if nmi_col:
            cfg["nmi_col"] = nmi_col
        else:
            cfg["const_nmi"] = _extract_nmi(filename)
        if exp_col:
            cfg["soe_col"] = exp_col
        return cfg

    # ── Datetime + Net Grid Import kWh (no NMI column, NMI from filename) ───
    net_import_col = _contains(nm, "net grid import")
    if _get(nm, "datetime") and net_import_col:
        return {
            "file_type": ft,
            "const_nmi": _extract_nmi(filename) or Path(filename).stem[:10],
            "datetime_col": _get(nm, "datetime"),
            "datetime_kind": "string_start",
            "ce_col": net_import_col,
            "interval_minutes": 30,
        }

    # ── 13. NMI + E col (CE) + some datetime ────────────────────────────────
    if _get(nm, "nmi") and _get(nm, "e"):
        dt_col = _get(nm, "datetime", "date", "period", "timestamp")
        if dt_col:
            return {
                "file_type": ft,
                "nmi_col":      _get(nm, "nmi"),
                "datetime_col": dt_col,
                "datetime_kind": "string_start",
                "ce_col":       _get(nm, "e"),
                "soe_col":      _get(nm, "b"),
                "interval_minutes": 30,
            }

    # ── 14. Wide NMIs: DateTime column + 3+ 10-digit NMI column headers ──────
    dt_norm = next(
        (k for k in cols if k in ("datetime", "date time", "timestamp", "date")), None
    )
    if dt_norm:
        nmi_header_cols = [
            k for k in cols
            if re.match(r"^\d{10}$", k) or re.match(r"^n\d{9}$", k.lower())
        ]
        if len(nmi_header_cols) >= 3:
            return {
                "file_type": ft,
                "layout":        "wide_nmis",
                "datetime_col":  nm[dt_norm],
                "datetime_kind": "string_start",
                "ce_col":        "_wide_nmis_",
                "interval_minutes": 30,
            }

    # ── Undetected ────────────────────────────────────────────────────────────
    return None
