"""
clean_data.py - Energy Data Standardisation
============================================
Reads format definitions from formats.json and standardises every input file
into one combined Clean_Output_Combined_YYYYMMDD_HHMMSS.csv.

OUTPUT SCHEMA
  NMI, Date_time, the_date, the_interval, net_kWh, CE_kWh, SOE_kWh, notes

KEY RULES
  - Date_time    = END of the 30-min interval  (e.g. interval 1 -> 0:30)
  - the_date     = trading/settlement date (start-of-interval date)
  - the_interval = 1-48
  - SOE_kWh      <= 0 always (negated if source is positive)
  - net_kWh      = CE_kWh - abs(SOE_kWh)  (always recalculated)
  - notes        = quality code lowercased; "a" or blank -> "actual"
  - NMI          = first 10 characters only
  - Max 17 520 rows per NMI; if exceeded -> latest 12 calendar months kept

TO ADD A NEW FORMAT
  Add one JSON object to formats.json  -  no Python changes needed.

FORMATS.JSON FIELDS (all optional except file_name, file_type, datetime_col, ce_col)
  file_name        path relative to Input/ directory
  file_type        xlsx | csv | tsv | xlsb | html
  datetime_col     column with timestamp (or date part when time_col also set)
  datetime_kind    excel_end | excel_start | string_end | string_start
  ce_col           consumption kWh column name (or NMI column name for col_header)
  skip             true -> skip this file (moved to unprocessed/)
  nmi_col          NMI column (when nmi_source="col", the default)
  nmi_source       "col" (default) | "filename" | "col_header"
  soe_col          export kWh column
  quality_col      quality/status column
  date_col         separate date column (when date and time are split)
  time_col         separate time column (when date and time are split)
  interval_col     column with interval number 1-48
  datetime_format  explicit strptime format (null = auto-detect AU / ISO / EU)
  header_row       0-indexed row with column headers (default 0)
  interval_minutes source interval in minutes (default 30); aggregated to 30-min
  layout           "standard" (default) | "long" | "wide_intervals" | "nem12"
  type_col         for layout=long: column holding register/stream type
  ce_type_value    value in type_col identifying CE rows
  soe_type_value   value in type_col identifying SOE rows
"""

import csv
import io
import json
import re
import shutil
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR        = Path(__file__).parent
INPUT_DIR       = BASE_DIR / "Input"
PROCESSED_DIR   = INPUT_DIR / "processed"
UNPROCESSED_DIR = INPUT_DIR / "unprocessed"
OUTPUT_DIR      = BASE_DIR / "Output"
FORMATS_FILE    = BASE_DIR / "formats.json"

EXCEL_EPOCH = datetime(1899, 12, 30)
MAX_ROWS    = 17_520    # 365 days x 48 half-hour intervals

# NMI letter-prefix validation (Australian NEM state codes only)
_VALID_NMI_PREFIXES = frozenset("NQSVAT")

# ---------------------------------------------------------------------------
# Date / time helpers
# ---------------------------------------------------------------------------

def excel_serial_to_dt(val) -> datetime:
    dt = EXCEL_EPOCH + timedelta(days=float(val))
    # Round to nearest minute — stored serials have limited decimal precision
    # (e.g. 0.0208333333 days ≈ 29 min 59.999 s instead of exactly 30 min)
    return (dt + timedelta(seconds=30)).replace(second=0, microsecond=0)


_DT_FORMATS = [
    "%d/%m/%Y %I:%M:%S %p",  # AU 12-hr AM/PM   01/05/2025 12:30:00 AM
    "%d/%m/%Y %I:%M %p",     # AU 12-hr no secs 01/05/2025 12:30 AM
    "%d/%m/%Y %H:%M:%S",     # AU with seconds  01/05/2025 00:30:00
    "%d/%m/%Y %H:%M",        # AU               01/05/2025 00:30
    "%d/%m/%Y",              # AU date-only
    "%Y-%m-%d %H:%M:%S.%f",  # ISO microseconds 2025-05-01 00:30:00.000000
    "%Y-%m-%d %H:%M:%S",     # ISO with seconds 2025-05-01 00:30:00
    "%Y-%m-%dT%H:%M:%S",     # ISO T-separator  2025-05-01T00:30:00
    "%Y-%m-%d %H:%M",        # ISO
    "%Y-%m-%d",              # ISO date-only
    "%d.%m.%Y %H:%M:%S",     # EU with seconds  01.05.2025 00:30:00
    "%d.%m.%Y %H:%M",        # EU
    "%d.%m.%Y",              # EU date-only
    "%d %b %Y %H:%M:%S",     # Month-name       01 Apr 2025 00:30:00
    "%d %b %Y %H:%M",        # Month-name short 01 Apr 2025 00:30
    "%d %b %Y",              # Month-name date  01 Apr 2025
    "%d-%b-%y %H:%M",        # Short month-year 02-Jan-25 00:30
    "%d-%b-%Y %H:%M",        # Short month-year 02-Jan-2025 00:30
]


def parse_dt(val, fmt: str = None) -> datetime:
    s = str(val).strip()
    # Normalise am/pm to uppercase so %p matches on all platforms
    s = re.sub(r"\b(am|pm)\b", lambda m: m.group(0).upper(), s, flags=re.IGNORECASE)
    if fmt:
        return datetime.strptime(s, fmt)
    for f in _DT_FORMATS:
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    raise ValueError(f"Cannot parse datetime: {val!r}")


def parse_au_dt(val) -> datetime:
    return parse_dt(val)


def to_end_dt(raw, kind: str, dt_fmt: str = None, interval_minutes: int = 30) -> datetime:
    """Return interval END as a datetime regardless of source format/kind."""
    if isinstance(raw, (datetime, pd.Timestamp)):
        dt = raw if isinstance(raw, datetime) else raw.to_pydatetime()
        return dt + timedelta(minutes=interval_minutes) if kind.endswith("_start") else dt

    if kind == "excel_end":
        return excel_serial_to_dt(raw)
    if kind == "excel_start":
        return excel_serial_to_dt(raw) + timedelta(minutes=interval_minutes)

    if kind in ("string_end", "string_start"):
        dt = parse_dt(raw, fmt=dt_fmt)
        return dt + timedelta(minutes=interval_minutes) if kind == "string_start" else dt

    raise ValueError(f"Unknown datetime_kind: {kind!r}")


def get_interval(dt: datetime) -> int:
    """1-48 from interval END datetime."""
    if dt.hour == 0 and dt.minute == 0:
        return 48
    return (dt.hour * 60 + dt.minute) // 30


def fmt_dt(dt: datetime) -> str:
    return f"{dt.day}/{dt.month:02d}/{dt.year} {dt.hour}:{dt.minute:02d}"


def fmt_date(dt: datetime) -> str:
    return f"{dt.day}/{dt.month:02d}/{dt.year}"


def is_empty(val) -> bool:
    if val is None:
        return True
    try:
        if pd.isna(val):
            return True
    except (TypeError, ValueError):
        pass
    return str(val).strip() in ("", "nan", "NaT", "None", "NaN")


def get_trading_date(end_dt: datetime, date_raw, kind: str, dt_fmt: str = None) -> str:
    if not is_empty(date_raw):
        if isinstance(date_raw, (datetime, pd.Timestamp)):
            d = date_raw if isinstance(date_raw, datetime) else date_raw.to_pydatetime()
            return fmt_date(d)
        if "excel" in kind:
            return fmt_date(excel_serial_to_dt(date_raw))
        return fmt_date(parse_dt(date_raw, fmt=dt_fmt))
    return fmt_date(end_dt - timedelta(minutes=30))


def get_notes(val) -> str:
    if is_empty(val):
        return "actual"
    v = str(val).strip().lower()
    return "actual" if v == "a" else v


def safe_float(val) -> float:
    if is_empty(val):
        return 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


# ---------------------------------------------------------------------------
# File reading
# ---------------------------------------------------------------------------

def read_source(src, file_type: str, header_row: int = 0) -> pd.DataFrame:
    """Read an energy data file.  src can be a Path or a BytesIO object."""
    if file_type == "xlsx":
        return pd.read_excel(src, engine="openpyxl", dtype=object, header=header_row)
    if file_type in ("tsv", "xlsb"):
        return pd.read_csv(src, sep="\t", dtype=object, header=header_row,
                           comment="#", encoding="utf-8", encoding_errors="replace")
    if file_type == "html":
        # pd.read_html needs a string path; BytesIO is also accepted by lxml
        tables = pd.read_html(src if not isinstance(src, Path) else str(src), header=header_row)
        return tables[0].astype(object)
    return pd.read_csv(src, dtype=object, header=header_row,
                       encoding="utf-8-sig", encoding_errors="replace")


# ---------------------------------------------------------------------------
# NMI helpers
# ---------------------------------------------------------------------------

def extract_nmi_from_filename(file_name: str) -> str:
    """Extract NMI from a file path stem. Matches state-prefix NMIs and 10-digit NMIs."""
    stem = Path(file_name).stem
    m = re.search(r"[NQSVAT][A-Z0-9]{9}|\d{10}", stem, re.IGNORECASE)
    return m.group(0)[:10].upper() if m else ""


# ---------------------------------------------------------------------------
# Layout transforms
# ---------------------------------------------------------------------------

def transform_long(df: pd.DataFrame, cfg: dict) -> tuple:
    """
    Convert long-format data (CE and SOE on separate rows keyed by a type column)
    into a standard wide-per-interval DataFrame.  Returns (df, patched_cfg).
    """
    type_col    = cfg["type_col"]
    ce_val      = str(cfg["ce_type_value"])
    soe_val     = str(cfg["soe_type_value"]) if cfg.get("soe_type_value") else None
    value_col   = cfg["ce_col"]
    quality_col = cfg.get("quality_col")

    nmi_col  = cfg.get("nmi_col", "")
    dt_col   = cfg.get("datetime_col", "")
    time_col = cfg.get("time_col")
    key_cols = [c for c in [nmi_col, dt_col, time_col] if c and c in df.columns]

    use_startswith = cfg.get("type_startswith", False)
    if use_startswith:
        ce_df = df[df[type_col].astype(str).str.strip().str.startswith(ce_val)].copy()
    else:
        ce_df = df[df[type_col].astype(str).str.strip() == ce_val].copy()
    keep  = [c for c in key_cols + [value_col] + ([quality_col] if quality_col else [])
             if c in ce_df.columns]
    ce_df = ce_df[keep].rename(columns={value_col: "__ce__"})

    if soe_val:
        if use_startswith:
            soe_df = df[df[type_col].astype(str).str.strip().str.startswith(soe_val)].copy()
        else:
            soe_df = df[df[type_col].astype(str).str.strip() == soe_val].copy()
        soe_keep   = [c for c in key_cols + [value_col] if c in soe_df.columns]
        soe_df     = soe_df[soe_keep].rename(columns={value_col: "__soe__"})
        merge_keys = [c for c in key_cols if c in ce_df.columns and c in soe_df.columns]
        merged     = ce_df.merge(soe_df, on=merge_keys, how="left")

        # Misaligned timestamps (e.g. E1 at :00, B1 at :05): merge finds no matches.
        # Fall back: treat each row independently — E1 carries CE, B1 carries SOE.
        if merged["__soe__"].isna().all():
            ce_part           = ce_df.copy()
            ce_part["__soe__"] = 0.0
            soe_part           = soe_df.copy()
            soe_part["__ce__"] = 0.0
            target_cols        = list(ce_part.columns)
            merged = pd.concat([
                ce_part.reindex(columns=target_cols, fill_value=None),
                soe_part.reindex(columns=target_cols, fill_value=None),
            ], ignore_index=True)
    else:
        merged = ce_df.copy()
        merged["__soe__"] = None

    new_cfg = dict(cfg)
    new_cfg["ce_col"]  = "__ce__"
    new_cfg["soe_col"] = "__soe__"
    return merged, new_cfg


def transform_wide_intervals(df: pd.DataFrame, cfg: dict) -> tuple:
    """
    Convert format_47-style data (interval times "HH:MM - HH:MM" as column headers)
    into a standard per-interval DataFrame.  Returns (df, patched_cfg).
    """
    nmi_col  = cfg["nmi_col"]
    date_col = cfg["datetime_col"]     # DATE column
    type_col = cfg["type_col"]
    ce_val   = cfg["ce_type_value"]
    soe_val  = cfg.get("soe_type_value")
    qual_col = cfg.get("quality_col")

    interval_cols = [c for c in df.columns
                     if re.match(r"\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", str(c))]
    if not interval_cols:
        return df, cfg

    id_vars = [c for c in df.columns if c not in interval_cols]
    melted  = df.melt(id_vars=id_vars, value_vars=interval_cols,
                      var_name="__int_lbl__", value_name="__energy__")

    def label_to_end(row):
        end_hm = str(row["__int_lbl__"]).split("-")[1].strip()
        return f"{row[date_col]} {end_hm}"

    melted["__datetime__"] = melted.apply(label_to_end, axis=1)

    ce_df  = melted[melted[type_col].astype(str).str.strip() == str(ce_val)].copy()
    ce_keep = [nmi_col, "__datetime__", "__energy__"] + ([qual_col] if qual_col and qual_col in ce_df.columns else [])
    ce_df   = ce_df[ce_keep].rename(columns={"__energy__": "__ce__"})

    if soe_val:
        soe_df = melted[melted[type_col].astype(str).str.strip() == str(soe_val)].copy()
        soe_df = soe_df[[nmi_col, "__datetime__", "__energy__"]].rename(columns={"__energy__": "__soe__"})
        result = ce_df.merge(soe_df, on=[nmi_col, "__datetime__"], how="left")
    else:
        result = ce_df.copy()
        result["__soe__"] = None

    new_cfg = dict(cfg)
    new_cfg["datetime_col"]  = "__datetime__"
    new_cfg["datetime_kind"] = "string_end"
    new_cfg["ce_col"]        = "__ce__"
    new_cfg["soe_col"]       = "__soe__"
    new_cfg["date_col"]      = None
    new_cfg["time_col"]      = None
    new_cfg["layout"]        = "standard"
    return result, new_cfg


def transform_wide_nmis(df: pd.DataFrame, cfg: dict) -> tuple:
    """
    Convert a wide-NMI format: one DateTime column + one column per NMI containing CE values.
    Duplicate NMI columns (pandas adds .1/.2) are aggregated by stripping the suffix and summing.
    Returns (melted_df, patched_cfg).
    """
    dt_col = cfg["datetime_col"]
    value_cols = [c for c in df.columns if c != dt_col]

    # Build mapping: base_nmi -> [col, col, ...] (handles .1 .2 duplicates)
    nmi_map = {}
    for col in value_cols:
        base = re.sub(r"\.\d+$", "", str(col)).strip()[:10]
        nmi_map.setdefault(base, []).append(col)

    rows = []
    for _, row in df.iterrows():
        for base_nmi, cols in nmi_map.items():
            total = sum(safe_float(row[c]) for c in cols)
            rows.append({"__nmi__": base_nmi, "__dt__": row[dt_col], "__ce__": total})

    result = pd.DataFrame(rows)

    new_cfg = dict(cfg)
    new_cfg["nmi_col"]      = "__nmi__"
    new_cfg["datetime_col"] = "__dt__"
    new_cfg["ce_col"]       = "__ce__"
    new_cfg["soe_col"]      = None
    new_cfg["time_col"]     = None
    new_cfg["layout"]       = "standard"
    return result, new_cfg


# ---------------------------------------------------------------------------
# Wide-daily parser (one row per day, H:MM columns = interval start times)
# ---------------------------------------------------------------------------

def transform_wide_daily(df: pd.DataFrame, cfg: dict) -> list:
    """
    Convert wide-daily format into standard clean output rows.
    Each source row is one day; column headers are interval start times (H:MM).
    Returns complete row dicts directly (like the NEM12 parser).
    """
    date_col      = cfg.get("datetime_col", "Date/Time")
    interval_cols = [c for c in df.columns if re.match(r"^\d{1,2}:\d{2}$", str(c).strip())]
    qual_col      = next((c for c in df.columns if "quality" in str(c).lower()), None)
    nmi           = str(cfg.get("const_nmi", ""))[:10]

    output = []
    for _, row in df.iterrows():
        date_str = str(row.get(date_col, "")).strip().split(".")[0]
        try:
            base_date = datetime.strptime(date_str, "%Y%m%d")
        except ValueError:
            continue

        notes = get_notes(row[qual_col]) if qual_col and qual_col in row.index else "actual"

        for time_label in interval_cols:
            val = safe_float(row[time_label])
            try:
                h, m = map(int, str(time_label).strip().split(":"))
                start_dt = base_date.replace(hour=h, minute=m)
                end_dt   = start_dt + timedelta(minutes=30)
            except (ValueError, TypeError):
                continue

            ce = round(val, 4)
            output.append({
                "NMI":          nmi,
                "Date_time":    fmt_dt(end_dt),
                "the_date":     fmt_date(start_dt),
                "the_interval": get_interval(end_dt),
                "net_kWh":      ce,
                "CE_kWh":       ce,
                "SOE_kWh":      0.0,
                "notes":        notes,
            })

    return output


# ---------------------------------------------------------------------------
# NEM12 parser
# ---------------------------------------------------------------------------

def _parse_nem12_rows(rows_iter, src_mins: int) -> list:
    """Core NEM12 parsing — accepts any iterable of row sequences (CSV or xlsx)."""
    bucket_size = 30 // src_mins
    per_day     = 48 * bucket_size

    buckets        = defaultdict(lambda: {"ce": 0.0, "soe": 0.0, "notes": "actual"})
    current_nmi    = None
    current_stream = None

    for record in rows_iter:
        if not record:
            continue
        rec_type = str(record[0]).strip()

        if rec_type == "200":
            current_nmi = str(record[1]).strip()[:10] if len(record) > 1 else None
            # NMI suffix (stream ID) can appear at index 2, 3, or 4 depending on variant
            current_stream = "E1"
            for idx in (3, 4, 2):
                if len(record) > idx:
                    val = str(record[idx]).strip()
                    if re.match(r"^[BEQ][1-9]$", val):
                        current_stream = val
                        break

        elif rec_type == "300" and current_nmi:
            if len(record) < 3:
                continue
            date_str = str(record[1]).strip()
            values   = [record[i] for i in range(2, min(2 + per_day, len(record)))]
            qual_idx = 2 + per_day
            quality  = str(record[qual_idx]).strip() if len(record) > qual_idx else ""
            notes    = get_notes(quality)
            try:
                base_date = datetime.strptime(date_str, "%Y%m%d")
            except ValueError:
                continue

            for bucket_idx in range(48):
                start  = bucket_idx * bucket_size
                chunk  = values[start : start + bucket_size]
                if not chunk:
                    continue
                energy = round(sum(safe_float(v) for v in chunk), 4)
                end_dt = base_date + timedelta(minutes=30 * (bucket_idx + 1))
                key    = (current_nmi, end_dt)

                if current_stream == "E1":
                    buckets[key]["ce"]    = energy
                    buckets[key]["notes"] = notes
                elif current_stream == "B1":
                    buckets[key]["soe"]   = -abs(energy)

    output = []
    for (nmi, end_dt), vals in sorted(buckets.items()):
        ce  = round(vals["ce"],  4)
        soe = round(vals["soe"], 4)
        output.append({
            "NMI":          nmi,
            "Date_time":    fmt_dt(end_dt),
            "the_date":     fmt_date(end_dt - timedelta(minutes=30)),
            "the_interval": get_interval(end_dt),
            "net_kWh":      round(ce - abs(soe), 4),
            "CE_kWh":       ce,
            "SOE_kWh":      soe,
            "notes":        vals["notes"],
        })
    return output


def parse_nem12(src, cfg: dict) -> list:
    """Parse NEM12-format CSV.  src can be a Path or a BytesIO object."""
    src_mins = cfg.get("interval_minutes", 30)
    try:
        if isinstance(src, Path):
            with open(src, newline="", encoding="utf-8", errors="replace") as fh:
                return _parse_nem12_rows(csv.reader(fh), src_mins)
        else:
            text = src.read().decode("utf-8", errors="replace")
            return _parse_nem12_rows(csv.reader(io.StringIO(text)), src_mins)
    except Exception as exc:
        print(f"  [WARN]  NEM12 parse error: {exc}")
        return []


def parse_nem12_xlsx(src, cfg: dict) -> list:
    """Parse NEM12 data stored row-by-row in an xlsx file.  src can be a Path or BytesIO."""
    import openpyxl
    src_mins = cfg.get("interval_minutes", 30)
    try:
        wb   = openpyxl.load_workbook(src, read_only=True, data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        return _parse_nem12_rows(rows, src_mins)
    except Exception as exc:
        print(f"  [WARN]  NEM12 xlsx parse error: {exc}")
        return []


# ---------------------------------------------------------------------------
# Interval resampling  (5-min or 15-min -> 30-min)
# ---------------------------------------------------------------------------

def resample_to_30min(rows: list) -> list:
    """
    Aggregate sub-30-min rows into 30-min buckets.
    Uses ceiling-to-nearest-30-min for end-time snapping.
    """
    buckets = defaultdict(lambda: {"ce": 0.0, "soe": 0.0, "notes": "actual", "nmi": ""})

    for r in rows:
        end_dt    = datetime.strptime(r["Date_time"], "%d/%m/%Y %H:%M")
        total_min = end_dt.hour * 60 + end_dt.minute

        if total_min == 0:
            bucket_end = end_dt
        else:
            bm = ((total_min - 1) // 30 + 1) * 30
            if bm >= 1440:
                bucket_end = (end_dt + timedelta(days=1)).replace(
                    hour=0, minute=0, second=0, microsecond=0)
            else:
                bucket_end = end_dt.replace(
                    hour=bm // 60, minute=bm % 60, second=0, microsecond=0)

        key = (r["NMI"], bucket_end)
        buckets[key]["ce"]  += r["CE_kWh"]
        buckets[key]["soe"] += r["SOE_kWh"]
        buckets[key]["nmi"]  = r["NMI"]
        if r["notes"] != "actual":
            buckets[key]["notes"] = r["notes"]

    output = []
    for (nmi, end_dt), vals in sorted(buckets.items()):
        ce  = round(vals["ce"],  4)
        soe = round(vals["soe"], 4)
        output.append({
            "NMI":          nmi,
            "Date_time":    fmt_dt(end_dt),
            "the_date":     fmt_date(end_dt - timedelta(minutes=30)),
            "the_interval": get_interval(end_dt),
            "net_kWh":      round(ce - abs(soe), 4),
            "CE_kWh":       ce,
            "SOE_kWh":      soe,
            "notes":        vals["notes"],
        })
    return output


# ---------------------------------------------------------------------------
# Forecast / gap-fill  (weekday-average profile, per NMI)
# ---------------------------------------------------------------------------

def _forecast_nmi(nmi_rows: list) -> list:
    """
    Build a (weekday, interval) average profile from ALL actual rows for one NMI,
    then emit a complete 365-day window ending at max_dt, filling any missing
    30-min slots with the weekday-average forecast.
    """
    if not nmi_rows:
        return []

    dts    = [datetime.strptime(r["Date_time"], "%d/%m/%Y %H:%M") for r in nmi_rows]
    max_dt = max(dts)
    nmi    = nmi_rows[0]["NMI"]

    # 12-month window: (window_start, max_dt]
    try:
        window_start = max_dt.replace(year=max_dt.year - 1)
    except ValueError:
        window_start = max_dt.replace(year=max_dt.year - 1, day=28)

    # Build profile from ALL actual data (including data before the window)
    ce_sums  = defaultdict(float)
    soe_sums = defaultdict(float)
    counts   = defaultdict(int)
    for dt, r in zip(dts, nmi_rows):
        key = (dt.weekday(), r["the_interval"])
        ce_sums[key]  += r["CE_kWh"]
        soe_sums[key] += r["SOE_kWh"]
        counts[key]   += 1

    avg = {
        k: {
            "ce":  round(ce_sums[k]  / counts[k], 4),
            "soe": round(soe_sums[k] / counts[k], 4),
        }
        for k in counts
    }

    # Index actual rows that fall inside the window
    actual = {dt: r for dt, r in zip(dts, nmi_rows) if dt > window_start}

    # Walk every 30-min slot in the window
    output       = []
    forecast_cnt = 0
    slot         = window_start + timedelta(minutes=30)

    while slot <= max_dt:
        if slot in actual:
            output.append(actual[slot])
        else:
            interval = get_interval(slot)
            key      = (slot.weekday(), interval)
            ce       = avg.get(key, {}).get("ce",  0.0)
            soe      = avg.get(key, {}).get("soe", 0.0)
            output.append({
                "NMI":          nmi,
                "Date_time":    fmt_dt(slot),
                "the_date":     fmt_date(slot - timedelta(minutes=30)),
                "the_interval": interval,
                "net_kWh":      round(ce - abs(soe), 4),
                "CE_kWh":       ce,
                "SOE_kWh":      soe,
                "notes":        "forecast",
            })
            forecast_cnt += 1
        slot += timedelta(minutes=30)

    if forecast_cnt:
        print(f"  [FORECAST]  NMI {nmi}: {forecast_cnt:,} slots filled "
              f"({len(output):,} total, {len(actual):,} actual)")
    return output


def fill_to_year(rows: list) -> list:
    """Split rows by NMI, apply _forecast_nmi to each, return combined result."""
    if not rows:
        return rows
    groups = defaultdict(list)
    for r in rows:
        groups[r["NMI"]].append(r)
    result = []
    for nmi_rows in groups.values():
        result.extend(_forecast_nmi(nmi_rows))
    return result


# ---------------------------------------------------------------------------
# Per-format processing
# ---------------------------------------------------------------------------

def process_format(cfg: dict, file_path: Path = None, file_bytes: bytes = None) -> tuple:
    """
    Transform one input file into a list of clean output row dicts.
    Returns (rows, success).

    Supply exactly one of:
      file_bytes  raw bytes from an upload (fully in-memory, nothing written to disk)
      file_path   explicit Path override (CLI / temp file)
      neither     falls back to INPUT_DIR / cfg["file_name"]
    """
    if file_bytes is not None:
        file_src = io.BytesIO(file_bytes)
        label = cfg.get("file_name", "uploaded file")
    elif file_path is not None:
        file_src = file_path
        label = file_path.name
    else:
        file_path = INPUT_DIR / cfg["file_name"]
        file_src = file_path
        label = cfg["file_name"]

    print(f"\n--- {label} ---")

    if cfg.get("skip", False):
        print("  [SKIP]  Flagged skip=true in formats.json")
        return [], False

    if isinstance(file_src, Path) and not file_src.exists():
        print(f"  [SKIP]  File not found: {file_src}")
        return [], False

    # NEM12: dedicated parser returns complete rows
    if cfg.get("layout") == "nem12":
        if cfg.get("file_type") == "xlsx":
            rows = parse_nem12_xlsx(file_src, cfg)
        else:
            rows = parse_nem12(file_src, cfg)
        print(f"  Cleaned rows : {len(rows):,}")
        return rows, len(rows) > 0

    # Wide daily: dedicated parser returns complete rows
    if cfg.get("layout") == "wide_daily":
        df_wd = read_source(file_src, cfg["file_type"], header_row=cfg.get("header_row", 3))
        print(f"  Source rows  : {len(df_wd):,}")
        rows = transform_wide_daily(df_wd, cfg)
        print(f"  Cleaned rows : {len(rows):,}")
        return rows, len(rows) > 0

    df = read_source(file_src, cfg["file_type"], header_row=cfg.get("header_row", 0))
    print(f"  Source rows  : {len(df):,}")

    # Skip leading metadata rows after the header
    data_skip = cfg.get("data_skip_rows", 0)
    if data_skip > 0:
        df = df.iloc[data_skip:].reset_index(drop=True)

    # Pre-filter rows by a column value before layout transforms
    filter_col = cfg.get("filter_col")
    filter_val = cfg.get("filter_value")
    if filter_col and filter_val and filter_col in df.columns:
        df = df[df[filter_col].astype(str).str.strip() == str(filter_val)].reset_index(drop=True)

    # Layout transforms
    if cfg.get("layout") == "long":
        df, cfg = transform_long(df, cfg)
    elif cfg.get("layout") == "wide_intervals":
        df, cfg = transform_wide_intervals(df, cfg)
    elif cfg.get("layout") == "wide_nmis":
        df, cfg = transform_wide_nmis(df, cfg)

    # Resolve constant NMI for filename / col_header / explicit sources
    nmi_source = cfg.get("nmi_source", "col")
    const_nmi  = None
    if cfg.get("const_nmi"):
        const_nmi = str(cfg["const_nmi"])[:10]
    elif nmi_source == "filename":
        const_nmi = extract_nmi_from_filename(cfg["file_name"])
        if not const_nmi:
            # Fallback: first segment of stem before space/underscore/hyphen
            stem = Path(cfg["file_name"]).stem
            const_nmi = re.split(r"[ _\-]", stem)[0][:10]
            if not const_nmi:
                print(f"  [WARN]  Could not extract NMI from filename: {cfg['file_name']}")
    elif nmi_source == "col_header":
        const_nmi = str(cfg["ce_col"])[:10]

    kind         = cfg.get("datetime_kind", "string_end")
    dt_fmt       = cfg.get("datetime_format")
    src_mins     = cfg.get("interval_minutes", 30)
    time_col     = cfg.get("time_col")
    dt_col       = cfg["datetime_col"]

    # Sniff datetime format once from the first parseable row (avoids 18 trials per row)
    if dt_fmt is None and kind in ("string_end", "string_start"):
        for _, probe in df.head(10).iterrows():
            try:
                raw = probe[dt_col]
                if time_col:
                    raw = str(probe[dt_col]).strip() + " " + str(probe[time_col]).strip()
                raw = str(raw).strip()
                for _f in _DT_FORMATS:
                    try:
                        datetime.strptime(raw, _f)
                        dt_fmt = _f
                        break
                    except ValueError:
                        pass
                if dt_fmt:
                    break
            except Exception:
                pass

    rows    = []
    skipped = 0

    for _, row in df.iterrows():

        # NMI
        if const_nmi:
            nmi = const_nmi
        elif cfg.get("nmi_col"):
            nmi = str(row[cfg["nmi_col"]]).strip()
        else:
            skipped += 1
            continue
        if len(nmi) > 10:
            nmi = nmi[:10]
        if not nmi or nmi in ("nan", "None"):
            skipped += 1
            continue
        # Reject NMIs whose first character is a letter outside NEM state codes
        if nmi[0].isalpha() and nmi[0].upper() not in _VALID_NMI_PREFIXES:
            skipped += 1
            continue

        # Interval END datetime
        try:
            if time_col:
                if kind.startswith("excel"):
                    raw_dt = safe_float(row[dt_col]) + safe_float(row[time_col])
                else:
                    d_val = row[dt_col]
                    t_val = row[time_col]
                    # xlsx may return Python datetime/time objects directly
                    if isinstance(d_val, (datetime, pd.Timestamp)) and hasattr(t_val, "hour"):
                        base   = d_val.to_pydatetime() if isinstance(d_val, pd.Timestamp) else d_val
                        raw_dt = base.replace(hour=int(t_val.hour), minute=int(t_val.minute),
                                              second=0, microsecond=0)
                    else:
                        raw_dt = str(d_val).strip() + " " + str(t_val).strip()
            else:
                raw_dt = row[dt_col]
            end_dt = to_end_dt(raw_dt, kind, dt_fmt=dt_fmt, interval_minutes=src_mins)
        except Exception:
            skipped += 1
            continue

        # Interval number 1-48
        if cfg.get("interval_col"):
            try:
                interval = int(float(row[cfg["interval_col"]]))
            except (TypeError, ValueError):
                interval = get_interval(end_dt)
        else:
            interval = get_interval(end_dt)

        # Trading date
        date_raw = row[cfg["date_col"]] if cfg.get("date_col") else None
        t_date   = get_trading_date(end_dt, date_raw, kind, dt_fmt=dt_fmt)

        # Energy values
        ce  = safe_float(row[cfg["ce_col"]])
        soe = safe_float(row[cfg["soe_col"]]) if cfg.get("soe_col") else 0.0
        if soe > 0:
            soe = -soe
        net = round(ce - abs(soe), 4)
        ce  = round(ce,  4)
        soe = round(soe, 4)

        # Notes / quality
        q_raw = row[cfg["quality_col"]] if cfg.get("quality_col") else None
        notes = get_notes(q_raw)

        rows.append({
            "NMI":          nmi,
            "Date_time":    fmt_dt(end_dt),
            "the_date":     t_date,
            "the_interval": interval,
            "net_kWh":      net,
            "CE_kWh":       ce,
            "SOE_kWh":      soe,
            "notes":        notes,
        })

    if skipped:
        print(f"  [WARN]  Skipped {skipped} rows (unparseable datetime or blank NMI)")
    print(f"  Cleaned rows : {len(rows):,}")

    if rows:
        pre = len(rows)
        keys = [(r["NMI"], r["Date_time"]) for r in rows]
        if src_mins != 30 or len(keys) != len(set(keys)):
            rows = resample_to_30min(rows)
            if len(rows) != pre or src_mins != 30:
                print(f"  Aggregated to: {len(rows):,} rows (30-min)")

    return rows, len(rows) > 0


# ---------------------------------------------------------------------------
# Trim & gap report
# ---------------------------------------------------------------------------

def limit_to_year(rows: list) -> list:
    if len(rows) <= MAX_ROWS:
        return rows
    dts    = [datetime.strptime(r["Date_time"], "%d/%m/%Y %H:%M") for r in rows]
    max_dt = max(dts)
    try:
        cutoff = max_dt.replace(year=max_dt.year - 1)
    except ValueError:
        cutoff = max_dt.replace(year=max_dt.year - 1, day=28)
    kept = [r for r, dt in zip(rows, dts) if dt > cutoff]
    print(f"  [TRIM]  {len(rows):,} -> {len(kept):,} rows  (cutoff: {fmt_date(cutoff)})")
    return kept


def gap_report(rows: list, nmi: str) -> None:
    n   = len(rows)
    gap = MAX_ROWS - n
    if gap <= 0:
        print(f"  [OK]   NMI {nmi} : {n:,} rows - full 12 months")
    else:
        days = round(gap / 48, 1)
        print(f"  [GAP]  NMI {nmi} : {n:,} rows - {gap} intervals missing (~{days} days)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def move_file(src: Path, dest_dir: Path, label: str) -> None:
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    if dest.exists():
        dest.unlink()
    shutil.move(str(src), str(dest))
    print(f"  [MOVE -> {label}/]  {src.name}")


def main() -> None:
    print()
    print("=" * 65)
    print("  Energy Data Standardisation")
    print("=" * 65)
    print(f"  Input    : {INPUT_DIR}")
    print(f"  Formats  : {FORMATS_FILE}")
    print("=" * 65)

    run_ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
    OUTPUT_FILE = OUTPUT_DIR / f"Clean_Output_Combined_{run_ts}.csv"

    if not FORMATS_FILE.exists():
        print(f"\n[ERROR] formats.json not found at {FORMATS_FILE}")
        sys.exit(1)

    with open(FORMATS_FILE, encoding="utf-8") as f:
        format_configs = json.load(f)

    print(f"  Loaded {len(format_configs)} format(s) from formats.json")

    run_processed_dir = PROCESSED_DIR / run_ts
    run_processed_dir.mkdir(parents=True, exist_ok=True)
    UNPROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    format_configs       = [c for c in format_configs if c.get("file_name")]
    registered_basenames = {Path(cfg["file_name"]).name for cfg in format_configs}
    all_rows = []

    for cfg in format_configs:
        rows, success = process_format(cfg)
        rows = fill_to_year(rows)
        nmi_label = rows[0]["NMI"] if rows else "unknown"
        gap_report(rows, nmi_label)
        all_rows.extend(rows)

        src = INPUT_DIR / cfg["file_name"]
        if src.exists():
            move_file(src, run_processed_dir if success else UNPROCESSED_DIR,
                      f"processed/{run_ts}" if success else "unprocessed")

    # Move unregistered top-level Input files to unprocessed
    for f in INPUT_DIR.iterdir():
        if f.is_file() and f.name not in registered_basenames:
            move_file(f, UNPROCESSED_DIR, "unprocessed")

    OUTPUT_DIR.mkdir(exist_ok=True)
    columns = ["NMI", "Date_time", "the_date", "the_interval", "net_kWh", "CE_kWh", "SOE_kWh", "notes"]
    pd.DataFrame(all_rows, columns=columns).to_csv(OUTPUT_FILE, index=False)

    print()
    print("=" * 65)
    print("  COMPLETE")
    print("=" * 65)
    print(f"  Total rows    : {len(all_rows):,}")
    distinct_nmis = sorted(set(r["NMI"] for r in all_rows))
    print(f"  Distinct NMIs : {len(distinct_nmis)}")
    for nmi in distinct_nmis:
        cnt  = sum(1 for r in all_rows if r["NMI"] == nmi)
        warn = f"  [WARN] NMI is {len(nmi)} chars (expected 10)" if len(nmi) != 10 else ""
        print(f"    {nmi}  ({cnt:,} rows){warn}")
    print()
    print(f"  Output file   : {OUTPUT_FILE}")
    print("=" * 65)


if __name__ == "__main__":
    main()
