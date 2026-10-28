"""Trucks Report (late trucks) builder - replaces the Tableau "Trucks Report_KW" download.

Input  : one or more Tableau exports - the custom-SQL data source CSV (one row per TLM, with I/M version columns
         and Tableau's RED/GREEN colors), the "Download > Data" CSV (3 rows per TLM, one per pane), a crosstab
         export, the output of sql/*.sql, or a folder of them. Column names are matched by alias.
Output : <out>/Trucks Report <as-of>.xlsx  (the report window, default last 7 days)
            Action List        open items to act on now: not departed, arrival overdue, projected late, fix record
            Trucks Report      one row per TLM, red/green kept as real cell formatting, delay minutes + status text
            Late - Report Card late trucks in the exact 2026 Truck Report Card column order, ready to paste
            Summary            KPIs, then by vendor, origin station, lane, repeat late trucks and day
            Data Quality       actuals keyed on the wrong date, arrivals before departures, missing actuals
            About              definitions, inputs, counts
         <out>/late_trucks_feed.json and .js   compact feed for the Coordinator Hub / Trucking Pulse
         <history>/late_trucks_history.csv     every truck ever loaded, one row per TLM (for Power BI / other reports)
         <history>/late_trucks_daily.csv       daily rollup by lane and vendor
         <history>/Trucks History Summary.xlsx monthly trend, vendor and station by month

Usage  : py late_trucks.py "export.csv" [--days 7 | --from 2026-10-01 --to 2026-10-18 | --all]
                          [--origin JFK,ORD] [--dest LHR] [--vendor "FLY US"] [--threshold 59]
                          [--report-card "2026 Truck Report Card.xlsx" | --report-card none] [--out Output]
                          [--arr-basis system|team] [--history-dir History | --no-history]
         py late_trucks.py "old exports folder" "2025.csv" --history-only     (backfill the history store)
Needs  : Python 3.9+, pandas, xlsxwriter, openpyxl (report-card cross-check only).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:  # formula_book.py sits next to this file
    sys.path.insert(0, str(HERE))
DEFAULT_REPORT_CARD = HERE / "Inputs" / "2026 Truck Report Card.xlsx"
LATE_MIN = 59               # Tableau rule: DATEDIFF('minute', scheduled, actual) > 59 -> RED
LATE_CHECK_MIN = 2 * 1440   # actual more than 2 days after schedule -> probably keyed on the wrong date
EARLY_CHECK_MIN = 12 * 60   # departure more than 12 h before schedule -> probably keyed on the wrong date
MIN_TRANSIT_MIN = 10        # arrival less than 10 min after departure (or before it) -> arrival actual is wrong
GRACE_MIN = 120             # no actual yet within 2 h of schedule = still Pending (data entry / extract lag)
LOOKBACK_H = 48             # Action List covers events scheduled in the last N hours
DEFAULT_DAYS = 7            # report window when no --days/--from/--to/--all is given
DATA_DAYS = 91              # Data tab of the daily workbook: last 13 weeks (trends + lane medians)
MIN_LANE_TRIPS = 3          # clean trips needed before a lane's median transit is used
LOCAL_TZ = "America/Chicago"   # time zone of the PC clock (as-of time)
SOURCE_LABEL = "TABLEAU"    # value written to the Report Card "Source" column (matches existing entries)
FEED_DAYS = 3               # days of truck-level rows in the feed (KPIs cover the whole window)
FEED_BACKLOG_DAYS = 7       # plus late trucks not yet on the Report Card from the last N days
HISTORY_FILE = "late_trucks_history.csv"

RED, GREEN, AMBER, GREY = "#E15759", "#59A14F", "#B07A00", "#8C8C8C"   # Tableau palette
HEADER_BLUE = "#4D77A4"

ALIASES = {
    "sched_date": ["Scheduled Departure Date", "Scheduled Departure Date (Exclude Nulls)", "SCHD_DEP_DT"],
    "truck": ["Truck", "MKT_FLIGHT_NBR", "fltnum", "TRK NBR"],
    "vendor_code": ["Vendor Code", "trkcmpcod"],
    "vendor": ["Vendor Name", "trkcmpnam"],
    "origin": ["Origin", "DEP_AIRPRT_IATA_CD"],
    "dest": ["Destination", "ARVL_AIRPRT_IATA_CD"],
    "ac_type": ["AC Type", "SUBFLEET_CD"],
    "sched_dep_time": ["Scheduled Departure Time", "SCHD_DEP_TM"],
    "sched_dep_dttm": ["Schedule Departure Date/Time", "Scheduled Departure", "SCHD_DEP_DTTM"],
    "sched_arr": ["Scheduled Arrival Date", "Scheduled Arrival", "SCHD_ARVL_DTTM"],
    "tlm": ["TLM", "trkordnum"],
    "act_dep": ["Actual Departure Date", "Actual Departure", "ACTL_DEP_DTTM"],
    "act_arr": ["Actual Arrival Date", "Actual Arrival", "ACTL_ARVL_DTTM"],
    "order_date": ["trkorddat", "Order Date"],
    "version": ["I", "Record Version"],
    "version_max": ["M"],
    "tab_dep_color": ["Actual Dep Color", "Tableau Dep Color"],
    "tab_arr_color": ["Actual Arrival Color", "Tableau Arr Color"],
    "source_file": ["Source File"],
}
REQUIRED = ["sched_date", "truck", "origin", "dest", "sched_arr", "act_dep", "act_arr"]
BASE_COLS = list(ALIASES)

# Station -> IANA time zone, used to tell "not departed yet" from "overdue" in station-local time.
# Stations not listed fall back to the airportsdata package if installed, otherwise to the PC's time zone.
_TZ_GROUPS = {
    "America/New_York": "ATL BDL BOS BWI CHS CLE CLT CMH CVG EWR GSO IAD JFK MCO MIA OPF ORF PHL PIT RDU RIC ROC "
                        "SUA TPA NYC LGA DCA JAX FLL RSW PBI BUF SYR ALB PVD SAV CAE GSP MYR DAY TYS CHA RIC "
                        "ABE MDT AVL ILM",
    "America/Detroit": "DTW GRR",
    "America/Indiana/Indianapolis": "IND",
    "America/Kentucky/Louisville": "SDF",
    "America/Chicago": "AUS BNA DFW IAH LRD MCI MEM MFE MKE MSP MSY OMA ORD STL TUL SAT HOU DAL BHM HSV JAN LIT "
                       "OKC ICT DSM MDW XNA CRP BRO",
    "America/Denver": "ABQ DEN ELP SLC COS",
    "America/Boise": "BOI",
    "America/Phoenix": "PHX TUS",
    "America/Los_Angeles": "LAS LAX SAN SEA SFO OAK SJC PDX ONT SMF RNO",
    "America/Anchorage": "ANC",
    "Pacific/Honolulu": "HNL",
    "America/Puerto_Rico": "SJU",
    "America/Toronto": "YUL YYZ YOW",
    "America/Vancouver": "YVR",
    "America/Edmonton": "YYC YEG",
    "America/Mexico_City": "GDL MEX QRO",
    "America/Monterrey": "MTY",
    "America/Cancun": "CUN",
    "America/Sao_Paulo": "CNF CWB GRU POA VCP GIG",
    "Europe/London": "ABZ BFS BHX EDI EMA GLA LHR MAN NCL LGW STN LTN LON BRS",
    "Europe/Dublin": "DUB ORK SNN",
    "Europe/Lisbon": "LIS OPO",
    "Europe/Madrid": "AGP BCN BIO MAD VLC VIT ZAZ SVQ",
    "Europe/Amsterdam": "AMS EIN RTM",
    "Europe/Berlin": "BER CGN DRS DUS FRA HAJ HAM MUC NUE STR LEJ",
    "Europe/Copenhagen": "BLL CPH",
    "Europe/Rome": "AOI BLQ FCO FLR LIN MIL MXP NAP VCE BGY PSA TRN",
    "Europe/Paris": "BOD CDG LIL LYS MRS NCE NTE ORY SXB TLS PAR",
    "Europe/Zurich": "BRN BSL ZRH GVA",
    "Europe/Brussels": "BRU LGG",
    "Europe/Luxembourg": "LUX",
    "Europe/Vienna": "VIE",
    "Europe/Budapest": "BUD",
    "Europe/Prague": "PRG",
    "Europe/Warsaw": "WAW WRO LCJ KTW",
    "Europe/Stockholm": "ARN GOT",
    "Europe/Oslo": "OSL",
    "Europe/Helsinki": "HEL",
    "Europe/Athens": "ATH",
    "Asia/Tokyo": "FUK HND KIX NGO NRT TYO",
    "Asia/Seoul": "ICN",
    "Asia/Hong_Kong": "HKG",
    "Asia/Shanghai": "PVG PEK",
    "Australia/Adelaide": "ADL",
    "Australia/Brisbane": "BNE",
    "Australia/Melbourne": "MEL",
    "Australia/Perth": "PER",
    "Australia/Sydney": "SYD",
}
STATION_TZ = {code: tz for tz, codes in _TZ_GROUPS.items() for code in codes.split()}

# AC type prefix -> (truck type, arrival rule). Training guide: transfer trucks never have arrival times; mail
# truck arrivals are entered as the STA. RFS trucks carry no arrival on the flight record either.
#   None = arrival never measured, STA = arrival keyed as STA (not measured; still chase a missing one)
TRUCK_TYPES = {"XFR": ("Transfer", "None"), "RFS": ("RFS", "None"), "MAL": ("Mail", "STA")}

RC_COLUMNS = ["Vendor", "Vendor code ", "Month", "Departure Date", "TRK NBR", "ORG", "DEST",
              "Scheduled Departure", "Scheduled Arrival", "Actual Departure", "Actual Arrival",
              "Source", "Source", "GCX", "EPX", "EPXTC", "MAIL", "AWB #", "Reason", "Fault", "Summary",
              "Agent", "Agent completed by"]
RC_REASONS = ["MECH", "DRIVER", "VENDOR", "CUSTOMS", "AA OPS", "AA SALES", "WX"]
RC_FAULTS = ["VN", "CB", "AA", "WX"]

# Action List: (code, priority, label, next step)
ACTIONS = [
    ("not_departed", 1, "Not departed", "Call vendor: confirm departure, delay or no-show"),
    ("arr_overdue", 1, "Arrival overdue", "Confirm arrival with destination / vendor; update the record"),
    ("proj_late", 2, "Projected late arrival", "Warn destination and customer of expected arrival"),
    ("dep_late", 3, "Departed late", "Watch; log on the Report Card if it arrives late"),
    ("fix_record", 4, "Fix record", "Correct the actual on the flight record"),
]


# --------------------------------------------------------------------------------------------- load
def read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=str, keep_default_na=False)
    # Tableau "Download Data" is UTF-8 CSV; "Crosstab" CSV is UTF-16 tab-separated.
    for enc, sep in (("utf-8-sig", ","), ("utf-16", "\t"), ("cp1252", ",")):
        try:
            df = pd.read_csv(path, dtype=str, encoding=enc, sep=sep, keep_default_na=False)
            if df.shape[1] > 3:
                return df
        except (UnicodeError, pd.errors.ParserError):
            continue
    raise SystemExit(f"Could not read {path} as CSV.")


def map_columns(df: pd.DataFrame, name: str = "input") -> pd.DataFrame:
    """Standard column names. For each field use the first alias column that has any values."""
    lower = {}
    for c in df.columns:
        lower.setdefault(str(c).strip().lower(), c)
    out = pd.DataFrame(index=df.index)
    present = set()
    for key, names in ALIASES.items():
        out[key] = None
        for n in names:
            src = lower.get(n.lower())
            if src is None:
                continue
            present.add(key)
            s = df[src].astype(str).str.strip()
            s = s.where(~s.isin(["", "nan", "None", "NaT", "NULL", "null"]))
            if s.notna().any():
                out[key] = s
                break
    missing = [k for k in REQUIRED if k not in present]
    if missing:
        raise SystemExit(f"{name} is missing required columns: {', '.join(missing)}.\nColumns found: {list(df.columns)}")
    return out


def to_dt(s: pd.Series) -> pd.Series:
    s = s.astype(object).where(s.notna(), None)
    have = s.notna()
    if not have.any():
        return pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for fmt in ("%m/%d/%Y %H:%M", "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%d-%b-%Y %H:%M:%S",
                "%d-%b-%Y"):
        r = pd.to_datetime(s, format=fmt, errors="coerce")
        if r.notna().sum() >= have.sum() * 0.98:
            rest = have & r.isna()
            if rest.any():
                r[rest] = pd.to_datetime(s[rest], format="mixed", errors="coerce")
            return r.astype("datetime64[ns]")
    return pd.to_datetime(s, format="mixed", errors="coerce").astype("datetime64[ns]")


def pick_rows(raw: pd.DataFrame, stats: dict) -> pd.DataFrame:
    """Drop superseded record versions (I <> M) and, when a TLM matched more than one flight leg (the Tableau join
    has no date key), keep only the leg whose scheduled date is closest to the truck order date."""
    if raw["version"].notna().any() and raw["version_max"].notna().any():
        keep = raw["version"].isna() | raw["version_max"].isna() | (raw["version"] == raw["version_max"])
        stats["old_versions"] = stats.get("old_versions", 0) + int((~keep).sum())
        raw = raw[keep]
    if raw["tlm"].isna().all():
        return raw
    leg = (raw["sched_date"].fillna("") + "|" + raw["origin"].fillna("") + "|" + raw["dest"].fillna("")
           + "|" + raw["truck"].fillna(""))
    n_leg = leg.groupby(raw["tlm"]).transform("nunique")
    multi = raw["tlm"].notna() & (n_leg > 1)
    if multi.any():
        sub = raw[multi]
        dist = (to_dt(sub["sched_date"]).dt.normalize() - to_dt(sub["order_date"]).dt.normalize()).abs()
        dist = dist.dt.days.fillna(9999)
        chosen = (pd.DataFrame({"tlm": sub["tlm"], "leg": leg[multi], "dist": dist})
                  .sort_values("dist", kind="stable").groupby("tlm")["leg"].first())
        drop = multi & (leg != raw["tlm"].map(chosen))
        stats["multi_leg_tlms"] = stats.get("multi_leg_tlms", 0) + int(sub["tlm"].nunique())
        raw = raw[~drop]
    return raw


def collapse(raw: pd.DataFrame) -> pd.DataFrame:
    """One row per TLM. The Tableau download repeats each TLM once per pane (departure, spacer, arrival)
    and only fills the actual for its own pane, so take the first non-blank value of every column."""
    key = raw["tlm"].fillna(raw["truck"].fillna("") + "|" + raw["sched_date"].fillna("") + "|" + raw["origin"].fillna(""))
    return raw.groupby(key.rename("_key"), sort=False).first().reset_index(drop=True)


def expand_inputs(items) -> list[Path]:
    files = []
    for it in items:
        p = Path(it)
        if p.is_dir():
            files += [f for f in p.iterdir() if f.suffix.lower() in (".csv", ".txt", ".xlsx")
                      and not f.name.startswith(("~$", "Trucks Report", "Trucks History", "late_trucks_"))]
        elif p.exists():
            files.append(p)
        else:
            raise SystemExit(f"Input not found: {p}")
    if not files:
        raise SystemExit("No input files found.")
    return sorted(set(files), key=lambda f: f.stat().st_mtime)


def load_sources(files: list[Path], history: Path | None, log=print) -> tuple[pd.DataFrame, dict]:
    """All sources -> one row per TLM. Later sources win: history store first, then inputs oldest to newest;
    within that, the higher record version (I) wins."""
    stats = {"files": [], "raw_rows": 0}
    frames = []
    sources = ([(history, 0)] if history and history.exists() else []) + [(f, i + 1) for i, f in enumerate(files)]
    for path, rank in sources:
        raw = map_columns(read_any(path), path.name)
        n = len(raw)
        df = collapse(pick_rows(raw, stats))
        df["_rank"] = rank
        df["_src"] = df["source_file"].fillna(path.name)
        # record version (I) decides between copies of a TLM; rows without one count as of the file's time
        df["_ver"] = to_dt(df["version"]).fillna(pd.Timestamp(datetime.fromtimestamp(path.stat().st_mtime)))
        frames.append(df)
        stats["files"].append((path.name, n, len(df)))
        if rank:
            stats["raw_rows"] += n
        log(f"  {path.name}: {n:,} rows -> {len(df):,} trucks" + ("  (history store)" if rank == 0 else ""))
    df = pd.concat(frames, ignore_index=True)
    key = df["tlm"].fillna(df["truck"].fillna("") + "|" + df["sched_date"].fillna("") + "|" + df["origin"].fillna(""))
    df = df.assign(_key=key).sort_values(["_ver", "_rank"], kind="stable")
    before = len(df)
    df = df.drop_duplicates("_key", keep="last").drop(columns=["_key", "_ver"]).reset_index(drop=True)
    stats["overlap"] = before - len(df)
    return df, stats


# -------------------------------------------------------------------------------------------- prepare
def station_now(stations: pd.Series, as_of: datetime, local_tz: str) -> pd.Series:
    """As-of time expressed in each station's local clock (station times on the flight record are local)."""
    try:
        import airportsdata  # optional: pip install airportsdata
        extra = airportsdata.load("IATA")
    except Exception:
        extra = {}
    base = pd.Timestamp(as_of)
    try:
        here = base.tz_localize(local_tz)
    except Exception:
        return pd.Series(base, index=stations.index)
    cache = {}
    for st in stations.dropna().unique():
        tz = STATION_TZ.get(st) or (extra.get(st) or {}).get("tz")
        try:
            cache[st] = here.tz_convert(tz).tz_localize(None) if tz else base
        except Exception:
            cache[st] = base
    return stations.map(cache).fillna(base).astype("datetime64[ns]")


def prepare(df: pd.DataFrame, as_of: datetime, local_tz: str = LOCAL_TZ) -> pd.DataFrame:
    df = df.copy()
    df["sched_date"] = to_dt(df["sched_date"]).dt.normalize()
    sdt = to_dt(df["sched_dep_dttm"]) if df["sched_dep_dttm"].notna().any() else pd.Series(pd.NaT, index=df.index)
    if sdt.isna().any() and df["sched_dep_time"].notna().any():
        tod = to_dt(df["sched_dep_time"])
        sdt = sdt.fillna(df["sched_date"] + pd.to_timedelta(tod.dt.hour * 60 + tod.dt.minute, unit="m"))
    df["sched_dep"] = sdt.fillna(df["sched_date"]).astype("datetime64[ns]")
    for c in ("sched_arr", "act_dep", "act_arr"):
        df[c] = to_dt(df[c])
    df["truck"] = norm_truck(df["truck"])
    for c in ("origin", "dest"):
        df[c] = df[c].str.upper()
    df["sched_arr_sys"] = df["sched_arr"]
    df["team_arr"] = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    df["lane"] = df["origin"].fillna("") + "-" + df["dest"].fillna("")
    pre = df["ac_type"].fillna("").str.upper().str[:3]
    df["truck_type"] = pre.map({k: v[0] for k, v in TRUCK_TYPES.items()}).fillna("Standard")
    df["arr_rule"] = pre.map({k: v[1] for k, v in TRUCK_TYPES.items()}).fillna("Normal")
    df["now_o"] = station_now(df["origin"], as_of, local_tz)
    df["now_d"] = station_now(df["dest"], as_of, local_tz)
    return df


# ---------------------------------------------------------------------------------------------- score
def norm_truck(s):
    """Truck numbers as text: '7841', '7841.0' and 7841 all become '7841'; anonymized IDs are kept as they are."""
    if isinstance(s, pd.Series):
        return s.astype("string").str.strip().str.replace(r"\.0+$", "", regex=True).astype(object).where(s.notna(), None)
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    return re.sub(r"\.0+$", "", str(s).strip())


def _minutes(a: pd.Series, b: pd.Series) -> pd.Series:
    """Whole minutes between two times, counted like Tableau's DATEDIFF('minute', b, a) (minute boundaries)."""
    return ((a.dt.floor("min") - b.dt.floor("min")).dt.total_seconds() / 60).round().astype("Int64")


def score(df: pd.DataFrame, threshold: int, grace: int = GRACE_MIN) -> pd.DataFrame:
    """Departure / arrival status: Late, On time, Check date (actual keyed wrong), No actual (overdue), Pending."""
    df["dep_delay"] = _minutes(df["act_dep"], df["sched_dep"])
    df["arr_delay"] = _minutes(df["act_arr"], df["sched_arr"])
    g = pd.Timedelta(minutes=grace)

    dd = df["dep_delay"]
    dep = pd.Series("On time", index=df.index, dtype=object)
    dep[(dd > threshold).fillna(False)] = "Late"
    dep[((dd > LATE_CHECK_MIN) | (dd < -EARLY_CHECK_MIN)).fillna(False)] = "Check date"
    miss = df["act_dep"].isna()
    dep[miss] = "Pending"
    dep[miss & (df["sched_dep"] + g < df["now_o"])] = "No actual"
    df["dep_status"] = dep

    # Arrival sanity is judged on the trip itself, not the distance from the scheduled arrival: weekend
    # schedules commit to Monday, so a Saturday arrival 2+ days "early" is fine, but an arrival before
    # the truck left (or a month later) is a keying error.
    dep_ok = df["dep_status"].isin(["Late", "On time"])
    base = df["act_dep"].where(dep_ok, df["sched_dep"])
    df["transit_min"] = _minutes(df["act_arr"], base)
    ad = df["arr_delay"]
    arr = pd.Series("On time", index=df.index, dtype=object)
    arr[(ad > threshold).fillna(False)] = "Late"
    arr[((ad > LATE_CHECK_MIN) | (df["transit_min"] < MIN_TRANSIT_MIN)).fillna(False)] = "Check date"
    miss = df["act_arr"].isna()
    arr[miss] = "Pending"
    arr[miss & (df["sched_arr"] + g < df["now_d"])] = "No actual"
    rule = df["arr_rule"] if "arr_rule" in df else pd.Series("Normal", index=df.index)
    arr[(rule == "None") | ((rule == "STA") & ~miss)] = "Not tracked"
    df["arr_status"] = arr
    df["late"] = (df["dep_status"] == "Late") | (df["arr_status"] == "Late")
    return df


def lane_transit(df: pd.DataFrame, start=None, end=None, min_obs: int = 3) -> pd.Series:
    """Median actual transit (minutes) per lane over clean trips (departure and arrival both Late/On time)
    scheduled between start and end - the same rows as the workbook's Data tab."""
    ok = df["dep_status"].isin(["Late", "On time"]) & df["arr_status"].isin(["Late", "On time"])
    if start is not None:
        ok &= df["sched_date"] >= pd.Timestamp(start)
    if end is not None:
        ok &= df["sched_date"] <= pd.Timestamp(end)
    g = df[ok].groupby("lane")["transit_min"].agg(["median", "count"])
    return g.loc[g["count"] >= min_obs, "median"].astype(float)


def project(df: pd.DataFrame, transit: pd.Series, threshold: int) -> pd.DataFrame:
    """Projected arrival for trucks that have departed but not arrived: actual departure + the lane's usual
    transit (or the scheduled transit when the lane has fewer than 3 clean trips)."""
    open_ = df["dep_status"].isin(["Late", "On time"]) & df["act_arr"].isna() & (df["arr_rule"] == "Normal")
    lane_min = df["lane"].map(transit)
    planned = (df["sched_arr"] - df["sched_dep"]).dt.total_seconds() / 60
    use = lane_min.fillna(planned)
    use = np.sign(use) * np.floor(use.abs() + 0.5)                       # whole minutes, Excel ROUND
    proj = df["act_dep"] + pd.to_timedelta(use, unit="m")
    df["proj_arr"] = proj.where(open_).astype("datetime64[ns]")
    df["proj_basis"] = pd.Series(None, index=df.index, dtype=object)
    df.loc[open_ & lane_min.notna(), "proj_basis"] = "lane median"
    df.loc[open_ & lane_min.isna() & planned.notna(), "proj_basis"] = "scheduled transit"
    d = (df["proj_arr"] - df["sched_arr"]).dt.total_seconds() / 60       # Excel ROUND: half away from zero
    df["proj_arr_delay"] = (np.sign(d) * np.floor(d.abs() + 0.5 + 1e-9)).astype("Int64")
    return df


def issues(df: pd.DataFrame, threshold: int, lookback_h: int = LOOKBACK_H) -> pd.DataFrame:
    """Per-truck issue text plus the Action List code for open items in the lookback window."""
    departed = df["dep_status"].isin(["Late", "On time", "Check date"])
    dep_bad = df["dep_status"] == "Check date"
    arr_bad = df["arr_status"] == "Check date"
    arr_before = arr_bad & (df["transit_min"] < 0).fillna(False)
    arr_same = arr_bad & (df["transit_min"] >= 0).fillna(False) & (df["transit_min"] < MIN_TRANSIT_MIN).fillna(False)
    no_dep = df["dep_status"] == "No actual"
    no_arr = departed & (df["arr_status"] == "No actual")
    proj_late = (df["proj_arr_delay"] > threshold).fillna(False) & (df["arr_status"] == "Pending")

    def days_off(m):
        return (m.abs() / 1440).round(0).astype("Int64").astype(str)

    parts = [
        (no_dep, "No departure recorded"),
        (dep_bad, "Departure actual " + days_off(df["dep_delay"]) + " d off schedule"),
        (arr_before, pd.Series("Arrival before departure", index=df.index)),
        (arr_same, pd.Series("Arrival keyed same as departure", index=df.index)),
        (arr_bad & ~arr_before & ~arr_same, "Arrival actual " + days_off(df["arr_delay"]) + " d off schedule"),
        (no_arr, "Arrival not recorded"),
        (df["dep_status"] == "Late", pd.Series("Late departure", index=df.index)),
        (df["arr_status"] == "Late", pd.Series("Late arrival", index=df.index)),
        (proj_late, pd.Series("Projected late arrival", index=df.index)),
    ]
    txt = pd.Series("", index=df.index, dtype=object)
    for mask, label in parts:
        lab = label if isinstance(label, pd.Series) else pd.Series(label, index=df.index)
        txt = txt.where(~mask, txt.where(txt == "", txt + "; ") + lab)
    df["issue"] = txt
    df["data_issue"] = (dep_bad | arr_bad | no_dep | no_arr)

    lb = pd.Timedelta(hours=lookback_h)
    recent_dep = df["sched_dep"] >= df["now_o"] - lb
    recent_arr = df["sched_arr"] >= df["now_d"] - lb
    code = pd.Series(None, index=df.index, dtype=object)
    # assigned lowest priority first so the highest-priority match wins
    for c, m in [
        ("fix_record", (dep_bad & recent_dep) | (arr_bad & recent_arr)),
        ("dep_late", (df["dep_status"] == "Late") & (df["arr_status"] == "Pending") & ~proj_late),
        ("proj_late", proj_late & departed),
        ("arr_overdue", no_arr & recent_arr),
        ("not_departed", no_dep & recent_dep),
    ]:
        code[m] = c
    df["action"] = code
    wl = pd.Series("", index=df.index, dtype=object)
    for lab, m in reversed([
        ("No ATD - possible no-show", no_dep),
        ("No ATA - arrival overdue", (df["arr_status"] == "No actual") & ~no_dep),
        ("Fix record - date keyed wrong", dep_bad | arr_bad),
        ("Late arrival", df["arr_status"] == "Late"),
        ("Projected late", proj_late),
        ("Late departure", df["dep_status"] == "Late"),
    ]):
        wl[m.fillna(False)] = lab
    df["wl_issue"] = wl
    return df


# ------------------------------------------------------------------------------ report-card cross-check
def _hhmm_offset(v):
    """'400+1' / 1800 / time(18,0) -> minutes from midnight of day 0."""
    if v is None:
        return None
    if hasattr(v, "hour"):
        return v.hour * 60 + v.minute
    m = re.match(r"^\s*(\d{1,2}):?(\d{2})?\s*(?:\+\s*(\d))?\s*$", str(v).replace(".0", ""))
    if not m:
        m2 = re.match(r"^\s*(\d{1,4})\s*(?:\+\s*(\d))?\s*$", str(v).replace(".0", ""))
        if not m2:
            return None
        hhmm = int(m2.group(1))
        return (hhmm // 100) * 60 + hhmm % 100 + 1440 * int(m2.group(2) or 0)
    return int(m.group(1)) * 60 + int(m.group(2) or 0) + 1440 * int(m.group(3) or 0)


def _read_schedule(ws) -> dict:
    """Report Card 'Schedule' tab -> {(truck, origin, dest): (etd_min, eta_min)}; ambiguous keys are dropped."""
    out, hdr, bad = {}, None, set()
    for row in ws.iter_rows(values_only=True, max_col=12):
        if hdr is None:
            cells = [str(c).strip().upper() if c is not None else "" for c in row]
            if "ORIGIN" in cells and "ETD" in cells:
                hdr = {k: cells.index(k) for k in ("ORIGIN", "DESTINATION", "ETD", "ETA")}
                hdr["TRUCK"] = next(i for i, c in enumerate(cells) if c.startswith("TRUCK"))
            continue
        if row[hdr["TRUCK"]] is None:
            continue
        try:
            k = (norm_truck(row[hdr["TRUCK"]]), str(row[hdr["ORIGIN"]]).strip().upper(),
                 str(row[hdr["DESTINATION"]]).strip().upper())
        except (TypeError, ValueError):
            continue
        v = (_hhmm_offset(row[hdr["ETD"]]), _hhmm_offset(row[hdr["ETA"]]))
        if None in v:
            continue
        if k in out and out[k] != v:
            bad.add(k)
        out[k] = v
    for k in bad:
        out.pop(k, None)
    return out


def load_report_card(path: Path | None, log=print) -> tuple[set, dict]:
    """(truck, departure date) pairs already logged on the Truck Report Card, plus its Schedule tab. The sheets
    declare ~1M rows of formatting, so read in read-only mode and stop after a long run of blank rows."""
    if not path or not Path(path).exists():
        return set(), {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # locked / no openpyxl - never block the report
        log(f"  report card not read ({e}); 'On Report Card' and team-schedule checks left blank")
        return set(), {}
    schedule = {}
    for ws in wb.worksheets:
        if ws.title.strip().lower() == "schedule":
            try:
                schedule = _read_schedule(ws)
            except Exception as e:
                log(f"  schedule tab not read ({e})")
    keys = set()
    for ws in [w for w in wb.worksheets if w.title.lower().startswith("report card")]:
        hdr, blank = None, 0
        for row in ws.iter_rows(values_only=True, max_col=30):
            if hdr is None:
                cells = [str(c).strip().upper() if c is not None else "" for c in row]
                if "TRK NBR" in cells and "DEPARTURE DATE" in cells:
                    hdr = (cells.index("TRK NBR"), cells.index("DEPARTURE DATE"))
                continue
            trk, dt = row[hdr[0]], row[hdr[1]]
            if trk is None and dt is None:
                blank += 1
                if blank > 200:
                    break
                continue
            blank = 0
            try:
                keys.add((norm_truck(trk), pd.Timestamp(dt).normalize()))
            except Exception:
                pass
    wb.close()
    log(f"  report card: {len(keys):,} logged truck-days, {len(schedule):,} schedule lanes read from {Path(path).name}")
    return keys, schedule


def apply_team_schedule(df: pd.DataFrame, schedule: dict, basis: str, threshold: int) -> pd.DataFrame:
    """Scheduled arrival per the team's Report Card 'Schedule' tab (ETA with +N day). The flight record's
    SCHD_ARVL_DTTM is sometimes later than the published schedule, which hides late arrivals.
    basis='team' scores arrivals against the team schedule where the lane is on it."""
    if not schedule or not len(df):
        return df
    eta = [schedule.get((t, o, d), (None, None))[1] if t is not None else None
           for t, o, d in zip(df["truck"], df["origin"].fillna(""), df["dest"].fillna(""))]
    df["team_arr"] = (df["sched_date"] + pd.to_timedelta(pd.Series(eta, index=df.index, dtype="float64"), unit="m")
                      ).astype("datetime64[ns]")
    if basis == "team":
        df["sched_arr"] = df["team_arr"].fillna(df["sched_arr_sys"])
        df = score(df, threshold)
    return df


# ----------------------------------------------------------------------------------------- helpers
def rc_time(ts, base) -> str:
    """Report Card style: 2030, 1305+1 (day offset from the scheduled departure date)."""
    if pd.isna(ts) or pd.isna(base):
        return ""
    d = (pd.Timestamp(ts).normalize() - pd.Timestamp(base).normalize()).days
    return f"{pd.Timestamp(ts):%H%M}" + (f"+{d}" if d > 0 else (str(d) if d < 0 else ""))


def _intish(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return "" if v is None or (isinstance(v, float) and pd.isna(v)) else v


def summarize(df: pd.DataFrame, by, label, top=None) -> pd.DataFrame:
    by = [by] if isinstance(by, str) else list(by)
    label = [label] if isinstance(label, str) else list(label)
    x = df.assign(
        _dl=df.dep_status.eq("Late"), _al=df.arr_status.eq("Late"), _lt=df.late,
        _am=df.arr_status.isin(["Late", "On time"]),
        _dq=df.dep_status.isin(["Check date", "No actual"]) | df.arr_status.eq("Check date"),
        _lm=df.arr_delay.where(df.arr_status.eq("Late")).astype("Float64"),
    )
    g = x.groupby(by, dropna=False)
    out = pd.DataFrame({
        "Trucks": g.size(), "Late departures": g["_dl"].sum(), "Late arrivals": g["_al"].sum(),
        "Late (dep or arr)": g["_lt"].sum(),
        "Arrivals measured": g["_am"].sum(), "Median min late (late arrivals)": g["_lm"].median(),
        "Check date / no departure": g["_dq"].sum(),
    })
    meas = out["Arrivals measured"].where(out["Arrivals measured"] > 0)
    out["On-time arrival %"] = 1 - out["Late arrivals"] / meas
    out["Late departure %"] = out["Late departures"] / out["Trucks"]
    out = out.sort_values(["Late arrivals", "Trucks"], ascending=False)
    out = (out.head(top) if top else out).reset_index().rename(columns=dict(zip(by, label)))
    return out


def repeat_late(df: pd.DataFrame, min_late: int = 3) -> pd.DataFrame:
    """Truck numbers late (departure or arrival) on min_late or more days in the window."""
    if not len(df):
        return pd.DataFrame(columns=["Truck", "Lane", "Vendor", "Trips", "Late trips", "Late %",
                                     "Late departures", "Late arrivals", "Median min late"])
    x = df[df["truck"].notna()].assign(
        _lm=df[["dep_delay", "arr_delay"]].where(df[["dep_status", "arr_status"]].eq("Late").values)
        .max(axis=1).astype("Float64"))
    g = x.groupby(["truck", "lane"], dropna=False)
    out = pd.DataFrame({
        "Vendor": g["vendor"].agg(lambda s: s.dropna().mode().iat[0] if s.notna().any() else ""),
        "Trips": g.size(), "Late trips": g["late"].sum(),
        "Late departures": g["dep_status"].agg(lambda s: (s == "Late").sum()),
        "Late arrivals": g["arr_status"].agg(lambda s: (s == "Late").sum()),
        "Median min late": g["_lm"].median(),
    })
    out = out[out["Late trips"] >= min_late]
    out.insert(3, "Late %", out["Late trips"] / out["Trips"])
    out = out.sort_values(["Late trips", "Late %"], ascending=False).reset_index()
    return out.rename(columns={"truck": "Truck", "lane": "Lane"})[
        ["Truck", "Lane", "Vendor", "Trips", "Late trips", "Late %", "Late departures", "Late arrivals",
         "Median min late"]]


def tableau_check(df: pd.DataFrame) -> dict:
    """Compare our Late/On time with Tableau's exported RED/GREEN colors (system basis)."""
    out = {}
    for side, col in (("dep", "tab_dep_color"), ("arr", "tab_arr_color")):
        if col not in df or df[col].isna().all():
            continue
        c = df[col].fillna("").str.upper()
        st = df[f"{side}_status"]
        both = c.isin(["RED", "GREEN"]) & st.isin(["Late", "On time"])
        agree = both & ((c == "RED") == (st == "Late"))
        out[side] = {"compared": int(both.sum()), "agree": int(agree.sum()),
                     "check_date_tableau_colored": int((c.isin(["RED", "GREEN"]) & (st == "Check date")).sum())}
    return out


# --------------------------------------------------------------------------------------------- history
HISTORY_COLUMNS = [  # (output header, field) - headers are also input aliases so the file can be re-read
    ("Scheduled Departure Date", "sched_date"), ("Truck", "truck"), ("Vendor Code", "vendor_code"),
    ("Vendor Name", "vendor"), ("Origin", "origin"), ("Destination", "dest"), ("Lane", "lane"),
    ("AC Type", "ac_type"), ("Truck Type", "truck_type"), ("Scheduled Departure", "sched_dep"), ("Scheduled Arrival", "sched_arr"),
    ("TLM", "tlm"), ("Actual Departure", "act_dep"), ("Actual Arrival", "act_arr"),
    ("Dep Delay (min)", "dep_delay"), ("Dep Status", "dep_status"), ("Arr Delay (min)", "arr_delay"),
    ("Arr Status", "arr_status"), ("Late", "late_i"), ("Late Departure", "dep_late_i"),
    ("Late Arrival", "arr_late_i"), ("Arrival Measured", "arr_meas_i"), ("Planned Transit (h)", "plan_h"),
    ("Actual Transit (h)", "transit_h"), ("Issue", "issue"), ("Data Issue", "data_issue_i"),
    ("Year", "year"), ("Month", "month"), ("Week Start", "week"), ("Weekday", "weekday"),
    ("Order Date", "order_date"), ("Record Version", "version"), ("Tableau Dep Color", "tab_dep_color"),
    ("Tableau Arr Color", "tab_arr_color"), ("Source File", "_src"),
]


def history_frame(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    x["late_i"] = x["late"].astype(int)
    x["dep_late_i"] = (x["dep_status"] == "Late").astype(int)
    x["arr_late_i"] = (x["arr_status"] == "Late").astype(int)
    x["arr_meas_i"] = x["arr_status"].isin(["Late", "On time"]).astype(int)
    x["data_issue_i"] = x["data_issue"].astype(int)
    x["plan_h"] = ((x["sched_arr"] - x["sched_dep"]).dt.total_seconds() / 3600).round(2)
    x["transit_h"] = (x["transit_min"].astype("Float64") / 60).round(2).where(x["arr_meas_i"] == 1)
    x["year"] = x["sched_date"].dt.year
    x["month"] = x["sched_date"].dt.strftime("%Y-%m")
    x["week"] = (x["sched_date"] - pd.to_timedelta(x["sched_date"].dt.weekday, unit="D")).dt.strftime("%Y-%m-%d")
    x["weekday"] = x["sched_date"].dt.strftime("%a")
    x["sched_date"] = x["sched_date"].dt.strftime("%Y-%m-%d")
    for c in ("sched_dep", "sched_arr", "act_dep", "act_arr"):
        x[c] = x[c].dt.strftime("%Y-%m-%d %H:%M")
    out = x[[f for _, f in HISTORY_COLUMNS]].copy()
    out.columns = [h for h, _ in HISTORY_COLUMNS]
    return out


def write_history(df: pd.DataFrame, folder: Path, as_of: datetime, log=print) -> dict:
    folder.mkdir(parents=True, exist_ok=True)
    df = df.sort_values(["sched_date", "truck", "tlm"]).reset_index(drop=True)
    h = history_frame(df)
    path = folder / HISTORY_FILE
    tmp = path.with_suffix(".tmp")
    h.to_csv(tmp, index=False, encoding="utf-8-sig")
    tmp.replace(path)
    out = {"history": path}
    try:
        h.to_parquet(folder / "late_trucks_history.parquet", index=False)
        out["parquet"] = folder / "late_trucks_history.parquet"
    except Exception:
        pass

    daily = summarize(df, ["sched_date", "origin", "dest", "vendor"], ["Date", "Origin", "Destination", "Vendor"])
    daily = daily.sort_values(["Date", "Origin", "Destination"]).reset_index(drop=True)
    daily["Date"] = pd.to_datetime(daily["Date"]).dt.strftime("%Y-%m-%d")
    daily.insert(3, "Lane", daily["Origin"].fillna("") + "-" + daily["Destination"].fillna(""))
    daily.to_csv(folder / "late_trucks_daily.csv", index=False, encoding="utf-8-sig")
    out["daily"] = folder / "late_trucks_daily.csv"

    try:
        import formula_book as FB
        out["summary"] = FB.build_history_book(df, folder / "Trucks History Summary.xlsx", as_of)
    except PermissionError:
        log("  Trucks History Summary.xlsx is open - close it to refresh")
    log(f"  history: {len(df):,} trucks {df['sched_date'].min():%Y-%m-%d} to {df['sched_date'].max():%Y-%m-%d} "
        f"-> {path}")
    return out


# ----------------------------------------------------------------------------------------------- feed
def build_feed(df: pd.DataFrame, meta: dict) -> dict:
    def iso(x):
        return None if pd.isna(x) else pd.Timestamp(x).strftime("%Y-%m-%dT%H:%M")

    def ival(x):
        return None if pd.isna(x) else int(x)

    rows = df.iloc[0:0]
    if len(df):
        recent_from = df["sched_date"].max() - pd.Timedelta(days=FEED_DAYS - 1)
        backlog_from = df["sched_date"].max() - pd.Timedelta(days=FEED_BACKLOG_DAYS - 1)
        rows = df[(df["sched_date"] >= recent_from) | df["action"].notna()
                  | (df["late"] & ~df["logged"] & (df["sched_date"] >= backlog_from))]
    stations = {}
    for t in df.itertuples(index=False):
        for role, st in (("out", t.origin), ("in", t.dest)):
            stations.setdefault(st, {"station": st, "out": 0, "in": 0, "dep_late": 0, "arr_late": 0})[role] += 1
        stations[t.origin]["dep_late"] += t.dep_status == "Late"
        stations[t.dest]["arr_late"] += t.arr_status == "Late"
    labels = {c: (p, lab) for c, p, lab, _ in ACTIONS}

    def rec(frame):
        return json.loads(frame.to_json(orient="records", date_format="iso"))

    return {
        "schema": "aa.late_trucks.v1", "generated": meta["as_of"], "source": meta["source"],
        "window": meta["window"], "late_threshold_min": meta["threshold"], "kpis": meta["kpis"],
        "by_station": sorted(stations.values(), key=lambda s: -(s["dep_late"] + s["arr_late"])),
        "by_vendor": rec(meta["summaries"][0][1]), "by_lane": rec(meta["summaries"][2][1]),
        "repeat_late": rec(meta["summaries"][3][1]),
        "trucks": [{
            "date": iso(t.sched_date)[:10] if pd.notna(t.sched_date) else None,
            "truck": int(t.truck) if str(t.truck).isdigit() else t.truck, "tlm": t.tlm,
            "vendor": t.vendor, "vendor_code": t.vendor_code, "o": t.origin, "d": t.dest, "type": t.ac_type,
            "sched_dep": iso(t.sched_dep), "act_dep": iso(t.act_dep), "dep_delay": ival(t.dep_delay),
            "dep_status": t.dep_status, "sched_arr": iso(t.sched_arr), "act_arr": iso(t.act_arr),
            "arr_delay": ival(t.arr_delay), "arr_status": t.arr_status,
            "proj_arr": iso(t.proj_arr), "proj_arr_delay": ival(t.proj_arr_delay),
            "late": bool(t.late), "on_report_card": bool(t.logged), "issue": t.issue or None,
            "action": labels[t.action][1] if isinstance(t.action, str) else None,
            "priority": labels[t.action][0] if isinstance(t.action, str) else None,
        } for t in rows.itertuples(index=False)],
    }


# ----------------------------------------------------------------------------------------------- main
def window_bounds(df: pd.DataFrame, a) -> tuple:
    """Report window (scheduled departure dates). Default: the last N days through the as-of day (or the last day in
    the data if older), plus tomorrow's schedule."""
    if not (a.days or a.date_from or a.date_to or getattr(a, "all", False)):
        a.days = DEFAULT_DAYS
    if not len(df):
        return pd.Timestamp(a.as_of_dt).normalize(), pd.Timestamp(a.as_of_dt).normalize()
    anchor = min(df["sched_date"].max(), pd.Timestamp(a.as_of_dt).normalize())
    if a.days:
        a.date_from = a.date_from or (anchor - pd.Timedelta(days=a.days - 1)).strftime("%Y-%m-%d")
        a.date_to = a.date_to or (anchor + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    lo = pd.Timestamp(a.date_from) if a.date_from else df["sched_date"].min()
    hi = pd.Timestamp(a.date_to) if a.date_to else min(df["sched_date"].max(), anchor + pd.Timedelta(days=1))
    return lo, hi


def filter_rows(df: pd.DataFrame, a) -> pd.DataFrame:
    for col, val in (("origin", a.origin), ("dest", a.dest)):
        if val:
            df = df[df[col].fillna("").str.upper().isin([v.strip().upper() for v in val.split(",")])]
    if a.vendor:
        df = df[df["vendor"].fillna("").str.upper().str.contains(a.vendor.upper(), regex=False)]
    if a.truck:
        df = df[df["truck"].fillna("").isin([norm_truck(t) for t in a.truck.split(",")])]
    return df


def station_info(stations, as_of: datetime, local_tz: str) -> tuple[dict, dict]:
    """{station: time zone}, {station: hours its clock is ahead of the as-of clock}."""
    try:
        import airportsdata
        extra = airportsdata.load("IATA")
    except Exception:
        extra = {}
    st = pd.Series(sorted(set(stations)), dtype=object)
    now = station_now(st, as_of, local_tz)
    tz = {s: STATION_TZ.get(s) or (extra.get(s) or {}).get("tz") or "" for s in st}
    off = {s: (n - pd.Timestamp(as_of)).total_seconds() / 3600 for s, n in zip(st, now)}
    return tz, off


def previous_entries(out: Path, exclude: Path, log=print) -> tuple[dict, dict]:
    """Coordinator entries (Worklist) and contacts (Vendors / Stations emails) from the newest earlier Trucks Report in
    the output folder."""
    import formula_book as FB
    prev = sorted((p for p in out.glob("Trucks Report *.xlsx") if p != exclude and not p.name.startswith("~$")),
                  key=lambda p: p.stat().st_mtime)
    if not prev:
        return {}, {}
    notes, contacts = FB.read_notes(prev[-1]), FB.read_contacts(prev[-1])
    if notes or contacts:
        log(f"  carried {len(notes):,} worklist entries and {len(contacts):,} contacts forward from {prev[-1].name}")
    return notes, contacts


def build(a, log=print) -> dict:
    import formula_book as FB
    items = getattr(a, "inputs", None) or [a.input]
    files = expand_inputs(items)
    as_of = (datetime.strptime(a.as_of, "%Y-%m-%d %H:%M") if a.as_of
             else datetime.fromtimestamp(max(f.stat().st_mtime for f in files)))
    a.as_of_dt = as_of
    threshold = a.threshold
    local_tz = getattr(a, "local_tz", None) or LOCAL_TZ
    grace = getattr(a, "grace", None)
    grace = GRACE_MIN if grace is None else grace
    lookback = getattr(a, "lookback", None) or LOOKBACK_H
    data_days = getattr(a, "data_days", None) or DATA_DAYS
    hist_dir = None if getattr(a, "no_history", False) else Path(getattr(a, "history_dir", None) or HERE / "History")
    hist_path = hist_dir / HISTORY_FILE if hist_dir else None

    base, st = load_sources(files, hist_path, log)
    df_all = prepare(base, as_of, local_tz)
    no_flight = int(df_all["sched_date"].isna().sum())
    df_all = df_all[df_all["sched_date"].notna()].reset_index(drop=True)  # Tableau 'Exclude Nulls': no flight leg
    if not a.as_of and len(df_all):
        # Guard: if departures were recorded well after the file's time stamp (copied file, shifted/anonymized dates),
        # use the latest plausible actual departure as the clock instead.
        ok = df_all["act_dep"].notna() & ((df_all["act_dep"] - df_all["sched_dep"]).abs() <= pd.Timedelta(days=2))
        latest = df_all.loc[ok, "act_dep"].max()
        if pd.notna(latest) and latest > pd.Timestamp(as_of) + pd.Timedelta(days=1):
            as_of = (latest + pd.Timedelta(hours=1)).floor("h").to_pydatetime()
            a.as_of_dt = as_of
            log(f"  NOTE: the export has departures up to {latest:%m/%d/%Y %H:%M}, after the file time; "
                f"as-of set to {as_of:%m/%d/%Y %H:%M} (use --as-of to override)")
            df_all = prepare(base, as_of, local_tz)
            df_all = df_all[df_all["sched_date"].notna()].reset_index(drop=True)
    df_all = score(df_all, threshold, grace)
    log(f"  {len(df_all):,} trucks loaded ({no_flight:,} TLMs with no flight leg dropped"
        + (f", {st['overlap']:,} duplicates across files" if st.get("overlap") else "") + ")")

    # Data tab rows: the report window plus enough history for daily / weekly trends and lane medians
    win_from, win_to = window_bounds(df_all, a)
    anchor = min(win_to, pd.Timestamp(as_of).normalize()) if len(df_all) else win_to
    data_from = min(win_from, anchor - pd.Timedelta(days=data_days - 1))
    df_data = filter_rows(df_all[(df_all["sched_date"] >= data_from) & (df_all["sched_date"] <= win_to)], a).copy()
    rc_path = (None if (getattr(a, "report_card", None) or "").lower() == "none"
               else Path(a.report_card) if getattr(a, "report_card", None) else DEFAULT_REPORT_CARD)
    keys, schedule = load_report_card(rc_path, log) if (rc_path and not getattr(a, "history_only", False)) else (set(), {})
    df_data = apply_team_schedule(df_data, schedule, a.arr_basis, threshold)
    transit = lane_transit(df_data, min_obs=MIN_LANE_TRIPS)
    df_data = issues(project(df_data, transit, threshold), threshold, lookback)
    df_data["logged"] = [t is not None and (t, d) in keys for t, d in zip(df_data["truck"], df_data["sched_date"])]
    df_data["in_window"] = (df_data["sched_date"] >= win_from) & (df_data["sched_date"] <= win_to)
    df_all = issues(project(df_all, transit, threshold), threshold, lookback)
    df_all["logged"] = False

    result = {"df_all": df_all, "as_of": as_of}
    if hist_dir:
        result.update(write_history(df_all, hist_dir, as_of, log))
    if getattr(a, "history_only", False):
        return result

    df = df_data[df_data["in_window"]].sort_values(["sched_date", "truck"], ascending=[False, True])
    n = len(df)
    meas = int(df["arr_status"].isin(["Late", "On time"]).sum())
    kpis = {
        "trucks": n,
        "dep_late": int((df["dep_status"] == "Late").sum()),
        "arr_late": int((df["arr_status"] == "Late").sum()),
        "late": int(df["late"].sum()),
        "ontime_arr_pct": round(1 - (df["arr_status"] == "Late").sum() / meas, 4) if meas else None,
        "late_not_logged": int((df["late"] & ~df["logged"]).sum()),
        "no_dep": int((df["dep_status"] == "No actual").sum()),
        "no_arr": int(((df["arr_status"] == "No actual") & (df["dep_status"] != "No actual")).sum()),
        "check_date": int(((df["dep_status"] == "Check date") | (df["arr_status"] == "Check date")).sum()),
        "to_work": int((df["wl_issue"] != "").sum()),
        "not_departed": int((df["action"] == "not_departed").sum()),
        "arr_overdue": int((df["action"] == "arr_overdue").sum()),
        "proj_late": int((df["action"] == "proj_late").sum()),
        "dep_late_open": int((df["action"] == "dep_late").sum()),
        "fix_record": int((df["action"] == "fix_record").sum()),
        "sched_arr_later_than_team": int(((df["sched_arr_sys"] - df["team_arr"]).dt.total_seconds() >= 3600).sum()),
        "arr_basis": a.arr_basis,
    }
    window = {"from": f"{win_from:%Y-%m-%d}", "to": f"{win_to:%Y-%m-%d}"}
    summaries = [("By vendor", summarize(df, "vendor", "Vendor")),
                 ("By origin station", summarize(df, "origin", "Origin")),
                 ("By lane (top 40 by late arrivals)", summarize(df, "lane", "Lane", top=40)),
                 ("Repeat late trucks (late on 3+ days)", repeat_late(df))]
    tc = tableau_check(df_data)
    tc_txt = "; ".join(f"{'departures' if s == 'dep' else 'arrivals'}: {v['agree']:,} of {v['compared']:,} agree "
                       f"with Tableau's RED/GREEN ({v['agree'] / v['compared']:.1%})"
                       + (f", {v['check_date_tableau_colored']:,} that Tableau colored are Check date here"
                          if v["check_date_tableau_colored"] else "")
                       for s, v in tc.items() if v["compared"]) or "No Tableau color columns in the input."
    src_names = ", ".join(f"{name} ({n_raw:,} rows)" for name, n_raw, _ in st["files"])
    meta = {"as_of": as_of.strftime("%m/%d/%Y %I:%M %p"), "source": ", ".join(f.name for f in files),
            "window": window, "threshold": threshold, "kpis": kpis, "summaries": summaries}

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx = out / f"Trucks Report {as_of:%Y-%m-%d %H%M}.xlsx"
    found_notes, found_contacts = previous_entries(out, xlsx, log)
    notes = FB.merge_notes(out / "worklist_notes.csv", found_notes, as_of)
    contacts = FB.merge_contacts(out / "contacts.csv", found_contacts)
    tz, off = station_info(list(df_data["origin"].dropna()) + list(df_data["dest"].dropna())
                           + list(FB.OFFLINE_REGION), as_of, local_tz)
    s = {"as_of": pd.Timestamp(as_of), "as_of_txt": meta["as_of"], "local_tz": local_tz, "threshold": threshold,
         "grace": grace, "early_check": EARLY_CHECK_MIN, "late_check": LATE_CHECK_MIN, "min_transit": MIN_TRANSIT_MIN,
         "min_lane_trips": MIN_LANE_TRIPS, "win_from": win_from, "win_to": win_to,
         "data_span": f"scheduled {data_from:%m/%d/%Y} to {win_to:%m/%d/%Y}", "lane_median": transit,
         "station_tz": tz, "station_offset": off, "has_rc": bool(keys), "sources": src_names,
         "n_data": len(df_data), "check_txt": tc_txt, "contacts": contacts}
    from shift_compare import compare_snapshots, make_snapshot, save_snapshot, select_prior
    snapshot_dir = Path(getattr(a, "snapshots_dir", None) or out.parent / "snapshots")
    snapshot = make_snapshot(df_data, s)
    previous = select_prior(snapshot_dir, as_of, getattr(a, "vs", "prior-shift"))
    if previous is not None:
        s["shift_comparison"] = compare_snapshots(previous, snapshot)
    else:
        log("  no earlier snapshot for this comparison; this run will establish the baseline")
    try:
        FB.build_daily_book(df_data, xlsx, s, notes)
    except PermissionError:
        raise SystemExit(f"{xlsx} is open in Excel - close it and run again.")
    snapshot_path = save_snapshot(snapshot, snapshot_dir)
    log(f"  saved {snapshot_path}")
    feed = build_feed(df, meta)
    (out / "late_trucks_feed.json").write_text(json.dumps(feed, indent=1), encoding="utf-8")
    (out / "late_trucks_feed.js").write_text("window.LATE_TRUCKS_FEED = " + json.dumps(feed) + ";\n", encoding="utf-8")
    log(f"  window {window['from']} to {window['to']}: {kpis['trucks']:,} trucks, {kpis['dep_late']:,} late departures, "
        f"{kpis['arr_late']:,} late arrivals, {kpis['check_date']:,} suspect dates")
    log(f"  worklist: {kpis['to_work']:,} trucks to work ({kpis['no_dep']} no ATD, {kpis['no_arr']} no ATA)")
    log(f"  Tableau check - {tc_txt}")
    log(f"  wrote {xlsx}")
    log(f"  wrote {out / 'late_trucks_feed.json'} (+ .js)")
    if getattr(a, "hub", None):
        hub = Path(a.hub)
        hub = hub.parent if hub.suffix.lower() in (".html", ".htm") else hub
        try:
            hub.mkdir(parents=True, exist_ok=True)
            (hub / "late_trucks_feed.js").write_text((out / "late_trucks_feed.js").read_text(encoding="utf-8"), encoding="utf-8")
            log(f"  copied feed to {hub} (Coordinator Hub reads it on reload)")
        except OSError as e:
            log(f"  could not copy feed to {hub}: {e}")
    result.update({"xlsx": xlsx, "df": df, "df_data": df_data, "feed": feed, "settings": s, "notes": notes,
                   "snapshot": snapshot_path, "shift_comparison": s.get("shift_comparison")})
    return result


def parse_args(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("inputs", nargs="+", help="Tableau export(s) (CSV/XLSX), SQL result files, or folders of them")
    p.add_argument("--from", dest="date_from", help="first scheduled departure date (YYYY-MM-DD)")
    p.add_argument("--to", dest="date_to", help="last scheduled departure date (YYYY-MM-DD)")
    p.add_argument("--days", type=int, help=f"last N days (default {DEFAULT_DAYS} when no other window is given)")
    p.add_argument("--all", action="store_true", help="report on every date in the input (large files: slow)")
    p.add_argument("--origin", help="comma list, e.g. JFK,ORD")
    p.add_argument("--dest", help="comma list")
    p.add_argument("--vendor", help="text contained in the vendor name")
    p.add_argument("--truck", help="comma list of truck numbers")
    p.add_argument("--threshold", type=int, default=LATE_MIN, help="late if more than N minutes (default 59)")
    p.add_argument("--grace", type=int, default=GRACE_MIN,
                   help=f"minutes after schedule before a missing actual counts as No actual (default {GRACE_MIN})")
    p.add_argument("--lookback", type=int, default=LOOKBACK_H, help=f"Action List hours (default {LOOKBACK_H})")
    p.add_argument("--local-tz", default=LOCAL_TZ, help=f"time zone of this PC's clock (default {LOCAL_TZ})")
    p.add_argument("--report-card", help="Truck Report Card workbook, or 'none' to skip the cross-check")
    p.add_argument("--arr-basis", choices=["system", "team"], default="system",
                   help="score arrivals vs flight-record schedule (system, = Tableau) or the Report Card Schedule tab ETA")
    p.add_argument("--as-of", help="'YYYY-MM-DD HH:MM' (default: newest input file time)")
    p.add_argument("--out", default=str(HERE / "Output"), help="output folder")
    p.add_argument("--hub", help="Coordinator Hub folder (or .html): also write late_trucks_feed.js there")
    p.add_argument("--history-dir", default=str(HERE / "History"), help="history store folder (default History)")
    p.add_argument("--no-history", action="store_true", help="do not read or update the history store")
    p.add_argument("--history-only", action="store_true", help="only update the history store (backfill)")
    p.add_argument("--data-days", type=int, default=DATA_DAYS,
                   help=f"days of trucks on the workbook's Data tab (default {DATA_DAYS})")
    p.add_argument("--snapshots-dir", help="snapshot folder (default: snapshots next to Output)")
    p.add_argument("--vs", default="prior-shift", help="compare to prior-shift, prior-day, or an earlier report/snapshot path")
    return p.parse_args(argv)


if __name__ == "__main__":
    print("Trucks Report builder")
    build(parse_args())
