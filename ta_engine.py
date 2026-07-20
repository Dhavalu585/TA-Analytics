"""
TA Dashboard Generator - Recruiting Intelligence Engine
=======================================================

Locked business rules
---------------------

Candidate Joined detail report
  * Column H  = Primary Recruiter
  * Column W  = Job_App_for_Hire_1: Source          (Hiring Source)
  * Column X  = Job_App_for_Hire_1: Hiring Manager  (as text)
  * Column AA = Department Name
  * Column AB = Regular Hire flag  -> STRICT filter: keep ONLY rows containing "regular"
  * Column AD = Head of Department (Business Owner)
  * Column AK = Total Days to Onboard  (TTF numerator)
  * Column AM = Region

Job Req & open Pos
  * Column A  = Open Job Requisition: Reference ID -> Open Reqs (unique)
  * Column B  = Position ID                        -> Positions open WITHOUT a Req raised
  * Column P  = Business Head
  * Column X  = Worker Type                        -> STRICT Regular filter
  * Column AK = Candidate stage                    -> Pure Open vs Advanced Stage
  * Column AM = Days bracket                       -> Aging bucket
  * Column AO = Region

Candidate by Status         = Advanced-stage candidate pipeline
Disposition Reasons Report  = Did-not-join after offer

TTF formula:
    Portfolio TTF = SUM(Col AK Total Days to Onboard for Regular hires)
                    / COUNT(Regular onboards with valid TTF)
Cuts: Region, Primary Recruiter, Department, Head of Department, Job Family.
Median TTF has been REMOVED per user request.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
SHEET_MAP = {
    "joined":           "Candidate Joined detail report",
    "open_reqs":        "Job Req Report & Open Pos",
    "candidate_status": "Candidate by Status",
    "dispositions":     "Disposition Reasons Report",
}

# 0-based column indices from user
COLPOS = {
    "joined": {
        "primary_recruiter":     7,   # H
        "hiring_source":        22,   # W = Job_App_for_Hire_1: Source
        "hiring_manager":       23,   # X = Job_App_for_Hire_1: Hiring Manager (as text)
        "department":           26,   # AA
        "regular_flag":         27,   # AB - STRICT Regular filter
        "head_of_department":   29,   # AD - Business Owner
        "total_days_to_onboard":36,   # AK - TTF numerator
        "region":               38,   # AM
        "hire_date":            1,    # B - kept for period grouping
    },
    "open_reqs": {
        "req_id":               0,   # A
        "position_title":       1,   # B - "Position" (title text, e.g. "#39492 Senior ... (Unfilled)")
        "position_id":          2,   # C - "Position ID" (e.g. "#39492")
        "primary_recruiter":    12,  # M
        "business_head":        15,  # P
        "worker_type":          23,  # X
        "declines_advanced":    31,  # AF - "Number of declines from advanced stage"
        "candidate_stage":      36,  # AK
        "days_bracket":         38,  # AM
        "region":               40,  # AO
    },
    "dispositions": {
        "disposition_reason":   7,   # H
    },
}

HEADER_ALIASES = {
    "department":         ["department name", "department", "function", "business unit"],
    "region":             ["region", "country/region", "geo", "country"],
    "quarter":            ["quarter", "qtr"],
    "month":              ["month"],
    "hiring_manager":     ["job_app_for_hire_1: hiring manager", "hiring manager", "manager"],
    "job_family":         ["job family", "job family group", "job profile"],
    "requisition_type":   ["requisition type", "req type", "job type"],
    "recruiting_manager": ["recruiting manager", "recruiter manager", "ta manager", "ta lead"],
    "target_hire_date":   ["target hire date", "target start date", "expected hire date"],
    "approval_status":    ["approval status", "req status", "requisition status"],
    "req_id":             ["open job requisition: reference id", "job requisition: reference id",
                           "job requisition", "requisition id", "req id", "reference id"],
    "primary_recruiter":  ["primary recruiter", "recruiter", "lead recruiter"],
    "hiring_source":      ["job_app_for_hire_1: source", "hiring source", "source"],
    "candidate_stage":    ["candidate stage", "stage", "current stage", "application stage"],
    "head_of_department": ["head of department", "hod", "department head"],
}

# STRICT Regular filter — Option A: keep only rows explicitly containing "regular"
REGULAR_INCLUDE_TERMS = ["regular"]

ADVANCED_STAGE_KEYWORDS = [
    "reference check", "reference",
    "background check", "background", "bgc",
    "offer",
    "ea", "executive assessment", "executive approval",
    "ready for hire", "ready", "pre-hire", "prehire",
]
PURE_OPEN_KEYWORDS = ["open", "no candidate", "sourcing", "screening in progress",
                       "not started", "new"]

# Ordered stage funnel for Column AK (Candidate stage) on Job Req & open Pos.
# Order matters for the funnel chart - roughly earliest -> latest.
STAGE_DETAIL_ORDER = [
    "Open / No Candidate",
    "Screening / Interview",
    "Reference Check",
    "Background Check",
    "Offer",
    "Executive Assessment (EA)",
    "Ready for Hire",
    "Other",
]

STAGE_DETAIL_KEYWORDS = [
    ("Ready for Hire",              ["ready for hire", "ready", "pre-hire", "prehire"]),
    ("Executive Assessment (EA)",   ["executive assessment", "executive approval", " ea "]),
    ("Offer",                       ["offer"]),
    ("Background Check",            ["background check", "background", "bgc"]),
    ("Reference Check",             ["reference check", "reference"]),
    ("Screening / Interview",       ["screen", "phone", "interview", "panel", "assessment", "shortlist"]),
    ("Open / No Candidate",         ["open", "no candidate", "sourcing", "screening in progress",
                                      "not started", "new", "", "nan", "none", "unknown"]),
]


# ---------------------------------------------------------------------------
def _clean(x) -> str:
    return re.sub(r"\s+", " ", str(x).strip().lower())


def _match_sheet(xl: pd.ExcelFile, expected: str) -> Optional[str]:
    ec = _clean(expected)
    for s in xl.sheet_names:
        if _clean(s) == ec:
            return s
    for s in xl.sheet_names:
        if ec in _clean(s) or _clean(s) in ec:
            return s
    # Fallback: token-overlap match. Handles renamed sheets like
    # "Job Req & Open Pos" -> "Job Req Report & Open Pos" where an extra
    # word breaks the plain substring check above.
    ec_tokens = set(re.findall(r"[a-z0-9]+", ec)) - {"the", "and", "&"}
    if not ec_tokens:
        return None
    best, best_score = None, 0.0
    for s in xl.sheet_names:
        s_tokens = set(re.findall(r"[a-z0-9]+", _clean(s))) - {"the", "and", "&"}
        if not s_tokens:
            continue
        overlap = len(ec_tokens & s_tokens) / len(ec_tokens | s_tokens)
        if overlap > best_score:
            best, best_score = s, overlap
    return best if best_score >= 0.5 else None


def col_by_pos(df: pd.DataFrame, sheet_key: str, logical: str) -> Optional[str]:
    idx = COLPOS.get(sheet_key, {}).get(logical)
    if idx is not None and idx < len(df.columns):
        return df.columns[idx]
    return None


def col_by_header(df: pd.DataFrame, logical: str) -> Optional[str]:
    aliases = HEADER_ALIASES.get(logical, [logical])
    lookup = {_clean(c): c for c in df.columns}
    for alias in aliases:
        a = _clean(alias)
        for k, orig in lookup.items():
            if a == k:
                return orig
    for alias in aliases:
        a = _clean(alias)
        for k, orig in lookup.items():
            if a in k or k in a:
                return orig
    return None


def as_str(df: pd.DataFrame, colname: Optional[str], fillna: str = "Unknown") -> pd.Series:
    if colname and colname in df.columns:
        return df[colname].astype(str).str.strip().replace(
            {"nan": fillna, "": fillna, "None": fillna, "NaT": fillna})
    return pd.Series([fillna] * len(df), index=df.index)


def as_datetime(df: pd.DataFrame, colname: Optional[str]) -> pd.Series:
    if colname and colname in df.columns:
        return pd.to_datetime(df[colname], errors="coerce")
    return pd.Series(pd.NaT, index=df.index)


def as_numeric(df: pd.DataFrame, colname: Optional[str]) -> pd.Series:
    if colname and colname in df.columns:
        return pd.to_numeric(df[colname], errors="coerce")
    return pd.Series(np.nan, index=df.index)


# ---------------------------------------------------------------------------
@dataclass
class TABundle:
    raw: Dict[str, pd.DataFrame] = field(default_factory=dict)
    detected: Dict[str, str] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)


def load_workbook(path: str | Path) -> TABundle:
    xl = pd.ExcelFile(path, engine="openpyxl")
    bundle = TABundle()
    for key, expected in SHEET_MAP.items():
        actual = _match_sheet(xl, expected)
        if not actual:
            bundle.warnings.append(f"Sheet not found: {expected}")
            bundle.raw[key] = pd.DataFrame()
            continue
        df = xl.parse(actual)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        bundle.raw[key] = df
        bundle.detected[key] = actual
    return bundle


# ---------------------------------------------------------------------------
def apply_strict_regular_filter(df: pd.DataFrame, filter_col: Optional[str]) -> Tuple[pd.DataFrame, str]:
    """STRICT Option A: keep ONLY rows where filter_col explicitly contains 'regular'."""
    if df.empty:
        return df, "No data to filter."
    if not filter_col or filter_col not in df.columns:
        return df.iloc[0:0].copy(), f"STRICT filter: column '{filter_col}' NOT found -> 0 rows kept."
    s = df[filter_col].astype(str).str.strip().str.lower().fillna("")
    mask = s.str.contains("regular", na=False)
    kept = int(mask.sum())
    return df.loc[mask].copy(), f"STRICT Regular filter on '{filter_col}' -> {kept} rows kept."


# ---------------------------------------------------------------------------
# Prepare: Joined (Candidate Joined detail report)
# ---------------------------------------------------------------------------
def prepare_joined(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    notes: List[str] = []
    if df.empty:
        return df, ["Candidate Joined detail report missing/empty."]

    out = df.copy()

    recruiter_c   = col_by_pos(out, "joined", "primary_recruiter")
    source_c      = col_by_pos(out, "joined", "hiring_source")
    hm_c          = col_by_pos(out, "joined", "hiring_manager")
    dept_c        = col_by_pos(out, "joined", "department")
    reg_flag_c    = col_by_pos(out, "joined", "regular_flag")
    hod_c         = col_by_pos(out, "joined", "head_of_department")
    tto_c         = col_by_pos(out, "joined", "total_days_to_onboard")
    region_c      = col_by_pos(out, "joined", "region")
    hire_date_c   = col_by_pos(out, "joined", "hire_date")

    out["_primary_recruiter"]     = as_str(out, recruiter_c)
    out["_hiring_source"]         = as_str(out, source_c)
    out["_hiring_manager"]        = as_str(out, hm_c)
    out["_department"]            = as_str(out, dept_c)
    out["_head_of_department"]    = as_str(out, hod_c)
    out["_region"]                = as_str(out, region_c)
    out["_hire_date"]             = as_datetime(out, hire_date_c)
    out["_total_days_to_onboard"] = as_numeric(out, tto_c)

    # Optional headers
    out["_recruiting_manager"] = as_str(out, col_by_header(out, "recruiting_manager"))
    out["_job_family"]         = as_str(out, col_by_header(out, "job_family"))
    out["_requisition_type"]   = as_str(out, col_by_header(out, "requisition_type"))

    # STRICT Regular Hire filter via Column AB
    notes.append(f"[Joined] Column AB detected as: '{reg_flag_c}'")
    out, m = apply_strict_regular_filter(out, reg_flag_c)
    notes.append(f"[Joined] {m}")

    # Clean TTF (remove blank / negative / >365)
    out["_ttf_valid"]   = out["_total_days_to_onboard"]
    out["_ttf_outlier"] = out["_ttf_valid"].isna() | (out["_ttf_valid"] < 0) | (out["_ttf_valid"] > 365)
    out["_ttf_clean"]   = out["_ttf_valid"].where(~out["_ttf_outlier"], np.nan)

    # Time periods based on Hire Date
    if "_hire_date" in out.columns and out["_hire_date"].notna().any():
        out["_month"]   = out["_hire_date"].dt.to_period("M").astype(str)
        out["_quarter"] = out["_hire_date"].dt.to_period("Q").astype(str)
        out["_year"]    = out["_hire_date"].dt.year

    return out, notes


# ---------------------------------------------------------------------------
# TTF calculators (Median removed)
# ---------------------------------------------------------------------------
def portfolio_ttf(df: pd.DataFrame) -> Dict[str, float]:
    """Portfolio TTF = SUM(Col AK) / COUNT(Regular onboards with valid TTF)."""
    col = "_ttf_clean"
    if df.empty or col not in df.columns:
        return {"Total Days": 0, "Onboards": 0, "TTF": np.nan}
    vals = pd.to_numeric(df[col], errors="coerce").dropna()
    if len(vals) == 0:
        return {"Total Days": 0, "Onboards": 0, "TTF": np.nan}
    return {
        "Total Days": float(vals.sum()),
        "Onboards": int(len(vals)),
        "TTF": round(float(vals.sum() / len(vals)), 1),
    }


def ttf_by_dimension(df: pd.DataFrame, dim_col: str, label: str,
                      top: int = 25) -> pd.DataFrame:
    """Portfolio TTF sliced by any dimension. Median TTF removed."""
    if df.empty or dim_col not in df.columns or "_ttf_clean" not in df.columns:
        return pd.DataFrame(columns=[label, "Onboards", "Total Days", "TTF"])
    tmp = df[[dim_col, "_ttf_clean"]].copy()
    tmp["_ttf_clean"] = pd.to_numeric(tmp["_ttf_clean"], errors="coerce")
    tmp = tmp.dropna(subset=["_ttf_clean"])
    if tmp.empty:
        return pd.DataFrame(columns=[label, "Onboards", "Total Days", "TTF"])
    g = tmp.groupby(dim_col)["_ttf_clean"].agg(
        Onboards="count",
        Total_Days="sum",
        TTF="mean",
    ).round(1).reset_index()
    g.columns = [label, "Onboards", "Total Days", "TTF"]
    return g.sort_values("TTF", ascending=False).head(top)


# ---------------------------------------------------------------------------
def _normalize_aging(x) -> str:
    s = str(x).strip()
    if s.lower() in ("", "nan", "none", "unknown"):
        return "Unknown"
    # Real format: "1, 0-89 Days" / "2, 90-179 Days" / "3, 180+ Days" -
    # keep the label, drop the leading sort-index prefix.
    m = re.match(r"^\s*\d+\s*,\s*(.+?)\s*$", s)
    if m:
        return m.group(1)
    # Fallback for exports that only give a raw day count or a different format.
    nums = [int(n) for n in re.findall(r"\d+", s)]
    if len(nums) >= 2:
        hi = nums[1]
        if hi <= 89:  return "0-89 Days"
        if hi <= 179: return "90-179 Days"
        return "180+ Days"
    if nums:
        n = nums[0]
        if n >= 180: return "180+ Days"
        if n >= 90:  return "90-179 Days"
        return "0-89 Days"
    if "+" in s or ">" in s: return "180+ Days"
    return s


AGING_ORDER = ["0-89 Days", "90-179 Days", "180+ Days", "Unknown"]


def _is_aged_bucket(bucket: str) -> bool:
    """True for the two oldest buckets, matched by substring so it survives
    minor label formatting differences across exports."""
    s = str(bucket)
    return ("90-179" in s) or ("180+" in s)


def _classify_open_stage(stage: str) -> str:
    s = str(stage).strip().lower()
    if s in ("", "nan", "none", "unknown"):
        return "Pure Open"
    for kw in ADVANCED_STAGE_KEYWORDS:
        if kw in s:
            return "Advanced Stage"
    for kw in PURE_OPEN_KEYWORDS:
        if kw and kw in s:
            return "Pure Open"
    if any(k in s for k in ["screen", "phone", "interview", "panel", "assessment", "shortlist"]):
        return "Interview / Active"
    return "Other"


def _classify_stage_detail(stage: str) -> str:
    """Full stage-funnel classification for Col AK (Candidate stage) on Open Reqs.
    Distinguishes Open / Screening / Reference / Background / Offer / EA / Ready for Hire,
    per the user's description of the stages actually present in that column."""
    s = f" {str(stage).strip().lower()} "
    for label, kws in STAGE_DETAIL_KEYWORDS:
        for kw in kws:
            if kw and (kw in s if kw.strip() else s.strip() == ""):
                return label
    return "Other"


def prepare_open_reqs(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    notes: List[str] = []
    if df.empty:
        return df, ["Job Req & open Pos missing/empty."]
    out = df.copy()

    out["_req_id"]           = as_str(out, col_by_pos(out, "open_reqs", "req_id"), fillna="")
    out["_position_id"]      = as_str(out, col_by_pos(out, "open_reqs", "position_id"), fillna="")
    out["_declines_advanced"] = as_numeric(out, col_by_pos(out, "open_reqs", "declines_advanced"))
    out["_business_head"]    = as_str(out, col_by_pos(out, "open_reqs", "business_head"))
    out["_candidate_stage"]  = as_str(out, col_by_pos(out, "open_reqs", "candidate_stage"), fillna="")
    out["_days_bracket_raw"] = as_str(out, col_by_pos(out, "open_reqs", "days_bracket"))
    out["_region"]           = as_str(out, col_by_pos(out, "open_reqs", "region"))
    out["_aging_bucket"]     = out["_days_bracket_raw"].apply(_normalize_aging)

    out["_department"]         = as_str(out, col_by_header(out, "department"))
    out["_hiring_manager"]     = as_str(out, col_by_header(out, "hiring_manager"))
    # Primary Recruiter = Column M (position first, header fallback if M is missing/blank)
    rec_col = col_by_pos(out, "open_reqs", "primary_recruiter") or col_by_header(out, "primary_recruiter")
    out["_primary_recruiter"]  = as_str(out, rec_col)
    out["_recruiting_manager"] = as_str(out, col_by_header(out, "recruiting_manager"))
    out["_job_family"]         = as_str(out, col_by_header(out, "job_family"))
    out["_requisition_type"]   = as_str(out, col_by_header(out, "requisition_type"))
    out["_approval_status"]    = as_str(out, col_by_header(out, "approval_status"))
    out["_target_hire_date"]   = as_datetime(out, col_by_header(out, "target_hire_date"))

    # STRICT Regular filter via Column X
    reg_col = col_by_pos(out, "open_reqs", "worker_type")
    notes.append(f"[Open Reqs] Column X detected as: '{reg_col}'")
    out, m = apply_strict_regular_filter(out, reg_col)
    notes.append(f"[Open Reqs] {m}")

    out["_stage_class"]  = out["_candidate_stage"].apply(_classify_open_stage)
    out["_stage_detail"] = out["_candidate_stage"].apply(_classify_stage_detail)

    req_blank = out["_req_id"].astype(str).str.strip().isin(["", "nan", "None", "Unknown"])
    pos_present = ~out["_position_id"].astype(str).str.strip().isin(["", "nan", "None", "Unknown"])
    out["_position_without_req"] = req_blank & pos_present

    out["_no_candidate_activity"] = out["_stage_class"].eq("Pure Open")
    if out["_target_hire_date"].notna().any():
        out["_past_target_hire"] = out["_target_hire_date"] < pd.Timestamp.today()
    else:
        out["_past_target_hire"] = np.nan
    return out, notes


# ---------------------------------------------------------------------------
def prepare_funnel(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    if df.empty:
        return df, ["Candidate by Status missing/empty."]
    out = df.copy()
    stage_c  = col_by_header(out, "candidate_stage")
    source_c = col_by_header(out, "hiring_source")

    out["_stage_raw"]     = as_str(out, stage_c, fillna="Advanced")
    out["_hiring_source"] = as_str(out, source_c)
    out["_department"]     = as_str(out, col_by_header(out, "department"))
    out["_region"]         = as_str(out, col_by_header(out, "region"))
    out["_hiring_manager"] = as_str(out, col_by_header(out, "hiring_manager"))
    out["_recruiter"]      = as_str(out, col_by_header(out, "primary_recruiter"))

    s = out["_stage_raw"].str.lower()
    out["_funnel_stage"] = "Advanced / Active"
    out.loc[s.str.contains("screen|phone", na=False), "_funnel_stage"]    = "Screen"
    out.loc[s.str.contains("interview|panel", na=False), "_funnel_stage"] = "Interview"
    out.loc[s.str.contains("offer", na=False), "_funnel_stage"]           = "Offer"
    out.loc[s.str.contains("ready|pre-hire|prehire|background", na=False), "_funnel_stage"] = "Ready for Hire"
    out.loc[s.str.contains("hire|joined", na=False), "_funnel_stage"]     = "Hired"
    return out, []


def prepare_dispositions(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    if df.empty:
        return df, ["Disposition Reasons Report missing/empty."]
    out = df.copy()
    reason_c = col_by_pos(out, "dispositions", "disposition_reason")
    out["_disposition_reason"] = as_str(out, reason_c, fillna="Unknown")
    out["_did_not_join_after_offer"] = out["_disposition_reason"].str.strip().str.lower().replace(
        {"unknown": ""}) != ""
    req_c = col_by_header(out, "req_id")
    if req_c:
        out["_req_id"] = as_str(out, req_c)
    # Enrich with dimensions (header-detected, best-effort) so declines can be
    # correlated against aging/department/region/recruiter/manager cuts.
    out["_department"]        = as_str(out, col_by_header(out, "department"))
    out["_region"]            = as_str(out, col_by_header(out, "region"))
    out["_hiring_manager"]    = as_str(out, col_by_header(out, "hiring_manager"))
    out["_primary_recruiter"] = as_str(out, col_by_header(out, "primary_recruiter"))
    out["_hiring_source"]     = as_str(out, col_by_header(out, "hiring_source"))
    return out, []


# ---------------------------------------------------------------------------
# INSIGHT ENGINE
# ---------------------------------------------------------------------------
# Implements the "So What?" leadership-relevance filter from the
# recruitment-dashboard skill: every insight below only surfaces when it
# clears a materiality bar (a minimum sample size + a minimum effect size),
# and is phrased as Observation -> Measured Change -> Business Impact.

AGED_BUCKET_MATCH = ("90-179", "180+")   # kept for backward compat / readability
MIN_SAMPLE = 5          # ignore cuts with too few reqs/candidates to mean anything
MIN_AGED_SHARE_FLAG = 0.35   # flag a cut only if >=35% of its reqs are aged 90+


def aging_vs_declines(open_reqs: pd.DataFrame, disp: pd.DataFrame,
                       dim: str, label: str) -> pd.DataFrame:
    """Correlate requisition aging with post-offer declines for a given cut
    (Department, Region, Hiring Manager, Recruiter, Source).

    Uses TWO decline signals:
      1. '_declines_advanced' - the native per-req "Number of declines from
         advanced stage" column on the Open Reqs tab itself (most reliable -
         no join needed, always in sync with the aging bucket on that row).
      2. Offer Declines from the Disposition Reasons Report, joined by
         whatever dimension value matches (secondary / supporting signal,
         since not every open req has a matching disposition record).

    Returns one row per dimension value: open reqs, % aged 90+ days,
    declines from both signals, used to answer: 'which reqs are aged, and
    did that correlate with offer declines?'
    """
    if open_reqs.empty or dim not in open_reqs.columns:
        return pd.DataFrame()
    ou = open_reqs.drop_duplicates("_req_id")
    ou = ou[ou["_req_id"].astype(str).str.strip() != ""]

    agg = {"Open_Reqs": ("_req_id", "nunique"),
           "Aged_Reqs": ("_aging_bucket", lambda s: s.isin(
               [b for b in AGING_ORDER if _is_aged_bucket(b)]).sum())}
    if "_declines_advanced" in ou.columns:
        agg["Declines_On_Reqs"] = ("_declines_advanced", "sum")
    base = ou.groupby(dim).agg(**agg).reset_index()
    cols = [label, "Open Reqs", "Aged Reqs (90+)"]
    if "_declines_advanced" in ou.columns:
        cols.append("Declines (on open reqs)")
    base.columns = cols
    base["Aged Share"] = (base["Aged Reqs (90+)"] / base["Open Reqs"]).round(2)

    if not disp.empty and dim in disp.columns:
        dc = disp.groupby(dim)["_did_not_join_after_offer"].sum().reset_index()
        dc.columns = [label, "Offer Declines (disposition report)"]
        base = base.merge(dc, on=label, how="left")
    else:
        base["Offer Declines (disposition report)"] = 0
    base["Offer Declines (disposition report)"] = base["Offer Declines (disposition report)"].fillna(0).astype(int)
    if "Declines (on open reqs)" in base.columns:
        base["Declines (on open reqs)"] = base["Declines (on open reqs)"].fillna(0).astype(int)
        base["Total Declines"] = base["Declines (on open reqs)"] + base["Offer Declines (disposition report)"]
    else:
        base["Total Declines"] = base["Offer Declines (disposition report)"]
    base["Decline Rate vs Open Reqs"] = (base["Total Declines"] / base["Open Reqs"]).round(2)

    base = base[base["Open Reqs"] >= MIN_SAMPLE]
    return base.sort_values("Aged Share", ascending=False)


def _pct_change(new: float, old: float) -> Optional[float]:
    if old in (0, None) or pd.isna(old) or pd.isna(new):
        return None
    return round((new - old) / old * 100, 1)


def qoq_trend(df: pd.DataFrame, value_type: str = "count",
              value_col: Optional[str] = None) -> pd.DataFrame:
    """Quarter-over-quarter series. value_type='count' counts rows per _quarter;
    value_type='mean' averages value_col (e.g. TTF) per _quarter."""
    if df.empty or "_quarter" not in df.columns:
        return pd.DataFrame(columns=["Quarter", "Value"])
    if value_type == "mean" and value_col:
        g = df.groupby("_quarter")[value_col].mean().round(1).reset_index()
    else:
        g = df.groupby("_quarter").size().reset_index(name="Value_tmp")
        g.columns = ["_quarter", "Value_tmp"]
    g.columns = ["Quarter", "Value"]
    return g.sort_values("Quarter")


def trend_insights(joined: pd.DataFrame, open_reqs: pd.DataFrame,
                    disp: pd.DataFrame, dims: List[Tuple[str, str]]) -> List[str]:
    """Generate QoQ narrative insights per dimension, gated by the 'So What?'
    materiality filter: only report a cut if it has a sustained (not one-off)
    and sizeable (effect + sample) shift in volume, TTF, or decline rate.
    Candidates are scored and only the top few per section are returned -
    a flat threshold alone still produces too many items when a whole org
    is aging (as here), so ranking matters as much as gating.
    Returns ready-to-display strings, Observation -> Change -> Impact."""
    # Hiring volume + TTF trend by cut, QoQ.
    # The most recent quarter in the data may still be in progress (e.g. today
    # is 3 weeks into the quarter) - comparing a partial quarter to a complete
    # one manufactures a false "decline" for almost every cut. Only compare
    # two quarters that are both complete.
    trend_candidates: List[Tuple[float, str]] = []
    if not joined.empty and "_quarter" in joined.columns:
        quarters = sorted(joined["_quarter"].dropna().unique())
        current_q = str(pd.Timestamp.today().to_period("Q"))
        comparable = [q for q in quarters if q != current_q]
        if len(comparable) >= 2:
            q_last, q_prev = comparable[-1], comparable[-2]
            for dim, label in dims:
                if dim not in joined.columns:
                    continue
                cur = joined[joined["_quarter"] == q_last]
                prv = joined[joined["_quarter"] == q_prev]
                g_cur = cur.groupby(dim).size()
                g_prv = prv.groupby(dim).size()
                for key in g_cur.index:
                    n_cur, n_prv = g_cur.get(key, 0), g_prv.get(key, 0)
                    if n_cur < MIN_SAMPLE and n_prv < MIN_SAMPLE:
                        continue
                    pct = _pct_change(n_cur, n_prv)
                    if pct is None or abs(pct) < 25:
                        continue
                    ttf_cur = cur[cur[dim] == key]["_ttf_clean"].mean()
                    ttf_prv = prv[prv[dim] == key]["_ttf_clean"].mean()
                    ttf_note = ""
                    if pd.notna(ttf_cur) and pd.notna(ttf_prv) and abs(ttf_cur - ttf_prv) >= 5:
                        direction = "increased" if ttf_cur > ttf_prv else "improved"
                        ttf_note = (f" Portfolio TTF for {key} also {direction} from "
                                    f"{ttf_prv:.0f} to {ttf_cur:.0f} days.")
                    direction = "increased" if pct > 0 else "declined"
                    text = (
                        f"**{label}: {key}** — hiring volume {direction} from {n_prv} to {n_cur} "
                        f"({pct:+.1f}%) {q_prev} -> {q_last}.{ttf_note}"
                    )
                    score = abs(pct) * min(1.0, (n_cur + n_prv) / 20) + (10 if ttf_note else 0)
                    trend_candidates.append((score, text))

    # Aging vs decline correlation, flagged only when aged share clears the bar,
    # then ranked so only the highest-risk cuts surface.
    aging_candidates: List[Tuple[float, str]] = []
    if not open_reqs.empty:
        for dim, label in dims:
            corr = aging_vs_declines(open_reqs, disp, dim, label)
            if corr.empty:
                continue
            risky = corr[(corr["Aged Share"] >= MIN_AGED_SHARE_FLAG) & (corr["Total Declines"] > 0)]
            for _, row in risky.iterrows():
                text = (
                    f"**{label}: {row[label]}** — {row['Aged Reqs (90+)']} of {row['Open Reqs']} open reqs "
                    f"({row['Aged Share']*100:.0f}%) are aged 90+ days, alongside {row['Total Declines']} "
                    f"decline(s) in this cut — aging and offer breakage are compounding here, "
                    f"raising delivery risk."
                )
                score = row["Aged Share"] * 50 + min(row["Total Declines"], 20)
                aging_candidates.append((score, text))

    out: List[str] = []
    out += [t for _, t in sorted(trend_candidates, key=lambda x: -x[0])[:6]]
    out += [t for _, t in sorted(aging_candidates, key=lambda x: -x[0])[:6]]
    return out

def aged_reqs_with_dispositions(open_reqs: pd.DataFrame, disp: pd.DataFrame) -> pd.DataFrame:
    """Direct Req ID join: which aged (90+ day) open reqs have a matching
    record in the Disposition Reasons Report (i.e. this req already saw a
    candidate decline / drop out at some point). This is the most precise
    version of the aging-vs-decline question - no dimension aggregation,
    just the actual matching requisitions."""
    if open_reqs.empty or disp.empty or "_req_id" not in open_reqs.columns \
            or "_req_id" not in disp.columns:
        return pd.DataFrame()
    ou = open_reqs.drop_duplicates("_req_id")
    aged = ou[ou["_aging_bucket"].apply(_is_aged_bucket)]
    aged = aged[aged["_req_id"].astype(str).str.strip() != ""]
    if aged.empty:
        return pd.DataFrame()
    d = disp[disp["_did_not_join_after_offer"]]
    if d.empty:
        return pd.DataFrame()
    reasons = d.groupby("_req_id")["_disposition_reason"].apply(
        lambda s: "; ".join(sorted(set(s)))).reset_index()
    reasons.columns = ["_req_id", "Disposition Reason(s)"]
    merged = aged.merge(reasons, on="_req_id", how="inner")
    cols = ["_req_id", "_department", "_region", "_primary_recruiter",
            "_aging_bucket", "_candidate_stage", "Disposition Reason(s)"]
    cols = [c for c in cols if c in merged.columns]
    out = merged[cols].rename(columns={
        "_req_id": "Req ID", "_department": "Department", "_region": "Region",
        "_primary_recruiter": "Recruiter", "_aging_bucket": "Aging Bucket",
        "_candidate_stage": "Current Stage",
    })
    return out


def recruiter_ttf_outliers(joined: pd.DataFrame) -> pd.DataFrame:
    """Recruiters whose Portfolio TTF is meaningfully worse than the team
    average, with enough onboards to be a real signal (not noise)."""
    t = ttf_by_dimension(joined, "_primary_recruiter", "Recruiter", top=1000)
    if t.empty:
        return t
    t = t[t["Onboards"] >= MIN_SAMPLE]
    if t.empty:
        return t
    team_avg = t["TTF"].mean()
    t["Vs Team Avg (days)"] = (t["TTF"] - team_avg).round(1)
    return t[t["Vs Team Avg (days)"] >= 10].sort_values("Vs Team Avg (days)", ascending=False)


# ---------------------------------------------------------------------------
def unique_count(df: pd.DataFrame, colname: str) -> int:
    if df.empty or colname not in df.columns:
        return 0
    s = df[colname].dropna().astype(str).str.strip()
    s = s[(s != "") & (s.str.lower() != "nan") & (s.str.lower() != "unknown")]
    return int(s.nunique())


def group_count(df: pd.DataFrame, colname: str, label: str,
                unique_col: Optional[str] = None, top: int = 25) -> pd.DataFrame:
    if df.empty or colname not in df.columns:
        return pd.DataFrame(columns=[label, "Count"])
    if unique_col and unique_col in df.columns:
        g = df.groupby(colname)[unique_col].nunique().reset_index()
        g.columns = [label, "Count"]
    else:
        g = df[colname].fillna("Unknown").astype(str).value_counts().rename_axis(label)\
                                                                    .reset_index(name="Count")
    return g.sort_values("Count", ascending=False).head(top)


def apply_dimension_filters(df: pd.DataFrame, filters: Dict[str, List[str]]) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    for column, values in filters.items():
        if not values:
            continue
        if column in out.columns:
            out = out[out[column].astype(str).isin([str(v) for v in values])]
    return out


def build_audit(bundle: TABundle, joined: pd.DataFrame, open_reqs: pd.DataFrame,
                funnel: pd.DataFrame, disp: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for key, expected in SHEET_MAP.items():
        rdf = bundle.raw.get(key, pd.DataFrame())
        rows.append({"Area": key, "Expected Sheet": expected,
                     "Detected Sheet": bundle.detected.get(key, "MISSING"),
                     "Raw Rows": len(rdf), "Columns": len(rdf.columns)})
    if not joined.empty:
        ttf = portfolio_ttf(joined)
        rows.append({"Area": "Regular Onboards (after strict Col AB filter)",
                     "Expected Sheet": "Candidate Joined detail report",
                     "Detected Sheet": "(strict 'regular' filter)",
                     "Raw Rows": len(joined), "Columns": ""})
        rows.append({"Area": "TTF: Total Days (Col AK sum)",
                     "Expected Sheet": "",
                     "Detected Sheet": "Total Days to Onboard",
                     "Raw Rows": int(ttf["Total Days"]), "Columns": ""})
        rows.append({"Area": "TTF: Onboards used in denominator",
                     "Expected Sheet": "",
                     "Detected Sheet": "valid Regular with TTF in [0,365]",
                     "Raw Rows": int(ttf["Onboards"]), "Columns": ""})
        rows.append({"Area": "TTF: Portfolio Average",
                     "Expected Sheet": "",
                     "Detected Sheet": "SUM / COUNT",
                     "Raw Rows": ttf["TTF"] if pd.notna(ttf["TTF"]) else "N/A",
                     "Columns": ""})
        rows.append({"Area": "TTF outliers excluded",
                     "Expected Sheet": "",
                     "Detected Sheet": "blank / <0 / >365",
                     "Raw Rows": int(joined["_ttf_outlier"].sum()) if "_ttf_outlier" in joined else 0,
                     "Columns": ""})
    if not open_reqs.empty:
        rows.append({"Area": "Open Requisitions (unique Col A)",
                     "Expected Sheet": "Job Req & open Pos",
                     "Detected Sheet": "unique Req ID",
                     "Raw Rows": unique_count(open_reqs, "_req_id"), "Columns": ""})
        rows.append({"Area": "  · Pure Open",
                     "Expected Sheet": "",
                     "Detected Sheet": "Column AK = Pure Open",
                     "Raw Rows": int((open_reqs["_stage_class"] == "Pure Open").sum()),
                     "Columns": ""})
        rows.append({"Area": "  · Advanced Stage",
                     "Expected Sheet": "",
                     "Detected Sheet": "Ref/BGC/Offer/EA/Ready",
                     "Raw Rows": int((open_reqs["_stage_class"] == "Advanced Stage").sum()),
                     "Columns": ""})
        rows.append({"Area": "Positions without Requisition",
                     "Expected Sheet": "Job Req & open Pos",
                     "Detected Sheet": "Col B populated, A blank",
                     "Raw Rows": int(open_reqs["_position_without_req"].sum())
                                 if "_position_without_req" in open_reqs else 0,
                     "Columns": ""})
    if not disp.empty:
        rows.append({"Area": "Did-not-join after offer",
                     "Expected Sheet": "Disposition Reasons Report",
                     "Detected Sheet": "Column H non-blank",
                     "Raw Rows": int(disp["_did_not_join_after_offer"].sum()),
                     "Columns": ""})
    return pd.DataFrame(rows)
