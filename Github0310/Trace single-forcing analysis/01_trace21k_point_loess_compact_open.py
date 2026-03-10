# -*- coding: utf-8 -*-
import re
from pathlib import Path

import numpy as np
import pandas as pd

# ====== Edit here only if needed ======
SCRIPT_DIR = Path(__file__).resolve().parent
INPUT_XLSX = SCRIPT_DIR / "TraCE21K_TS_point_28.7N_117.2E_annual_merged.xlsx"
SHEET_NAME = 0
SPAN = 0.20
DEGREE = 1
# ======================================

OUT_XLSX = SCRIPT_DIR / f"{INPUT_XLSX.stem}_loess{SPAN:.2f}_compact.xlsx"

# Candidate age-axis columns: prefer time_kyr_* first, then year_like
YEARLIKE_CANDIDATES = ["year_like"]
TIME_PREFIX = "time_kyr"
TIME_MERGE_TOL = 1e-6  # kyr; 1e-6 kyr = 0.001 yr, used to judge whether time_kyr_* columns can be treated as identical

def loess_skmisc(x: np.ndarray, y: np.ndarray, span: float, degree: int) -> np.ndarray:
    """Use scikit-misc LOESS; positions with NaN in y remain NaN in the output."""
    try:
        from skmisc.loess import loess
    except ImportError as e:
        raise ImportError(
            "scikit-misc is required for true LOESS:\n"
            "pip install scikit-misc\n"
        ) from e

    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    out = np.full_like(y, np.nan, dtype=float)
    m = np.isfinite(x) & np.isfinite(y)
    if m.sum() < 10:
        return out

    xs = x[m]
    ys = y[m]

    order = np.argsort(xs)
    xs2 = xs[order]
    ys2 = ys[order]

    model = loess(xs2, ys2, span=span, degree=degree)
    model.fit()
    pred_sorted = model.predict(xs2).values

    idx_m = np.where(m)[0]
    out[idx_m[order]] = pred_sorted
    return out

def _max_abs_diff(a: np.ndarray, b: np.ndarray) -> float:
    m = np.isfinite(a) & np.isfinite(b)
    if m.sum() == 0:
        return 0.0
    return float(np.nanmax(np.abs(a[m] - b[m])))

def pick_yearlike_col(df: pd.DataFrame) -> str | None:
    for c in YEARLIKE_CANDIDATES:
        if c in df.columns:
            return c
    return None

def list_time_cols(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if str(c).startswith(TIME_PREFIX)]

def is_ts_series_col(c: str) -> bool:
    # Only capture TS_*_C_* temperature series columns (exclude loess columns)
    s = str(c)
    if not s.startswith("TS_"):
        return False
    if "_C_" not in s:
        return False
    if "loess" in s.lower():
        return False
    return True

def guess_forcing_suffix(ts_col: str) -> str | None:
    # Identify the forcing suffix at the end: ..._GHG / ..._ORB / ..._ICE / ..._FWF
    m = re.search(r"_(GHG|ORB|ICE|FWF)\b", ts_col, flags=re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return None

def main():
    df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME)

    # 1) Find all TS series columns
    series_cols = [c for c in df.columns if is_ts_series_col(str(c))]
    if not series_cols:
        raise ValueError(f"No TS_*_C_* series columns were found. Current columns: {list(df.columns)}")

    # 2) Find time columns and try merging time_kyr_*
    time_cols = list_time_cols(df)
    yearlike_col = pick_yearlike_col(df)

    use_merged_time = False
    merged_time_col = None

    if time_cols:
        # Prefer time_kyr_GHG as the reference
        base = "time_kyr_GHG" if "time_kyr_GHG" in time_cols else time_cols[0]
        base_arr = pd.to_numeric(df[base], errors="coerce").to_numpy(dtype=float)

        all_ok = True
        for c in time_cols:
            arr = pd.to_numeric(df[c], errors="coerce").to_numpy(dtype=float)
            if _max_abs_diff(base_arr, arr) > TIME_MERGE_TOL:
                all_ok = False
                break

        if all_ok:
            # Merge into one time_kyr column (use the reference column)
            df["time_kyr"] = base_arr
            use_merged_time = True
            merged_time_col = "time_kyr"

    # 3) Output suffix: use fixed _loess02 when span=0.20
    if abs(SPAN - 0.20) < 1e-12:
        out_suffix = "_loess02"
    else:
        out_suffix = f"_loess{str(SPAN).replace('.', '')}"

    # 4) Select the x column for each series (keep it as consistent as possible)
    # - If all time_kyr_* columns are identical => use time_kyr for all series
    # - Otherwise: GHG uses time_kyr_GHG (if available), ORB uses time_kyr_ORB, etc.; fall back to year_like, then to row index
    x_for_series: dict[str, str] = {}

    if use_merged_time:
        for c in series_cols:
            x_for_series[c] = merged_time_col
    else:
        for c in series_cols:
            suf = guess_forcing_suffix(str(c))
            candidate = None
            if suf is not None:
                tc = f"time_kyr_{suf}"
                if tc in df.columns:
                    candidate = tc
            if candidate is None and yearlike_col is not None:
                candidate = yearlike_col
            if candidate is None:
                if "row_index" not in df.columns:
                    df["row_index"] = np.arange(len(df), dtype=float)
                candidate = "row_index"
            x_for_series[c] = candidate

    # 5) Compute loess for all series
    loess_results = {}
    for c in series_cols:
        xcol = x_for_series[c]
        x = pd.to_numeric(df[xcol], errors="coerce").to_numpy(dtype=float)
        y = pd.to_numeric(df[c], errors="coerce").to_numpy(dtype=float)
        loess_results[c + out_suffix] = loess_skmisc(x, y, span=SPAN, degree=DEGREE)

    # 6) Build the compact output: merge age columns when possible + include all loess columns
    out_cols = []

    # Age columns: prefer merged time_kyr, otherwise keep all used time_kyr_* columns (deduplicated), then add year_like if present
    if use_merged_time:
        out_cols.append(merged_time_col)
    else:
        used_time = []
        for c in series_cols:
            xc = x_for_series[c]
            if xc.startswith(TIME_PREFIX) and xc not in used_time:
                used_time.append(xc)
        out_cols.extend(used_time)

    if yearlike_col is not None and yearlike_col not in out_cols:
        out_cols.append(yearlike_col)

    out_df = pd.DataFrame({col: df[col] for col in out_cols})
    # Order loess columns consistently: four forcings first, then mean3/mean4, then others
    def sort_key(name: str):
        u = name.upper()
        if "_C_GHG" in u: return (0, name)
        if "_C_ORB" in u: return (1, name)
        if "_C_ICE" in u: return (2, name)
        if "_C_FWF" in u: return (3, name)
        if "MEAN3" in u:  return (4, name)
        if "MEAN4" in u:  return (5, name)
        return (9, name)

    for k in sorted(loess_results.keys(), key=sort_key):
        out_df[k] = loess_results[k]

    # 7) Write output
    sheet = f"loess_{SPAN:g}_compact"
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
        out_df.to_excel(w, index=False, sheet_name=sheet)

    print(f"Input: {INPUT_XLSX}")
    print(f"Detected TS series: {len(series_cols)} columns")
    print(f"Time columns merged: {use_merged_time} (tol={TIME_MERGE_TOL} kyr)")
    print(f"Saved: {OUT_XLSX}")

if __name__ == "__main__":
    main()
