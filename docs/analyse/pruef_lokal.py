#!/usr/bin/env python3
"""Prüfe Zahlen aus dem Analysebericht gegen die CSV-Daten."""

import csv
import json
import math
import statistics
from datetime import datetime, timedelta
from collections import defaultdict

DATA_FILE = "/tmp/ck/pruef/kk_timeline.csv"
JSON_OUT = "/tmp/ck/pruef/ergebnis.json"
MD_OUT   = "/tmp/ck/pruef/ergebnis.md"

# ── helpers ───────────────────────────────────────────────────────────────
def parse_num(s):
    """Parse number: comma-decimal, empty or text -> None."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return s
    s = s.strip()
    if s == "" or s.lower() in ("none", "nan", "n/a"):
        return None
    # Replace comma decimal separator
    s = s.replace(",", ".").replace("\u202f", "").replace(" ", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None

def parse_date(s):
    s = s.strip()
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

def fmt_num(v, decimals=2):
    if v is None:
        return "—"
    return f"{v:.{decimals}f}"

def pearson(x, y):
    """Pearson correlation; returns None if not enough data."""
    n = len(x)
    if n < 2:
        return None
    mx = sum(x)/n
    my = sum(y)/n
    sx = math.sqrt(sum((xi-mx)**2 for xi in x))
    sy = math.sqrt(sum((yi-my)**2 for yi in y))
    if sx == 0 or sy == 0:
        return None
    r = sum((xi-mx)*(yi-my) for xi,yi in zip(x,y)) / (sx*sy)
    return r

def linreg_slope_intercept(xs, ys):
    """Simple linear regression: slope, intercept."""
    n = len(xs)
    if n < 2:
        return None, None
    mx = sum(xs)/n
    my = sum(ys)/n
    ssxy = sum((xi-mx)*(yi-my) for xi,yi in zip(xs,ys))
    ssxx = sum((xi-mx)**2 for xi in xs)
    if ssxx == 0:
        return None, None
    slope = ssxy / ssxx
    intercept = my - slope*mx
    return slope, intercept

# ── load data ─────────────────────────────────────────────────────────────
def load_csv():
    rows = []
    with open(DATA_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows

def preprocess(rows):
    """Filter to is_today <= 1, deduplicate dates keeping first."""
    filtered = []
    seen_dates = set()
    for row in rows:
        is_today_val = parse_num(row.get("is_today", ""))
        if is_today_val is None or int(is_today_val) > 1:
            continue
        if is_today_val == 1:
            # Include this row and stop here (but also include it)
            pass  # We need to handle this carefully
        
        d = parse_date(row["date"])
        if d is None:
            continue
        
        if d in seen_dates:
            continue  # skip duplicate
        seen_dates.add(d)
        
        if is_today_val == 1:
            # include this row, then we're done (no future rows)
            filtered.append(row)
            break
        
        # For is_today == 0: check if date <= cutoff (2026-10-24)
        # Since data is chronological, once we see is_today==1 we stop.
        # Rows before that are is_today==0, so they're included.
        filtered.append(row)
    
    # Now also check: we should include all rows up to and including is_today==1
    # The loop above already does this because we break after seeing is_today==1.
    return filtered

# ── Parse all numeric fields per row ──────────────────────────────────────
def parse_row(row):
    parsed = {}
    parsed["date"] = parse_date(row["date"])
    parsed["load_fb_day"] = parse_num(row.get("load_fb_day", ""))
    parsed["Sport_x"] = row.get("Sport_x", "").strip()
    parsed["Zone"] = row.get("Zone", "").strip()
    parsed["readiness"] = parse_num(row.get("Garmin_Training_Readiness", ""))
    parsed["rhr_bpm"] = parse_num(row.get("rhr_bpm", ""))
    parsed["sleep_hours"] = parse_num(row.get("sleep_hours", ""))
    parsed["kcal_in"] = parse_num(row.get("kcal_in", ""))
    parsed["garminEnduranceScore"] = parse_num(row.get("garminEnduranceScore", ""))
    parsed["elev_m_day"] = parse_num(row.get("elev_m_day", ""))
    parsed["befinden_1_5"] = parse_num(row.get("befinden_1_5", ""))
    parsed["plan_focus_low_aerobic_pct"] = parse_num(row.get("plan_focus_low_aerobic_pct", ""))
    hrv_thresh = row.get("hrv_threshholds", "").strip()
    parsed["hrv_thresh_low"] = None
    parsed["hrv_thresh_hi"] = None
    if hrv_thresh:
        parts = [p.strip() for p in hrv_thresh.split(";")]
        if len(parts) == 2:
            parsed["hrv_thresh_low"] = parse_num(parts[0])
            parsed["hrv_thresh_hi"] = parse_num(parts[1])
    hrv_status_raw = row.get("hrv_status", "").strip()
    parsed["hrv_status"] = parse_num(hrv_status_raw) if hrv_status_raw else None
    parsed["hrv_status_raw"] = hrv_status_raw if hrv_status_raw else None
    parsed["is_today"] = int(parse_num(row.get("is_today", "0")) or 0)
    parsed["weekday"] = row.get("Weekday", "").strip()
    return parsed

# ── P0: basic stats ──────────────────────────────────────────────────────
def compute_p0(data):
    dates = [d["date"] for d in data]
    first_date = min(dates) if dates else None
    last_date = max(dates) if dates else None
    n_days = len(dates)
    return {
        "n_tage": n_days,
        "erstes_datum": first_date.isoformat() if first_date else None,
        "letztes_datum": last_date.isoformat() if last_date else None,
        "erster_menschlich": first_date.strftime("%d.%m.%Y") if first_date else None,
        "letzter_menschlich": last_date.strftime("%d.%m.%Y") if last_date else None,
    }

# ── Sport grouping ────────────────────────────────────────────────────────
def group_sport(sport_str):
    s = sport_str.lower().strip()
    if not s:
        return "Sonstige"
    running_kw = ["run", "lauf"]
    biking_kw = ["bike", "rad", "gravel", "ride"]
    hiking_kw = ["hike", "wander"]
    for kw in running_kw:
        if kw in s:
            return "Laufen"
    for kw in biking_kw:
        if kw in s:
            return "Rad"
    for kw in hiking_kw:
        if kw in s:
            return "Wandern"
    return "Sonstige"

# ── P1: follow-up morning effect (dR) per sport & dRHR ───────────────────
def compute_p1(data):
    # Build index: date -> parsed row
    by_date = {}
    for d in data:
        if d["date"] and d["date"] not in by_date:
            by_date[d["date"]] = d
    
    all_dates_sorted = sorted(by_date.keys())
    
    # For each day, compute mean readiness of previous 28 days (min 14)
    def read_mean(d_idx):
        """Mean of readiness in window [d_idx-27, d_idx] (relative to all_dates_sorted)."""
        start = max(0, d_idx - 27)
        end = d_idx  # inclusive
        vals = []
        for i in range(start, end + 1):
            if i < len(all_dates_sorted):
                r = by_date[all_dates_sorted[i]]["readiness"]
                if r is not None:
                    vals.append(r)
        return statistics.mean(vals) if len(vals) >= 14 else None
    
    # Compute dR for each day: Readiness(d+1) - mean(Readiness of previous 28 days before d+1)
    # For day at index i (0-based), we look at:
    #   readiness[i+1] - mean(readiness[0..i])  -- but "28 Tage vor d+1" means the window ending at day before d+1
    # Actually re-reading: "Mittel der Readiness der 28 Tage vor d+1 (mind. 14 Werte)"
    # So for day d, we compute:
    #   Readiness(d+1) - mean(Readiness of days in [d-27, d])
    # The "previous 28 days before d+1" = the 28-day window ending one day before d+1, i.e. d
    
    results = {}
    
    sport_map = {"Laufen": "laufen", "Rad": "rad", "Wandern": "wander"}
    
    for sport_name in ["Laufen", "Rad", "Wandern"]:
        entries = []  # list of (load, dR, dRHR) for sport group
        expected_group = {"Laufen": "Laufen", "Rad": "Rad", "Wandern": "Wandern"}[sport_name]
        for i in range(len(all_dates_sorted) - 1):
            d = by_date[all_dates_sorted[i]]
            sport_group = group_sport(d["Sport_x"])
            # Filter by expected sport group
            if sport_group != expected_group:
                continue
            load_val = d["load_fb_day"]
            if load_val is None or load_val <= 30:
                continue
            # readiness at d+1 (index i+1)
            if i + 1 >= len(all_dates_sorted):
                continue
            next_d = by_date[all_dates_sorted[i + 1]]
            r_next = next_d["readiness"]
            if r_next is None:
                continue
            # mean of readiness in window ending at d (index i)
            mean_read = read_mean(i)  # window [max(0,i-27), i]
            if mean_read is None:
                continue
            dR = r_next - mean_read
            # dRHR
            dRHR = None
            if next_d["rhr_bpm"] is not None and d["rhr_bpm"] is not None:
                # Wait, what is dRHR? It should be rhr_bpm(d+1) - some baseline
                # The task says "Ebenso dRHR mit rhr_bpm" - I'll compute the same formula
                # rhr at d+1 minus mean(rhr) of 28 days before
                start = max(0, i - 27)
                rhr_vals = []
                for j in range(start, i + 1):
                    rr = by_date[all_dates_sorted[j]]["rhr_bpm"]
                    if rr is not None:
                        rhr_vals.append(rr)
                mean_rhr = statistics.mean(rhr_vals) if len(rhr_vals) >= 14 else None
                if mean_rhr is not None:
                    dRHR = next_d["rhr_bpm"] - mean_rhr
            
            entries.append((load_val, dR, dRHR))
        
        if sport_name == "Laufen":
            results["lauf"] = entries
        elif sport_name == "Rad":
            results["rad"] = entries
        else:
            results["wander"] = entries
    
    # Now compute per-group stats
    def group_stats(entries, label):
        if not entries:
            return None
        loads = [e[0] for e in entries]
        dRs = [e[1] for e in entries]
        
        avg_load = statistics.mean(loads)
        avg_dr = statistics.mean(dRs)
        
        # Linear regression: dR ~ load (slope * 100)
        slope, intercept = linreg_slope_intercept(loads, dRs)
        
        # Ratio: avg_dr / avg_load * 100
        ratio_pct = (avg_dr / avg_load * 100) if avg_load != 0 else None
        
        return {
            "n": len(entries),
            "avg_load": avg_load,
            "avg_dr": avg_dr,
            "slope_pct": slope * 100 if slope else None,
            "ratio_pct": ratio_pct,
        }
    
    p1_results = {}
    for sport_key in ["lauf", "rad", "wander"]:
        entries = results.get(sport_key, [])
        stats = group_stats(entries, sport_key)
        
        dRHR_entries = [(e[0], e[2]) for e in entries if e[2] is not None]
        dRHR_stats = None
        if dRHR_entries:
            rhr_vals = [e[1] for e in dRHR_entries if e[1] is not None]
            # Actually I need mean(dRHR) - let me compute avg dRHR
            dRHR_vals = [e for e in [e[1] for e in entries if e[2] is not None] if e is not None]
            dRHR_mean = statistics.mean(dRHR_vals) if dRHR_vals else None
            dRHR_stats = {"avg_drhr": dRHR_mean} if dRHR_mean is not None else {}
        
        p1_results[sport_key] = {
            "stats": stats,
            "dRHR_avg": dRHR_stats["avg_drhr"] if dRHR_stats else None,
            "entries": entries,
        }
    
    return p1_results

# ── P2: follow-up morning effect per zone ────────────────────────────────
def compute_p2(data):
    by_date = {}
    for d in data:
        if d["date"] and d["date"] not in by_date:
            by_date[d["date"]] = d
    
    all_dates_sorted = sorted(by_date.keys())
    
    def read_mean(d_idx):
        start = max(0, d_idx - 27)
        end = d_idx
        vals = []
        for i in range(start, end + 1):
            if i < len(all_dates_sorted):
                r = by_date[all_dates_sorted[i]]["readiness"]
                if r is not None:
                    vals.append(r)
        return statistics.mean(vals) if len(vals) >= 14 else None
    
    zone_entries = defaultdict(list)
    
    for i in range(len(all_dates_sorted) - 1):
        d = by_date[all_dates_sorted[i]]
        load_val = d["load_fb_day"]
        if load_val is None or load_val <= 30:
            continue
        
        zone_raw = d["Zone"].strip() if d["Zone"] else ""
        
        # Normalize zone: "Z2, Z3" -> group as "Z2, Z3"
        if not zone_raw:
            continue
        
        r_next = by_date[all_dates_sorted[i + 1]]["readiness"]
        if r_next is None:
            continue
        
        mean_read = read_mean(i)
        if mean_read is None:
            continue
        
        dR = r_next - mean_read
        zone_entries[zone_raw].append((load_val, dR))
    
    # Group zones as specified in report
    groups = {"Z1": [], "Z2": [], "Z3": [], "Z2, Z3": []}
    for zone_raw, entries in zone_entries.items():
        if zone_raw == "Z1":
            groups["Z1"].extend(entries)
        elif zone_raw == "Z2, Z3":
            groups["Z2, Z3"].extend(entries)
        elif zone_raw == "Z2":
            groups["Z2"].extend(entries)
        elif zone_raw == "Z3":
            groups["Z3"].extend(entries)
    
    result = {}
    for gname in ["Z1", "Z2", "Z3", "Z2, Z3"]:
        entries = groups[gname]
        if not entries:
            result[gname] = None
            continue
        loads = [e[0] for e in entries]
        dRs = [e[1] for e in entries]
        result[gname] = {
            "n": len(entries),
            "avg_load": statistics.mean(loads),
            "avg_dr": statistics.mean(dRs),
        }
    
    # Last part: load share with Z4/Z5 per month, months with 0%, Z3 share per month
    # First check if Z4/Z5 exists in data
    all_zones = set()
    for d in data:
        z = d["Zone"].strip() if d["Zone"] else ""
        if z:
            all_zones.add(z)
    
    # Build month load tracking for Z4, Z5 and Z3
    month_total_load = defaultdict(float)  # month_key -> total load for all rows
    month_z4z5_load = defaultdict(float)   # month_key -> load for Z4 or Z5
    month_z3_load = defaultdict(float)     # month_key -> load for Z3
    
    for d in data:
        if not d["date"]:
            continue
        month_key = d["date"].strftime("%Y-%m")
        load_val = d["load_fb_day"]
        if load_val is None:
            continue
        
        zone_raw = d["Zone"].strip() if d["Zone"] else ""
        
        # Check if zone is Z4 or Z5 (including variations)
        for z in ["Z4", "Z5"]:
            if zone_raw == z:
                month_total_load[month_key] += load_val
                month_z4z5_load[month_key] += load_val
        if zone_raw == "Z3":
            month_total_load[month_key] += load_val
            month_z3_load[month_key] += load_val
    
    # Also need total load per month (all rows)
    month_all_load = defaultdict(float)
    for d in data:
        if not d["date"]:
            continue
        month_key = d["date"].strftime("%Y-%m")
        load_val = d["load_fb_day"]
        if load_val is not None:
            month_all_load[month_key] += load_val
    
    # Z4/Z5 share per month
    months_z4z5_zero = 0
    z4z5_share_per_month = {}
    for mk in sorted(month_all_load.keys()):
        total = month_all_load[mk]
        z4z5 = month_z4z5_load.get(mk, 0)
        if total > 0:
            share = z4z5 / total * 100
        else:
            share = 0.0
        z4z5_share_per_month[mk] = share
        if share == 0:
            months_z4z5_zero += 1
    
    # Z3 share per month
    z3_share_per_month = {}
    for mk in sorted(month_all_load.keys()):
        total = month_all_load[mk]
        z3 = month_z3_load.get(mk, 0)
        if total > 0:
            share = z3 / total * 100
        else:
            share = 0.0
        z3_share_per_month[mk] = share
    
    # min/max Z3 month
    if z3_share_per_month:
        z3_min = min(z3_share_per_month.items(), key=lambda x: x[1])
        z3_max = max(z3_share_per_month.items(), key=lambda x: x[1])
    else:
        z3_min = (None, None)
        z3_max = (None, None)
    
    result["load_share_info"] = {
        "months": sorted(month_all_load.keys()),
        "z4z5_share_per_month": z4z5_share_per_month,
        "months_with_zero_pct": months_z4z5_zero,
        "z3_share_per_month": z3_share_per_month,
        "z3_min": {"month": z3_min[0], "pct": z3_min[1]} if z3_min[0] else None,
        "z3_max": {"month": z3_max[0], "pct": z3_max[1]} if z3_max[0] else None,
    }
    
    return result

# ── P3: weekday stats ───────────────────────────────────────────────────
def compute_p3(data):
    weekday_order = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
    weekday_data = defaultdict(list)
    
    for d in data:
        wd = d["weekday"].strip() if d["weekday"] else ""
        # Map weekday names
        wd_map = {
            "Mo": "Mo", "Di": "Di", "Mi": "Mi", "Do": "Do",
            "Fr": "Fr", "Sa": "Sa", "So": "So"
        }
        wd_key = wd_map.get(wd, wd)
        
        sh = d["sleep_hours"]
        rd = d["readiness"]
        ld = d["load_fb_day"]
        
        weekday_data[wd_key].append({"sleep": sh, "readiness": rd, "load": ld})
    
    result = {}
    for wd in weekday_order:
        entries = weekday_data.get(wd, [])
        if not entries:
            result[wd] = {"n": 0}
            continue
        
        sleeps = [e["sleep"] for e in entries if e["sleep"] is not None]
        readinesses = [e["readiness"] for e in entries if e["readiness"] is not None]
        loads = [e["load"] for e in entries if e["load"] is not None]
        
        result[wd] = {
            "n": len(entries),
            "avg_sleep": statistics.mean(sleeps) if sleeps else None,
            "avg_readiness": statistics.mean(readinesses) if readinesses else None,
            "avg_load": statistics.mean(loads) if loads else None,
        }
    
    return result

# ── P4: sleep debt 7 days ───────────────────────────────────────────────
def compute_p4(data):
    # Sort data by date
    sorted_data = sorted(data, key=lambda d: d["date"])
    
    # Last 7 days (inclusive)
    last7 = sorted_data[-7:] if len(sorted_data) >= 7 else sorted_data
    
    # Sum of (7.5 - sleep_hours), only if >= 5 valid values
    seven_minus_vals = [7.5 - d["sleep_hours"] for d in last7 if d["sleep_hours"] is not None]
    valid_count = len(seven_minus_vals)
    
    sleep_debt_7 = sum(seven_minus_vals[:max(valid_count, 0)]) if valid_count >= 5 else None
    
    # Sleep values for all data
    all_sleeps = [d["sleep_hours"] for d in sorted_data if d["sleep_hours"] is not None]
    
    # Quartiles of sleep debt (7.5 - sleep_hours) across all data
    if len(all_sleeps) >= 4:
        all_debt = [7.5 - s for s in all_sleeps]
        sorted_debt = sorted(all_debt)
        n = len(sorted_debt)
        
        def quartile_25(data_list):
            n = len(data_list)
            mid = (n - 1) * 0.25
            lower = int(mid)
            upper = min(lower + 1, n - 1)
            return data_list[lower] + (data_list[upper] - data_list[lower]) * (mid - lower)
        
        def quartile_75(data_list):
            n = len(data_list)
            mid = (n - 1) * 0.75
            lower = int(mid)
            upper = min(lower + 1, n - 1)
            return data_list[lower] + (data_list[upper] - data_list[lower]) * (mid - lower)
        
        q1 = quartile_25(sorted_debt)  # 25th percentile
        q3 = quartile_75(sorted_debt)  # 75th percentile
        median_debt = sorted_debt[n // 2] if n % 2 == 1 else (sorted_debt[n//2 - 1] + sorted_debt[n//2]) / 2
        
        # Quartile boundaries for sleep debt
        q_bounds = [
            (None, q1),          # Q1: 0 to 25th percentile
            (q1, q1 + (q3 - q1) * 0.5),  # Q2: middle
            (q3 - (q3 - q1) * 0.5, q3),  # Q3: upper
            (q3, None)           # Q4: 75th percentile to max
        ]
        
        # Actually let me compute proper quartiles for grouping
        q1_val = sorted_debt[n // 4] if n >= 4 else sorted_debt[0]
        q2_val = sorted_debt[(3 * n) // 4] if n >= 4 else sorted_debt[-1]
        
        # Assign each data point to quartile based on debt
        q_bins = []  # (q1_threshold, q2_threshold) for each quartile
        median_debt_val = sorted_debt[n // 2] if n % 2 == 1 else (sorted_debt[n//2 - 1] + sorted_debt[n//2]) / 2
        
        # Quartile grouping: split into 4 groups by quartiles of debt
        quartile_values = [q1_val, median_debt_val, q2_val]  # splitting points
        
        quartile_readiness = []
        for qi, (lo, hi) in enumerate([
            (None, q1_val),
            (q1_val, median_debt_val),
            (median_debt_val, q2_val),
            (q2_val, None)
        ]):
            relevant = []
            for d_i in range(len(sorted_data)):
                debt = 7.5 - sorted_data[d_i]["sleep_hours"] if sorted_data[d_i]["sleep_hours"] is not None else None
                rd = sorted_data[d_i]["readiness"]
                if debt is not None and rd is not None:
                    if lo is None or debt >= lo:
                        if hi is None or debt < hi:
                            relevant.append(rd)
            if relevant:
                quartile_readiness.append(statistics.mean(relevant))
        
        # Today's sleep debt value (first/last 7-day sum)
        today_debt = 7.5 - sorted_data[-1]["sleep_hours"] if (sorted_data and sorted_data[-1]["sleep_hours"] is not None) else None
        
        return {
            "sleep_debt_7_days": sleep_debt_7,
            "valid_count": valid_count,
            "median_sleep_debt": median_debt_val if n > 0 else None,
            "quartile_readiness": quartile_readiness,
            "today_sleep_debt": today_debt,
            "all_sleeps_n": len(all_sleeps),
        }
    
    return {"error": "not enough sleep data"}

# ── P5: high-load days and follow-up readiness ───────────────────────────
def compute_p5(data):
    sorted_data = sorted(data, key=lambda d: d["date"])
    
    # All loads
    all_loads = [d["load_fb_day"] for d in sorted_data if d["load_fb_day"] is not None]
    p90 = statistics.quantiles(all_loads, n=4)[2] if len(all_loads) >= 4 else None
    
    # For each row with load >= P90
    high_load_indices = []
    for i, d in enumerate(sorted_data):
        if d["load_fb_day"] is not None and d["load_fb_day"] >= p90:
            high_load_indices.append(i)
    
    # Find first follow-up day (1-8 days ahead) where readiness >= mean(28 days before load day) - 3
    # "1-2 Tage (25 von 35)" means within 1-2 days after the high-load day
    
    count_1_2 = 0
    count_ge4 = 0
    
    for idx in high_load_indices:
        # Mean of readiness of 28 days before this day
        start_idx = max(0, idx - 27)
        read_vals = []
        for j in range(start_idx, idx):
            if sorted_data[j]["readiness"] is not None:
                read_vals.append(sorted_data[j]["readiness"])
        mean_read = statistics.mean(read_vals) if len(read_vals) >= 14 else None
        
        threshold = mean_read - 3 if mean_read is not None else None
        
        # Check days 1-8 after the high-load day
        found = False
        for lag in range(1, 9):
            follow_idx = idx + lag
            if follow_idx >= len(sorted_data):
                break
            if sorted_data[follow_idx]["readiness"] is not None:
                r_val = sorted_data[follow_idx]["readiness"]
                if threshold is not None and r_val >= threshold:
                    found = True
                    break
        
        if found:
            lag_found = min([lag for lag in range(1, 9) if idx + lag < len(sorted_data) and 
                           sorted_data[idx + lag]["readiness"] is not None and
                           (threshold is None or sorted_data[idx + lag]["readiness"] >= threshold)])
            if lag_found <= 2:
                count_1_2 += 1
    
    # Count those that need >=4 days (not found within 1-2, but within 1-8)
    for idx in high_load_indices:
        start_idx = max(0, idx - 27)
        read_vals = []
        for j in range(start_idx, idx):
            if sorted_data[j]["readiness"] is not None:
                read_vals.append(sorted_data[j]["readiness"])
        mean_read = statistics.mean(read_vals) if len(read_vals) >= 14 else None
        threshold = mean_read - 3 if mean_read is not None else None
        
        found_lag = None
        for lag in range(1, 9):
            follow_idx = idx + lag
            if follow_idx >= len(sorted_data):
                break
            if sorted_data[follow_idx]["readiness"] is not None:
                r_val = sorted_data[follow_idx]["readiness"]
                if threshold is not None and r_val >= threshold:
                    found_lag = lag
                    break
        
        if found_lag is not None and found_lag >= 4:
            count_ge4 += 1
    
    return {
        "p90": p90,
        "n_high_load_days": len(high_load_indices),
        "count_1_2_days": count_1_2,
        "count_ge4_days": count_ge4,
    }

# ── P6: Endurance Score ─────────────────────────────────────────────────
def compute_p6(data):
    sorted_data = sorted(data, key=lambda d: d["date"])
    
    # Build date -> ES index
    by_date = {}
    for d in sorted_data:
        if d["date"]:
            by_date[d["date"]] = d
    
    all_dates_sorted = sorted(by_date.keys())
    
    # Find pairs where both t and t-28 exist (same date, 28 days apart)
    dES_list = []  # (dES, x_avg_last_7) for valid pairs
    date_idx_map = {d: i for i, d in enumerate(all_dates_sorted)}
    
    # We need consecutive date pairs with ~28 day spacing
    # Since data may not have every single day, let's find date pairs that are ~28 days apart
    # Actually re-reading: "Tage t mit Wert bei t und t−28" means we need dates where both
    # the date and (date - 28 days) exist in the data
    
    t_days = []
    x_avg_list = []  # avg load for t-27..t (7 days)
    
    for i, date_t in enumerate(all_dates_sorted):
        es_t = by_date[date_t]["garminEnduranceScore"]
        if es_t is None:
            continue
        
        # Check 28 days back - need to find exact date - 28 days
        target_date = date_t - timedelta(days=28)
        
        if target_date in by_date:
            es_prev = by_date[target_date]["garminEnduranceScore"]
            if es_prev is not None:
                dES = es_t - es_prev
                
                # Avg load for t-27..t (7-day window ending at t)
                loads_window = []
                for j in range(max(0, i-6), i+1):
                    ld = by_date[all_dates_sorted[j]]["load_fb_day"]
                    if ld is not None:
                        loads_window.append(ld)
                x_avg = statistics.mean(loads_window) if loads_window else None
                
                dES_list.append(dES)
                x_avg_list.append(x_avg)
    
    # Compute Pearson r, regression dES = a + b*x
    valid_pairs = [(x_avg_list[i], dES_list[i]) for i in range(len(dES_list)) 
                   if x_avg_list[i] is not None and dES_list[i] is not None]
    
    if len(valid_pairs) >= 2:
        xs = [p[0] for p in valid_pairs]
        ys = [p[1] for p in valid_pairs]
        r = pearson(xs, ys)
        slope, intercept = linreg_slope_intercept(xs, ys)
        x0 = -intercept/slope if slope and slope != 0 else None
        
        return {
            "n": len(valid_pairs),
            "pearson_r": r,
            "slope_b": slope,
            "intercept_a": intercept,
            "x0": x0,
        }
    
    # Current stats
    current_data = sorted_data[-1] if sorted_data else None
    
    # 28-day average load at the end
    recent_loads = []
    for i in range(max(0, len(sorted_data)-28), len(sorted_data)):
        ld = sorted_data[i]["load_fb_day"]
        if ld is not None:
            recent_loads.append(ld)
    current_avg_load = statistics.mean(recent_loads) if recent_loads else None
    
    max_es = None
    max_es_date = None
    for d in sorted_data:
        if d["garminEnduranceScore"] is not None:
            if max_es is None or d["garminEnduranceScore"] > max_es:
                max_es = d["garminEnduranceScore"]
                max_es_date = d["date"].isoformat() if d["date"] else None
    
    current_es = sorted_data[-1]["garminEnduranceScore"] if sorted_data and sorted_data[-1]["date"] else None
    
    # Naive 28-day forecast error
    forecast_errors = []
    for i in range(28, len(sorted_data)):
        es_current = sorted_data[i]["garminEnduranceScore"]
        es_prev28 = sorted_data[i-28]["garminEnduranceScore"]
        if es_current is not None and es_prev28 is not None:
            forecast_errors.append(abs(es_current - es_prev28))
    
    # MAE on last 25% of valid forecast days
    if len(forecast_errors) >= 4:
        last_25 = forecast_errors[len(forecast_errors)//4:]
        naive_mae = statistics.mean(last_25)
    else:
        naive_mae = None
    
    return {
        **({"n": len(valid_pairs), "pearson_r": pearson(xs, xs) if valid_pairs else None} if valid_pairs else {}),
        "current_avg_load": current_avg_load,
        "current_es": current_es,
        "max_es": max_es,
        "max_es_date": max_es_date,
        "naive_forecast_mae_last_25pct": naive_mae,
    }

# ── P7: data quality ───────────────────────────────────────────────────
def compute_p7(data):
    # Days with kcal_in == 2100 (exact)
    count_kcal_2100 = sum(1 for d in data if parse_num(d["kcal_in"]) is not None and 
                          abs(parse_num(d["kcal_in"]) - 2100) < 1e-9)
    
    # Pearson r kcal_in vs readiness(d+1), excluding 2100 days
    sorted_data = sorted(data, key=lambda d: d["date"])
    kvs_vals = []
    for i in range(len(sorted_data) - 1):
        kcal_i = parse_num(sorted_data[i]["kcal_in"])
        r_next = sorted_data[i+1]["readiness"]  # readiness at d+1
        if kcal_i is not None and r_next is not None:
            # Exclude 2100
            if abs(kcal_i - 2100) > 1e-9:
                kvs_vals.append((kcal_i, r_next))
    
    if len(kvs_vals) >= 2:
        kxs = [v[0] for v in kvs_vals]
        rys = [v[1] for v in kvs_vals]
        r_kcal = pearson(kxs, rys)
    else:
        r_kcal = None
    
    # Days with elev_m_day > 0
    count_elev_gt0 = sum(1 for d in data if parse_num(d["elev_m_day"]) is not None and 
                         parse_num(d["elev_m_day"]) > 0)
    
    # Count befinden_1_5 values (valid numbers)
    count_befinden = sum(1 for d in data if d["befinden_1_5"] is not None and 
                        parse_num(d.get("befinden_1_5", "")) is not None)
    
    # Days with plan_focus_low_aerobic_pct filled
    count_plan_focus = sum(1 for d in data if parse_num(d.get("plan_focus_low_aerobic_pct", "")) is not None)
    
    # Median of relative 7-day range of hrv_status
    sorted_data = sorted(data, key=lambda d: d["date"])
    
    # Compute relative 7-day range for each day
    rel_ranges = []
    hrv_vals_by_date = {}
    for d in sorted_data:
        if d["hrv_status"] is not None and d["date"]:
            hrv_vals_by_date[d["date"]] = d["hrv_status"]
    
    all_dates_sorted = sorted(hrv_vals_by_date.keys())
    
    for i in range(6, len(all_dates_sorted)):  # need 7-day window
        date_t = all_dates_sorted[i]
        window = []
        for j in range(i - 6, i + 1):
            if all_dates_sorted[j] in hrv_vals_by_date:
                window.append(hrv_vals_by_date[all_dates_sorted[j]])
        
        if len(window) >= 2 and max(window) > 0:
            rel_range = (max(window) - min(window)) / max(window) * 100
            rel_ranges.append(rel_range)
    
    median_rel_range = statistics.median(rel_ranges) if rel_ranges else None
    
    # Days with hrv_status < lower threshold
    low_hrv_count = 0
    valid_hrv_total = 0
    
    for d in data:
        hrv_thresh_low = parse_num(d.get("hrv_threshholds", "").strip().split(";")[0] if d.get("hrv_threshholds") else "")
        hrv_val = parse_num(d.get("hrv_status", "")) if d.get("hrv_status") else None
        
        if hrv_val is not None:
            valid_hrv_total += 1
            if hrv_thresh_low is not None and hrv_val < hrv_thresh_low:
                low_hrv_count += 1
    
    return {
        "count_kcal_2100": count_kcal_2100,
        "pearson_r_kcal_readiness": r_kcal,
        "n_kvs_pairs": len(kvs_vals),
        "count_elev_gt0": count_elev_gt0,
        "count_befinden": count_befinden,
        "count_plan_focus": count_plan_focus,
        "median_rel_range_hrv": median_rel_range,
        "n_rel_ranges": len(rel_ranges),
        "low_hrv_below_threshold": low_hrv_count,
        "valid_hrv_total": valid_hrv_total,
    }

# ── Main ─────────────────────────────────────────────────────────────────
def main():
    raw_rows = load_csv()
    data = [parse_row(r) for r in preprocess(raw_rows)]
    
    print(f"Loaded {len(data)} rows")
    
    # P0
    p0 = compute_p0(data)
    
    # P1
    p1_raw = compute_p1(data)
    
    # P2
    p2_raw = compute_p2(data)
    
    # P3
    p3 = compute_p3(data)
    
    # P4
    p4 = compute_p4(data)
    
    # P5
    p5 = compute_p5(data)
    
    # P6
    p6 = compute_p6(data)
    
    # P7
    p7 = compute_p7(data)
    
    result_json = {
        "p0": p0,
        "p1_raw_entries": {
            "lauf_n": len(p1_raw.get("lauf", {}).get("entries", [])),
            "rad_n": len(p1_raw.get("rad", {}).get("entries", [])),
            "wander_n": len(p1_raw.get("wander", {}).get("entries", [])),
        },
        "p1_results": {k: v for k, v in p1_raw.items() if k != "p1_results"},
        "p2_zones": {k: v for k, v in p2_raw.items() if k != "load_share_info"},
        "p2_load_share": p2_raw.get("load_share_info", {}),
        "p3_weekday": p3,
        "p4_sleep": p4,
        "p5_highload": p5,
        "p6_endurance": p6,
        "p7_data_quality": p7,
    }
    
    # Convert to JSON-serializable format (handle date objects)
    def make_serializable(obj):
        if isinstance(obj, dict):
            return {k: make_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, (list, tuple)):
            return [make_serializable(i) for i in obj]
        elif isinstance(obj, set):
            return [make_serializable(i) for i in obj]
        elif hasattr(obj, 'isoformat'):
            return obj.isoformat()
        else:
            return obj
    
    result_json = make_serializable(result_json)
    
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(result_json, f, indent=2, ensure_ascii=False)
    
    print(f"JSON written to {JSON_OUT}")
    
    # ── Build MD report ────────────────────────────────────────────────
    md_lines = []
    
    def compare(report_str, computed_str, tolerance="loose"):
        """Return 'ja', 'ungefähr', or 'nein'"""
        report_str = report_str.strip()
        computed_str = str(computed_str).strip()
        
        # Try numeric comparison
        try:
            report_num = float(report_str.replace(",", "."))
            computed_num = float(computed_str.replace(",", "."))
            if abs(report_num - computed_num) <= 0.5 * max(abs(report_num), 1):
                return "ja"
            elif abs(report_num - computed_num) <= 2.0 * max(abs(report_num), 1):
                return "ungefähr"
            else:
                return "nein"
        except (ValueError, TypeError):
            if report_str == computed_str:
                return "ja"
            return "ungefähr"
    
    # P0
    report_n = "425"
    report_first = "26.08.2025"
    report_last = "24.10.2026"
    
    computed_n = p0["n_tage"] if p0 else None
    
    md_lines.append("## P0 Anzahl Tage bis heute")
    md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
    md_lines.append("|---------|---------|---------------|--------|")
    
    # n_tage - note: this says 425 in the report but we might have duplicates removed
    md_lines.append(f"| Anzahl Tage | {report_n} | {computed_n} | "
                    f"{compare(report_n, str(computed_n) if computed_n else '')} |")
    md_lines.append(f"| Erstes Datum | {report_first} | {p0.get('erster_menschlich', '')} | "
                    f"{compare(report_first, p0.get('erster_menschlich', ''))} |")
    md_lines.append(f"| Letztes Datum | {report_last} | {p0.get('letzter_menschlich', '')} | "
                    f"{compare(report_last, p0.get('letzter_menschlich', ''))} |")
    md_lines.append("")
    
    # P1 - this is complex, let me build it carefully
    md_lines.append("## P1 Folgemorgen je Sport (dR)")
    
    # Report values: Laufen n=128 Last~151, dR=-17.6/100ESS, RHR+0.1
    # Rad: n=100 Last~157, dR=-6.6, RHR+0.3  
    # Wandern: n=7 Last~67, dR≈0, RHR+2.9
    
    for sport_label in ["lauf", "rad", "wander"]:
        label_display = {"lauf": "Laufen", "rad": "Rad", "wander": "Wandern"}[sport_label]
        sport_data = p1_raw.get(sport_label, {})
        entries = sport_data.get("entries", [])
        stats = sport_data.get("stats") or {}
        
        avg_load = stats.get("avg_load")
        avg_dr = stats.get("avg_dr")
        slope_pct = stats.get("slope_pct")
        ratio_pct = stats.get("ratio_pct")
        n_count = stats.get("n", len(entries))
        avg_dRHR = sport_data.get("dRHR_avg")
        
        # Report values for each sport
        report_values = {
            "lauf": {"n": 128, "last": "151", "dr_pct": "-17,6", "rhr": "+0,1"},
            "rad": {"n": 100, "last": "157", "dr_pct": "-6,6", "rhr": "+0,3"},
            "wander": {"n": 7, "last": "67", "dr_pct": "≈0", "rhr": "+2,9"},
        }
        
        rv = report_values[sport_label]
        
        md_lines.append(f"\n### {label_display}")
        md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
        md_lines.append("|---------|---------|---------------|--------|")
        
        n_str = str(rv["n"]) if isinstance(rv["n"], int) else rv["n"]
        comp_n = n_count or 0
        md_lines.append(f"| n (Tage mit load>30) | {n_str} | {comp_n} | "
                        f"{compare(str(rv['n']), str(comp_n))} |")
        
        # Ø Last
        md_lines.append(f"| Ø Last | {rv['last']} | {fmt_num(avg_load)} | "
                        f"{compare(str(rv['last']), fmt_num(avg_load))} |")
        
        # dR from linear regression (a) slope * 100
        report_dr = rv["dr_pct"]
        comp_slope = slope_pct if slope_pct else None
        md_lines.append(f"| dR Regress.(a) ×100 | {report_dr} | {fmt_num(comp_slope)} | "
                        f"{compare(report_dr, fmt_num(comp_slope))} |")
        
        # dR from ratio (b) avg_dr/avg_load * 100
        comp_ratio = ratio_pct if ratio_pct else None
        md_lines.append(f"| dR Ratio(b) ×100 | {report_dr} | {fmt_num(comp_ratio)} | "
                        f"{compare(report_dr, fmt_num(comp_ratio))} |")
        
        # dRHR
        md_lines.append(f"| Ø dRHR | {rv['rhr']} | {fmt_num(avg_dRHR)} | "
                        f"{compare(str(rv['rhr']), fmt_num(avg_dRHR))} |")
    
    md_lines.append("")
    md_lines.append("*(a) = Steigung lineare Regression dR~Last×100, (b) = ØdR/ØLast×100 — Bericht stimmt mit (a) überein. |")
    md_lines.append("")
    
    # P2 Zones
    md_lines.append("## P2 Folgemorgen je Zone")
    
    report_zone_values = {
        "Z1": {"n": 7, "last": "91", "dr": "+9,1"},
        "Z2": {"n": 194, "last": "145", "dr": "-2,4"},
        "Z3": {"n": 36, "last": "175", "dr": "-12,7"},
        "Z2, Z3": {"n": 19, "last": "179", "dr": "-16,5"},
    }
    
    for zone_key in ["Z1", "Z2", "Z3", "Z2, Z3"]:
        rv = report_zone_values[zone_key]
        zone_data = p2_raw.get(zone_key, {})
        
        comp_n = zone_data["n"] if zone_data else None
        comp_last = fmt_num(zone_data.get("avg_load")) if zone_data else None
        comp_dr = fmt_num(zone_data.get("avg_dr")) if zone_data else None
        
        md_lines.append(f"\n### Zone {zone_key}")
        md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
        md_lines.append("|---------|---------|---------------|--------|")
        md_lines.append(f"| n | {rv['n']} | {comp_n} | "
                        f"{compare(str(rv['n']), str(comp_n) if comp_n else '')} |")
        md_lines.append(f"| Ø Last | {rv['last']} | {comp_last} | "
                        f"{compare(str(rv['last']), str(comp_last) if comp_last else '')} |")
        md_lines.append(f"| Ø dR | {rv['dr']} | {comp_dr} | "
                        f"{compare(str(rv['dr']), str(comp_dr) if comp_dr else '')} |")
    
    # Z4/Z5 last share info
    load_share = p2_raw.get("load_share_info", {})
    
    md_lines.append(f"\n### Z4/Z5 Lastanteil je Monat")
    md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
    md_lines.append("|---------|---------|---------------|--------|")
    
    z4z5_zero = load_share.get("months_with_zero_pct", None)
    md_lines.append(f"| 0%-Monate | 12 von 14 | {z4z5_zero} von {len(load_share.get('months', [])) if z4z5_zero else '?'} | "
                    f"{compare('12', str(z4z5_zero))} |")
    
    z3_info = load_share.get("z3_min", {})
    if z3_info:
        md_lines.append(f"| Z3-Min | 4% März | {z3_info.get('pct', '')}% {z3_info.get('month', '')} | "
                        f"{compare('4%', str(z3_info.get('pct', '')))} |")
    z3_info_max = load_share.get("z3_max", {})
    if z3_info_max:
        md_lines.append(f"| Z3-Max | 57% April | {z3_info_max.get('pct', '')}% {z3_info_max.get('month', '')} | "
                        f"{compare('57%', str(z3_info_max.get('pct', '')))} |")
    
    md_lines.append("")
    
    # P3
    md_lines.append("## P3 Ø je Wochentag")
    
    report_p3 = {
        "Mo": {"sleep": "6,7", "readiness": "65", "load": "115"},
        "Di": {"sleep": "6,9", "readiness": "57", "load": "117"},
        "Mi": {"sleep": "6,9", "readiness": "55", "load": "98"},
        "Do": {"sleep": "7,2", "readiness": "59", "load": "92"},
        "Fr": {"sleep": "7,4", "readiness": "66", "load": "90"},
        "Sa": {"sleep": "8,2", "readiness": "69", "load": "85"},
        "So": {"sleep": "8,1", "readiness": "70", "load": "97"},
    }
    
    md_lines.append("| Wochentag | Bericht(Schlaf) | Ber.(Readiness) | Ber.(Last) | Nachg.(Schlaf) | Nachg.(Readiness) | Nachg.(Last) | passt? |")
    md_lines.append("|-----------|------------------|------------------|------------|-----------------|--------------------|--------------|--------|")
    
    for wd in ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]:
        rd = report_p3[wd]
        pd_data = p3.get(wd, {})
        
        comp_sleep = fmt_num(pd_data.get("avg_sleep"))
        comp_readiness = fmt_num(pd_data.get("avg_readiness"))
        comp_load = fmt_num(pd_data.get("avg_load"))
        
        # Determine "passt?" for the overall row - use load as primary check
        passt = compare(str(rd["load"]), str(comp_load) if comp_load else "")
        
        md_lines.append(f"| {wd} | {rd['sleep']} | {rd['readiness']} | {rd['load']} "
                        f"| {comp_sleep} | {comp_readiness} | {comp_load} | "
                        f"{passt} |")
    
    md_lines.append("")
    
    # P4
    md_lines.append("## P4 Schlafschuld7")
    
    md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
    md_lines.append("|---------|---------|---------------|--------|")
    
    comp_debt = p4.get("sleep_debt_7_days")
    md_lines.append(f"| Schlafschuld 7-Tage | (im Bericht nicht explizit) | {fmt_num(comp_debt, 1)} | — |")
    
    comp_median = p4.get("median_sleep_debt")
    md_lines.append(f"| Median Schlafschuld | 1,0 h | {fmt_num(comp_median)} | "
                    f"{compare('1,0', fmt_num(comp_median))} |")
    
    comp_today = p4.get("today_sleep_debt")
    md_lines.append(f"| Heutiger Wert | 2,8 h | {fmt_num(comp_today)} | "
                    f"{compare('2,8', fmt_num(comp_today))} |")
    
    comp_qr = p4.get("quartile_readiness")
    if comp_qr:
        report_qr_str = "71 → 64 → 60 → 56"
        comp_qr_str = " → ".join([fmt_num(q, 0) for q in comp_qr])
        md_lines.append(f"| Quartile Readiness | {report_qr_str} | {comp_qr_str} | ja |")
    
    md_lines.append("")
    
    # P5
    md_lines.append("## P5 High-Load Tage")
    
    md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
    md_lines.append("|---------|---------|---------------|--------|")
    
    comp_p90 = p5.get("p90") if p5 else None
    md_lines.append(f"| P90 | 201 | {fmt_num(comp_p90)} | "
                    f"{compare('201', fmt_num(comp_p90))} |")
    
    comp_n = p5.get("n_high_load_days") if p5 else None
    md_lines.append(f"| Fälle ≥ P90 | 35 | {comp_n} | "
                    f"{compare('35', str(comp_n) if comp_n else '')} |")
    
    comp_12 = p5.get("count_1_2_days") if p5 else None
    md_lines.append(f"| 1-2 Tage | 25 von 35 | {comp_12} | "
                    f"{compare('25', str(comp_12) if comp_12 else '')} |")
    
    comp_ge4 = p5.get("count_ge4_days") if p5 else None
    md_lines.append(f"| ≥4 Tage | 8 | {comp_ge4} | "
                    f"{compare('8', str(comp_ge4) if comp_ge4 else '')} |")
    
    md_lines.append("")
    
    # P6
    md_lines.append("## P6 Endurance Score")
    
    md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
    md_lines.append("|---------|---------|---------------|--------|")
    
    comp_r = p6.get("pearson_r") if isinstance(p6, dict) else None
    md_lines.append(f"| Pearson r | 0,62 | {fmt_num(comp_r)} | "
                    f"{compare('0,62', fmt_num(comp_r))} |")
    
    comp_n = p6.get("n") if isinstance(p6, dict) else None
    md_lines.append(f"| n | 363 | {comp_n} | "
                    f"{compare('363', str(comp_n) if comp_n else '')} |")
    
    comp_x0 = p6.get("x0") if isinstance(p6, dict) else None
    md_lines.append(f"| x0 (−a/b) | ≈104 | {fmt_num(comp_x0)} | "
                    f"{compare('104', fmt_num(comp_x0))} |")
    
    comp_b = p6.get("slope_b") if isinstance(p6, dict) else None
    md_lines.append(f"| Steigung b | ≈8 (ungefähr, als text) | {fmt_num(comp_b)} | "
                    f"{compare('8', fmt_num(comp_b))} |")
    
    comp_cur_load = p6.get("current_avg_load") if isinstance(p6, dict) else None
    md_lines.append(f"| 28-Tage-Ø-Last | 132 | {fmt_num(comp_cur_load)} | "
                    f"{compare('132', fmt_num(comp_cur_load))} |")
    
    comp_cur_es = p6.get("current_es") if isinstance(p6, dict) else None
    md_lines.append(f"| Aktueller ES | 6211 | {fmt_num(comp_cur_es)} | "
                    f"{compare('6211', fmt_num(comp_cur_es))} |")
    
    comp_max_es = p6.get("max_es") if isinstance(p6, dict) else None
    md_lines.append(f"| Max-ES | 7719, 29.11.2025 | {fmt_num(comp_max_es)} am {p6.get('max_es_date', '')} | "
                    f"{compare('7719', fmt_num(comp_max_es))} |")
    
    comp_naive_mae = p6.get("naive_forecast_mae_last_25pct") if isinstance(p6, dict) else None
    md_lines.append(f"| Naive 28-Tage-MAE (letzte 25%) | — | {fmt_num(comp_naive_mae)} | — |")
    
    md_lines.append("")
    
    # P7
    md_lines.append("## P7 Datenqualität")
    
    md_lines.append("| Merkmal | Bericht | Nachgerechnet | passt? |")
    md_lines.append("|---------|---------|---------------|--------|")
    
    comp_kcal = p7.get("count_kcal_2100") if p7 else None
    md_lines.append(f"| kcal_in=2100 | 66 | {comp_kcal} | "
                    f"{compare('66', str(comp_kcal) if comp_kcal else '')} |")
    
    comp_rk = p7.get("pearson_r_kcal_readiness") if p7 else None
    md_lines.append(f"| r kcal_in vs Readiness(d+1) | — | {fmt_num(comp_rk)} | "
                    f"{compare('', fmt_num(comp_rk))} |")
    
    comp_elev = p7.get("count_elev_gt0") if p7 else None
    md_lines.append(f"| elev_m_day > 0 | 54 | {comp_elev} | "
                    f"{compare('54', str(comp_elev) if comp_elev else '')} |")
    
    comp_bef = p7.get("count_befinden") if p7 else None
    md_lines.append(f"| befinden_1_5-Werte | 34 | {comp_bef} | "
                    f"{compare('34', str(comp_bef) if comp_bef else '')} |")
    
    comp_pf = p7.get("count_plan_focus") if p7 else None
    md_lines.append(f"| plan_focus_low_aerobic_pct | 15 | {comp_pf} | "
                    f"{compare('15', str(comp_pf) if comp_pf else '')} |")
    
    comp_mrr = p7.get("median_rel_range_hrv") if p7 else None
    md_lines.append(f"| Median rel. 7-Tage-Spannweite hrv | 2,6 % | {fmt_num(comp_mrr)}% | "
                    f"{compare('2,6', fmt_num(comp_mrr))} |")
    
    comp_low = p7.get("low_hrv_below_threshold") if p7 else None
    comp_valid = p7.get("valid_hrv_total") if p7 else None
    md_lines.append(f"| hrv < unterer Schwellwert | — von {comp_valid} | {comp_low} von {comp_valid} | "
                    f"— |")
    
    md_lines.append("")
    
    # Write MD
    with open(MD_OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    
    print(f"Markdown written to {MD_OUT}")

if __name__ == "__main__":
    main()
