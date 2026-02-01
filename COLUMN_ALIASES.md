# Coach_Kira — Column Aliases / Header Normalization Baseline

Coach_Kira currently uses **multiple header spellings/casing** for the same concept.
This is a common source of bugs.

This document defines:
- a **canonical name** for each concept
- the **known aliases** found in code
- migration notes

---

## 1) Canonical naming recommendation

Recommended convention:
- lowercase `snake_case`

Example:
- `sport` instead of `Sport_x`
- `zone` instead of `Zone` / `coach_zone`
- `coache_ess_day` instead of mixed `coachE_ESS_day`

However, the sheet currently mixes styles. Until a migration is done, treat aliases as valid.

---

## 2) Known aliases (from code)

### 2.1 Date
- Canonical: `date`
- Aliases observed:
  - `date`
  - `datum`
  - `day` (fallback detection exists)

### 2.2 Today marker
- Canonical: `is_today`
- Aliases observed:
  - `is_today`

### 2.3 Sport / activity type
- Canonical: `sport_x`
- Aliases observed:
  - `Sport_x`
  - `sport_x`

### 2.4 Zone
- Canonical: `zone`
- Aliases observed:
  - `Zone`
  - `zone`
  - `coach_zone`

### 2.5 Planned load / ESS
- Canonical: `coache_ess_day`
- Aliases observed:
  - `coachE_ESS_day`
  - `coache_ess_day`

### 2.6 TE targets
- Canonical:
  - `target_aerobic_te`
  - `target_anaerobic_te`
- Aliases observed:
  - `Target_Aerobic_TE`
  - `Target_Anaerobic_TE`
  - `target_aerobic_te`
  - `target_anaerobic_te`

### 2.7 Forecast metrics
- Canonical:
  - `coache_atl_forecast`
  - `coache_ctl_forecast`
  - `coache_acwr_forecast`
- Aliases observed:
  - `coachE_ATL_forecast`
  - `coachE_CTL_forecast`
  - `coachE_ACWR_forecast`
  - `coache_atl_forecast`
  - `coache_ctl_forecast`

### 2.8 Observed metrics
- Canonical:
  - `load_fb_day`
  - `fbatl_obs`
  - `fbctl_obs`
  - `fbacwr_obs`
- Aliases observed:
  - `load_fb_day`
  - `fbATL_obs` / `fbatl_obs`
  - `fbCTL_obs` / `fbctl_obs`
  - `fbACWR_obs` / `fbacwr_obs`

### 2.9 Week phase
- Canonical: `week_phase`
- Aliases observed:
  - `Week_Phase`
  - `week_phase`

---

## 3) Migration plan (safe)

1) Implement a helper:
   - `getCol_(headers, ...aliases)` returning the first existing index.
2) Change all logic to use alias lookup.
3) Add a validator warning when non-canonical headers are used.
4) Once stable, migrate sheet headers to canonical names and remove legacy aliases.

---

*Last updated: 2026-02-01*
