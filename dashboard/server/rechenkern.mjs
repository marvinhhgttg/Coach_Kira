// server/rechenkern.ts
var leer = (v) => v === null || v === void 0 || v === "";
var zahl = (v) => {
  if (leer(v)) return NaN;
  if (typeof v === "number") return v;
  return Number(String(v).replace(",", "."));
};
var nullWenigstens = (v) => leer(v) ? 0 : zahl(v);
var KERN_SPALTEN = [
  "fKEI",
  "load_star",
  "ctl_star",
  "ctl_gap",
  "coachE_Smart_Gains",
  "SG_Flags",
  "Monotony7",
  "Strain7",
  "coachE_ATL_morning",
  "coachE_CTL_morning",
  "TE_Balance_Trend"
];
function monotonie(w) {
  if (w.length < 2) return null;
  const mu = w.reduce((s, x) => s + x, 0) / w.length;
  const sd = Math.sqrt(w.reduce((s, x) => s + (x - mu) ** 2, 0) / w.length);
  return sd === 0 ? null : mu / sd;
}
function rechneZeile(rows, i, k) {
  const r = rows[i];
  const vor = (n) => rows[i - n];
  const c = (name) => zahl(k[name]);
  const fKEI = Math.min(1.15, Math.max(0.85, 1 + 0.05 * nullWenigstens(r.coachE_Smart_Gains)));
  const load_star = nullWenigstens(r.coachE_ESS_day) * nullWenigstens(r.fKEI);
  const nach = rows[i + 1];
  const ctl_star = leer(vor(1)?.ctl_star ?? null) && i === 8 ? nullWenigstens(r.coachE_CTL_forecast) : (1 - c("alpha_CTL")) * nullWenigstens(vor(1)?.ctl_star ?? null) + c("alpha_CTL") * (c("S_CTL") * nullWenigstens(r.load_star) + c("B_CTL") + c("wH_CTL") * nullWenigstens(nach?.sleep_hours ?? null) + c("wS_CTL") * nullWenigstens(nach?.sleep_score_0_100 ?? null));
  const ctl_gap = nullWenigstens(rows[8]?.coachE_CTL_forecast ?? null) - nullWenigstens(r.ctl_star);
  let coachE_Smart_Gains = null;
  const ah7 = vor(7)?.coachE_CTL_forecast;
  if (!leer(r.coachE_CTL_forecast) && !leer(ah7 ?? null) && !leer(r.Monotony7)) {
    coachE_Smart_Gains = (zahl(r.coachE_CTL_forecast) - zahl(ah7 ?? null)) * 10 / (Math.max(nullWenigstens(r.coachE_ACWR_forecast), 0.8) * (1 + zahl(r.Monotony7))) * 0.1;
  }
  let SG_Flags = "";
  try {
    const u = String(r.hrv_threshholds ?? "");
    const p = u.indexOf(";");
    if (p < 0) throw new Error("kein ;");
    const unten = Number(u.slice(0, p).replace(",", "."));
    if (!Number.isFinite(unten)) throw new Error("keine Zahl");
    let f = nullWenigstens(r.hrv_status) < unten - 2 ? "H " : "";
    f += leer(r.Monotony7) || zahl(r.Monotony7) > 1.6 ? "M" : "";
    SG_Flags = f.trim();
  } catch {
    SG_Flags = "";
  }
  const fenster7 = [];
  for (let j = Math.max(0, i - 6); j <= i; j++) {
    const v = rows[j]?.coachE_ESS_day;
    if (!leer(v ?? null) && Number.isFinite(zahl(v ?? null))) fenster7.push(zahl(v ?? null));
  }
  const Monotony7 = monotonie(fenster7);
  const fensterGestern = [];
  for (let j = Math.max(0, i - 7); j <= i - 1; j++) {
    const v = rows[j]?.coachE_ESS_day;
    if (!leer(v ?? null) && Number.isFinite(zahl(v ?? null))) fensterGestern.push(zahl(v ?? null));
  }
  const mono2 = monotonie(fensterGestern);
  const Strain7 = mono2 === null ? null : mono2 * fensterGestern.reduce((s, x) => s + x, 0);
  const v2 = vor(2);
  const coachE_ATL_morning = !v2 || leer(v2.coachE_ESS_day) ? null : -3.442725 + 0.815463593 * nullWenigstens(v2.coachE_ATL_forecast) + 1.088519637 * zahl(v2.coachE_ESS_day);
  const coachE_CTL_morning = !v2 || leer(v2.coachE_CTL_forecast) ? null : -3.95042 + 0.97283824 * zahl(v2.coachE_CTL_forecast);
  let summe = 0;
  let gewichtet = 0;
  for (let j = Math.max(0, i - 27); j <= i; j++) {
    const z = rows[j];
    const af = nullWenigstens(z?.coachE_ESS_day ?? null);
    summe += af;
    const treffer = nullWenigstens(z?.Target_Anaerobic_TE ?? null) >= 1 || nullWenigstens(z?.Target_Aerobic_TE ?? null) >= 4;
    if (treffer) gewichtet += af;
  }
  const TE_Balance_Trend = summe === 0 ? 0 : gewichtet / summe;
  return { fKEI, load_star, ctl_star, ctl_gap, coachE_Smart_Gains, SG_Flags, Monotony7, Strain7, coachE_ATL_morning, coachE_CTL_morning, TE_Balance_Trend };
}

// server/prognose.ts
var leer2 = (v) => v === null || v === void 0 || v === "";
var n0 = (v) => leer2(v) ? 0 : Number(v);
function schnitt(rows, i, spalte) {
  const w = [];
  for (let k = Math.max(0, i - 27); k <= i; k++) {
    const v = rows[k]?.[spalte];
    if (!leer2(v ?? null) && Number.isFinite(Number(v))) w.push(Number(v));
  }
  return w.length ? w.reduce((s, x) => s + x, 0) / w.length : null;
}
var atl = (vorher, last) => -8.819849 + 0.83645661 * n0(vorher) + 1.24059231 * Number(last);
function ctl(rows, i, vorher, spalteLast) {
  const s = schnitt(rows, i, spalteLast);
  if (s === null) return null;
  return 0.7 * (-3.95042 + 0.97283824 * n0(vorher) + 0.2336734 * Number(rows[i][spalteLast])) + 0.3 * 7.1 * s;
}
var quote = (a, b) => n0(b) === 0 ? null : n0(a) / n0(b);
var PROGNOSE_SPALTEN = ["coachE_ATL_forecast", "coachE_CTL_forecast", "coachE_ACWR_forecast", "fbATL_obs", "fbCTL_obs", "fbACWR_obs"];
function rechnePrognose(rows, i) {
  const r = rows[i];
  const v = rows[i - 1];
  const af = r.coachE_ESS_day;
  const e = r.load_fb_day;
  return {
    coachE_ATL_forecast: leer2(af) ? null : atl(v?.coachE_ATL_forecast, af),
    coachE_CTL_forecast: leer2(af) ? null : ctl(rows, i, v?.coachE_CTL_forecast, "coachE_ESS_day"),
    coachE_ACWR_forecast: leer2(af) ? null : quote(r.coachE_ATL_forecast, r.coachE_CTL_forecast),
    fbATL_obs: leer2(e) ? null : atl(v?.fbATL_obs, e),
    fbCTL_obs: leer2(e) ? null : ctl(rows, i, v?.fbCTL_obs, "load_fb_day"),
    fbACWR_obs: leer2(e) ? null : quote(r.fbATL_obs, r.fbCTL_obs)
  };
}
function simuliere(rows, start, ess = {}) {
  const kopie = rows.map((r) => ({ ...r }));
  const aus = [];
  for (let i = start; i < kopie.length; i++) {
    const r = kopie[i];
    const d = String(r.date);
    if (d in ess) r.coachE_ESS_day = ess[d];
    if (leer2(r.coachE_ESS_day)) break;
    const p = rechnePrognose(kopie, i);
    r.coachE_ATL_forecast = p.coachE_ATL_forecast;
    r.coachE_CTL_forecast = p.coachE_CTL_forecast;
    r.coachE_ACWR_forecast = quote(r.coachE_ATL_forecast, r.coachE_CTL_forecast);
    aus.push({ date: r.date, ess: r.coachE_ESS_day, atl: r.coachE_ATL_forecast, ctl: r.coachE_CTL_forecast, acwr: r.coachE_ACWR_forecast });
  }
  return aus;
}

// server/rechenkern_cli.ts
var ALLE_SPALTEN = [...KERN_SPALTEN, ...PROGNOSE_SPALTEN];
var PROGNOSE = new Set(PROGNOSE_SPALTEN);
function normFormel(f, zeile) {
  return f.replace(/(?<![A-Z$_])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g, (m, d1, col, d2, n) => d2 === "$" ? m : `${d1}${col}[${Number(n) - zeile}]`).replace(/\s+/g, "");
}
var leer3 = (v) => v === null || v === void 0 || v === "";
var istFehler = (v) => typeof v === "string" && /^#(REF|VALUE|N\/A|DIV\/0|NUM|NAME|ERROR)/.test(v);
function abgleich(tl, kk, heute) {
  const zu = (r) => r.zeilen.map((z) => Object.fromEntries(r.spalten.map((s, j) => [s, z[j]])));
  const rows = zu(tl);
  const iHeute = rows.findIndex((r) => r.date === heute);
  const ergebnis = [];
  const cacheKern = /* @__PURE__ */ new Map();
  const rechne = (i, k) => {
    if (PROGNOSE.has(k)) return rechnePrognose(rows, i)[k];
    if (!cacheKern.has(i)) cacheKern.set(i, rechneZeile(rows, i, tl.konstanten));
    return cacheKern.get(i)[k];
  };
  const zeileHeute = iHeute >= 0 ? Number(rows[iHeute].row) : Infinity;
  for (const k of ALLE_SPALTEN) {
    const formeln = tl.formelnAlle?.[k] || {};
    const zeilenMitFormel = Object.keys(formeln).map(Number).sort((a, b) => a - b);
    const bisHeute = zeilenMitFormel.filter((z) => z <= zeileHeute);
    const vz = bisHeute.length ? bisHeute[bisHeute.length - 1] : zeilenMitFormel[zeilenMitFormel.length - 1];
    const vorlage = vz ? normFormel(formeln[String(vz)], vz) : null;
    let zukunft = 0;
    let geprueft = 0, treffer = 0, sheetFehler = 0, alteVersion = 0, festwert = 0, maxDiff = 0;
    const beispiele = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r.date) > heute && !PROGNOSE.has(k)) continue;
      const f = formeln[String(r.row)];
      if (!f) {
        festwert++;
        continue;
      }
      if (!vorlage || normFormel(f, Number(r.row)) !== vorlage) {
        alteVersion++;
        continue;
      }
      const soll = r[k];
      if (istFehler(soll)) {
        sheetFehler++;
        continue;
      }
      const ist = rechne(i, k);
      geprueft++;
      if (String(r.date) > heute) zukunft++;
      let gleich;
      if (leer3(soll) || leer3(ist)) gleich = leer3(soll) && leer3(ist);
      else if (typeof soll === "number" && typeof ist === "number") {
        const d = Math.abs(soll - ist);
        maxDiff = Math.max(maxDiff, d);
        gleich = d <= Math.max(0.05, Math.abs(soll) * 5e-3);
      } else gleich = String(soll).trim() === String(ist).trim();
      if (gleich) treffer++;
      else if (beispiele.length < 5) beispiele.push(`${r.date}: Sheet ${JSON.stringify(soll)} \xB7 Kern ${JSON.stringify(ist)}`);
    }
    ergebnis.push({ spalte: k, geprueft, treffer, sheetFehler, alteVersion, festwert, maxDiff, beispiele, zukunft, art: PROGNOSE.has(k) ? "prognose" : "kern" });
  }
  let drift = [];
  if (kk?.ok) {
    const kkRows = new Map(zu(kk).map((r) => [String(r.row), r]));
    const spalten = ["coachE_ESS_day", "load_fb_day", ...PROGNOSE_SPALTEN, ...KERN_SPALTEN];
    drift = spalten.map((s) => {
      let n = 0, ab = 0;
      const bsp = [];
      for (const r of rows) {
        if (String(r.date) > heute) continue;
        const k2 = kkRows.get(String(r.row));
        if (!k2) continue;
        const a = r[s], b = k2[s];
        n++;
        const gleich = typeof a === "number" && typeof b === "number" ? Math.abs(a - b) <= Math.max(0.05, Math.abs(a) * 5e-3) : String(a ?? "").trim() === String(b ?? "").trim() || istFehler(a) && istFehler(b);
        if (!gleich) {
          ab++;
          if (bsp.length < 3) bsp.push(`${r.date}: timeline ${JSON.stringify(a)} \xB7 KK ${JSON.stringify(b)}`);
        }
      }
      return { spalte: s, verglichen: n, abweichend: ab, beispiele: bsp };
    });
  }
  const fa = tl.formelnAlle?.coachE_ATL_forecast || {};
  const start = rows.findIndex((r) => String(r.date) > heute && fa[String(r.row)]);
  let kette = null;
  if (start >= 0) {
    const sim = simuliere(rows, start);
    let exakt = 0, maxAcwr = 0;
    for (const x of sim) {
      const r = rows.find((z) => z.date === x.date);
      const d = Math.abs(Number(r.coachE_ACWR_forecast) - (x.acwr ?? NaN));
      const ok = Math.abs(Number(r.coachE_ATL_forecast) - (x.atl ?? NaN)) < 1e-6 && Math.abs(Number(r.coachE_CTL_forecast) - (x.ctl ?? NaN)) < 1e-6 && d < 1e-9;
      if (ok) exakt++;
      if (Number.isFinite(d)) maxAcwr = Math.max(maxAcwr, d);
    }
    kette = { start: String(rows[start].date), tage: sim.length, exakt, maxAcwr };
  }
  return { ok: true, heute, zeilen: rows.length, spalten: ergebnis, drift, kette };
}
export {
  abgleich,
  normFormel,
  simuliere
};
