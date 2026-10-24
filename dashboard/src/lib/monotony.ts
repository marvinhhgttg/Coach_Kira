/**
 * Monotonie-Regel (Coach Kira, Stand 22.10.2026)
 *
 * Grundsatz: Ein einzelner hoher Monotonie-Wert ist meist kein Problem. Relevant ist die
 * 7-Tage-Monotonie (Ø/σ der Tageslasten) im Verlauf – zusammen mit der Zahl fordernder Tage
 * ohne Entlastung, der Wochenlast (Strain = Wochenlast × Monotonie) und den Erholungsdaten.
 *
 * Zonen der 7-Tage-Monotonie:
 *   < 1,5 gute Variation · 1,5–2,0 normal · > 2,0 aufmerksam · > 2,5 ungünstig (wenn anhaltend) · > 3,0 Warnzone
 *
 * Steuerung (kein automatisches Abbruch-Signal, sondern Verlaufssignal):
 *   - hoher Einzelwert, wenige fordernde Tage, Erholung ok → keine Einschränkung
 *   - > 2,0 und schon 4–5 fordernde Tage gesammelt → nächster Tag sehr locker oder frei
 *   - bleibt > 2,5 (≥ 2 Tage) → innerhalb 24–48 h Ruhetag oder sehr kurzer Regenerationsreiz
 *   - > 3,0 → mindestens ein echter Entlastungstag bzw. deutliche Reduktion
 *   - hohe Monotonie + schlechte Erholungswerte → keine Qualitätseinheit „durchdrücken“
 *   - 5–7 Tage ohne deutliche Entlastung → Entlastung einplanen
 */

export type MonoStufe = 'gut' | 'normal' | 'aufmerksam' | 'unguenstig' | 'warnung';
export type MonoCap = 'TRAIN' | 'EASY' | null;

export const MONO = { GUT: 1.5, NORMAL: 2.0, UNGUENSTIG: 2.5, WARNUNG: 3.0 } as const;

export function monoStufe(v: number | null | undefined): MonoStufe | null {
  if (v == null || !Number.isFinite(v)) return null;
  if (v > MONO.WARNUNG) return 'warnung';
  if (v > MONO.UNGUENSTIG) return 'unguenstig';
  if (v > MONO.NORMAL) return 'aufmerksam';
  if (v >= MONO.GUT) return 'normal';
  return 'gut';
}

export const MONO_STUFE_LABEL: Record<MonoStufe, string> = {
  gut: 'gute Variation',
  normal: 'normal',
  aufmerksam: 'aufmerksam',
  unguenstig: 'ungünstig, wenn anhaltend',
  warnung: 'Warnzone',
};

/** Tailwind-Textfarbe je Zone (gut/normal ohne Farbe) */
export function monoTone(v: number | null | undefined): string {
  const s = monoStufe(v);
  return s === 'warnung' ? 'text-red-300' : s === 'unguenstig' ? 'text-orange-300' : s === 'aufmerksam' ? 'text-yellow-300' : '';
}

/** Referenz-Tageslast: Ø der letzten bis zu 28 Tage (inkl. Ruhetage) */
export function refLoad(loads: number[]): number {
  const xs = loads.slice(-28).filter((x) => Number.isFinite(x));
  return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;
}
/** Fordernder Tag: mindestens 60 ESS und mindestens 70 % der Referenz-Tageslast */
export function isDemanding(load: number, ref: number) {
  return load >= Math.max(60, 0.7 * ref);
}
/** Deutliche Entlastung: höchstens 40 ESS oder höchstens 50 % der Referenz-Tageslast */
export function isRelief(load: number, ref: number) {
  return load <= Math.max(40, 0.5 * ref);
}

/** 7-Tage-Monotonie wie im Sheet: Ø / σ (Population) der letzten 7 Tageslasten */
export function monotony7(loads: number[]): number | null {
  const w = loads.slice(-7);
  if (w.length < 7) return null;
  const m = w.reduce((a, b) => a + b, 0) / 7;
  const sd = Math.sqrt(w.reduce((a, b) => a + (b - m) ** 2, 0) / 7);
  if (sd === 0) return m > 0 ? 9.99 : 0;
  return m / sd;
}

export interface MonoInput {
  /** 7-Tage-Monotonie bis einschließlich des letzten Tages in loads */
  mono: number | null;
  /** Tageslasten chronologisch, letzter Eintrag = letzter abgeschlossener Tag */
  loads: number[];
  /** 7-Tage-Monotonie der Vortage, gleiche Reihenfolge (optional, für „anhaltend“) */
  monoHistory?: (number | null)[];
  /** Erholung schlecht (z. B. HRV unter Normalband, Ruhepuls erhöht, Schlaf schlecht, schwere Beine/Befinden ≤ 2) */
  recoveryBad?: boolean;
}

export interface MonoAssessment {
  stufe: MonoStufe | null;
  cap: MonoCap; // max. Votum für den nächsten Tag
  restWithin48h: boolean; // Ruhetag oder sehr kurzer Regenerationsreiz binnen 24–48 h
  streakNoRelief: number; // Tage in Folge ohne deutliche Entlastung
  demandingInStreak: number; // davon fordernde Tage
  daysAbove25: number; // Tage in Folge mit Monotonie > 2,5
  weekLoad: number;
  strain: number | null;
  weekLoadHigh: boolean; // Wochenlast ≥ 110 % des 4-Wochen-Schnitts
  text: string; // kurze Einordnung (Deutsch)
}

const f2 = (v: number) => v.toFixed(2).replace('.', ',');

export function assessMonotony(inp: MonoInput): MonoAssessment {
  const loads = inp.loads.filter((x) => Number.isFinite(x));
  const ref = refLoad(loads);
  let streak = 0;
  let demanding = 0;
  for (let i = loads.length - 1; i >= 0; i--) {
    if (isRelief(loads[i], ref)) break;
    streak++;
    if (isDemanding(loads[i], ref)) demanding++;
  }
  const hist = [...(inp.monoHistory || []), inp.mono];
  let above25 = 0;
  for (let i = hist.length - 1; i >= 0; i--) {
    const v = hist[i];
    if (v == null || v <= MONO.UNGUENSTIG) break;
    above25++;
  }
  const weekLoad = loads.slice(-7).reduce((a, b) => a + b, 0);
  const four = loads.slice(-28);
  const avgWeek = four.length >= 14 ? (four.reduce((a, b) => a + b, 0) / four.length) * 7 : weekLoad;
  const weekLoadHigh = avgWeek > 0 && weekLoad >= 1.1 * avgWeek;
  const m = inp.mono;
  const stufe = monoStufe(m);
  const strain = m != null ? Math.round(weekLoad * m) : null;
  const base = { stufe, streakNoRelief: streak, demandingInStreak: demanding, daysAbove25: above25, weekLoad, strain, weekLoadHigh };

  const res = (cap: MonoCap, rest: boolean, text: string): MonoAssessment => ({ ...base, cap, restWithin48h: rest, text });

  if (m == null) {
    if (streak >= 5 && demanding >= 3) return res('EASY', false, `${streak} Tage ohne deutliche Entlastung → nächster Tag sehr locker oder frei.`);
    return res(null, false, '');
  }
  // Entlastung in den letzten 48 h schon erfolgt? (streak = Tage seit der letzten deutlichen Entlastung)
  const reliefDone = streak < 2;
  if (m > MONO.WARNUNG && !reliefDone)
    return res('EASY', true, `Monotonie ${f2(m)} > 3,0 (Warnzone) → echten Entlastungstag oder deutliche Reduktion einplanen.`);
  if (m > MONO.UNGUENSTIG && (above25 >= 2 || demanding >= 4) && !reliefDone)
    return res('EASY', true, `Monotonie bleibt über 2,5 (${above25 >= 2 ? `${above25} Tage` : `${demanding} fordernde Tage`}) → innerhalb 24–48 h Ruhetag oder sehr kurzer Regenerationsreiz.`);
  if (m > MONO.UNGUENSTIG && reliefDone && (above25 >= 2 || m > MONO.WARNUNG))
    return res('TRAIN', false, `Monotonie ${f2(m)} noch hoch, Entlastung in den letzten 48 h ist erfolgt → keine Qualitätseinheit, Rhythmus hart/locker beibehalten.`);
  if (m > MONO.NORMAL && demanding >= 4)
    return res('EASY', false, `Monotonie ${f2(m)} > 2,0 nach ${demanding} fordernden Tagen → nächster Tag sehr locker oder frei.`);
  if (m > MONO.NORMAL && inp.recoveryBad)
    return res('TRAIN', false, `Monotonie ${f2(m)} > 2,0 und Erholungswerte schlechter → keine Qualitätseinheit durchdrücken.`);
  if (streak >= 5 && demanding >= 3)
    return res('EASY', false, `${streak} Tage ohne deutliche Entlastung → Entlastung einplanen (Monotonie ${f2(m)}).`);
  if (demanding >= 3 && inp.recoveryBad)
    return res('EASY', false, `${demanding} fordernde Tage in Folge und Erholungswerte schlechter → Belastung reduzieren.`);
  if (m > MONO.NORMAL && weekLoadHigh)
    return res('TRAIN', false, `Monotonie ${f2(m)} > 2,0 bei hoher Wochenlast (Strain ${strain}) → keine zusätzliche Qualität.`);
  if (m > MONO.NORMAL)
    return res(null, false, `Monotonie ${f2(m)} – ${demanding === 1 ? '1 fordernder Tag' : `${demanding} fordernde Tage`} in Folge: beobachten, keine Einschränkung.`);
  return res(null, false, stufe === 'normal' ? `Monotonie ${f2(m)} normal.` : '');
}

/**
 * Plan-Prüfung (14 Tage): Verstöße gegen die Monotonie-Regel.
 * days: Plan chronologisch inkl. Monotonie der Simulation; history: Ist-Lasten vor dem ersten Plantag.
 */
export function monotonyPlanViolations(
  days: { label: string; load: number; mono: number | null; locked?: boolean }[],
  history: number[] = [],
): string[] {
  const out: string[] = [];
  const all = [...history];
  const ref = refLoad([...history, ...days.map((d) => d.load)]);
  const relief = (i: number) => i < days.length && isRelief(days[i].load, ref);
  let streak = 0;
  for (let k = history.length - 1; k >= 0 && !isRelief(history[k], ref); k--) streak++;
  let run25 = 0;
  for (let i = 0; i < days.length; i++) {
    const d = days[i];
    all.push(d.load);
    streak = isRelief(d.load, ref) ? 0 : streak + 1;
    run25 = d.mono != null && d.mono > MONO.UNGUENSTIG ? run25 + 1 : 0;
    const nextRelief48 = relief(i + 1) || relief(i + 2);
    const lastDays = i + 2 >= days.length; // Planende: nicht bewertbar
    if (streak === 5) out.push(`${d.label}: 5 Tage in Folge ohne deutliche Entlastung`);
    if (!lastDays && d.mono != null && d.mono > MONO.WARNUNG && !nextRelief48)
      out.push(`${d.label}: Monotonie ${f2(d.mono)} > 3,0 ohne Entlastungstag in den nächsten 48 h`);
    else if (!lastDays && run25 === 2 && !nextRelief48)
      out.push(`${d.label}: Monotonie seit 2 Tagen > 2,5 ohne Ruhetag in den nächsten 48 h`);
    if (i + 1 < days.length && d.mono != null && d.mono > MONO.NORMAL) {
      // fordernde Tage bis hier zählen
      let dem = 0;
      for (let k = all.length - 1; k >= 0 && !isRelief(all[k], ref); k--) if (isDemanding(all[k], ref)) dem++;
      const nxt = days[i + 1];
      if (dem >= 4 && !nxt.locked && !isRelief(nxt.load, ref) && !(nxt.load <= 0.6 * ref))
        out.push(`${nxt.label}: nach Monotonie ${f2(d.mono)} und ${dem} fordernden Tagen sollte dieser Tag sehr locker oder frei sein`);
    }
  }
  return out;
}

/** Regeltext für KI-Prompts (Chat, Briefings, Plan-Schleife) */
export const MONOTONIE_REGEL_TEXT = [
  'Monotonie-Regel: Ein einzelner hoher Monotonie-Wert ist meist kein Problem. Maßgeblich ist die rollierende 7-Tage-Monotonie (Ø/σ der Tageslasten) im Verlauf.',
  'Zonen: < 1,5 gute Variation; 1,5–2,0 normal; > 2,0 aufmerksam (vor allem bei hoher Wochenlast); > 2,5 ungünstig, wenn es länger anhält; > 3,0 Warnzone.',
  'Muster: 1–2 ähnliche Tage normal; 3–4 ähnliche, fordernde Tage beobachten (bei schlechterem Schlaf, HRV, Ruhepuls, schweren Beinen oder weniger Motivation reduzieren); 5–7 Tage ohne deutliche Entlastung relevant; mehrere Wochen hohe Monotonie = Risiko für Ermüdung, Stagnation, Überlastung, Infekte.',
  'Steuerung: Über etwa 2,0 und schon 4–5 fordernde Tage → nächster Tag sehr locker oder frei. Bleibt der Wert über 2,5 → innerhalb 24–48 h echter Ruhetag oder sehr kurzer Regenerationsreiz. Über 3,0 → mindestens ein echter Entlastungstag oder deutliche Reduktion. Hohe Monotonie plus schlechte Erholungswerte → keine Qualitätseinheit durchdrücken.',
  'Entscheidend ist die Kombination mit der Gesamtlast (Strain = Wochenlast × Monotonie). Monotonie senkst du durch klaren Rhythmus: harte Reize klar hart, lockere Tage wirklich locker, mindestens ein echter Erholungstag bei hoher Gesamtlast – nicht durch weniger Gesamtlast.',
].join('\n');
