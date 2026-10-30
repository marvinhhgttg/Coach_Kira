/**
 * Workout ↔ ESS: hält Kiras strukturierte Workouts an Marcs persönliches Verhältnis von Last zu Dauer.
 * Grundlage ist der ESS-Rechner (gelernt aus den eigenen Garmin-Einheiten):
 *   ln ESS = a + b·ln(Minuten) + c·Ø HF
 * Das Workout wird als ganze Einheit bewertet – Gesamtdauer und zeitgewichtete Ø HF aus den Zonen.
 * Liegt Kiras Vorschlag daneben, werden die Hauptteil-Schritte deterministisch gestreckt bzw. gekürzt.
 */
import type { Modell, ProfilId } from './essRechner';

export type WZone = 'Z1' | 'Z2' | 'Z3' | 'Z4' | 'Z5' | 'offen';
export interface WStep { name: string; minuten: number; art: string; zone: WZone; hinweis?: string }
export interface WBlock { wiederholungen: number; schritte: WStep[] }
export interface WWorkout { name: string; beschreibung?: string; bloecke: WBlock[] }
export type HfZone = { zone: string; low: number; high: number };

/** Profil des ESS-Rechners passend zur Planzeile. */
export function profilFuer(sport: string, text: string): ProfilId | null {
  const t = text.toLowerCase();
  if (sport === 'cycling') {
    if (/indoor|rolle|zwift|trainer|ergo/.test(t)) return 'bike_indoor';
    if (/gravel|mtb|gelände/.test(t)) return 'bike_gravel';
    return 'bike_road';
  }
  if (sport === 'running') return 'run';
  if (sport === 'hiking' || sport === 'walking') return 'hike';
  if (/ruder|row/.test(t)) return 'row';
  return null;
}

/** Typische Ø HF eines Schritts in dieser Zone (Z1: oberer Bereich, weil Ein-/Auslaufen nicht bei 100 bpm liegt). */
export function zoneHf(zones: HfZone[], z: WZone, m?: Modell): number {
  const x = zones.find((q) => q.zone === z);
  if (!x) return m ? m.hf.p25 : 135;
  if (z === 'Z1') return Math.max(x.low, x.high - 10);
  return (x.low + x.high) / 2;
}

const essAus = (m: Modell, min: number, hf: number) => Math.exp(m.base[0] + m.base[1] * Math.log(min) + m.base[2] * hf);

export function workoutWerte(w: WWorkout, zones: HfZone[], m: Modell) {
  let min = 0, hfMin = 0;
  for (const b of w.bloecke) {
    const r = Math.max(1, b.wiederholungen || 1);
    for (const s of b.schritte) {
      const d = (Number(s.minuten) || 0) * r;
      min += d;
      hfMin += d * zoneHf(zones, s.zone, m);
    }
  }
  const hf = min > 0 ? hfMin / min : 0;
  return { min, hf, ess: min > 0 ? essAus(m, min, hf) : 0 };
}

/** Gesamtdauer, die bei dieser Ø HF die Ziel-Last ergibt. */
export function zielDauer(m: Modell, ess: number, hf: number): number {
  return Math.exp((Math.log(ess) - m.base[0] - m.base[2] * hf) / m.base[1]);
}

/** Dauer-Empfehlung für den Prompt: Ø HF ≈ Planzone, etwas darunter wegen Ein-/Auslaufen. */
export function dauerHinweis(m: Modell, ess: number, zones: HfZone[], planZone: string) {
  const z = (/Z[1-5]/.exec(String(planZone || '').toUpperCase())?.[0] || 'Z2') as WZone;
  const hfHaupt = zoneHf(zones, z, m);
  const hf = Math.round(hfHaupt - (z === 'Z1' ? 0 : 3));
  const min = zielDauer(m, ess, hf);
  return { min: Math.round(min), hf, zone: z };
}

const istRand = (s: WStep) => s.art === 'warmup' || s.art === 'cooldown';

function mitFaktor(w: WWorkout, f: number, rand: number): WWorkout {
  return {
    ...w,
    bloecke: w.bloecke.map((b) => ({
      ...b,
      schritte: b.schritte.map((s) => {
        const ff = istRand(s) ? rand : f;
        if (ff === 1) return s;
        // kurze Intervalle (< 3 min) behalten ihre Länge – nur die Dauerteile werden angepasst
        if (!istRand(s) && s.minuten < 3 && b.wiederholungen > 1) return s;
        const neu = Math.max(istRand(s) ? 5 : 1, Math.round(s.minuten * ff));
        return { ...s, minuten: neu };
      }),
    })),
  };
}

/**
 * Passt das Workout an die Ziel-ESS an (±8 %). Zuerst der Hauptteil, reicht das nicht (z. B. Ein- und
 * Auslaufen allein schon zu lang), werden auch Ein- und Auslaufen gekürzt (mindestens 5 min).
 */
export function anZielAnpassen(w: WWorkout, zones: HfZone[], m: Modell, ziel: number) {
  const vorher = workoutWerte(w, zones, m);
  if (!(ziel > 0) || !(vorher.min > 0)) return { workout: w, vorher, nachher: vorher, angepasst: false };
  if (Math.abs(vorher.ess - ziel) / ziel <= 0.08) return { workout: w, vorher, nachher: vorher, angepasst: false };
  const suche = (rand: number) => {
    let lo = 0.1, hi = 4;
    let best = w;
    for (let i = 0; i < 30; i++) {
      const mid = (lo + hi) / 2;
      const cand = mitFaktor(w, mid, rand);
      const e = workoutWerte(cand, zones, m).ess;
      best = cand;
      if (e > ziel) hi = mid; else lo = mid;
    }
    return best;
  };
  let neu = suche(1);
  let n = workoutWerte(neu, zones, m);
  if (Math.abs(n.ess - ziel) / ziel > 0.08 && n.ess > ziel) {
    // Hauptteil schon minimal – Ein- und Auslaufen mit kürzen
    for (const rand of [0.75, 0.5, 0.35]) {
      neu = suche(rand);
      n = workoutWerte(neu, zones, m);
      if (Math.abs(n.ess - ziel) / ziel <= 0.08) break;
    }
  }
  return { workout: neu, vorher, nachher: n, angepasst: true };
}
