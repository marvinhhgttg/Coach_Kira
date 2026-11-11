/**
 * Erfahrungsmodell „Lernen aus Nichtbefolgung“ – fest berechnet, ohne Sprachmodell.
 *
 * Frage: Wenn Marc härter trainiert hat als das Tagesvotum erlaubte – hatte das physiologische Folgen,
 * oder hat sein Körper genauso gut reagiert wie an Tagen, an denen dieselbe Belastung erlaubt war?
 *
 * Methode (je Tag D mit Regel-Votum V und tatsächlichem Anspruch A > V):
 *  - Reaktion an D+1, D+2, D+3 = gleiche Formel wie die Treffsicherheit
 *    (ΔReadiness/10 + ΔHRV/5 − ΔRuhepuls/3 + ΔSchlafscore/10 gegen das 7-Tage-Mittel bis D).
 *  - Vergleich gegen Referenztage mit GLEICHEM Anspruch A, die das Votum erlaubt hat und
 *    ähnlicher Ausgangsbereitschaft (±7 Punkte, bei zu wenig Tagen ±10/±14).
 *  - Folgekosten K = Mittel der Differenzen D+1…D+3; dazu Anteil „Tief“ (Readiness < 55 in D+1…D+3) gegen Referenz.
 *  - Gruppiert nach begrenzendem Grund (bind) und Sprung V→A; 90-%-Intervall per Bootstrap (fester Seed).
 * Einstufung (ab 8 Fällen): verträglich · kostet · unklar.
 * Persönliches Votum (nur heute): eine Stufe höher, wenn der passende Sprung „verträglich“ ist – nie bei
 * Schutzsignalen (HRV-Alarm, Ruhepuls+Readiness, Befinden, rotes Einzelsignal).
 */
import { buildDays, demandOf, VOTE_LABEL, VOTE_ORDER, type RecoveryDay, type Vote, type VoteGrund } from '../screens/RecoveryPanel';

export const GRUND_LABEL: Record<VoteGrund, string> = {
  readiness: 'niedriger Bereitschaft',
  rot: 'rotes Einzelsignal',
  rhr_schlaf: 'Ruhepuls hoch + wenig Schlaf',
  hrv: 'HRV-Alarm',
  warn2: 'zwei Warnsignale',
  rhr_ready: 'Ruhepuls hoch + niedrige Bereitschaft',
  acwr: 'ACWR',
  mono: 'Monotonie',
  befinden: 'Befinden',
};
/** Schutzsignale: werden nie durch Erfahrung gelockert */
export const SCHUTZ: VoteGrund[] = ['hrv', 'rhr_ready', 'befinden', 'rot'];
export const MIN_N = 8;
const MIN_N_ALLE = 12;

export type Klasse = 'vertraeglich' | 'kostet' | 'unklar';
export type Fall = {
  datum: string;
  votum: Vote;
  anspruch: Vote;
  bind: VoteGrund;
  readiness: number;
  d: [number | null, number | null, number | null];
  k: number;
  tief: boolean;
  refTief: number;
  nRef: number;
};
export type Gruppe = {
  key: string;
  bind: VoteGrund | '*';
  von: Vote;
  nach: Vote;
  n: number;
  d1: number | null;
  d2: number | null;
  d3: number | null;
  k: number;
  kLo: number;
  kHi: number;
  tief: number;
  refTief: number;
  klasse: Klasse;
};
export type Erfahrung = { stand: string; faelle: Fall[]; gruppen: Gruppe[]; nTage: number; nHaerter: number };

type Zeile = { i: number; datum: string; V: Vote; A: Vote; bind: VoteGrund; ready: number; s: (number | null)[]; tief: boolean };

const pos = (v: unknown): number | null => (typeof v === 'number' && Number.isFinite(v) && v > 0 ? v : null);
const mean = (xs: number[]) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0);
const r2 = (v: number) => Math.round(v * 100) / 100;

function zeilen(days: RecoveryDay[]): Zeile[] {
  const feld = (d: RecoveryDay | undefined, f: 'readiness' | 'hrv' | 'rhr' | 'sleepScore') => (d ? pos((d as any)[f]) : null);
  const score = (i: number, k: number): number | null => {
    const n = days[i + k];
    if (!n || feld(n, 'readiness') == null) return null;
    const win = days.slice(Math.max(0, i - 6), i + 1);
    const dl = (f: 'readiness' | 'hrv' | 'rhr' | 'sleepScore') => {
      const w = feld(n, f);
      const b = win.map((x) => feld(x, f)).filter((v): v is number => v != null);
      return w != null && b.length ? w - mean(b) : 0;
    };
    return dl('readiness') / 10 + dl('hrv') / 5 - dl('rhr') / 3 + dl('sleepScore') / 10;
  };
  const out: Zeile[] = [];
  days.forEach((d, i) => {
    if (!d.vote || d.readiness == null) return;
    // nur lückenlos aufeinanderfolgende Kalendertage werten
    const next = days[i + 1];
    if (!next || Date.parse(next.date) - Date.parse(d.date) !== 86400000) return;
    const last = d.actual.load ?? d.load;
    if (last == null) return;
    const s1 = score(i, 1);
    if (s1 == null) return;
    const A = demandOf(last, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
    const folge = days.slice(i + 1, i + 4);
    out.push({
      i,
      datum: d.date,
      V: d.vote,
      A,
      bind: (d.bind || 'readiness') as VoteGrund,
      ready: d.readiness,
      s: [s1, score(i, 2), score(i, 3)],
      tief: folge.some((x) => x.readiness != null && x.readiness < 55),
    });
  });
  return out;
}

function rng(seed: number) {
  let s = seed % 2147483647;
  return () => (s = (s * 16807) % 2147483647) / 2147483647;
}
function bootstrap(xs: number[], seed = 7, n = 1500): [number, number] {
  if (xs.length < 3) return [-Infinity, Infinity];
  const r = rng(seed);
  const m: number[] = [];
  for (let b = 0; b < n; b++) {
    let s = 0;
    for (let j = 0; j < xs.length; j++) s += xs[Math.floor(r() * xs.length)];
    m.push(s / xs.length);
  }
  m.sort((a, b) => a - b);
  return [m[Math.floor(0.05 * n)], m[Math.floor(0.95 * n)]];
}

export function klassifiziere(n: number, k: number, kLo: number, kHi: number, tiefDiff: number, minN = MIN_N): Klasse {
  if (n < minN) return 'unklar';
  if (k <= -0.5 || kHi < 0 || tiefDiff >= 0.25) return 'kostet';
  if (k >= 0.3 && kLo > -0.5 && tiefDiff <= 0.15) return 'vertraeglich';
  return 'unklar';
}

export function erfahrungAusTagen(days: RecoveryDay[]): Erfahrung {
  const z = zeilen(days);
  const ref = z.filter((r) => VOTE_ORDER.indexOf(r.A) <= VOTE_ORDER.indexOf(r.V));
  const haerter = z.filter((r) => VOTE_ORDER.indexOf(r.A) > VOTE_ORDER.indexOf(r.V));
  const faelle: Fall[] = [];
  for (const r of haerter) {
    let c: Zeile[] = [];
    for (const w of [7, 10, 14]) {
      c = ref.filter((x) => x.A === r.A && Math.abs(x.ready - r.ready) <= w && x.datum !== r.datum);
      if (c.length >= 4) break;
    }
    if (!c.length) continue;
    const d = [0, 1, 2].map((k) => {
      const cs = c.map((x) => x.s[k]).filter((v): v is number => v != null);
      return r.s[k] != null && cs.length ? r2((r.s[k] as number) - mean(cs)) : null;
    }) as [number | null, number | null, number | null];
    const ds = d.filter((v): v is number => v != null);
    faelle.push({
      datum: r.datum,
      votum: r.V,
      anspruch: r.A,
      bind: r.bind,
      readiness: r.ready,
      d,
      k: r2(mean(ds)),
      tief: r.tief,
      refTief: r2(mean(c.map((x) => (x.tief ? 1 : 0)))),
      nRef: c.length,
    });
  }
  const gruppe = (key: string, bind: VoteGrund | '*', von: Vote, nach: Vote, fs: Fall[], minN: number): Gruppe => {
    const ks = fs.map((f) => f.k);
    const [kLo, kHi] = bootstrap(ks, key.length * 31 + fs.length);
    const k = mean(ks);
    const tief = mean(fs.map((f) => (f.tief ? 1 : 0)));
    const refTief = mean(fs.map((f) => f.refTief));
    const dm = (j: 0 | 1 | 2) => {
      const v = fs.map((f) => f.d[j]).filter((x): x is number => x != null);
      return v.length ? r2(mean(v)) : null;
    };
    return { key, bind, von, nach, n: fs.length, d1: dm(0), d2: dm(1), d3: dm(2), k: r2(k), kLo: r2(kLo), kHi: r2(kHi), tief: r2(tief), refTief: r2(refTief), klasse: klassifiziere(fs.length, k, kLo, kHi, tief - refTief, minN) };
  };
  const gruppen: Gruppe[] = [];
  const keys = new Map<string, Fall[]>();
  const keysAlle = new Map<string, Fall[]>();
  for (const f of faelle) {
    const k = `${f.bind}|${f.votum}>${f.anspruch}`;
    keys.set(k, [...(keys.get(k) || []), f]);
    if (SCHUTZ.includes(f.bind)) continue; // Sammelgruppe ohne Schutzsignale
    const ka = `*|${f.votum}>${f.anspruch}`;
    keysAlle.set(ka, [...(keysAlle.get(ka) || []), f]);
  }
  for (const [k, fs] of keys) gruppen.push(gruppe(k, fs[0].bind, fs[0].votum, fs[0].anspruch, fs, MIN_N));
  for (const [k, fs] of keysAlle) gruppen.push(gruppe(k, '*', fs[0].votum, fs[0].anspruch, fs, MIN_N_ALLE));
  gruppen.sort((a, b) => b.n - a.n);
  return { stand: days.length ? days[days.length - 1].date : '', faelle, gruppen, nTage: z.length, nHaerter: faelle.length };
}

export function erfahrungAusRohdaten(raw: any[]): Erfahrung {
  return erfahrungAusTagen(buildDays(raw));
}

/** Passende Gruppe für einen Sprung: erst Grund+Sprung, sonst alle Gründe außer Schutzsignalen. */
export function gruppeFuer(e: Erfahrung, bind: VoteGrund, von: Vote, nach: Vote): Gruppe | null {
  const g = e.gruppen.find((x) => x.key === `${bind}|${von}>${nach}`);
  if (g && g.n >= MIN_N) return g;
  const a = e.gruppen.find((x) => x.key === `*|${von}>${nach}`);
  return a && a.n >= MIN_N_ALLE ? a : g || null;
}

const fz = (v: number | null, d = 1) => (v == null ? '–' : `${v > 0 ? '+' : v < 0 ? '−' : '±'}${Math.abs(v).toLocaleString('de-DE', { maximumFractionDigits: d, minimumFractionDigits: d })}`);
const pct = (v: number) => `${Math.round(v * 100)} %`;
const gName = (g: Gruppe) => `${VOTE_LABEL[g.von]}→${VOTE_LABEL[g.nach]}${g.bind === '*' ? '' : ` bei ${GRUND_LABEL[g.bind]}`}`;

export function gruppenText(g: Gruppe): string {
  const urteil = g.klasse === 'vertraeglich' ? 'verträglich' : g.klasse === 'kostet' ? 'hat Erholung gekostet' : 'kein klares Muster';
  return `${gName(g)} (${g.n}×): ${urteil} – Erholung Tag 1/2/3 ${fz(g.d1)}/${fz(g.d2)}/${fz(g.d3)} gegenüber vergleichbaren erlaubten Tagen, Tief unter 55 in ${pct(g.tief)} der Fälle (sonst ${pct(g.refTief)}).`;
}

/** Persönliches Votum für heute: Regel-Votum + Erfahrung. */
export function persoenlichesVotum(
  today: RecoveryDay,
  e: Erfahrung | null,
  anwenden = true,
): { vote: Vote | null; ruleVote: Vote | null; hinweise: string[] } {
  const rv = today.vote;
  const hinweise: string[] = [];
  if (!e || !rv || today.readiness == null) return { vote: rv, ruleVote: rv, hinweise };
  const bind = (today.bind || 'readiness') as VoteGrund;
  let vote = rv;
  const idx = VOTE_ORDER.indexOf(rv);
  const hoeher = idx < VOTE_ORDER.length - 1 ? VOTE_ORDER[idx + 1] : null;
  const schutz = SCHUTZ.includes(bind) || today.levels.hrv === 'rot' || (today.befinden != null && today.befinden <= 2);
  if (hoeher) {
    const g = gruppeFuer(e, bind, rv, hoeher);
    if (g && g.klasse === 'vertraeglich' && !schutz) {
      if (anwenden) vote = hoeher;
      hinweise.push(
        `Erfahrung: ${VOTE_LABEL[hoeher]} statt ${VOTE_LABEL[rv]} war bei dir bisher verträglich (${g.n}×, Erholung über 3 Tage ${fz(g.k)} gegenüber vergleichbaren erlaubten Tagen)` +
          (anwenden ? ` → persönliches Votum ${VOTE_LABEL[hoeher]}.` : ' – Anpassung ist ausgeschaltet.'),
      );
    } else if (g && g.klasse === 'vertraeglich' && schutz) {
      hinweise.push(`Erfahrung spräche für ${VOTE_LABEL[hoeher]}, aber heute greift ein Schutzsignal (${GRUND_LABEL[bind]}) – Regel-Votum bleibt.`);
    } else if (g && g.klasse === 'unklar') {
      hinweise.push(`Erfahrung: ${VOTE_LABEL[hoeher]} statt ${VOTE_LABEL[rv]}${g.bind === '*' ? '' : ` bei ${GRUND_LABEL[g.bind]}`} – bisher kein klares Muster (${g.n}×, Erholung über 3 Tage ${fz(g.k)}). Regel-Votum bleibt.`);
    }
  }
  // Warnung für die Stufe über dem (ggf. angepassten) Votum
  const vi = VOTE_ORDER.indexOf(vote);
  if (vi < VOTE_ORDER.length - 1) {
    const ueber = VOTE_ORDER[vi + 1];
    const g = gruppeFuer(e, bind, rv === vote ? rv : vote, ueber) || gruppeFuer(e, bind, rv, ueber);
    if (g && g.klasse === 'kostet')
      hinweise.push(
        `Erfahrung: ${VOTE_LABEL[ueber]} hat dich in dieser Lage bisher Erholung gekostet (${g.n}×: Tag 1/2/3 ${fz(g.d1)}/${fz(g.d2)}/${fz(g.d3)}, Tief unter 55 in ${pct(g.tief)} statt ${pct(g.refTief)}) – nicht drüber gehen.`,
      );
  }
  return { vote, ruleVote: rv, hinweise };
}

/** Kompakter Text für Kira (Report, Briefings, Chat, Telegram). Leer, wenn noch nichts gelernt ist. */
export function erfahrungText(e: Erfahrung | null, max = 8): string {
  if (!e) return '';
  const gs = e.gruppen.filter((g) => g.klasse !== 'unklar').sort((a, b) => Math.abs(b.k) * b.n - Math.abs(a.k) * a.n).slice(0, max);
  if (!gs.length) return '';
  return [
    `Stand ${e.stand}: ${e.nHaerter} Tage, an denen Marc härter trainiert hat als das Regel-Votum, verglichen mit Tagen gleicher Einheitsstufe und ähnlicher Ausgangsbereitschaft, an denen das Votum sie erlaubte.`,
    ...gs.map((g) => '- ' + gruppenText(g)),
    'Nutze das für Bewertung und Empfehlung: Wo Härter bei Marc verträglich war, darfst du Spielraum nennen; wo es gekostet hat, klar warnen und die Folgetage einplanen. Schutzsignale (HRV-Alarm, Ruhepuls hoch bei niedriger Bereitschaft, Befinden ≤ 2, rotes Einzelsignal) bleiben bindend.',
  ].join('\n');
}
