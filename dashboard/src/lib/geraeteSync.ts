/**
 * Etappe 5 (20.11.2026): diese Einstellungen gelten auf allen Geräten gleich (Server › /api/einstellungen).
 * Ablauf: nach der Anmeldung einmal laden (Server gewinnt; was nur lokal existiert, wird hochgeladen),
 * danach jede Änderung über localStorage.setItem automatisch an den Server melden. Beim Zurückkehren
 * in die App (sichtbar/Fokus) wird erneut geladen. Nur Live-Cockpit; in v2 aus.
 */
import { SERVER_BASE } from './cloudLLM';
import { setEngine, type Engine } from './engine';

export const GERAETE_SCHLUESSEL = ['ck_votum_erfahrung_v2', 'ck_kira_engine_v2', 'ck_es_datum', 'ck_es_ziel', 'ck_aufstehen', 'ck_lang_tage', 'ck_lang_sehr'] as const;
const MENGE = new Set<string>(GERAETE_SCHLUESSEL);
const EVT_ERFAHRUNG = 'ck-erfahrung-schalter';

let aktiv = false;
let ausstehend: Record<string, string> = {};
let timer: ReturnType<typeof setTimeout> | null = null;
let origSet: ((k: string, v: string) => void) | null = null;

function senden() {
  timer = null;
  const werte = ausstehend;
  ausstehend = {};
  if (!Object.keys(werte).length) return;
  fetch(`${SERVER_BASE}/api/einstellungen`, {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ werte }),
  }).catch(() => {
    ausstehend = { ...werte, ...ausstehend }; // beim nächsten Mal erneut
  });
}

function vormerken(k: string, v: string) {
  ausstehend[k] = v;
  if (!timer) timer = setTimeout(senden, 800);
}

function lokal(k: string): string | null {
  try { return localStorage.getItem(k); } catch { return null; }
}

/** Server-Werte übernehmen, ohne sie zurückzumelden. */
function uebernehmen(werte: Record<string, string>) {
  for (const k of GERAETE_SCHLUESSEL) {
    const v = werte[k];
    if (v == null || lokal(k) === v) continue;
    try { (origSet || localStorage.setItem.bind(localStorage))(k, v); } catch { /* blockiert */ }
    if (k === 'ck_kira_engine_v2') setEngine(v as Engine);
    if (k === 'ck_votum_erfahrung_v2') window.dispatchEvent(new Event(EVT_ERFAHRUNG));
  }
}

// 21.11.2026: Kiras Wochenplan geräteübergreifend (Server › /api/wochenplan). Neuerer Plan gewinnt.
export const WOCHENPLAN_KEY = 'ck_kira_wochenplan_v1';
export const EVT_WOCHENPLAN = 'ck-wochenplan';
const erstelltVon = (roh: string | null) => { try { const j = JSON.parse(roh || 'null'); return j && j.erstellt ? Date.parse(j.erstellt) || 0 : 0; } catch { return 0; } };
export async function wochenplanAbgleich(): Promise<void> {
  const r = await fetch(`${SERVER_BASE}/api/wochenplan?cb=${Date.now()}`, { cache: 'no-store', credentials: 'same-origin' });
  if (!r.ok) return;
  const j = await r.json();
  const hier = lokal(WOCHENPLAN_KEY);
  const tServer = j && j.plan && j.plan.erstellt ? Date.parse(j.plan.erstellt) || 0 : 0;
  const tHier = erstelltVon(hier);
  if (tServer > tHier) {
    try { (origSet || localStorage.setItem.bind(localStorage))(WOCHENPLAN_KEY, JSON.stringify(j.plan)); } catch { /* blockiert */ }
    window.dispatchEvent(new Event(EVT_WOCHENPLAN));
  } else if (tHier > tServer && hier) {
    await wochenplanSenden(JSON.parse(hier));
  }
}
export async function wochenplanSenden(plan: unknown): Promise<void> {
  try {
    await fetch(`${SERVER_BASE}/api/wochenplan`, {
      method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ plan }),
    });
  } catch { /* beim nächsten Abgleich erneut */ }
}

async function laden(): Promise<void> {
  wochenplanAbgleich().catch(() => {});
  const r = await fetch(`${SERVER_BASE}/api/einstellungen?cb=${Date.now()}`, { cache: 'no-store', credentials: 'same-origin' });
  if (!r.ok) return;
  const j = await r.json();
  const werte: Record<string, string> = (j && j.werte) || {};
  uebernehmen(werte);
  // Erstes Gerät: was der Server noch nicht kennt, aber hier gesetzt ist, hochladen
  for (const k of GERAETE_SCHLUESSEL) if (werte[k] == null && lokal(k) != null) vormerken(k, lokal(k)!);
}

/** Einmal nach erfolgreicher Anmeldung aufrufen. Wartet höchstens 3 s, damit das Cockpit nie hängt. */
export async function geraeteEinstellungenStarten(): Promise<void> {
  if (!aktiv) {
    aktiv = true;
    try {
      const proto = Storage.prototype;
      const orig = proto.setItem;
      origSet = (k, v) => orig.call(localStorage, k, v);
      proto.setItem = function (this: Storage, k: string, v: string) {
        const vorher = this === localStorage && MENGE.has(k) ? this.getItem(k) : null;
        orig.call(this, k, v);
        if (this === localStorage && MENGE.has(k) && vorher !== String(v)) vormerken(k, String(v));
      };
    } catch { /* ohne Speicher-Hook: nur Laden */ }
    const neu = () => { if (document.visibilityState === 'visible') laden().catch(() => {}); };
    document.addEventListener('visibilitychange', neu);
    window.addEventListener('focus', neu);
  }
  await Promise.race([laden().catch(() => {}), new Promise((r) => setTimeout(r, 3000))]);
}
