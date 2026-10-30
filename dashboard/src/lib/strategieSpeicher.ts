/**
 * Strategie-Ansicht dauerhaft speichern: Formspitzen-Events und Szenarien A/B/C
 * liegen im Sheet (AI_KIRA_EXTRAS, typ 'strategie', ein Eintrag) – nicht mehr nur im Browser.
 *
 * - Beim Start wird der Sheet-Stand geladen. Ist er neuer als der Browser-Stand, gewinnt das Sheet.
 * - Hat nur der Browser Daten (Altbestand), werden sie einmalig ins Sheet übernommen.
 * - Jede Änderung wird nach 1,5 s gespeichert; der Browser bleibt Zwischenspeicher (offline).
 * - Ohne Apps-Script-Patch („strategie“ in CKX_TYPES) bleibt alles wie bisher im Browser.
 */
import { useEffect, useState } from 'react';
import { fetchExtras, saveExtra } from './extras';

export const STRAT_DATUM = '2025-08-26'; // fester Schlüssel: genau ein Eintrag
const STAND_KEY = 'ck_strategie_stand_v1';
export const EV_KEY = 'ck_plan_events_v1';
export const SC_KEY = 'ck_plan_scenarios_v1';

export type StratStatus = 'laden' | 'gespeichert' | 'speichert' | 'nur_browser' | 'fehler';
export interface StratDaten { events: unknown[]; szenarien: unknown[]; stand: string }

let status: StratStatus = 'laden';
let statusText = '';
const statusSubs = new Set<() => void>();
const datenSubs = new Set<(d: StratDaten) => void>();
let timer: ReturnType<typeof setTimeout> | null = null;
let gestartet = false;

function setStatus(s: StratStatus, t = '') {
  status = s;
  statusText = t;
  statusSubs.forEach((f) => f());
}

function lies<T>(k: string, def: T): T {
  try {
    const v = localStorage.getItem(k);
    return v ? (JSON.parse(v) as T) : def;
  } catch {
    return def;
  }
}
function schreib(k: string, v: unknown) {
  try {
    localStorage.setItem(k, JSON.stringify(v));
  } catch {
    /* ignore */
  }
}

export function lokalerStand(): StratDaten {
  return { events: lies<unknown[]>(EV_KEY, []), szenarien: lies<unknown[]>(SC_KEY, []), stand: lies<string>(STAND_KEY, '') };
}

/** Wird von der Strategie-Ansicht bei jeder Änderung aufgerufen. */
export function strategieGeaendert() {
  schreib(STAND_KEY, new Date().toISOString());
  if (status === 'nur_browser') return;
  if (timer) clearTimeout(timer);
  setStatus('speichert');
  timer = setTimeout(() => void hochladen(), 1500);
}

async function hochladen() {
  const d = lokalerStand();
  const text = `${d.events.length} Events · ${d.szenarien.length} Szenarien`;
  try {
    await saveExtra('strategie', STRAT_DATUM, text, d, 'cockpit');
    setStatus('gespeichert', d.stand);
  } catch (e) {
    const msg = String((e as Error)?.message || e);
    if (/typ|type|unbekannt|unknown|erlaubt/i.test(msg)) setStatus('nur_browser', 'Apps-Script-Patch „Strategie speichern“ fehlt');
    else setStatus('fehler', msg);
  }
}

/** Einmal pro Seitenaufruf: Sheet-Stand holen und mit dem Browser abgleichen. */
export async function strategieStart() {
  if (gestartet) return;
  gestartet = true;
  const lokal = lokalerStand();
  try {
    const items = await fetchExtras<StratDaten>('strategie', 1);
    const srv = items.find((x) => x.datum === STRAT_DATUM)?.data;
    const sEv = Array.isArray(srv?.events) ? (srv!.events as any[]) : [];
    const sSc = Array.isArray(srv?.szenarien) ? (srv!.szenarien as any[]) : [];
    // Zusammenführen statt Überschreiben: Events je id vereinigen, Szenarien je Name – das jüngere gewinnt.
    const ev = new Map<string, any>();
    for (const e of [...sEv, ...(lokal.events as any[])]) if (e && e.id != null && !ev.has(String(e.id))) ev.set(String(e.id), e);
    const sc = new Map<string, any>();
    for (const x of [...sSc, ...(lokal.szenarien as any[])]) {
      if (!x || !x.name) continue;
      const alt = sc.get(x.name);
      if (!alt || String(x.savedAt || '') > String(alt.savedAt || '')) sc.set(x.name, x);
    }
    // Nur wenn der Browser bewusst neuer ist (Stand gesetzt und jünger), dürfen Löschungen aus dem Browser gelten.
    const browserNeuer = !!lokal.stand && (!srv?.stand || lokal.stand > srv.stand);
    const events = browserNeuer ? (lokal.events as any[]) : [...ev.values()].sort((a, b) => String(a.date).localeCompare(String(b.date)));
    const szenarien = browserNeuer ? (lokal.szenarien as any[]) : [...sc.values()].sort((a, b) => String(a.name).localeCompare(String(b.name)));
    schreib(EV_KEY, events);
    schreib(SC_KEY, szenarien);
    const gleich = srv && JSON.stringify(events) === JSON.stringify(sEv) && JSON.stringify(szenarien) === JSON.stringify(sSc);
    datenSubs.forEach((f) => f({ events, szenarien, stand: srv?.stand || lokal.stand }));
    if (gleich) {
      schreib(STAND_KEY, srv!.stand || '');
      setStatus('gespeichert', srv!.stand || '');
    } else {
      schreib(STAND_KEY, new Date().toISOString());
      await hochladen(); // prüft zugleich, ob der Patch aktiv ist
    }
  } catch {
    setStatus('fehler', 'Sheet nicht erreichbar – Änderungen bleiben vorerst im Browser');
  }
}

/** Neue Daten aus dem Sheet (z. B. vom anderen Gerät) an die Ansicht melden. */
export function onStrategieDaten(f: (d: StratDaten) => void) {
  datenSubs.add(f);
  return () => {
    datenSubs.delete(f);
  };
}

export function useStratStatus(): { status: StratStatus; text: string } {
  const [, setN] = useState(0);
  useEffect(() => {
    const f = () => setN((n) => n + 1);
    statusSubs.add(f);
    return () => {
      statusSubs.delete(f);
    };
  }, []);
  return { status, text: statusText };
}
