/**
 * Gemeinsamer KI-Schalter für Kira-Chat und alle Briefings.
 * 'local' = Ollama auf dem Mac (coach-kira), 'pplx' = Cloud (Perplexity über Apps Script).
 * Standard: lokal. Kein automatischer Wechsel in die Cloud.
 */
import { useEffect, useState } from 'react';
import { pingLocal } from './chatContext';

export type Engine = 'local' | 'pplx';
const KEY = 'ck_kira_engine_v2';

function read(): Engine {
  try {
    const v = localStorage.getItem(KEY);
    return v === 'pplx' ? 'pplx' : 'local';
  } catch {
    return 'local';
  }
}
let store: Engine = read();
const subs = new Set<(e: Engine) => void>();

export function setEngine(e: Engine) {
  store = e;
  try {
    localStorage.setItem(KEY, e);
  } catch {
    /* blockiert – nur im Speicher */
  }
  subs.forEach((f) => f(e));
}

export function useEngine(): [Engine, (e: Engine) => void] {
  const [e, setE] = useState<Engine>(store);
  useEffect(() => {
    subs.add(setE);
    return () => {
      subs.delete(setE);
    };
  }, []);
  return [e, setEngine];
}

export type LocalStatus = { ok: boolean; hasModel: boolean; error?: string } | null;
let statusStore: LocalStatus = null;
const statusSubs = new Set<(s: LocalStatus) => void>();
let pending: Promise<LocalStatus> | null = null;

/** Gemeinsame Verfügbarkeitsprüfung (wird zwischen allen Anzeigen geteilt). */
export async function checkLocal(): Promise<LocalStatus> {
  if (!pending) {
    pending = pingLocal()
      .then((s) => {
        statusStore = { ok: s.ok, hasModel: s.hasModel, error: s.error };
        statusSubs.forEach((f) => f(statusStore));
        return statusStore;
      })
      .finally(() => {
        pending = null;
      });
  }
  return pending;
}

export function useLocalStatus(active: boolean): [LocalStatus, () => void] {
  const [s, setS] = useState<LocalStatus>(statusStore);
  useEffect(() => {
    statusSubs.add(setS);
    return () => {
      statusSubs.delete(setS);
    };
  }, []);
  useEffect(() => {
    if (active) checkLocal();
  }, [active]);
  return [s, () => void checkLocal()];
}

export function localStatusText(s: LocalStatus): { text: string; tone: 'ok' | 'warn' | 'wait' } {
  if (!s) return { text: 'prüfe…', tone: 'wait' };
  if (s.ok && s.hasModel) return { text: 'Lokal bereit (coach-kira)', tone: 'ok' };
  if (s.ok) return { text: 'Ollama läuft, aber das Modell coach-kira fehlt', tone: 'warn' };
  return { text: `Lokales Modell nicht erreichbar${s.error ? ' – ' + s.error : ''}`, tone: 'warn' };
}
