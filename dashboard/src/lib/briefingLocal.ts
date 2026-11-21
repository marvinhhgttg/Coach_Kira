/**
 * Briefings erzeugen: Daten vom Apps Script (mode=briefingData), Prompt aus kiraPrompts.ts (gleich wie Mac),
 * gerechnet wird mit coach-kira auf dem Mac, gespeichert wird wieder im Apps Script (mode=saveBriefing).
 */
import { ladeErfahrung } from './useErfahrung';
import { erfahrungText } from './votumErfahrung';
import { BASE_URL } from './api';
import { LOCAL_MODEL, ollamaAufruf, pingLocal } from './chatContext';
import { getEngine } from './engine';
import { cloudChat } from './cloudLLM';
import { briefingMessages, type KiMsg } from './kiraPrompts';
import { fetchChartData } from './api';
import { rangeFor } from './windows';
import { SERVER_BASE } from './cloudLLM';
import { RECOVERY_FETCH_METRICS } from '../screens/RecoveryPanel';
import { kernVotumAusRohdaten, heuteBerlin, type KernVotumKi } from './kernVotumKi';
import { ladeP3 } from './paket3';
import { fetchGarminTage } from './garmin';
import { langKontext, parseLangEinstellung, type LangKontext } from './esLang';

/** 21.11.2026: lange Einheiten fürs Plan-Briefing (Einstellung aus den Geräte-Einstellungen). */
export async function ladeLangKontext(): Promise<LangKontext | null> {
  try {
    const [t, g] = await Promise.all([ladeP3(), fetchGarminTage()]);
    if (!t.length || !g.size) return null;
    let tage: string | null = null, sehr: string | null = null;
    try { tage = localStorage.getItem('ck_lang_tage'); sehr = localStorage.getItem('ck_lang_sehr'); } catch { /* Standard */ }
    return langKontext(t, g, parseLangEinstellung(tage, sehr), heuteBerlin());
  } catch {
    return null;
  }
}

/** Umschaltung 20.11.2026: Tagesvotum für den Prompt (Schalter vom Cockpit-Server, Standard kern). */
async function ladeKernVotumKi(): Promise<KernVotumKi> {
  let q = 'kern';
  try {
    const j: any = await (await fetch(`${SERVER_BASE}/api/texte/status?cb=${Date.now()}`, { cache: 'no-store' })).json();
    if (j && j.einstellung && j.einstellung.votumQuelle === 'ampel') q = 'ampel';
  } catch {
    /* Standard kern */
  }
  try {
    const res: any = await fetchChartData(rangeFor(30, 28, 90), [...RECOVERY_FETCH_METRICS] as any);
    return kernVotumAusRohdaten(res?.data || [], q);
  } catch {
    return kernVotumAusRohdaten([], q);
  }
}

export type BriefingType = 'status' | 'consolidated' | 'plan';
export type BriefingSource = { engine: 'local' | 'cloud'; model: string | null; timestamp: string | null };

async function postScript(mode: string, body: unknown) {
  const res = await fetch(`${BASE_URL}?mode=${mode}&cb=${Date.now()}`, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
  });
  const t = await res.text();
  try {
    return JSON.parse(t);
  } catch {
    throw new Error('Die Datenquelle antwortet gerade nicht wie erwartet.');
  }
}

export async function fetchBriefingSources(): Promise<Record<BriefingType, BriefingSource | null>> {
  try {
    const res = await fetch(`${BASE_URL}?mode=briefingSources&cb=${Date.now()}`, { redirect: 'follow' });
    const j = await res.json();
    return j && j.ok ? j.sources : ({} as any);
  } catch {
    return {} as any;
  }
}

/**
 * Erzeugt ein Briefing lokal. onProgress bekommt den bisherigen Text (Streaming).
 * Rückgabe = Antwort von saveBriefing (analysis bzw. briefing + recommendations).
 */
export async function runLocalBriefing(type: BriefingType, onProgress?: (text: string) => void) {
  const cloud = getEngine() === 'pplx';
  if (!cloud) {
    const st = await pingLocal();
    if (!st.ok || !st.hasModel) {
      throw new Error('Lokales Modell nicht erreichbar. Starte Ollama auf dem Mac oder stelle den Schalter auf Cloud.');
    }
  }
  // Stufe 3: Daten vom Apps Script (mode=briefingData), Prompt aus kiraPrompts (gleiche Quelle wie der Mac).
  // Übergang: kennt das Apps Script briefingData noch nicht, alter Weg über mode=briefingPrompt.
  let prompt: { ok: boolean; format?: string; messages: KiMsg[]; error?: string };
  let dj: any = null;
  try {
    const roh = await (await fetch(`${BASE_URL}?mode=briefingData&type=${type}&cb=${Date.now()}`, { redirect: 'follow' })).text();
    dj = JSON.parse(roh);
  } catch {
    dj = null; // Apps Script noch ohne briefingData (HTML-Fehlerseite ohne CORS) → alter Weg
  }
  if (dj && dj.ok && dj.daten) {
    const erfText = erfahrungText(await ladeErfahrung());
    const kernVotum = await ladeKernVotumKi();
    const langeEinheiten = type === 'plan' ? await ladeLangKontext() : null;
    const b = briefingMessages(type, { ...dj.daten, ...(erfText ? { erfahrung: erfText } : {}), kernVotum, ...(langeEinheiten ? { langeEinheiten } : {}) });
    prompt = { ok: true, format: b.format, messages: b.messages };
  } else {
    const pr = await (await fetch(`${BASE_URL}?mode=briefingPrompt&type=${type}&cb=${Date.now()}`, { redirect: 'follow' })).text();
    try {
      prompt = JSON.parse(pr);
    } catch {
      throw new Error('Für das Briefing liegen gerade keine Daten vor.');
    }
  }
  if (!prompt.ok) throw new Error(prompt.error || (dj && dj.error) || 'Briefing-Daten konnten nicht geladen werden.');

  if (cloud) {
    const r = await cloudChat(prompt.messages, { json: prompt.format === 'json', model: 'sonar-pro', maxTokens: 4000 });
    onProgress?.(r.answer);
    const savedC = await postScript('saveBriefing', { kind: 'saveBriefing', type, text: r.answer, model: r.model, engine: 'cloud' });
    if (!savedC.ok) throw new Error(savedC.error || 'Speichern im Apps Script fehlgeschlagen.');
    return savedC as { ok: true; analysis?: string; briefing?: string; recommendations?: any[]; timestamp?: string };
  }

  const { text } = await ollamaAufruf(
    {
      model: LOCAL_MODEL,
      think: false,
      keep_alive: '30m',
      ...(prompt.format === 'json' ? { format: 'json' } : {}),
      options: { temperature: 0.3, num_ctx: 32768 },
      messages: prompt.messages,
    },
    onProgress,
    'briefing',
  );
  if (!text.trim()) throw new Error('Das lokale Modell hat keinen Text geliefert.');
  const saved = await postScript('saveBriefing', { kind: 'saveBriefing', type, text, model: LOCAL_MODEL, engine: 'local' });
  if (!saved.ok) throw new Error(saved.error || 'Speichern im Apps Script fehlgeschlagen.');
  return saved as { ok: true; analysis?: string; briefing?: string; recommendations?: any[]; timestamp?: string };
}
