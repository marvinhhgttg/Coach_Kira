/**
 * Briefings erzeugen: Daten vom Apps Script (mode=briefingData), Prompt aus kiraPrompts.ts (gleich wie Mac),
 * gerechnet wird mit coach-kira auf dem Mac, gespeichert wird wieder im Apps Script (mode=saveBriefing).
 */
import { BASE_URL } from './api';
import { LOCAL_MODEL, OLLAMA_URL, pingLocal } from './chatContext';
import { getEngine } from './engine';
import { cloudChat } from './cloudLLM';
import { briefingMessages, type KiMsg } from './kiraPrompts';

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
    throw new Error('Apps Script antwortet nicht mit JSON (Briefing-Patch eingespielt?)');
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
    const b = briefingMessages(type, dj.daten);
    prompt = { ok: true, format: b.format, messages: b.messages };
  } else {
    const pr = await (await fetch(`${BASE_URL}?mode=briefingPrompt&type=${type}&cb=${Date.now()}`, { redirect: 'follow' })).text();
    try {
      prompt = JSON.parse(pr);
    } catch {
      throw new Error('Apps Script liefert weder briefingData noch briefingPrompt (Patch eingespielt und neue Version bereitgestellt?)');
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

  const res = await fetch(`${OLLAMA_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: LOCAL_MODEL,
      stream: true,
      think: false,
      keep_alive: '30m',
      ...(prompt.format === 'json' ? { format: 'json' } : {}),
      options: { temperature: 0.3, num_ctx: 32768 },
      messages: prompt.messages,
    }),
  });
  if (!res.ok || !res.body) throw new Error(`Ollama HTTP ${res.status}`);
  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let buf = '';
  let text = '';
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    let nl: number;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).trim();
      buf = buf.slice(nl + 1);
      if (!line) continue;
      const j = JSON.parse(line);
      if (j.error) throw new Error(j.error);
      const piece = j.message?.content || '';
      if (piece) {
        text += piece;
        onProgress?.(text);
      }
    }
  }
  if (!text.trim()) throw new Error('Das lokale Modell hat keinen Text geliefert.');
  const saved = await postScript('saveBriefing', { kind: 'saveBriefing', type, text, model: LOCAL_MODEL, engine: 'local' });
  if (!saved.ok) throw new Error(saved.error || 'Speichern im Apps Script fehlgeschlagen.');
  return saved as { ok: true; analysis?: string; briefing?: string; recommendations?: any[]; timestamp?: string };
}
