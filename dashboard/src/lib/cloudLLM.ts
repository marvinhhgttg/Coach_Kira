/**
 * Cloud-KI über den Cockpit-Server (POST /api/llm → Perplexity). Kein Apps Script.
 * Der Schlüssel liegt im Credential-Tresor des Servers, nie im Browser.
 */
import type { LlmMsg } from './localLLM';

const SENTINEL = '__PORT_5000__';
/** Basis-URL des Cockpit-Servers: veröffentlicht „/port/5000“, lokal gleiche Herkunft. */
export const SERVER_BASE = SENTINEL.startsWith('__PORT') ? '' : '/' + SENTINEL.replace(/^\/+/, '');

export type CloudOpts = { model?: 'sonar' | 'sonar-pro'; sources?: boolean; maxTokens?: number; temperature?: number; json?: boolean };
export type CloudAntwort = { answer: string; model: string; citations: string[]; used: number; limit: number };

export async function cloudChat(messages: LlmMsg[], opts: CloudOpts = {}): Promise<CloudAntwort> {
  const msgs = opts.json
    ? [{ role: 'system' as const, content: 'Antworte ausschließlich mit gültigem JSON – ohne Codeblock, ohne Text davor oder danach.' }, ...messages]
    : messages;
  const res = await fetch(`${SERVER_BASE}/api/llm`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ messages: msgs, model: opts.model || 'sonar-pro', sources: !!opts.sources, max_tokens: opts.maxTokens, temperature: opts.temperature }),
  });
  const t = await res.text();
  let j: any;
  try {
    j = JSON.parse(t);
  } catch {
    throw new Error(`Cockpit-Server antwortet nicht (HTTP ${res.status}).`);
  }
  if (!res.ok || !j.ok) throw new Error(j?.error || `Cloud-KI HTTP ${res.status}`);
  const answer = String(j.answer || '').trim();
  if (!answer) throw new Error('Die Cloud-KI hat keinen Text geliefert.');
  return { answer, model: j.model, citations: j.citations || [], used: j.used, limit: j.limit };
}

export async function cloudStatus(): Promise<{ ok: boolean; configured: boolean; used: number; limit: number; error?: string }> {
  try {
    const r = await fetch(`${SERVER_BASE}/api/llm/status`, { cache: 'no-store' });
    const j = await r.json();
    return { ok: !!j.ok, configured: !!j.configured, used: Number(j.used || 0), limit: Number(j.limit || 0) };
  } catch (e) {
    return { ok: false, configured: false, used: 0, limit: 0, error: e instanceof Error ? e.message : String(e) };
  }
}

/** JSON-Aufgaben (Wochenplan, Prüfschleife) in der Cloud. */
export async function cloudJson(msgs: LlmMsg[]): Promise<string> {
  return (await cloudChat(msgs, { json: true, maxTokens: 4000 })).answer;
}
