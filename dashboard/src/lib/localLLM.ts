/**
 * KI-Aufruf für alle Cockpit-Funktionen. Folgt dem globalen Schalter:
 *   lokal → coach-kira (Ollama auf dem Mac) direkt aus dem Browser, mit Streaming
 *   Cloud → Perplexity über den Cockpit-Server (/api/llm)
 * Kein automatischer Wechsel zwischen beiden.
 */
import { LOCAL_MODEL, ollamaAufruf, pingLocal } from './chatContext';
import { getEngine } from './engine';
import { cloudChat } from './cloudLLM';

export type LlmMsg = { role: 'system' | 'user' | 'assistant'; content: string };

export async function ensureLocal(): Promise<void> {
  const st = await pingLocal();
  if (!st.ok || !st.hasModel) {
    throw new Error(`Lokales Modell nicht erreichbar${st.error ? ' – ' + st.error : ''}. Ollama auf dem Mac starten oder Schalter prüfen.`);
  }
}

export type KiOpts = { format?: 'json' | Record<string, unknown>; numCtx?: number; temperature?: number; onToken?: (textSoFar: string) => void };

/** Engine-abhängiger Aufruf (Name aus historischen Gründen „localChat“). */
export async function localChat(messages: LlmMsg[], opts: KiOpts = {}): Promise<string> {
  if (getEngine() === 'pplx') {
    const r = await cloudChat(messages, { json: !!opts.format, temperature: opts.temperature, maxTokens: 4000 });
    opts.onToken?.(r.answer);
    return r.answer;
  }
  return ollamaChat(messages, opts);
}

/** Immer lokal (coach-kira). */
export async function ollamaChat(messages: LlmMsg[], opts: KiOpts = {}): Promise<string> {
  await ensureLocal();
  const { text } = await ollamaAufruf(
    {
      model: LOCAL_MODEL,
      think: false,
      keep_alive: '30m',
      ...(opts.format ? { format: opts.format } : {}),
      options: { temperature: opts.temperature ?? 0.3, num_ctx: opts.numCtx ?? 16384 },
      messages,
    },
    opts.onToken,
    'funktion',
  );
  if (!text.trim()) throw new Error('Das lokale Modell hat keinen Text geliefert.');
  return text.trim();
}

export function parseJsonLoose<T = any>(text: string): T {
  const t = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  try {
    return JSON.parse(t) as T;
  } catch {
    const a = t.indexOf('{');
    const b = t.lastIndexOf('}');
    if (a >= 0 && b > a) return JSON.parse(t.slice(a, b + 1)) as T;
    throw new Error('Antwort des Modells ist kein gültiges JSON.');
  }
}
