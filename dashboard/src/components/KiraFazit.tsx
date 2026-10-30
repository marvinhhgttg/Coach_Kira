import { useState } from 'react';
import { localChat } from '../lib/localLLM';
import { useEngine } from '../lib/engine';

/** KI-Fazit zu einem Bericht – lokal mit coach-kira, pro Gerät zwischengespeichert. */
const KEY = 'ck_kira_fazit_v1';
function readCache(): Record<string, { text: string; ts: string }> {
  try {
    return JSON.parse(localStorage.getItem(KEY) || '{}');
  } catch {
    return {};
  }
}
function writeCache(c: Record<string, { text: string; ts: string }>) {
  try {
    const keys = Object.keys(c);
    if (keys.length > 40) for (const k of keys.slice(0, keys.length - 40)) delete c[k];
    localStorage.setItem(KEY, JSON.stringify(c));
  } catch {
    /* ignore */
  }
}
function hash(s: string) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return String(h);
}

export function KiraFazit({ title, facts, data }: { title: string; facts: string[]; data?: unknown }) {
  const [engine] = useEngine();
  const payload = JSON.stringify({ title, facts, data: data ?? null });
  const id = title + ':' + hash(payload);
  const [text, setText] = useState<string | null>(() => readCache()[id]?.text ?? null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string | null>(null);

  async function run() {
    setBusy(true);
    setErr(null);
    setText('');
    try {
      const out = await localChat(
        [
          {
            role: 'system',
            content:
              'Du bist Coach Kira, Sportwissenschaftlerin. Du sprichst Marc direkt per Du an. Antworte auf Deutsch, ohne Überschrift, in 4 bis 6 Sätzen. Nutze nur die gegebenen Daten, erfinde keine Zahlen.',
          },
          {
            role: 'user',
            content: `Bericht: ${title}\n\nRegelbasierte Fakten:\n- ${facts.join('\n- ') || 'keine'}\n\nKennzahlen (JSON):\n${JSON.stringify(data ?? {}, null, 1)}\n\nAufgabe: Ziehe ein Fazit – was ist der Trend, was ist die wichtigste Lehre, und gib 1–2 konkrete Empfehlungen für die nächsten Tage.`,
          },
        ],
        { numCtx: 8192, onToken: setText },
      );
      setText(out);
      const c = readCache();
      c[id] = { text: out, ts: new Date().toISOString() };
      writeCache(c);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
      setText(readCache()[id]?.text ?? null);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="mt-2 rounded border border-border px-3 py-2">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="label">Kira-Fazit</div>
        <button className="btn btn-ghost text-xs px-2 py-1" onClick={run} disabled={busy}>
          {busy ? `Kira (${engine === 'local' ? 'lokal' : 'Cloud'}) schreibt…` : text ? 'Neu erstellen' : `Fazit erstellen (${engine === 'local' ? 'lokal' : 'Cloud'})`}
        </button>
      </div>
      {err && <p className="mt-1 text-xs text-red-300">{err}</p>}
      {text ? (
        <p className="mt-1 text-sm text-ink-muted leading-relaxed whitespace-pre-wrap">{text}</p>
      ) : (
        !busy && <p className="mt-1 text-2xs text-ink-dim">Kira deutet die Zahlen dieses Berichts – je nach Schalter lokal auf deinem Mac oder in der Cloud.</p>
      )}
    </div>
  );
}
