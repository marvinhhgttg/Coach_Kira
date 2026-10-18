import { useEffect, useRef, useState } from 'react';
import { Panel } from '../components/UI';
import { EngineSwitch } from '../components/EngineSwitch';
import { useEngine } from '../lib/engine';
import {
  buildChatContext,
  sendChat,
  pingLocal,
  sendChatLocal,
  logLocalChat,
  LOCAL_MODEL,
  type ChatContext,
  type ChatMsg,
} from '../lib/chatContext';

const STORAGE_KEY = 'ck_kira_chat_v1';
const MAX_MSGS = 40;
const CONTEXT_TTL_MS = 10 * 60 * 1000;

function readHistory(): ChatMsg[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function writeHistory(msgs: ChatMsg[]) {
  try {
    const trimmed = msgs.slice(-MAX_MSGS);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(trimmed));
  } catch {
    // localStorage blockiert – nichts tun.
  }
}

type ExtendedMsg = ChatMsg & { engine?: 'pplx' | 'local'; model?: string; tokPerS?: number | null; pending?: boolean };

function fmtTs(ts?: string): string {
  if (!ts) return '';
  try {
    return new Date(ts).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  } catch {
    return '';
  }
}

function renderInlineMarkdown(text: string): React.ReactNode[] {
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  return parts.map((part, i) => {
    if (part.startsWith('**') && part.endsWith('**')) {
      return <strong key={i} className="text-ink font-semibold">{part.slice(2, -2)}</strong>;
    }
    return <span key={i}>{part}</span>;
  });
}

function renderCitations(citations?: string[]): React.ReactNode[] {
  if (!citations || citations.length === 0) return [];
  return citations.map((u, idx) => {
    let host = '';
    try {
      host = new URL(u).hostname;
    } catch {
      host = u;
    }
    return (
      <a
        key={u}
        href={u}
        target="_blank"
        rel="noreferrer"
        className="text-2xs text-accent underline break-all"
      >
        [{idx + 1}] {host}
      </a>
    );
  });
}

function renderAssistantContent(content: string, citations?: string[]): React.ReactNode {
  const lines = content.split('\n');
  const children: React.ReactNode[] = [];
  let keyIdx = 0;
  for (const line of lines) {
    const trimmed = line.trimEnd();
    if (trimmed === '') {
      children.push(<div key={keyIdx++} className="h-2" />);
      continue;
    }
    if (trimmed.startsWith('- ')) {
      children.push(
        <div key={keyIdx++} className="flex gap-2 text-ink">
          <span className="mt-2 h-1.5 w-1.5 rounded-full bg-accent shrink-0" />
          <span className="max-w-none">{renderInlineMarkdown(trimmed.slice(2))}</span>
        </div>
      );
    } else {
      children.push(
        <div key={keyIdx++} className="max-w-none">{renderInlineMarkdown(trimmed)}</div>
      );
    }
  }
  const citeNodes = renderCitations(citations);
  if (citeNodes.length > 0) {
    children.push(
      <div key={keyIdx++} className="mt-2 space-y-0.5">
        {citeNodes}
      </div>
    );
  }
  return <>{children}</>;
}

export function KiraChat() {
  const [open, setOpen] = useState(false);
  const [msgs, setMsgs] = useState<ExtendedMsg[]>(readHistory as () => ExtendedMsg[]);
  const [input, setInput] = useState('');
  const [sending, setSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [sources, setSources] = useState(false);
  const [used, setUsed] = useState<number | null>(null);
  const [limit, setLimit] = useState<number | null>(null);
  const [ctx, setCtx] = useState<ChatContext | null>(null);
  const [ctxTs, setCtxTs] = useState<string | null>(null);
  const ctxRef = useRef<{ ctx: ChatContext | null; ts: number | null }>({ ctx: null, ts: null });

  const [engine] = useEngine();
  const [, setLocalStatus] = useState<{ ok: boolean; hasModel: boolean; error?: string } | null>(null);

  const listRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    writeHistory(msgs);
  }, [msgs]);

  useEffect(() => {
    if (!listRef.current) return;
    const el = listRef.current;
    requestAnimationFrame(() => {
      el.scrollTop = el.scrollHeight;
    });
  }, [msgs, sending, error]);

  async function ensureContext(): Promise<ChatContext> {
    const now = Date.now();
    const last = ctxRef.current;
    if (last.ctx && last.ts && (now - last.ts) < CONTEXT_TTL_MS) {
      return last.ctx;
    }
    const fresh = await buildChatContext();
    ctxRef.current = { ctx: fresh, ts: now };
    setCtx(fresh);
    setCtxTs(new Date(now).toISOString());
    return fresh;
  }

  const updatePending = (fn: (m: ExtendedMsg) => ExtendedMsg) =>
    setMsgs((prev) => {
      const i = prev.length - 1;
      if (i < 0 || !prev[i].pending) return prev;
      return [...prev.slice(0, i), fn(prev[i])];
    });

  async function send(textRaw: string) {
    const text = textRaw.trim();
    if (!text || sending) return;
    const userMsg: ExtendedMsg = { role: 'user', content: text, ts: new Date().toISOString() };
    // Nur echte Nachrichten (keine leeren/abgebrochenen Platzhalter) an das Modell schicken
    const next = [...msgs.filter((m) => !m.pending && m.content.trim() !== ''), userMsg];
    setMsgs((prev) => [...prev.filter((m) => !m.pending), userMsg]);
    setInput('');
    setSending(true);
    setError(null);
    try {
      const context = await ensureContext();
      const plain = next.map(({ role, content }) => ({ role, content }));
      if (engine === 'local') {
        const st = await pingLocal();
        setLocalStatus(st);
        if (!st.ok || !st.hasModel) {
          setError('Lokales Modell nicht erreichbar. Starte Ollama auf dem Mac oder wechsle zu Perplexity.');
          return;
        }
        setMsgs((prev) => [...prev, { role: 'assistant', content: '', engine: 'local', model: LOCAL_MODEL, ts: new Date().toISOString(), pending: true }]);
        const res = await sendChatLocal(plain, context, (t) => updatePending((m) => ({ ...m, content: t })));
        updatePending((m) => ({ ...m, content: res.answer, tokPerS: res.tokPerS, pending: false }));
        logLocalChat(text, res.answer, LOCAL_MODEL).catch(() => {});
      } else {
        setUsed(null);
        setLimit(null);
        const res = await sendChat(plain, context, sources);
        setMsgs((prev) => [...prev, { role: 'assistant', content: res.answer, citations: res.citations, ts: new Date().toISOString(), engine: 'pplx' }]);
        setUsed(res.used);
        setLimit(res.limit);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setMsgs((prev) => prev.filter((m) => !(m.pending && m.content === '')).map((m) => (m.pending ? { ...m, pending: false } : m)));
    } finally {
      setSending(false);
    }
  }

  function handleSend() {
    send(input);
  }

  function handleKeyDown(e: React.KeyboardEvent<HTMLTextAreaElement>) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  }

  function handleChip(text: string) {
    send(text);
  }

  function handleClear() {
    if (!window.confirm('Chat-Verlauf im Browser löschen? Der Eintrag in AI_CHAT_LOG bleibt.')) return;
    setMsgs([]);
    setUsed(null);
    setLimit(null);
    setError(null);
    ctxRef.current = { ctx: null, ts: null };
    setCtx(null);
    setCtxTs(null);
  }

  const chips = [
    'Kann ich morgen hart trainieren?',
    'Warum lautet der Tagesentscheid so?',
    'Wie sieht meine Belastung diese Woche aus?',
    'Was sollte ich heute für die Erholung tun?',
  ];

  const ctxLabel = ctxTs ? `Daten Stand ${fmtTs(ctxTs)}` : '';

  return (
    <Panel title="Frag Kira · Sportwissenschaftlerin" right={
      <button
        className="btn btn-ghost text-xs px-2 py-1"
        onClick={() => setOpen((o) => !o)}
      >
        {open ? 'Schließen' : 'Öffnen'}
      </button>
    }>
      {!open ? (
        <div className="text-xs text-ink-dim">
          Fragen zu Training, Erholung und Plan – Kira antwortet mit deinen Cockpit-Daten (Perplexity sonar-pro oder lokal auf deinem Mac).
        </div>
      ) : (
        <div className="space-y-3">
          {ctxLabel && <div className="text-2xs text-ink-dim">{ctxLabel}</div>}
          <EngineSwitch />
          {msgs.length === 0 && (
            <div className="flex flex-wrap gap-2">
              {chips.map((c) => (
                <button
                  key={c}
                  className="chip border border-border-strong bg-bg-subtle text-xs hover:text-ink"
                  onClick={() => handleChip(c)}
                  disabled={sending}
                >
                  {c}
                </button>
              ))}
            </div>
          )}
          <div ref={listRef} className="max-h-[28rem] overflow-y-auto space-y-3 pr-1">
            {msgs.filter((m) => !(m.pending && m.content === '')).map((m, i) => (
              <div key={i} className={m.role === 'user' ? 'ml-auto max-w-[85%]' : 'mr-auto max-w-[92%]'}>
                {m.role === 'assistant' && (
                  <div className="text-2xs text-ink-dim mb-1">
                    {m.engine === 'local'
                      ? `Kira · lokal (${m.model || LOCAL_MODEL})${m.tokPerS != null ? ` · ${m.tokPerS} Tok/s` : ''}`
                      : 'Kira · Perplexity'}
                  </div>
                )}
                <div
                  className={
                    m.role === 'user'
                      ? 'rounded-lg bg-accent/15 border border-accent/30 px-3 py-2 text-sm whitespace-pre-wrap'
                      : 'rounded-lg bg-bg-subtle border border-border px-3 py-2 text-sm whitespace-pre-wrap'
                  }
                >
                  {m.role === 'assistant'
                    ? renderAssistantContent(m.content, m.citations)
                    : <span className="whitespace-pre-wrap">{m.content}</span>}
                </div>
              </div>
            ))}
            {sending && !(msgs.length && msgs[msgs.length - 1].pending && msgs[msgs.length - 1].content !== '') && (
              <div className="mr-auto max-w-[92%]">
                <div className="rounded-lg bg-bg-subtle border border-border px-3 py-2 text-xs text-ink-dim animate-pulse">
                  {engine === 'local' ? 'Kira (lokal) denkt nach… beim ersten Mal lädt der Mac das Modell, das kann bis zu einer Minute dauern.' : 'Kira denkt nach…'}
                </div>
              </div>
            )}
            {error && (
              <div className="text-xs text-red-300 border border-red-400/40 rounded p-2">
                {error}
              </div>
            )}
            {used != null && limit != null && (
              <div className="text-2xs text-ink-dim">
                Heute {used}/{limit} Anfragen
              </div>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <textarea
              rows={2}
              className="w-full rounded border border-border bg-bg px-3 py-2 text-sm"
              placeholder="Frag Kira, z. B. „Kann ich morgen Intervalle machen?“"
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={handleKeyDown}
              disabled={sending}
            />
            <div className="flex flex-wrap items-center gap-2">
              <label className="text-2xs text-ink-dim hover:text-ink cursor-pointer select-none">
                <input
                  type="checkbox"
                  className="mr-1"
                  checked={sources}
                  onChange={(e) => setSources(e.target.checked)}
                  disabled={sending || engine === 'local'}
                  title={engine === 'local' ? 'nur mit Perplexity' : undefined}
                />
                Mit Fachquellen (wissenschaftliche Suche)
              </label>
              <button
                className="btn btn-primary text-xs px-3 py-1.5"
                onClick={handleSend}
                disabled={sending || !input.trim()}
              >
                Senden
              </button>
            </div>
          </div>
          <div className="flex justify-between items-center">
            <button
              className="text-2xs text-ink-dim hover:text-ink underline"
              onClick={handleClear}
            >
              Verlauf löschen
            </button>
          </div>
          <div className="text-2xs text-ink-dim">
            {engine === 'local'
              ? 'Kira (lokal) läuft auf deinem Mac – keine Cloud, kein Tageslimit. Verlauf: dieser Browser + AI_CHAT_LOG.'
              : 'Kira berät nur – Plan und Sheet ändert sie nicht. Verlauf: dieser Browser + AI_CHAT_LOG.'}
          </div>
        </div>
      )}
    </Panel>
  );
}
