/**
 * Etappe 3b – Anmeldung im Cockpit (20.11.2026).
 * Einmal je Gerät mit PIN anmelden; die Sitzung gilt 45 Tage ab letzter Nutzung.
 * Ersteinrichtung nur mit dem Einrichtungscode. In der Testinstanz v2 gibt es keine Anmeldung.
 */
import { useCallback, useEffect, useState, type ReactNode } from 'react';
import { SERVER_BASE } from '../lib/cloudLLM';
import { Logo } from './Logo';

type Status = { configured: boolean; authenticated: boolean } | null;

// Abgelaufene Sitzung irgendwo im Cockpit → Anmeldung wieder zeigen
if (typeof window !== 'undefined' && !(window as any).__ckAnmeldeWaechter) {
  (window as any).__ckAnmeldeWaechter = true;
  const orig = window.fetch.bind(window);
  window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
    const r = await orig(input, init);
    try {
      const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
      if (r.status === 401 && url.includes(`${SERVER_BASE}/api/`) && !/\/api\/(widget|auth)\b/.test(url)) {
        window.dispatchEvent(new Event('ck-abgemeldet'));
      }
    } catch { /* egal */ }
    return r;
  };
}

async function jsonAnfrage(pfad: string, body?: unknown) {
  const r = await fetch(`${SERVER_BASE}/api/${pfad}`, body === undefined
    ? { cache: 'no-store', credentials: 'same-origin' }
    : { method: 'POST', credentials: 'same-origin', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
  const t = await r.text();
  let j: any = null;
  try { j = JSON.parse(t); } catch { /* kein JSON */ }
  return { status: r.status, j };
}

const FEHLER: Record<string, string> = {
  invalid_pin: 'PIN falsch.',
  already_configured: 'Das Cockpit ist schon eingerichtet – bitte mit PIN anmelden.',
  setup_disabled: 'Die Einrichtung ist auf dem Server nicht freigeschaltet.',
};

export function AnmeldeTor({ children }: { children: ReactNode }) {
  const [frei, setFrei] = useState<boolean | null>(null);
  const [st, setSt] = useState<Status>(null);
  const [pin, setPin] = useState('');
  const [pin2, setPin2] = useState('');
  const [code, setCode] = useState('');
  const [fehler, setFehler] = useState('');
  const [laeuft, setLaeuft] = useState(false);

  const pruefen = useCallback(async () => {
    try {
      const inst = await jsonAnfrage('instanz');
      if (inst.j?.instanz && inst.j.instanz !== 'live') { setFrei(true); return; }
      const s = await jsonAnfrage('auth/status');
      if (!s.j || typeof s.j.authenticated !== 'boolean') { setFrei(true); return; } // ohne Server (z. B. Vorschau): nicht sperren
      setSt(s.j);
      setFrei(!!s.j.authenticated);
    } catch {
      setFrei(true);
    }
  }, []);

  useEffect(() => {
    pruefen();
    const ab = () => { setFrei(false); pruefen(); };
    window.addEventListener('ck-abgemeldet', ab);
    return () => window.removeEventListener('ck-abgemeldet', ab);
  }, [pruefen]);

  if (frei === null) return <div className="min-h-screen" />;
  if (frei) return <>{children}</>;

  const einrichten = st ? !st.configured : false;

  async function absenden(e: React.FormEvent) {
    e.preventDefault();
    setFehler('');
    if (einrichten) {
      if (!/^\d{6,12}$/.test(pin)) return setFehler('PIN: 6 bis 12 Ziffern.');
      if (pin !== pin2) return setFehler('Die beiden PINs stimmen nicht überein.');
      if (!code.trim()) return setFehler('Bitte den Einrichtungscode eingeben.');
    } else if (!pin) return setFehler('Bitte PIN eingeben.');
    setLaeuft(true);
    try {
      const r = einrichten ? await jsonAnfrage('auth/setup', { code: code.trim(), pin }) : await jsonAnfrage('auth/login', { pin });
      if (r.j?.authenticated) {
        setPin(''); setPin2(''); setCode('');
        window.dispatchEvent(new Event('ck-angemeldet'));
        setFrei(true);
      } else {
        setFehler(FEHLER[r.j?.error] || r.j?.error || `Anmeldung fehlgeschlagen (HTTP ${r.status}).`);
      }
    } catch (err) {
      setFehler(String(err));
    } finally {
      setLaeuft(false);
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center p-4">
      <form onSubmit={absenden} className="panel w-full max-w-sm p-6 space-y-4" aria-labelledby="anmelde-titel">
        <div className="flex items-center gap-3">
          <Logo />
          <div>
            <h1 id="anmelde-titel" className="text-lg font-semibold">Coach Kira</h1>
            <p className="text-sm opacity-70">{einrichten ? 'Cockpit einrichten' : 'Anmelden'}</p>
          </div>
        </div>
        {einrichten && (
          <label className="block">
            <span className="label">Einrichtungscode</span>
            <input className="input w-full mt-1" value={code} onChange={(e) => setCode(e.target.value)} autoComplete="one-time-code" autoCapitalize="none" spellCheck={false} />
          </label>
        )}
        <label className="block">
          <span className="label">{einrichten ? 'Neuer PIN (6 bis 12 Ziffern)' : 'PIN'}</span>
          <input className="input w-full mt-1 tnum" type="password" inputMode="numeric" pattern="[0-9]*" autoComplete={einrichten ? 'new-password' : 'current-password'} value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ''))} autoFocus />
        </label>
        {einrichten && (
          <label className="block">
            <span className="label">PIN wiederholen</span>
            <input className="input w-full mt-1 tnum" type="password" inputMode="numeric" pattern="[0-9]*" autoComplete="new-password" value={pin2} onChange={(e) => setPin2(e.target.value.replace(/\D/g, ''))} />
          </label>
        )}
        {fehler && <p className="text-sm" role="alert" style={{ color: 'var(--ck-red, #dc2626)' }}>{fehler}</p>}
        <button type="submit" className="btn btn-primary w-full" disabled={laeuft}>{laeuft ? 'Einen Moment …' : einrichten ? 'PIN festlegen und anmelden' : 'Anmelden'}</button>
        <p className="text-xs opacity-60">Einmal je Gerät. Die Anmeldung bleibt 45 Tage ab der letzten Nutzung bestehen. Widgets brauchen keine Anmeldung.</p>
      </form>
    </div>
  );
}
