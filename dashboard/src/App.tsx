import { useEffect, useState } from 'react';
import { CommandCenter } from './screens/CommandCenter';
import { PlanCockpit } from './screens/PlanCockpit';
import { ChartDeck } from './screens/ChartDeck';
import { TacticalLog } from './screens/TacticalLog';
import { Toast } from './components/UI';
import { Logo } from './components/Logo';
import { fmtTime } from './lib/format';
import { getProxyStatus, loginProxy, logoutProxy, setupProxy } from './lib/proxy';

type Tab = 'command' | 'plan' | 'charts' | 'logs';
type ThemeMode = 'dark' | 'garmin';

const TABS: { id: Tab; label: string; sub: string }[] = [
  { id: 'command', label: 'Command Center', sub: 'status · runStatus' },
  { id: 'plan', label: 'Plan Cockpit', sub: 'planJsonV2' },
  { id: 'charts', label: 'Chart Deck', sub: 'chartData' },
  { id: 'logs', label: 'Tactical Log', sub: 'activities' },
];

export function App() {
  // Token in React state ONLY — never persisted.
  const [token, setToken] = useState('');
  const [tab, setTab] = useState<Tab>('command');
  const [toast, setToast] = useState<
    { kind: 'ok' | 'err' | 'info'; text: string } | null
  >(null);
  const [now, setNow] = useState(new Date());
  const [proxy, setProxy] = useState({ available: false, configured: false, authenticated: false });
  const [pin, setPin] = useState('');
  const [setupPin, setSetupPin] = useState('');
  const [themeMode, setThemeMode] = useState<ThemeMode>('dark');

  useEffect(() => {
    if (!toast) return;
    const id = window.setTimeout(() => setToast(null), 6000);
    return () => window.clearTimeout(id);
  }, [toast]);

  useEffect(() => {
    const id = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(id);
  }, []);

  async function refreshProxyStatus() {
    try {
      const s = await getProxyStatus();
      setProxy({ available: true, configured: s.configured, authenticated: s.authenticated });
    } catch {
      setProxy({ available: false, configured: false, authenticated: false });
    }
  }

  useEffect(() => {
    refreshProxyStatus();
  }, []);

  function pushToast(kind: 'ok' | 'err' | 'info', text: string) {
    setToast({ kind, text });
  }

  async function onProxySetup() {
    try {
      const s = await setupProxy(token.trim(), setupPin.trim());
      setProxy({ available: true, configured: s.configured, authenticated: s.authenticated });
      setToken('');
      setSetupPin('');
      pushToast('ok', 'Privater Proxy eingerichtet. Token liegt jetzt serverseitig.');
    } catch (e: any) {
      pushToast('err', e?.message || String(e));
    }
  }

  async function onProxyLogin() {
    try {
      const s = await loginProxy(pin.trim());
      setProxy({ available: true, configured: s.configured, authenticated: s.authenticated });
      setPin('');
      pushToast('ok', 'Proxy-Session aktiv.');
    } catch (e: any) {
      pushToast('err', e?.message || String(e));
    }
  }

  async function onProxyLogout() {
    await logoutProxy();
    await refreshProxyStatus();
    pushToast('info', 'Proxy-Session beendet.');
  }

  return (
    <div className={`min-h-full flex flex-col ${themeMode === 'garmin' ? 'garmin-mode' : 'dark-mode'}`}>
      {/* Header */}
      <header className="border-b border-border bg-bg-raised/70 backdrop-blur-sm sticky top-0 z-30">
        <div className="max-w-[1400px] mx-auto px-5 py-3 flex items-center gap-5">
          <div className="flex items-center gap-2.5 text-ink">
            <Logo size={22} />
            <div className="leading-tight">
              <div className="text-sm font-semibold tracking-tight">Coach Kira</div>
              <div className="text-2xs text-ink-dim tnum">Cockpit · MVP</div>
            </div>
          </div>

          <nav className="hidden md:flex gap-0.5 ml-4" role="tablist">
            {TABS.map((t) => (
              <button
                key={t.id}
                role="tab"
                aria-selected={tab === t.id}
                onClick={() => setTab(t.id)}
                className={`px-3 py-1.5 rounded text-sm transition-colors ${
                  tab === t.id
                    ? 'bg-bg-subtle text-ink'
                    : 'text-ink-muted hover:text-ink hover:bg-bg-subtle/60'
                }`}
                data-testid={`tab-${t.id}`}
              >
                <span className="block">{t.label}</span>
                <span className="block text-2xs text-ink-dim tnum mt-0.5">{t.sub}</span>
              </button>
            ))}
          </nav>

          <div className="ml-auto flex items-center gap-3 text-2xs tnum text-ink-dim">
            <div
              className="theme-switch inline-flex items-center rounded-md border border-border bg-bg-panel p-0.5 shadow-sm"
              role="group"
              aria-label="Darstellungsmodus wählen"
              data-testid="theme-mode-switch"
            >
              <button
                onClick={() => setThemeMode('dark')}
                className={`theme-switch-option px-2.5 py-1 text-xs rounded transition-colors ${
                  themeMode === 'dark'
                    ? 'bg-bg-subtle text-ink'
                    : 'text-ink-muted hover:text-ink'
                }`}
                data-testid="button-theme-dark"
                aria-pressed={themeMode === 'dark'}
                title="Darkmode aktivieren"
              >
                Dark
              </button>
              <button
                onClick={() => setThemeMode('garmin')}
                className={`theme-switch-option px-2.5 py-1 text-xs rounded transition-colors ${
                  themeMode === 'garmin'
                    ? 'bg-bg-subtle text-ink'
                    : 'text-ink-muted hover:text-ink'
                }`}
                data-testid="button-theme-garmin"
                aria-pressed={themeMode === 'garmin'}
                title="Garmin-Mode aktivieren"
              >
                Garmin
              </button>
            </div>
            <span className="hidden sm:inline">{fmtTime(now.toISOString())}</span>
            <span className="hidden md:inline">Europe/Berlin</span>
          </div>
        </div>

        {/* mobile tab strip */}
        <nav className="md:hidden flex overflow-x-auto px-3 pb-2 gap-1" role="tablist">
          {TABS.map((t) => (
            <button
              key={t.id}
              role="tab"
              aria-selected={tab === t.id}
              onClick={() => setTab(t.id)}
              className={`shrink-0 px-3 py-1.5 rounded text-xs ${
                tab === t.id ? 'bg-bg-subtle text-ink' : 'text-ink-muted'
              }`}
            >
              {t.label}
            </button>
          ))}
        </nav>
      </header>

      <main className="flex-1 max-w-[1400px] w-full mx-auto px-5 py-6">
        {tab === 'command' && (
          <CommandCenter
            token={token}
            setToken={setToken}
            toast={pushToast}
            proxy={proxy}
            pin={pin}
            setPin={setPin}
            setupPin={setupPin}
            setSetupPin={setSetupPin}
            onProxySetup={onProxySetup}
            onProxyLogin={onProxyLogin}
            onProxyLogout={onProxyLogout}
          />
        )}
        {tab === 'plan' && <PlanCockpit token={token} toast={pushToast} proxyAuthenticated={proxy.authenticated} />}
        {tab === 'charts' && <ChartDeck />}
        {tab === 'logs' && <TacticalLog />}
      </main>

      <footer className="border-t border-border mt-6">
        <div className="max-w-[1400px] mx-auto px-5 py-3 flex flex-wrap items-center justify-between gap-2 text-2xs text-ink-dim tnum">
          <span>Coach Kira pptx.app · MVP-Prototyp · Live-Daten via Apps-Script-Endpunkte</span>
          <span>
            Proxy: {proxy.available ? (proxy.authenticated ? 'aktiv' : proxy.configured ? 'bereit' : 'Setup nötig') : 'nicht verfügbar'} ·
            Token: {token ? 'gesetzt (Session)' : 'nicht gesetzt'}
          </span>
        </div>
      </footer>

      <Toast toast={toast} onClose={() => setToast(null)} />
    </div>
  );
}
