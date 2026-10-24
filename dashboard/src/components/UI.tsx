import type { ReactNode } from 'react';

export function Panel({
  title,
  children,
  right,
  className = '',
}: {
  title?: string;
  children: ReactNode;
  right?: ReactNode;
  className?: string;
}) {
  return (
    <section className={`panel ${className}`}>
      {(title || right) && (
        <header className="flex flex-wrap items-center justify-between gap-2 px-4 py-2.5 border-b border-border">
          {title && <h2 className="label">{title}</h2>}
          {right}
        </header>
      )}
      <div className="p-4">{children}</div>
    </section>
  );
}

export function Skeleton({ className = '' }: { className?: string }) {
  return (
    <div
      className={`bg-bg-subtle/60 rounded animate-pulse ${className}`}
      aria-hidden
    />
  );
}

export function ErrorBox({ message, onRetry }: { message: string; onRetry?: () => void }) {
  return (
    <div className="rounded border border-ampel-rot/40 bg-ampel-rot/5 px-3 py-2 text-sm text-red-300 flex items-center justify-between gap-3">
      <span className="truncate">Fehler: {message}</span>
      {onRetry && (
        <button onClick={onRetry} className="btn btn-ghost text-xs px-2 py-1">
          Erneut
        </button>
      )}
    </div>
  );
}

export function StatPill({
  label,
  value,
  hint,
  emphasize,
}: {
  label: string;
  value: ReactNode;
  hint?: ReactNode;
  emphasize?: boolean;
}) {
  return (
    <div className="panel-raised px-4 py-3">
      <div className="label">{label}</div>
      <div
        className={`tnum ${emphasize ? 'text-3xl' : 'text-2xl'} font-semibold mt-1 leading-none`}
      >
        {value}
      </div>
      {hint && <div className="text-2xs text-ink-muted mt-1.5 tnum">{hint}</div>}
    </div>
  );
}

export function ArcGauge({
  label,
  value,
  sub,
  tone = 'auto',
}: {
  label: string;
  value: number | null | undefined;
  sub?: ReactNode;
  tone?: 'auto' | 'blue' | 'orange' | 'green' | 'red';
}) {
  const v = typeof value === 'number' && Number.isFinite(value) ? Math.max(0, Math.min(100, value)) : null;
  const circumference = 2 * Math.PI * 42;
  const arc = circumference * 0.72;
  const gap = circumference - arc;
  const offset = v === null ? arc : arc * (1 - v / 100);
  const color =
    tone === 'green'
      ? '#22c55e'
      : tone === 'orange'
        ? '#f97316'
        : tone === 'red'
          ? '#ef4444'
          : tone === 'blue'
            ? '#38bdf8'
            : v === null
              ? '#6b7280'
              : v < 30
                ? '#ef4444'
                : v < 50
                  ? '#f97316'
                  : v < 70
                    ? '#eab308'
                    : v < 85
                      ? '#38bdf8'
                      : '#22c55e';

  return (
    <div className="panel-raised p-4 min-h-[172px] flex flex-col items-center justify-between">
      <div className="label self-start">{label}</div>
      <div className="relative h-28 w-36">
        <svg viewBox="0 0 120 100" className="h-full w-full">
          <defs>
            <linearGradient id={`gauge-${label.replace(/\s+/g, '-')}`} x1="0" x2="1" y1="1" y2="0">
              <stop offset="0%" stopColor="#ef4444" />
              <stop offset="30%" stopColor="#f97316" />
              <stop offset="55%" stopColor="#eab308" />
              <stop offset="78%" stopColor="#22c55e" />
              <stop offset="100%" stopColor="#38bdf8" />
            </linearGradient>
          </defs>
          <path
            d="M18 70 A42 42 0 1 1 102 70"
            fill="none"
            stroke="#1f2731"
            strokeWidth="9"
            strokeLinecap="round"
          />
          <path
            d="M18 70 A42 42 0 1 1 102 70"
            fill="none"
            stroke={`url(#gauge-${label.replace(/\s+/g, '-')})`}
            strokeWidth="9"
            strokeLinecap="round"
            pathLength={100}
            strokeDasharray={`${v === null ? 0 : v} 100`}
          />
          {v !== null && (
            <circle
              cx={18 + (84 * v) / 100}
              cy={70 - Math.sin((Math.PI * v) / 100) * 42}
              r="4"
              fill={color}
              stroke="#0a0d12"
              strokeWidth="2"
            />
          )}
        </svg>
        <div className="absolute inset-x-0 bottom-5 text-center">
          <div className="tnum text-3xl font-semibold leading-none" style={{ color }}>
            {v === null ? '—' : Math.round(v)}
          </div>
        </div>
      </div>
      {sub && <div className="text-2xs text-ink-muted tnum text-center">{sub}</div>}
    </div>
  );
}

export function Toast({
  toast,
  onClose,
}: {
  toast: { kind: 'ok' | 'err' | 'info'; text: string } | null;
  onClose: () => void;
}) {
  if (!toast) return null;
  const colors =
    toast.kind === 'ok'
      ? 'border-ampel-gruen/50 bg-ampel-gruen/10 text-green-200'
      : toast.kind === 'err'
        ? 'border-ampel-rot/50 bg-ampel-rot/10 text-red-200'
        : 'border-accent-strong/40 bg-accent-strong/10 text-accent';
  return (
    <div className="fixed bottom-4 right-4 z-50 max-w-sm">
      <div className={`rounded border px-3.5 py-2.5 text-sm shadow-lg ${colors}`}>
        <div className="flex items-start gap-3">
          <div className="flex-1">{toast.text}</div>
          <button onClick={onClose} className="text-ink-muted hover:text-ink text-xs">
            ✕
          </button>
        </div>
      </div>
    </div>
  );
}

export function Modal({
  open,
  onClose,
  title,
  children,
}: {
  open: boolean;
  onClose: () => void;
  title: string;
  children: ReactNode;
}) {
  if (!open) return null;
  return (
    <div
      className="fixed inset-0 z-40 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4"
      onClick={onClose}
    >
      <div
        className="panel-raised w-full max-w-lg shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-center justify-between px-4 py-3 border-b border-border">
          <h3 className="text-sm font-semibold tracking-tight">{title}</h3>
          <button onClick={onClose} className="btn btn-ghost text-xs">
            Schliessen
          </button>
        </header>
        <div className="p-4 text-sm">{children}</div>
      </div>
    </div>
  );
}

export function Spinner({ size = 14 }: { size?: number }) {
  return (
    <span
      className="inline-block rounded-full border-2 border-border border-t-accent animate-spin"
      style={{ width: size, height: size }}
      aria-hidden
    />
  );
}
