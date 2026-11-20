import type { Ampel } from '../lib/api';

const COLORS: Record<string, { dot: string; bg: string; border: string; text: string; label: string }> = {
  'GRÜN':   { dot: '#22c55e', bg: 'rgba(34,197,94,0.10)',  border: 'rgba(34,197,94,0.35)',  text: '#86efac', label: 'GRUEN' },
  GRUEN:    { dot: '#22c55e', bg: 'rgba(34,197,94,0.10)',  border: 'rgba(34,197,94,0.35)',  text: '#86efac', label: 'GRUEN' },
  GELB:     { dot: '#eab308', bg: 'rgba(234,179,8,0.10)',  border: 'rgba(234,179,8,0.35)',  text: '#fde68a', label: 'GELB' },
  ORANGE:   { dot: '#f97316', bg: 'rgba(249,115,22,0.10)', border: 'rgba(249,115,22,0.35)', text: '#fdba74', label: 'ORANGE' },
  ROT:      { dot: '#ef4444', bg: 'rgba(239,68,68,0.10)',  border: 'rgba(239,68,68,0.40)',  text: '#fca5a5', label: 'ROT' },
  GRAU:     { dot: '#6b7280', bg: 'rgba(107,114,128,0.12)',border: 'rgba(107,114,128,0.40)',text: '#9ca3af', label: 'GRAU' },
  BLAU:     { dot: '#3b82f6', bg: 'rgba(59,130,246,0.12)', border: 'rgba(59,130,246,0.40)', text: '#93c5fd', label: 'BLAU' },
  LILA:     { dot: '#a855f7', bg: 'rgba(168,85,247,0.12)', border: 'rgba(168,85,247,0.40)', text: '#d8b4fe', label: 'LILA' },
};

function resolve(ampel?: Ampel) {
  if (!ampel) return COLORS.GRAU;
  const key = ampel.toString().toUpperCase().normalize('NFC');
  // Apps Script may return 'GRÜN' or 'GRUEN'
  return COLORS[key] || COLORS[key.replace('Ü', 'U')] || COLORS.GRAU;
}

export function AmpelDot({ ampel, size = 8 }: { ampel?: Ampel; size?: number }) {
  const c = resolve(ampel);
  return (
    <span
      aria-label={`Ampel ${c.label}`}
      className="inline-block rounded-full shrink-0"
      style={{
        width: size,
        height: size,
        backgroundColor: c.dot,
        boxShadow: `0 0 0 2px rgba(0,0,0,0.4), 0 0 12px ${c.dot}66`,
      }}
    />
  );
}

export function AmpelChip({ ampel }: { ampel?: Ampel }) {
  const c = resolve(ampel);
  return (
    <span
      className="chip tnum"
      style={{
        backgroundColor: c.bg,
        borderColor: c.border,
        color: c.text,
        border: '1px solid',
      }}
    >
      <AmpelDot ampel={ampel} size={6} />
      {c.label}
    </span>
  );
}

export function AmpelBig({ ampel, label }: { ampel?: Ampel; label?: string }) {
  const c = resolve(ampel);
  return (
    <div
      className="rounded-md p-4 flex items-center gap-4 border"
      style={{ backgroundColor: c.bg, borderColor: c.border }}
    >
      <span
        className="rounded-full block shrink-0"
        style={{
          width: 28,
          height: 28,
          backgroundColor: c.dot,
          boxShadow: `0 0 24px ${c.dot}88, inset 0 0 6px rgba(0,0,0,0.3)`,
        }}
      />
      <div className="min-w-0">
        <div className="label">{label || 'Gesamtampel'}</div>
        <div className="text-lg font-semibold tracking-tight" style={{ color: c.text }}>
          {c.label}
        </div>
      </div>
    </div>
  );
}
