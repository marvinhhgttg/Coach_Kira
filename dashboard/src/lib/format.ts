// Time + number formatting helpers.

export function fmtTime(iso?: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleTimeString('de-DE', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

export function fmtDateTime(iso?: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

export function fmtRelative(iso?: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const diff = Date.now() - d.getTime();
  const s = Math.floor(diff / 1000);
  if (s < 5) return 'gerade eben';
  if (s < 60) return `vor ${s}s`;
  const m = Math.floor(s / 60);
  if (m < 60) return `vor ${m}m`;
  const h = Math.floor(m / 60);
  if (h < 24) return `vor ${h}h`;
  const days = Math.floor(h / 24);
  return `vor ${days}d`;
}

export function fmtDateShort(iso?: string): string {
  if (!iso) return '—';
  // ISO yyyy-mm-dd or full ISO
  const datePart = iso.length >= 10 ? iso.slice(0, 10) : iso;
  const [y, m, d] = datePart.split('-');
  if (!y || !m || !d) return iso;
  return `${d}.${m}.${y.slice(2)}`;
}

export function fmtNum(
  v: number | null | undefined,
  digits = 0
): string {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  return v.toLocaleString('de-DE', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

export function fmtSigned(v: number | null | undefined, digits = 0): string {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  const sign = v > 0 ? '+' : '';
  return sign + fmtNum(v, digits);
}

export function fmtPct(v: number | null | undefined): string {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  const sign = v > 0 ? '+' : '';
  return sign + (v * 100).toFixed(0) + '%';
}
