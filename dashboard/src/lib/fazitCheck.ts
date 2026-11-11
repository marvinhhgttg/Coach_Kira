/** Nachprüfung von KI-Fazits gegen die Fakten. */
/** Zahlen aus dem Fazit, die in den Fakten nicht vorkommen (Rundung ±1 in der letzten Stelle erlaubt). */
export function fremdeZahlen(text: string, quelle: string): string[] {
  const norm = (s: string) => Number(s.replace(/\./g, '').replace(',', '.').replace('−', '-'));
  const zahlen = (s: string) => [...s.matchAll(/[-+−]?\d{1,3}(?:\.\d{3})*(?:,\d+)?|[-+−]?\d+(?:[.,]\d+)?/g)].map((m) => m[0]);
  const erlaubt = new Set<number>();
  for (const z of zahlen(quelle)) { const n = Math.abs(norm(z)); if (Number.isFinite(n)) erlaubt.add(Math.round(n * 100) / 100); }
  const out: string[] = [];
  for (const z of zahlen(text)) {
    const n = Math.abs(norm(z));
    const dezimal = /,\d/.test(z);
    if (!Number.isFinite(n) || (!dezimal && n <= 12) || (!dezimal && n >= 2024 && n <= 2030)) continue; // kleine ganze Zahlen (Zählwörter), Jahre
    const dez = (z.split(',')[1] || '').length;
    const tol = dez ? 10 ** -dez : 1;
    let ok = false;
    for (const e of erlaubt) if (Math.abs(e - n) <= tol + 1e-9) { ok = true; break; }
    if (!ok && !out.includes(z)) out.push(z);
  }
  return out;
}

