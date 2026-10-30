/** Garmin-Abruf auf Knopfdruck: Auftrag beim Cockpit-Server, der Mac holt ihn im 60-s-Takt ab. */
import { SERVER_BASE } from './cloudLLM';

export type AnfrageStatus = { ok: boolean; pending: boolean; requestedAt?: number; pickedAt?: number | null; doneAt?: number; result?: string; now?: number; error?: string };

async function j(res: Response): Promise<AnfrageStatus> {
  const t = await res.text();
  try {
    return JSON.parse(t);
  } catch {
    throw new Error(`Cockpit-Server antwortet nicht (HTTP ${res.status}).`);
  }
}

export async function garminAnfordern(): Promise<AnfrageStatus> {
  return j(await fetch(`${SERVER_BASE}/api/garmin/anfrage`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}' }));
}

export async function garminAnfrageStatus(): Promise<AnfrageStatus> {
  return j(await fetch(`${SERVER_BASE}/api/garmin/anfrage?cb=${Date.now()}`, { cache: 'no-store' }));
}
