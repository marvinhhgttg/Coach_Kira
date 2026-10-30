/** Workout an Garmin senden: Auftrag beim Cockpit-Server, der Mac lädt es hoch und legt es in den Garmin-Kalender. */
import { SERVER_BASE } from './cloudLLM';
import type { FitWorkout } from './fit';

export type WorkoutAuftrag = { id: string; datum: string; requestedAt: number; pickedAt: number | null; doneAt: number | null; ok: boolean | null; result: string; workoutId?: string };

async function j(res: Response): Promise<any> {
  const t = await res.text();
  let x: any;
  try { x = JSON.parse(t); } catch { throw new Error(`Cockpit-Server antwortet nicht (HTTP ${res.status}).`); }
  if (!x.ok) throw new Error(x.error || `HTTP ${res.status}`);
  return x;
}

export async function workoutSenden(datum: string, workout: FitWorkout & { beschreibung?: string }): Promise<WorkoutAuftrag> {
  const x = await j(await fetch(`${SERVER_BASE}/api/garmin/workout`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ datum, workout }) }));
  return x.auftrag;
}

export async function workoutAuftragStatus(id: string): Promise<WorkoutAuftrag> {
  const x = await j(await fetch(`${SERVER_BASE}/api/garmin/workout?id=${encodeURIComponent(id)}&cb=${Date.now()}`, { cache: 'no-store' }));
  return x.auftrag;
}
