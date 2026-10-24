/** Kira-Kommentare zu Aktivitäten aus AI_ACTIVITY_REVIEWS (Apps Script mode=activityReviews). */
import { BASE_URL } from './api';

export type ActivityReview = {
  date: string;
  text: string;
  engine: 'local' | 'cloud';
  model: string | null;
};

export async function fetchActivityReviews(): Promise<Map<string, ActivityReview>> {
  try {
    const res = await fetch(`${BASE_URL}?mode=activityReviews&limit=2000&cb=${Date.now()}`, { redirect: 'follow' });
    const j = await res.json();
    const list: ActivityReview[] = j && j.ok && Array.isArray(j.reviews) ? j.reviews : [];
    return new Map(list.map((r) => [r.date, r]));
  } catch {
    // Patch noch nicht eingespielt oder offline: einfach keine Kommentare
    return new Map();
  }
}
