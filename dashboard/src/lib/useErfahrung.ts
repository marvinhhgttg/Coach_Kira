/** Erfahrungsmodell einmal je Sitzung aus der ganzen Historie laden (ohne KI) + Schalter „ins Votum einrechnen“. */
import { useEffect, useState } from 'react';
import { fetchChartData } from './api';
import { ALL_RANGE } from './windows';
import { RECOVERY_FETCH_METRICS } from '../screens/RecoveryPanel';
import { erfahrungAusRohdaten, type Erfahrung } from './votumErfahrung';

let cache: Promise<Erfahrung | null> | null = null;
const LS_AN = 'ck_votum_erfahrung';
const EVT = 'ck-erfahrung-schalter';

export function ladeErfahrung(neu = false): Promise<Erfahrung | null> {
  if (!cache || neu)
    cache = fetchChartData(ALL_RANGE, [...RECOVERY_FETCH_METRICS])
      .then((res: any) => erfahrungAusRohdaten(res.data || []))
      .catch(() => {
        cache = null;
        return null;
      });
  return cache;
}

export function useErfahrung(): Erfahrung | null {
  const [e, setE] = useState<Erfahrung | null>(null);
  useEffect(() => {
    let off = false;
    ladeErfahrung().then((x) => !off && setE(x));
    return () => {
      off = true;
    };
  }, []);
  return e;
}

export function erfahrungAn(): boolean {
  try {
    return localStorage.getItem(LS_AN) !== 'aus';
  } catch {
    return true;
  }
}

export function setErfahrungAn(an: boolean) {
  try {
    localStorage.setItem(LS_AN, an ? 'an' : 'aus');
  } catch {
    /* ignore */
  }
  window.dispatchEvent(new Event(EVT));
}

export function useErfahrungAn(): [boolean, (an: boolean) => void] {
  const [an, setAn] = useState(erfahrungAn());
  useEffect(() => {
    const f = () => setAn(erfahrungAn());
    window.addEventListener(EVT, f);
    return () => window.removeEventListener(EVT, f);
  }, []);
  return [an, setErfahrungAn];
}
