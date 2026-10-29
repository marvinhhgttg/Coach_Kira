import { useMemo, useState } from 'react';
import { VOTE_LABEL } from '../screens/RecoveryPanel';
import { useErfahrung, useErfahrungAn } from '../lib/useErfahrung';
import { GRUND_LABEL, MIN_N, type Gruppe } from '../lib/votumErfahrung';
import { fmtDateShort } from '../lib/format';

const fz = (v: number | null) => (v == null ? '–' : `${v > 0 ? '+' : v < 0 ? '−' : '±'}${Math.abs(v).toLocaleString('de-DE', { maximumFractionDigits: 1, minimumFractionDigits: 1 })}`);
const pct = (v: number) => `${Math.round(v * 100)} %`;
const KL: Record<Gruppe['klasse'], { t: string; c: string }> = {
  vertraeglich: { t: 'verträglich', c: 'text-green-600 dark:text-green-400' },
  kostet: { t: 'kostet Erholung', c: 'text-orange-600 dark:text-orange-300' },
  unklar: { t: 'kein klares Muster', c: 'text-ink-dim' },
};

/** Was Kira aus Nichtbefolgung lernt – fest berechnet aus der ganzen Historie. */
export function ErfahrungTile() {
  const e = useErfahrung();
  const [an, setAn] = useErfahrungAn();
  const [alle, setAlle] = useState(false);
  const gruppen = useMemo(() => (e ? e.gruppen.filter((g) => (alle ? g.n >= 3 : g.n >= MIN_N)) : []), [e, alle]);

  return (
    <div className="mt-6" data-testid="erfahrung">
      <div className="label">Lernen aus Nichtbefolgung</div>
      <p className="text-2xs text-ink-dim mt-1">
        Wenn du härter trainiert hast als das Votum: Hatte das Folgen – oder hat dein Körper genauso gut reagiert wie an Tagen, an denen dieselbe Belastung erlaubt war?
        Verglichen wird mit Tagen gleicher Einheitsstufe und ähnlicher Ausgangsbereitschaft (±7 Punkte), über 3 Tage. Fest berechnet, ohne KI; lernt mit jedem neuen Tag weiter.
      </p>
      {!e && <div className="text-2xs text-ink-dim mt-2">lädt …</div>}
      {e && (
        <>
          <div className="mt-2 flex flex-wrap items-center gap-3 text-2xs">
            <span className="text-ink-dim">{e.nHaerter} Tage härter als das Votum ausgewertet · Stand {fmtDateShort(e.stand)}</span>
            <label className="flex items-center gap-1 cursor-pointer" data-testid="erfahrung-schalter">
              <input type="checkbox" checked={an} onChange={(x) => setAn(x.target.checked)} />
              Erfahrung ins heutige Votum einrechnen
            </label>
            <label className="flex items-center gap-1 cursor-pointer">
              <input type="checkbox" checked={alle} onChange={(x) => setAlle(x.target.checked)} />
              auch kleine Gruppen zeigen (ab 3)
            </label>
          </div>
          <table className="w-full text-2xs mt-3" data-testid="erfahrung-tabelle">
            <thead className="text-ink-dim text-left">
              <tr className="border-b border-border">
                <th className="pb-1">Sprung</th>
                <th>Grund</th>
                <th>Fälle</th>
                <th>Erholung Tag 1 / 2 / 3</th>
                <th>Tief &lt; 55 (sonst)</th>
                <th>Urteil</th>
              </tr>
            </thead>
            <tbody>
              {gruppen.map((g) => (
                <tr key={g.key} className="border-b border-border">
                  <td className="pt-1">{VOTE_LABEL[g.von]} → {VOTE_LABEL[g.nach]}</td>
                  <td className="pt-1">{g.bind === '*' ? 'alle (ohne Schutzsignale)' : GRUND_LABEL[g.bind]}</td>
                  <td className="pt-1 tnum">{g.n}</td>
                  <td className="pt-1 tnum">{fz(g.d1)} / {fz(g.d2)} / {fz(g.d3)}</td>
                  <td className="pt-1 tnum">{pct(g.tief)} ({pct(g.refTief)})</td>
                  <td className={`pt-1 ${g.n < MIN_N ? 'text-ink-dim' : KL[g.klasse].c}`}>{g.n < MIN_N ? 'zu wenige Fälle' : KL[g.klasse].t}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="text-2xs text-ink-dim mt-2">
            Erholung = Reaktion am Morgen (Bereitschaft, HRV, Ruhepuls, Schlafscore gegen das 7-Tage-Mittel) minus Reaktion an vergleichbaren erlaubten Tagen; positiv = besser erholt.
            Urteil ab {MIN_N} Fällen: verträglich, wenn die Erholung über 3 Tage mindestens +0,3 und das Tief-Risiko kaum höher ist; kostet, wenn sie −0,5 oder schlechter ist oder Tiefs deutlich häufiger sind.
          </p>
          <p className="text-2xs text-ink-dim mt-1">
            Wirkung: Ist der Sprung über dein heutiges Votum „verträglich“, hebt der Recovery Brief das Votum um eine Stufe an – nie bei Schutzsignalen (HRV-Alarm, Ruhepuls hoch bei
            niedriger Bereitschaft, Befinden ≤ 2, rotes Einzelsignal). „Kostet“ erscheint als Warnung. Kira bekommt dieselben Erkenntnisse im KI-Report, in Briefings, im Chat und bei Telegram.
          </p>
        </>
      )}
    </div>
  );
}
