/**
 * Coach Kira · Prompt-Werkstatt – EINZIGE Quelle aller Kira-Prompts.
 * Wird im Cockpit (Browser) genutzt und für den Mac als kira_prompts.mjs gebaut
 * (tools/mac-runner, `node kira_prompts.mjs <typ>` liest JSON von stdin).
 * Das Apps Script liefert nur noch Daten (supervisorJob.data, briefingData, telegramJobs).
 * Texte 1:1 aus dem Apps Script übernommen (Stand Build 2026-10-17__kira_chat, 26.10.2026).
 */
/* eslint-disable */

export type KiMsg = { role: 'system' | 'user' | 'assistant'; content: string };
export type BriefingTyp = 'status' | 'plan' | 'consolidated';

// ---------------------------------------------------------------------------
// Regeln & Wissen
// ---------------------------------------------------------------------------
export const MONOTONIE_REGEL = [
  "Monotonie-Regel: Ein einzelner hoher Monotonie-Wert ist meist kein Problem. Maßgeblich ist die rollierende 7-Tage-Monotonie (Ø/σ der Tageslasten) im Verlauf.",
  "Zonen: < 1,5 gute Variation; 1,5–2,0 normal; > 2,0 aufmerksam (vor allem bei hoher Wochenlast); > 2,5 ungünstig, wenn es länger anhält; > 3,0 Warnzone.",
  "Muster: 1–2 ähnliche Tage normal; 3–4 ähnliche, fordernde Tage beobachten (bei schlechterem Schlaf, HRV, Ruhepuls, schweren Beinen oder weniger Motivation reduzieren); 5–7 Tage ohne deutliche Entlastung relevant; mehrere Wochen hohe Monotonie = Risiko für Ermüdung, Stagnation, Überlastung, Infekte.",
  "Steuerung: Ein hoher Wert an einem Tag ist praktisch bedeutungslos. Liegt die 7-Tage-Monotonie an 2 Tagen in Folge über etwa 2,0 und sind schon 4–5 fordernde Tage gesammelt → nächster Tag sehr locker oder frei. Als Entlastung zählt ein lockerer oder freier Tag (höchstens 40 ESS bzw. 60 % der üblichen Tageslast). Bleibt der Wert über 2,5 → innerhalb 24–48 h echter Ruhetag oder sehr kurzer Regenerationsreiz. Über 3,0 → mindestens ein echter Entlastungstag oder deutliche Reduktion. Hohe Monotonie plus schlechte Erholungswerte → keine Qualitätseinheit durchdrücken.",
  "Entscheidend ist die Kombination mit der Gesamtlast (Strain = Wochenlast × Monotonie). Monotonie senkst du durch klaren Rhythmus: harte Reize klar hart, lockere Tage wirklich locker, mindestens ein echter Erholungstag bei hoher Gesamtlast – nicht durch weniger Gesamtlast."
].join('\n');

/** Erfahrung aus Nichtbefolgung (lib/votumErfahrung.ts, fest berechnet). Leer → Prompt unverändert. */
export function erfahrungBlock(text?: string): string {
  const t = String(text || '').trim();
  if (!t) return '';
  return '\n---\n**ERFAHRUNG AUS NICHTBEFOLGUNG (aus Marcs eigenen Daten, fest berechnet – berücksichtigen):**\n' + t + '\n---\n';
}

function monotonieBlock() {
  return '\n---\n**MONOTONIE-REGEL (verbindlich für Bewertung und Plan):**\n' + MONOTONIE_REGEL +
    '\nBewerte die Prognose-Monotonie oben nach dieser Regel: ein einzelner hoher Tageswert ist kein Grund zum Kürzen; ' +
    'entscheidend sind fordernde Tage ohne Entlastung, anhaltend hohe Werte und die Wochenlast.\n---\n';
}

const WISSENSBLOCK_V105 = `
--- KERNWISSEN (Firstbeat Whitepapers) ---
Du (Kira) musst diese Definitionen als "Quelle der Wahrheit" für deine Analysen verwenden.

1.  **Training Effect (TE) Skala (0.0-5.0)[cite: 1799, 2123]:**
    * **Basis:** Basiert auf EPOC (Excess Post-Exercise Oxygen Consumption)[cite: 1667, 2102].
    * **EPOC:** Misst die Störung der Homöostase (also den Erholungsbedarf) durch das Training[cite: 1721, 2102, 3460].
    * **TE 1.0-1.9 (Minor):** Erhält die Grundlagenausdauer oder fördert die Erholung[cite: 1809, 2123].
    * **TE 2.0-2.9 (Maintaining):** Hält die aktuelle kardiorespiratorische Fitness[cite: 1811, 2123].
    * **TE 3.0-3.9 (Improving):** Verbessert die kardiorespiratorische Fitness (empfohlen 2-4x pro Woche)[cite: 1813, 2123].
    * **TE 4.0-4.9 (Highly Improving):** Verbessert stark (empfohlen 1-2x pro Woche), erfordert mehr Aufmerksamkeit bei der Erholung[cite: 1815, 2123].
    * **TE 5.0 (Overreaching):** Extremer Reiz ("Overreaching")[cite: 1817, 2123]. Führt nur mit ausreichender Erholung zur Verbesserung[cite: 2123].


2.  **Training Load & ACWR (Belastungssteuerung):**
    * **Acute Load (ATL):** Die Summe des TRIMP der letzten 7 Tage.
    * **Chronic Load (CTL):** Der Durchschnitts-TRIMP der letzten 28 Tage.
    * **ACWR (Acute:Chronic Workload Ratio):** Das Verhältnis von Akut zu Chronisch.
    * **Zielzone ("Sweet Spot"):** ACWR zwischen **0.8 und 1.2**.
    * **Untertraining (ACWR < 0.8):** Deutet auf **Formverlust / Detraining** hin (NICHT auf Überforderung!).
    * **Überlastung (ACWR > 1.5):** Gefahrenzone! Verletzungsrisiko steigt signifikant.
    * **WICHTIG:** Es gibt **keine** mathematische Korrelation zwischen ACWR (Last) und HRV (Stress). Behandle sie als separate Signale.
// ...

3.  **HRV (Stress & Erholung)[cite: 1026]:**
    * **Grundlage:** Die HRV-Analyse spiegelt das Autonome Nervensystem (ANS) wider[cite: 1044, 1164, 1435].
    * **Stress-Reaktion:** Wird durch Dominanz des **Sympathischen Nervensystems** (fight or flight) angezeigt (HR steigt, HRV sinkt)[cite: 1033, 1154, 1304].
    * **Erholung:** Wird durch Dominanz des **Parasympathischen (Vagalen) Nervensystems** (rest and digest) angezeigt (HR sinkt, HRV steigt)[cite: 1034, 1158, 1298].
    * **Schlechte Erholung / Ermüdung:** Zeigt sich, wenn die sympathische Dominanz (Stress) *auch in Ruhephasen* (z.B. Schlaf) anhält[cite: 2669, 2898, 2901].

4.  **VO2max:**
    * **Definition:** Die "goldene Standardmessung" für die aerobe Fitness[cite: 3042, 3047].
--- ENDE KERNWISSEN ---
`;

/** KK_WISSEN: aktive Inhalte (Spalte C bei Marker 'x'); null = Blatt fehlt → Kernwissen */
function wissensText(wissen: string[] | null | undefined) {
  if (!Array.isArray(wissen)) return WISSENSBLOCK_V105;
  let t = "--- EXPERTEN-WISSENSDATENBANK (KIRA SYSTEM) ---\n";
  wissen.forEach((c) => { if (c) t += c + "\n\n"; });
  return t + "--- ENDE WISSENSDATENBANK ---\n";
}

function trefferBlock(tr: { datum: string; text: string } | null | undefined) {
  if (!tr || !tr.text) return '';
  return '\n---\n**KIRA-TREFFSICHERHEIT (Rückblick, Stand ' + tr.datum + '):**\n' +
    'So hat Marc in den letzten Wochen auf deine Empfehlungen reagiert. Berücksichtige das bei Plan und Empfehlung.\n' +
    String(tr.text).slice(0, 1500) + '\n---\n';
}

function planungsBlock(p: { heute: string; ziele: { datum: string; text: string }[]; wochen: any[] } | null | undefined) {
  if (!p) return '';
  const heute = p.heute;
  let out = '';
  const ziele = (p.ziele || []).filter((z) => z.datum >= heute).sort((a, b) => (a.datum < b.datum ? -1 : 1));
  if (ziele.length) out += '**ZIELE / TOUREN (darauf hinplanen):**\n' + ziele.slice(0, 5).map((r) => '- ' + r.datum + ': ' + String(r.text).slice(0, 300)).join('\n') + '\n';
  const wochen = p.wochen || [];
  if (wochen.length) {
    const w = wochen.filter((x: any) => x.start <= heute).pop() || wochen[0];
    const n = wochen[wochen.indexOf(w) + 1];
    if (w) out += '**SAISONRAHMEN diese Woche (ab ' + w.start + '):** Phase ' + w.phase + ', Ziel-Wochenlast ca. ' + w.ziel_ess + ' ESS, Fokus: ' + (w.fokus || '-') + (n ? ' · nächste Woche: ' + n.phase + ', ca. ' + n.ziel_ess + ' ESS' : '') + '\n';
  }
  return out ? '\n---\n' + out + 'Richte Plan und Empfehlung an Zielen und Saisonrahmen aus, ohne die Monotonie- und ACWR-Regeln zu verletzen.\n---\n' : '';
}

// ---------------------------------------------------------------------------
// Zahl-Helfer (wie im Apps Script)
// ---------------------------------------------------------------------------
function parseGermanFloat_(value: any): number {
  if (value === null || value === undefined) return NaN;
  const s = String(value).trim();
  if (!s) return NaN;
  const cleaned = s.replace(/\s/g, '').replace('%', '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.');
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : NaN;
}
function parseGermanFloat(value: any): any {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const n = parseGermanFloat_(value);
    return Number.isFinite(n) ? n : value;
  }
  return value;
}
function getBaselineNumber_(baselineObj: any, aliases: string[]) {
  const b = baselineObj || {};
  const keys = Object.keys(b);
  const norm = (x: any) => String(x || '').trim().toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/g, '');
  const aliasNorm = (aliases || []).map(norm);
  for (let i = 0; i < keys.length; i++) {
    const kNorm = norm(keys[i]);
    if (!kNorm) continue;
    if (aliasNorm.indexOf(kNorm) !== -1) {
      const n = parseGermanFloat_(b[keys[i]]);
      if (Number.isFinite(n)) return n;
    }
  }
  for (let i = 0; i < keys.length; i++) {
    const kNorm = norm(keys[i]);
    if (!kNorm) continue;
    for (let a = 0; a < aliasNorm.length; a++) {
      const wanted = aliasNorm[a];
      if (!wanted) continue;
      if (kNorm.indexOf(wanted) !== -1 || wanted.indexOf(kNorm) !== -1) {
        const n = parseGermanFloat_(b[keys[i]]);
        if (Number.isFinite(n)) return n;
      }
    }
  }
  return NaN;
}

// ---------------------------------------------------------------------------
// Supervisor (KI-Report): Datenpaket aus supervisorJob.data → Prompt + JSON-Schema
// ---------------------------------------------------------------------------
export const SUPERVISOR_SYSTEM = 'Du bist ein präziser Daten-Analyst. Deine Aufgabe ist es, die Daten zu analysieren und NUR das angeforderte JSON zu füllen.';

export const SUPERVISOR_SCHEMA: any = {
    "name": "fitness_report",
    "strict": true,
    "schema": {
      "type": "object",
      "properties": {
        "empfehlung_zukunft": {
          "type": "object",
          "properties": {
            "plan_status": { "type": "string", "enum": ["FREIGABE", "BEOBACHTUNG", "INTERVENTION"] },
            "text": { "type": "string", "description": "Ausführliche Gesamtanalyse" }
          },
          "required": ["plan_status", "text"],
          "additionalProperties": false
        },
        "bewertung_ernaehrung_7d": {
          "type": "object",
          "properties": {
            "text": { "type": "string", "description": "Kurzbewertung Ernährung" }
          },
          "required": ["text"],
          "additionalProperties": false
        },
        "einzelscore_kommentare": {
          "type": "array",
          "description": "Liste der Kommentare für jeden einzelnen Score",
          "items": {
            "type": "object",
            "properties": {
              "metrik": { "type": "string" },
              "text_info": { "type": "string", "description": "Kurzer Kommentar (1 Satz)" }
            },
            "required": ["metrik", "text_info"],
            "additionalProperties": false
          }
        },
        "empfohlener_plan": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "datum": { "type": "string" },
              "tag": { "type": "string" },
              "original_load_ess": { "type": "number" },
              "empfohlener_load_ess": { "type": "number" },
              "empfohlene_zone": { "type": "string" },
              "beispiel_training_1": { "type": "string" },
              "beispiel_training_2": { "type": "string" },
              "prognostizierte_auswirkung": { 
                "type": "string", 
                "description": "Detaillierte Erklärung." 
              },
              "wetterempfehlung": { "type": "string" }
            },
            "required": ["datum", "tag", "original_load_ess", "empfohlener_load_ess", "empfohlene_zone", "beispiel_training_1", "beispiel_training_2", "prognostizierte_auswirkung", "wetterempfehlung"],
            "additionalProperties": false
          }
        }
      },
      "required": ["empfehlung_zukunft", "bewertung_ernaehrung_7d", "einzelscore_kommentare", "empfohlener_plan"],
      "additionalProperties": false
    }
  };

export function supervisorPrompt(d: any): { system: string; prompt: string; schema: any } {
  return { system: SUPERVISOR_SYSTEM, prompt: supervisorText(d), schema: SUPERVISOR_SCHEMA.schema };
}

function supervisorText(d: any): string {
  const data = { heute: d.heute || {}, zukunft: { plan: d.zukunft_plan || [] }, baseline: d.baseline || {}, history: d.history || '', monotony_varianz: Number(d.monotony_varianz) || 0, rtp_status: String(d.rtp_status || '') };
  const fitnessScores: any[] = d.fitnessScores || [];
  const nutritionScore: any = d.nutritionScore || { raw_wert: 0, num_score: 0 };
  const gesamtScoreNum = d.gesamtScoreNum;
  const gesamtScoreAmpel = d.gesamtScoreAmpel;
  const zoneFactors = d.zoneFactors;
  const elevFactors = d.elevFactors;
  const weekConfigCSV = d.weekConfigCSV;
  const isActivityDone = !!d.isActivityDone;
  const wetterDaten = d.wetterDaten;
  
  const alleScores = fitnessScores.concat([nutritionScore]);
  const { heute, zukunft, baseline, history, monotony_varianz, rtp_status } = data; 

  // --- ACWR FIX: einzig wahre Quelle für Text_Info = coachE_ACWR_forecast (HEUTE) ---
const acwrRawHeute = (heute?.coachE_ACWR_forecast ?? heute?.coache_acwr_forecast);
const acwrForecastNum = parseGermanFloat_(acwrRawHeute);

const acwrFix = Number.isFinite(acwrForecastNum)
  ? acwrForecastNum.toFixed(2).replace('.', ',')  // "1,13"
  : String(acwrRawHeute ?? '').trim();

const acwrFixDot = acwrFix.replace(',', '.');      // "1.13" (falls du es numerisch brauchst)



  // TE-Balance heute kommt fertig aus dem Apps Script (AI_DATA_FORECAST), sonst monotony_varianz
  const fallbackTeBalanceHeute = (monotony_varianz * 100).toFixed(1).replace('.', ',') + "%";
  const teBalanceHeute = d.teBalanceHeute || fallbackTeBalanceHeute;

  // --- 1. DEFINITIONEN & DATUM FIX ---
  const dynamicKnowledge = wissensText(d.wissen);

  const cleanDate = String(d.heute_datum || '');
  const cleanTag = String(d.heute_tag || '');
  
  let heuteString = `- ${cleanDate} (${cleanTag}): `; 

  const heutigerIstLoad = parseGermanFloat(heute['load_fb_day']);
  const isActivityDone_Text = isActivityDone ? "true" : "false";
  const heutigerSollLoad = parseGermanFloat(heute['coachE_ESS_day']); 
  
  // Kontext-Header
  let statusHeuteHeader = isActivityDone 
      ? `✅ TRAINING HEUTE BEREITS ABSOLVIERT (Ist-Load: ${heutigerIstLoad})`
      : `⏳ TRAINING HEUTE NOCH OFFEN (Soll-Load: ${heutigerSollLoad})`;

  let scoresText = "";
  alleScores.forEach(s => { 
      // FIX ACWR: Komma zu Punkt, damit die KI die Zahl versteht (1.30 statt 1,30)
      let cleanVal = String(s.raw_wert).replace(',', '.');
      
      // FIX ACWR: Wenn Wert hoch ist (>0.8) aber Score 0 (weil rot), 
      // dann sag der KI explizit "RISIKO", sonst denkt sie "Daten fehlen".
      let scoreInfo = s.num_score;
      if (s.metrik.includes("ACWR") && parseFloat(cleanVal) > 0.8 && s.num_score < 10) {
         scoreInfo = `${s.num_score} (ACHTUNG: HOCHRISIKO-BEREICH!)`; 
      }
      
      scoresText += `- ${s.metrik}: ${cleanVal} (Score: ${scoreInfo}, ${s.ampel})\n`; 
  });
  const metricsChecklist = alleScores.map(s => `"${s.metrik}"`).join(', ');
  const keiMetricFromFitnessScores = fitnessScores.find(s => s && s.metrik === "KEI");
  const keiPromptValue = keiMetricFromFitnessScores ? String(keiMetricFromFitnessScores.raw_wert || "—") : "—";

  // -----------------------------------------------------------
  // 2. RTP INTELLIGENZ: MATRIX-BERECHNUNG
  // -----------------------------------------------------------
  const isRTP = (rtp_status && rtp_status.includes("RTP"));
  
  // A) Phase extrahieren
  let daysSinceHealthy = 0;
  if (isRTP) {
      try { daysSinceHealthy = parseInt(rtp_status.split('_')[2]); } catch (e) { daysSinceHealthy = 0; }
  }

  // B) Schweregrad ermitteln (Krankheitstage zählen)
  let sickDaysCount = 0;
  if (isRTP && history) {
      const rows = history.split('\n');
      for (let i = rows.length - 1; i >= 0; i--) {
          const rowStr = rows[i].toLowerCase();
          if (rowStr.includes("krank") || rowStr.includes("sick") || rowStr.includes("infekt")) {
              sickDaysCount++;
          } else if (sickDaysCount > 0 && rowStr.includes(",")) { 
              break; 
          }
      }
  }
  
  // C) FALLBACK
  if (isRTP && sickDaysCount === 0) {
      sickDaysCount = 7; 
  }

  // D) Smart Offset Berechnung (Start = 10 - Krank)
  let protocolOffset = 0;
  if (sickDaysCount < 9) {
      protocolOffset = 9 - sickDaysCount;
      if (protocolOffset < 0) protocolOffset = 0;
  }

  // E) Helper: Effektiven Tag berechnen
  const getEffectiveDay = (dayIndexFromNow: number) => {
      if (!isRTP) return 0;
      const arrayIndex = (daysSinceHealthy - 1) + protocolOffset + dayIndexFromNow; 
      const effectiveDay = arrayIndex + 1; 
      return (effectiveDay <= 9) ? effectiveDay : 0; 
  };

  // RTP RAMPE (9 Tage)
  const rtpRamp = [10, 30, 60, 0, 60, 0, 70, 80, 0]; 

  // -----------------------------------------------------------
  // 3. HEUTE: WAS IST DRAN?
  // -----------------------------------------------------------
  let todayPlan_LoadRec = 0;
  let todayPlan_Zone = "Ruhetag (Fixed)";
  let todayPlan_Sport1 = "Geplanter Ruhetag";
  let todayPlan_Sport2 = "Yoga / Stretching";
  let todayPlan_Text = "Regeneration.";
  let heuteZusatzInfo = "";

  const todayEffectiveDay = isRTP ? getEffectiveDay(0) : 0;

  if (todayEffectiveDay > 0) {
      const rtpLoad = rtpRamp[todayEffectiveDay - 1]; 
      
      if (isActivityDone) {
          heuteZusatzInfo = ` -> ✅ RTP TAG ${todayEffectiveDay}/9 ERLEDIGT.`;
          todayPlan_LoadRec = heutigerIstLoad; 
          todayPlan_Text = `RTP (Tag ${todayEffectiveDay}/9). Einheit absolviert.`;
          todayPlan_Zone = "Erledigt";
          todayPlan_Sport1 = "Absolviert";
      } else {
          heuteZusatzInfo = ` -> ⚠️ RTP TAG ${todayEffectiveDay}/9 (Smart Entry). Load: ${rtpLoad}.`;
          todayPlan_LoadRec = rtpLoad;
          todayPlan_Text = `RTP-Schutz (Tag ${todayEffectiveDay}/9). Protokoll aktiv.`;

          switch (todayEffectiveDay) {
              case 1: todayPlan_Zone="Z1/Rek"; todayPlan_Sport1="Spaziergang"; break;
              case 2: todayPlan_Zone="Z1 (Bike)"; todayPlan_Sport1="Diagnose-Rolle"; break;
              case 3: todayPlan_Zone="Z2 (Bike)"; todayPlan_Sport1="Rolle Intensiv"; break;
              case 4: todayPlan_Zone="Ruhetag"; todayPlan_Sport1="Integration (Ruhe)"; break;
              case 5: todayPlan_Zone="Z2 (Run)"; todayPlan_Sport1="Lauf (kurz)"; break;
              case 6: todayPlan_Zone="Ruhetag"; todayPlan_Sport1="Integration (Ruhe)"; break;
              case 7: todayPlan_Zone="Z2 (Run)"; todayPlan_Sport1="Lauf (Steigerung)"; break;
              case 8: todayPlan_Zone="Z2 (Bike)"; todayPlan_Sport1="Lange Rolle"; break;
              case 9: todayPlan_Zone="Ruhetag"; todayPlan_Sport1="Abschluss"; break;
          }
      }
  }

  let heutigerFixMarker = (heute['fix'] || ' ').toString().trim();
  let prognoseLoadText = "Original Load (ESS):\n";
  prognoseLoadText += `${heuteString}${heutigerSollLoad}${heuteZusatzInfo}\n`;
  let prognoseFixText = "Fixed Tage (fix='x'):\n" + `${heuteString}${heutigerFixMarker}\n`; 
  let prognoseMonoText = "Prognose Monotonie (Load):\n" + `${heuteString}${heute['Monotony7']} (Belastungsfokus Ratio: ${(monotony_varianz * 100).toFixed(1)}%)\n`; 
  let prognoseAcwrText = "Prognose ACWR:\n" + `${heuteString}${acwrFix}\n`;

  // -----------------------------------------------------------
  // 4. ZIELWERT BERECHNEN
  // -----------------------------------------------------------
  let peakCTL = 0;
  try {
      if (history) {
          const histLines = history.split('\n');
          for (let i = histLines.length - 1; i > 0; i--) {
              const parts = histLines[i].split(',').map((s: string) => s.replace(/"/g, '').trim());
              const histCTL = parseFloat(parts[5]); 
              if (!isNaN(histCTL) && histCTL > peakCTL) peakCTL = histCTL;
          }
      }
  } catch(e) {}
  const currentCTL = parseGermanFloat(heute['coachE_CTL_forecast']) || parseGermanFloat(heute['fbCTL_obs']) || 60;
  const anchorCTL = (peakCTL > 0) ? peakCTL : currentCTL;
  const targetRecoveryLoad = Math.round(anchorCTL * 0.6);

  // -----------------------------------------------------------
  // 5. DATEN & FAKTOREN
  // -----------------------------------------------------------
  const resting_calories_raw = baseline?.['Resting_Calories (kcal)'];
const resting_calories_num = parseGermanFloat(resting_calories_raw);
const resting_calories = Number.isFinite(resting_calories_num) ? resting_calories_num : 0;

  const rhr_default = getBaselineNumber_(baseline, ['RHR_default (bpm)', 'RHR_default(bpm)', 'rhr_default (bpm)', 'rhr_default_bpm', 'rhr_default']);
  const vo2max = parseGermanFloat(baseline['VO2max (mg/kg/min)']);
  const ftp_rad = parseGermanFloat(baseline['FTP_Rad (W/kg)']);
  const hill_score = parseGermanFloat(baseline['Hill_Score (0-100)']);
  const run_econ = parseGermanFloat(baseline['Running_Economy (184-225)']);
  
  const allgemeineZiele = baseline['Allgemeine Ziele'] || "Keine Ziele definiert.";

  let ernaehrungLabel = "";
  const rawDeficit = parseFloat(nutritionScore.raw_wert);
  if (rawDeficit > 0) ernaehrungLabel = "DEFIZIT";
  else if (rawDeficit < 0) ernaehrungLabel = "ÜBERSCHUSS";
  else ernaehrungLabel = "AUSGEGLICHEN";

  const problemScores = alleScores.filter(s => s.num_score < 80)
                                   .map(s => `${s.metrik}: ${s.raw_wert} (${s.ampel}/${s.num_score})`)
                                   .join('; ');

  const isRecoveryCritical = alleScores.some(s => 
      (s.metrik === "HRV Status" || s.metrik === "Training Readiness" || s.metrik === "RHR") && s.num_score < 40
  );
  const isRecoveryBad = alleScores.some(s => 
      (s.metrik === "HRV Status" || s.metrik === "Training Readiness" || s.metrik === "RHR") && s.num_score < 80
  );
  const isRecoverySuper = alleScores.every(s => 
      (s.metrik !== "HRV Status" && s.metrik !== "Training Readiness") || s.num_score >= 90
  );

  // -----------------------------------------------------------
  // 6. LOOP DURCH ZUKUNFT (AB MORGEN)
  // -----------------------------------------------------------
  zukunft.plan.forEach((tag: any, i: number) => { 
    let tagString = `${tag.datum} (${tag.tag})`; 
    let phase = (tag.week_phase) ? String(tag.week_phase).trim().toUpperCase() : "A"; 
    let phaseMarker = "🚀 [PHASE A: AUFBAU]";
    let loadInstruction = `Geplanter Load [ORIGINAL-PLAN: ${tag.original_load_ess}]`; 


    if (phase === "E") {
        phaseMarker = "🛑 [PHASE E: ENTLASTUNG]";
        loadInstruction = `Original Plan: ${tag.original_load_ess} -> ZIEL: Reduzieren auf ca. ${targetRecoveryLoad}`;
    }

    if (phase === "E") {
    phaseMarker = "🛑 [PHASE E: ENTLASTUNG]";
    loadInstruction = `Original Plan: ${tag.original_load_ess} -> ZIEL: Reduzieren auf ca. ${targetRecoveryLoad}`;
} else if (phase.startsWith("A") && phase !== "A") {
    // macht A1/A2/A3 im Text sichtbar, A bleibt wie bisher
    phaseMarker = `🚀 [PHASE ${phase}: AUFBAU]`;
}


    // RTP OVERRIDE
    if (isRTP) {
        const futureEffDay = getEffectiveDay(i + 1); 

        if (futureEffDay > 0) {
            phaseMarker = `🚑 [RTP TAG ${futureEffDay}/9]`;
            
            const isDone = (tag.original_act_load && tag.original_act_load > 5);
            if (isDone) {
                 loadInstruction = `✅ ERLEDIGT (Ist: ${tag.original_act_load}). [ORIGINAL-PLAN: ${tag.original_load_ess}]`;
            } else {
                const baseInst = `[ORIGINAL-PLAN: ${tag.original_load_ess}] (IGNORIEREN!) -> ⚠️ RTP VORGABE: `;
                switch (futureEffDay) {
                    case 1: loadInstruction = `${baseInst}Ziel ~10 (Walk). Zone: 'Z1/Rek'.`; break;
                    case 2: loadInstruction = `${baseInst}Ziel ~30 (Bike). Zone: 'Z1 (Bike)'.`; break;
                    case 3: loadInstruction = `${baseInst}Ziel ~60 (Bike). Zone: 'Z2 (Bike)'.`; break;
                    case 4: loadInstruction = `${baseInst}RUHETAG. Zone: 'Ruhetag'.`; break;
                    case 5: loadInstruction = `${baseInst}Ziel ~60 (Run). Zone: 'Z2 (Run)'.`; break;
                    case 6: loadInstruction = `${baseInst}RUHETAG. Zone: 'Ruhetag'.`; break;
                    case 7: loadInstruction = `${baseInst}Ziel ~70 (Run). Zone: 'Z2 (Run)'.`; break;
                    case 8: loadInstruction = `${baseInst}Ziel ~80 (Bike). Zone: 'Z2 (Bike)'.`; break;
                    case 9: loadInstruction = `${baseInst}RUHETAG. Zone: 'Ruhetag'.`; break;
                    default: loadInstruction = `RTP Fehler.`;
                }
            }
        } 
    }

    prognoseLoadText += `${phaseMarker} - ${tagString}: ${loadInstruction}\n`;
    prognoseFixText += `- ${tagString}: ${tag.original_fix_marker || ' '}\n`; 
    prognoseMonoText += `- ${tagString}: ${tag.original_monotony}\n`;
    prognoseAcwrText += `- ${tagString}: ${tag.original_acwr}\n`;
  });



  // --- PROMPT ZUSAMMENBAU ---
  let prompt = `Du bist Coach Kira, eine erfahrene KI-Sportwissenschaftlerin und Fitnessexpertin.
Analysiere **deine (Marcs)** Daten. Ich habe die Scores (0-100) berechnet.
Deine Aufgabe ist es, im JSON-Format Folgendes zu liefern:
1. 'empfehlung_zukunft': Eine ausführliche textliche Gesamtbewertung ("text") und einen Plan-Status ("plan_status").
2. 'bewertung_ernaehrung_7d': Eine **kurze (max. 2 Sätze)** textliche Bewertung der Ernährungsbilanz ("text").
3. 'empfohlener_plan': Einen angepassten 8-TAGE-PLAN (HEUTE + 7 TAGE ZUKUNFT).

4. 'einzelscore_kommentare': Ein Array von Objekten für JEDEN der 10 FITNESS-SCORES (NICHT Ernährung), mit "metrik" und "text_info".
**WICHTIG: Formuliere alle deine Analysen sehr detailliert und kontextbezogen ('text', 'text_info', 'prognostizierte_auswirkung', 'wetterempfehlung') direkt an **dich (Marc)** und verwende die 'Du'-Anrede.**

${dynamicKnowledge}

--- 🎯 DER ABSOLUTE DATEN-ANKER FÜR HEUTE (${cleanDate}) ---
! WICHTIG: Falls Historien-Tabellen abweichende Werte enthalten: IGNORIEREN !
! NUR DIESE WERTE SIND DIE WAHRHEIT FÜR DEINEN TEXT !
- Ruhepuls (RHR) heute: ${heute['rhr_bpm']} bpm
- HRV-Status heute: ${heute['hrv_status']} ms (Bereich: ${heute['hrv_threshholds']})
- Schlaf: ${heute['sleep_hours']}h (Score: ${heute['sleep_score_0_100']}/100)
- Trainingszustand: ${heute['trainingszustand']}
- Aktueller ACWR (Forecast): ${acwrFix}  (WERT IST FIX, genau so verwenden)
- Heutige Belastung (IST): ${heutigerIstLoad} ESS
- Heutige Belastung (SOLL laut Plan): ${heutigerSollLoad} ESS
- TE-Balance (% Intensiv): ${teBalanceHeute} %
- KEI (aus fitnessScores): ${keiPromptValue}
- Protein-Zufuhr: ${heute['protein_g']}g
- Energie-Bilanz (Ø 7 Tage): ${nutritionScore.raw_wert} kcal
- RTP-STATUS (System-Schutz): ${rtp_status}
-----------------------------------------------------------

---
**STATUS HEUTE: ${statusHeuteHeader}**
(Nutze diese Info zwingend für die Perspektive deiner Analyse!)
---

---
DEINE ALLGEMEINEN ZIELE (KONTEXT):
**${allgemeineZiele}**
---

---
**DEIN IDEALER WOCHENPLAN (ALS HINWEIS, NIEDRIGE PRIORITÄT)**
(CSV: day_of_week,load,sport,zone)
**${weekConfigCSV}**
---

---
**NEU (V42): WETTER PROGNOSE (HEUTE & MORGEN):**
${wetterDaten}
---
${trefferBlock(d.treffer)}
${monotonieBlock()}${erfahrungBlock(d && d.erfahrung)}
${planungsBlock(d.planung)}


---
**DEINE SCHÄTZ-FAKTOREN (ZUM BERECHNEN VON BEISPIELEN):**
Zonen-Faktoren (Punkte pro Stunde): ${JSON.stringify(zoneFactors)}
Höhen-Faktoren (Punkte pro 100hm): ${JSON.stringify(elevFactors)}
**FORMEL: Load = (Minuten / 60) * ZonenFaktor + (Höhenmeter / 100) * HöhenFaktor**
---

DATEN-BASELINE (PHYSIOLOGISCHES PROFIL)
- Ruheumsatz: ${resting_calories.toFixed(0)} kcal/Tag
- RHR-Basis: ${rhr_default} bpm
- VO2max: ${vo2max} (mg/kg/min)
- FTP (Rad): ${ftp_rad} (W/kg)
- Hill Score: ${hill_score} / 100
- Running Economy: ${run_econ} (niedriger ist besser)

---
**VERGANGENE 14 TAGE (HISTORIE & KONTEXT):**
(Zeigt den IST-Load (Actual_Load) und die Training Effect (TE) Verteilung)
${history}
---

BERECHNETER GESAMTSTATUS (HEUTE: ${cleanDate})
- Gesamtscore: ${gesamtScoreNum} / 100 (${gesamtScoreAmpel})
- HEUTIGER Original-Load (ESS): ${heutigerSollLoad}
- HEUTIGER IST-Load (load_fb_day): ${heutigerIstLoad}
- HEUTIGER Wochentag: ${cleanTag}
- HEUTIGE Monotonie (Load): ${heute['Monotony7']}
- HEUTIGE TE Balance: ${(monotony_varianz * 100).toFixed(1)}% (Dies ist das Ratio Intensiv/Basis)
- V106-INFO: isRecoveryCritical == ${isRecoveryCritical} (mindestens ein Bio-Marker < 40/ROT)
- V106-INFO: isRecoveryBad == ${isRecoveryBad} (mindestens ein Bio-Marker < 80/GELB)
- V106-INFO: isRecoverySuper == ${isRecoverySuper} (Bio-Marker Top > 90)
- **RTP-STATUS (Krankheit): ${rtp_status}** (Wenn != HEALTHY -> ALARM!)
- **RTP-SCHWEREGRAD:** ${sickDaysCount} Tage krank -> Start bei Tag ${todayEffectiveDay}/9.
- Problem-Scores (heute): ${problemScores || 'Keine'}

---
BERECHNETE EINZELSCORES (HEUTE):
${scoresText}
---
GEPLANTER TRAININGSPLAN & PROGNOSE (HEUTE + 7 TAGE)
(Dies ist der *alte* Plan. Du sollst ihn verbessern. Achte auf [FIXED] Marker!)
${prognoseLoadText}
${prognoseFixText}
${prognoseMonoText}
${prognoseAcwrText}
---
DEINE AUFGABE (MAXIMALER KONTEXT & SPORTWISSENSCHAFTLICHE ANALYSE):

**Berücksichtige bei allen Empfehlungen die 'Allgemeinen Ziele' UND die 'Vergangene 14 Tage (Historie)', um konsistent zu planen.**

1.  **ANALYSIERE ERNÄHRUNG ('bewertung_ernaehrung_7d'):**
    * Bewerte die 7-Tage-Bilanz (Score ${nutritionScore.num_score}, Wert ${nutritionScore.raw_wert} kcal) **in maximal 2 Sätzen**.
    * (REGEL: Der Wert ${nutritionScore.raw_wert} kcal ist ein ${ernaehrungLabel}).

2.  **ANALYSIERE GESAMTSTATUS & LEITE PLANSTATUS AB ('empfehlung_zukunft'):**
        * **AUFTRAG:** Schreibe eine **umfassende, tiefgehende Analyse (mindestens 150-200 Wörter)**.
        * **STRUKTUR (Halte dich grob daran):**
          A) **Status Quo (Der Ist-Zustand):** Wie stehen deine Bio-Marker (HRV/RHR/Schlaf) im Vergleich zur Basis? Was ist der Haupttreiber für den aktuellen Gesamtscore (${gesamtScoreAmpel})?
          B) **Die Verbindung (Korrelation):** Erkläre das "Warum". 
          C) **Strategie & Ausblick:** Was bedeutet das konkret für die nächsten 3-4 Tage? Beziehe die **Wochen-Phase** (Aufbau vs. Entlastung) und die **TE-Balance** mit ein.
          D) **ACWR-Kontext-Check (WICHTIG):**
             - Bewerte einen niedrigen ACWR (< 0.8) **NIEMALS** als "Überforderung". Das ist technisch falsch.
             - **Szenario 1 (Schlechte Bio-Marker / Krank):** Ein niedriger ACWR ist hier **GUT und ERWARTET** (notwendige Erholung). Lobe die Disziplin zur Pause.
             - **Szenario 2 (Gute Bio-Marker / Fit):** Ein niedriger ACWR ist hier **SCHLECHT** (Formverlust/Untertraining). Warne davor, dass wir Potenziale verschenken.

        * **STIL:** Sprich wie ein erfahrener Performance-Coach. Analytisch, präzise, aber motivierend. Keine Floskeln, sondern Fakten aus den Daten.
    
    **WICHTIG - KONTEXT-WEICHE:**
    * **FALL A (TRAINING BEREITS ABSOLVIERT):** Bewerte die durchgeführte Einheit rückblickend! War sie für deinen heutigen Zustand (HRV/Readiness) angemessen? War sie vielleicht zu hart? Oder genau richtig? Beachte, dass dadurch auch die Training Readiness naturgemäß niedrig ist. Richte den Blick dann auf die Regeneration für morgen.
    * **FALL B (TRAINING NOCH OFFEN):**
      Gib eine klare Handlungsanweisung und Motivation für das anstehende Training.

3.  **PASSE 8-TAGE-PLAN AN & VALIDIERE ('empfohlener_plan'):**
    * Erstelle einen **VOLLSTÄNDIGEN 8-TAGE-PLAN** (heute + 7 Folgetage).
    * **Entferne Kcal/Makros.** Der Plan enthält nur die Felder aus dem JSON-Beispiel unten.
    * **WICHTIG (V56-FIX):** Das Feld 'original_load_ess' in deinem JSON MUSS exakt mit dem 'Original Load (ESS)' aus dem Prognose-Block übereinstimmen.
    
    * **(V136-MOD) UNVERHANDELBARE REGELN (HIERARCHIE IST WICHTIG):**

      *** WICHTIGSTE REGEL (PRIORITÄT 1): PERIODISIERUNG & PHASEN-DISZIPLIN ***
      Schau dir für JEDEN EINZELNEN TAG im Abschnitt 'GEPLANTER TRAININGSPLAN' den Marker an.
      Der Marker entscheidet über Load-Strategie UND Text-Label.
      
      * **WENN '🚀 [PHASE A: AUFBAU - REIZ SETZEN!]':**
        * **TEXT-REGEL (STRENG):** Du darfst das Wort "Entlastungswoche" NIEMALS verwenden. Auch nicht, wenn der Gesamtstatus "VORSICHT" ist.
        * **LOAD-REGEL:** Du sollst "normale" Reize (Z3, Z4, Intervalle) planen.
        * **CAUTION-EXCEPTION:** Falls der Gesamtstatus "ANPASSUNG VORSICHT" ist, darfst du den Load leicht reduzieren (max. -20%), aber du darfst NICHT in den "Tapering/Recovery"-Modus wechseln. Nenne es "Angepasster Aufbau" oder "Moderater Reiz", aber NICHT Entlastung.
        
      * **WENN '🛑 [PHASE E: ENTLASTUNG - FÜSSE HOCH!]':**
        * **TEXT-REGEL:** Der Text MUSS zwingend mit "Entlastungswoche." beginnen.
        * **LOAD-REGEL:** Du MUSST den Load drastisch senken (Zielwert im Schnitt ca. ${targetRecoveryLoad}).
        * **VERBOT:** Plane KEINE hohen Reize (Z4/Z5) und keine langen Einheiten, egal wie gut die Readiness ist.

      * **REGEL 0: FIXED DAYS (PRIO 0):**
      * Tage mit 'x' sind unveränderlich. Übernimm den 'original_load_ess'.

      * **REGEL 0.8: KRANKHEITS-RÜCKKEHR (RETURN-TO-PLAY - HÖCHSTE SICHERHEIT):**
        * Status: **${rtp_status}**
        * **WICHTIG:** Wenn der Status 'RTP_PHASE_X' ist, gilt das Protokoll. Wir steigen SMART ein (basierend auf Krankheitsdauer).
        * **DEINE PFLICHT:**
          1. **DATEN-INTEGRITÄT (EXTREM WICHTIG):** Das Feld 'original_load_ess' darf NIEMALS verändert werden! Es MUSS exakt der Wert sein, der oben im Text hinter "[ORIGINAL-PLAN: ...]" steht.
          2. Schreibe den RTP-Zielwert AUSSCHLIESSLICH in das Feld 'empfohlener_load_ess'.
          3. **RAMP-UP LOGIK (ZWINGEND):** Halte dich strikt an die Anweisungen im Abschnitt 'GEPLANTER TRAININGSPLAN' oben (z.B. RTP Tag 8 -> Bike Load 80).
          4. **WICHTIG:** Wenn in der Anweisung oben steht "Setze Zone X", dann DARFST du in 'empfohlene_zone' NICHT 'Ruhetag' schreiben (außer es ist wirklich ein Ruhetag)!
          5. Begründung MUSS lauten: "RTP-Schutz (Tag X/9). Protokoll aktiv."
        * Ein Plan mit Z4 oder hohem Load in dieser Phase ist ein FEHLER von dir!

      * **REGEL 1: HEUTIGER STATUS (isActivityDone == ${isActivityDone_Text}):**
      * WENN 'true': Setze 'empfohlener_load_ess' auf **${heutigerIstLoad}**. Tag ist vorbei.

      * **REGEL 2: PROBLEMLÖSUNG (PRIO 4):**
      * -> **PRIO A (PHYSIO-VETO - KRITISCH):**
        * Gilt, wenn 'isRecoveryCritical == true' (siehe Info oben). 
        * **AKTION:** Load drastisch senken (Max 30% des Originals). Begründung: "Physio-Veto".
      
      * -> **PRIO B (TE BALANCE ZU NIEDRIG < 30%):**
        * Wenn Veto inaktiv: Schlage Intensität vor (Z4/Z5).

      * -> **PRIO C (ACWR ZU HOCH > 1.5):**
        * Reduziere Load.

      * **REGEL 7: DETAIL-TIEFE & BEGRÜNDUNG ('prognostizierte_auswirkung'):**
        * Sei nicht wortkarg!
        * **Inhalt:** Erkläre den physiologischen Nutzen (z.B. Mitochondrien, Laktattoleranz).
        
        * **ABWEICHUNGS-CHECK (WICHTIG):**
          * Vergleiche deinen 'empfohlener_load_ess' mit dem 'original_load_ess'.
          * Ist dein Vorschlag **> 20% niedriger** als das Original?
          * **DANN MUSS** der Text mit der Begründung beginnen!
        
        * **In PHASE E:** Text MUSS mit "Entlastungswoche." beginnen.
        * **In PHASE A:** Text darf NICHT "Entlastungswoche" enthalten.

4.  **ANALYSIERE EINZELSCORES ('einzelscore_kommentare'):**
      * Erstelle ein Array von Objekten für **JEDEN** der 10 Fitness-Scores.
      * **WICHTIG (V110):** Du MUSST für **JEDEN** der folgenden Keys einen Kommentar liefern (exakte Schreibweise):
      * [${metricsChecklist}]
      
      *** SPEZIAL-BRIEFING "KEI" (Der Wahrheits-Detektor): ***
    * Dieser Wert misst: "Fitness-Gewinn (CTL-Trend) MINUS Aufwand (Strain)".
    * **INTERPRETATIONSHILFE FÜR DICH (KEI-LEGENDEN-SKALA):**
      * **Wert >= 8:** "High Efficiency" (sehr effizient, loben).
      * **Wert 3 bis < 8:** "Productive" (solider Aufbau).
      * **Wert 0 bis < 3:** "Maintenance / Transition" (Erhalt/Übergang).
      * **Wert < 0:** Ineffizienter Reiz / Warnsignal.
        -> Schreib klar: "Aufwand hoch, Ertrag niedrig. Plan anpassen."
    * Gib für jeden Score einen **aussagekräftigen Kommentar (1-2 Sätze)**.

REGELN FÜR 'plan_status':
- 'INTERVENTION': Wenn 'isRecoveryCritical == true' ODER RTP_STATUS aktiv.
- 'BEOBACHTUNG': Wenn 'isRecoveryBad == true' ODER ACWR > 1.3.
- 'FREIGABE': Wenn alle Scores GRÜN sind.

ANTWORTE AUSSCHLIESSLICH IM FOLGENDEN JSON-FORMAT (ohne Markdown):
{
  "empfehlung_zukunft": {
    "plan_status": "FREIGABE | BEOBACHTUNG | INTERVENTION", 
    "text": "DEINE AUSFÜHRLICHE GESAMTBEWERTUNG..."
  },
  "bewertung_ernaehrung_7d": {
    "text": "DEINE KURZE (MAX 2 SÄTZE) BEWERTUNG DER ERNÄHRUNG..."
  },
  "einzelscore_kommentare": [
    { "metrik": "RHR", "text_info": "Dein Kommentar zu RHR..." },
    { "metrik": "ACWR (Forecast)", "text_info": "Dein Kommentar zum ACWR..." }
    /* ... und ALLE anderen aus der Liste ... */
  ],
  "empfohlener_plan": [
    {
      "datum": "${cleanDate}",
      "tag": "${cleanTag}",
      "original_load_ess": 0, // KOPIERE HIER STUMPFE DEN WERT AUS "[ORIGINAL-PLAN: ...]"! NICHT ÄNDERN!
      "empfohlener_load_ess": ${todayPlan_LoadRec},
      "empfohlene_zone": "${todayPlan_Zone}", 
      "beispiel_training_1": "${todayPlan_Sport1}",
      "beispiel_training_2": "${todayPlan_Sport2}",
      "prognostizierte_auswirkung": "${todayPlan_Text}",
      "wetterempfehlung": "Wetter egal."
    }
    /* ... 7 weitere Tage ... */
  ]
}


`;
  return prompt;
}

// ---------------------------------------------------------------------------
// Briefings (Status, Plan, Konsolidiert): Daten aus briefingData
// ---------------------------------------------------------------------------
export function briefingMessages(type: BriefingTyp, daten: any): { format: 'text' | 'json'; messages: KiMsg[] } {
  const res = briefingRoh(type, daten);
  if (res.messages.length && res.messages[0].role === 'system') res.messages[0].content = String(res.messages[0].content) + '\n\n' + MONOTONIE_REGEL + erfahrungBlock(daten && daten.erfahrung);
  return res;
}

function briefingRoh(type: BriefingTyp, daten: any): { format: 'text' | 'json'; messages: KiMsg[] } {
  if (type === 'status') {
    const status = daten.status;
    return {
      format: 'text' as const,
      messages: [
        { role: 'system' as const, content: 'Du bist ein präziser deutschsprachiger Ausdauer-, Trainings- und Recovery-Coach.' },
        { role: 'user' as const, content: [
          'Du bist Coach Kira.',
          'Analysiere die folgenden Status-Items knapp, konkret und handlungsorientiert.',
          'Sprich Marc direkt per Du an.',
          'Fokus: Training, Recovery, Schlaf, ACWR, Readiness, Ernährung.',
          'Gib maximal 6 Bulletpoints plus eine klare Empfehlung für heute.',
          '',
          JSON.stringify(status, null, 2)
        ].join('\n') }
      ]
    };
  }

  if (type === 'plan') {
    const status = daten.status;
    const sim = daten.sim;
    return {
      format: 'json' as const,
      messages: [
        { role: 'system' as const, content: [
          'Du bist Coach Kira, eine deutschsprachige Sportwissenschaftlerin, Ausdauercoach und Recovery-Analystin.',
          'Du gibst präzise, konservative, handlungsorientierte Empfehlungen für Marc.',
          'Du nutzt ACWR, ATL, CTL, Training Readiness, Schlaf, HRV, RHR, TE und KEI/Smart Gains.',
          'Du sprichst Marc direkt per Du an.',
          'Du antwortest als valides JSON ohne Markdown-Codeblock.'
        ].join(' ') },
        { role: 'user' as const, content: [
          'Erstelle ein Plan-Briefing für das Plan Cockpit.',
          'Nutze Command-Center-Status, aktuelle PlanApp-Daten und Garmin-/Coach-Kira-Kontext.',
          '',
          'Output-Format exakt:',
          '{',
          '  "briefing": "mehrzeiliger deutscher Text mit konkreten Empfehlungen",',
          '  "recommendations": [',
          '    {"date":"YYYY-MM-DD", "day":"Mo", "recommended_load_ess":130, "recommended_zone":"Z2", "rationale":"kurze Begründung"}',
          '  ]',
          '}',
          '',
          'Regeln:',
          '- recommendations muss genau die sichtbaren 14 Tage enthalten.',
          '- ESS nur realistische Werte 0 bis 300.',
          '- Wenn Recovery/Readiness kritisch ist: Last reduzieren.',
          '- Wenn ACWR > 1.3: keine zusätzliche Steigerung.',
          '- Wenn Schlafdauer oder Readiness schlecht ist: Z2 oder Off bevorzugen.',
          '- Keine medizinischen Diagnosen.',
          '',
          'STATUS:',
          JSON.stringify(status, null, 2),
          '',
          'PLAN_CONTEXT:',
          JSON.stringify(sim, null, 2)
        ].join('\n') }
      ]
    };
  }

  if (type === 'consolidated') {
    const context = daten.context;
    const skillPerspectives = [
      { name: '/garmin-gesamtanalyse', role: 'Gesamtbild aus Trainings-, Recovery-, Schlaf-, HRV-, Ernährungs- und Belastungsdaten.' },
      { name: '/garmin-anomalie-erkennung', role: 'Auffälligkeiten, Datenbrüche, unerwartete Sprünge, mögliche Überlastung oder Recovery-Risiken.' },
      { name: '/garmin-ernaehrung-trainingskorrelation', role: 'Zusammenhang zwischen Energie-/Makroaufnahme, Defizit/Überschuss, Load, Recovery und Schlaf.' },
      { name: '/garmin-wochenplanung', role: 'Adaptive 7- bis 14-Tage-Planung mit Load-, Zone- und Ruhetagslogik.' },
      { name: '/garmin-trainingsstatus', role: 'Trainingszustand, ATL/CTL/ACWR, TE Balance, KEI, Monotony/Strain und Fortschrittslogik.' },
      { name: '/garmin-erholungsstatus', role: 'Schlafdauer, Schlafscore, HRV, RHR, Training Readiness, Stress und Erholungsbedarf.' },
      { name: '/garmin-gesundheitsstatus', role: 'Gesundheitsnahe, nicht-diagnostische Risikoeinschätzung und Watchlist.' }
    ];
    return {
      format: 'text' as const,
      messages: [
        { role: 'system' as const, content: [
          'Du bist Coach Kira: eine deutschsprachige Sportwissenschaftlerin, Ausdauercoach, Recovery-Analystin und Garmin-Datenexpertin.',
          'Du synthetisierst mehrere Analyseperspektiven, statt nur Statuswerte nachzuerzählen.',
          'Du sprichst Marc direkt per Du an.',
          'Keine medizinischen Diagnosen. Bei Gesundheitsrisiken vorsichtig formulieren.',
          'Nutze klare Überschriften und Bulletpoints. Keine JSON-Antwort.'
        ].join(' ') },
        { role: 'user' as const, content: [
          'Erstelle eine konsolidierte Coach-Kira-Lagebeurteilung.',
          '',
          'Die folgenden Garmin-Skill-Perspektiven sollen gedanklich verwendet, synthetisiert und zusammengefasst werden:',
          JSON.stringify(skillPerspectives, null, 2),
          '',
          'Pflichtstruktur:',
          '## Gesamtbild',
          '## Auffälligkeiten und Anomalien',
          '## Erholung und Gesundheitsstatus',
          '## Training und Belastungssteuerung',
          '## Ernährung und Trainingskorrelation',
          '## Plan 14 Tage',
          '## Empfehlung für heute',
          '## Watchlist',
          '',
          'Konkrete Regeln:',
          '- Nenne konkrete Werte, wenn sie entscheidungsrelevant sind.',
          '- Gib klare Handlungsprioritäten.',
          '- Weise auf Widersprüche hin, z.B. guter KEI aber niedrige Readiness.',
          '- Verdichte, nicht ausufern.',
          '- Maximal 900 Wörter.',
          '',
          'DATENKONTEXT:',
          JSON.stringify(context, null, 2)
        ].join('\n') }
      ]
    };
  }

  throw new Error('Unbekannter Briefing-Typ: ' + type);
}


// ---------------------------------------------------------------------------
// Telegram: freie Frage an Kira (Mac, lokal)
// ---------------------------------------------------------------------------
export function telegramMessages(frage: string, ctx: { status?: any; vorausblick?: string; erfahrung?: string; verlauf?: { frage: string; antwort: string }[] }): KiMsg[] {
  const sys = [
    'Du bist Coach Kira, Marcs Sportwissenschaftlerin und Trainingscoach – wie eine gute Freundin, direkt und per Du.',
    'Du antwortest auf Telegram: kurz (max. 6 Sätze), konkret, auf Deutsch, ohne Markdown-Tabellen und ohne Emojis.',
    'Nutze die Daten unten. Erfinde keine Werte. Keine medizinischen Diagnosen.',
    MONOTONIE_REGEL,
  ].join('\n') + erfahrungBlock(ctx.erfahrung);
  const daten = [
    ctx.status ? 'STATUS (Cockpit):\n' + JSON.stringify(ctx.status).slice(0, 6000) : '',
    ctx.vorausblick ? 'VORAUSBLICK:\n' + String(ctx.vorausblick).slice(0, 3000) : '',
  ].filter(Boolean).join('\n\n');
  const msgs: KiMsg[] = [{ role: 'system', content: sys + (daten ? '\n\n' + daten : '') }];
  (ctx.verlauf || []).slice(-3).forEach((v) => {
    msgs.push({ role: 'user', content: v.frage });
    msgs.push({ role: 'assistant', content: v.antwort });
  });
  msgs.push({ role: 'user', content: frage });
  return msgs;
}
