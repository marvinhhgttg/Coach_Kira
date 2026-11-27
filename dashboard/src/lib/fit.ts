/**
 * Minimaler FIT-Encoder für strukturierte Workouts (Garmin).
 * Nachrichten: file_id (0), workout (26), workout_step (27).
 * Herzfrequenz-Ziele als absolute bpm (FIT: Wert + 100), Leistung als Watt (FIT: Wert + 1000).
 */

export type FitSport = 'running' | 'cycling' | 'walking' | 'hiking' | 'training' | 'generic';
export type FitIntensity = 'active' | 'rest' | 'warmup' | 'cooldown' | 'recovery' | 'interval';

export interface FitStep {
  name: string;
  minutes: number; // 0 = offen (Lap-Taste)
  intensity: FitIntensity;
  hrLow?: number | null;
  hrHigh?: number | null;
  powerLow?: number | null;
  powerHigh?: number | null;
  notes?: string;
}

export interface FitBlock {
  repeat: number; // 1 = einfach
  steps: FitStep[];
}

export interface FitWorkout {
  name: string;
  sport: FitSport;
  blocks: FitBlock[];
}

const SPORT: Record<FitSport, number> = { generic: 0, running: 1, cycling: 2, training: 10, walking: 11, hiking: 17 };
const INTENSITY: Record<FitIntensity, number> = { active: 0, rest: 1, warmup: 2, cooldown: 3, recovery: 4, interval: 5 };

// Basistypen: enum 0x00, uint16 0x84, uint32 0x86, string 0x07, uint32z 0x8C
type Field = { num: number; size: number; base: number; value: number | string };

const CRC_TABLE = [
  0x0000, 0xcc01, 0xd801, 0x1400, 0xf001, 0x3c00, 0x2800, 0xe401, 0xa001, 0x6c00, 0x7800, 0xb401, 0x5000, 0x9c01, 0x8801,
  0x4400,
];
function crc16(bytes: Uint8Array, start = 0, end = bytes.length): number {
  let crc = 0;
  for (let i = start; i < end; i++) {
    const b = bytes[i];
    let tmp = CRC_TABLE[crc & 0xf];
    crc = (crc >> 4) & 0x0fff;
    crc = crc ^ tmp ^ CRC_TABLE[b & 0xf];
    tmp = CRC_TABLE[crc & 0xf];
    crc = (crc >> 4) & 0x0fff;
    crc = crc ^ tmp ^ CRC_TABLE[(b >> 4) & 0xf];
  }
  return crc;
}

function asciiBytes(s: string, size: number): number[] {
  // FIT-Strings sind UTF-8; Umlaute ersetzen wir sicherheitshalber durch ASCII
  const clean = s
    .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue')
    .replace(/Ä/g, 'Ae').replace(/Ö/g, 'Oe').replace(/Ü/g, 'Ue').replace(/ß/g, 'ss')
    .replace(/[^\x20-\x7e]/g, '');
  const out = new Array(size).fill(0);
  for (let i = 0; i < Math.min(clean.length, size - 1); i++) out[i] = clean.charCodeAt(i);
  return out;
}

class Writer {
  bytes: number[] = [];
  u8(v: number) { this.bytes.push(v & 0xff); }
  u16(v: number) { this.u8(v); this.u8(v >> 8); }
  u32(v: number) { this.u16(v & 0xffff); this.u16((v >>> 16) & 0xffff); }
  definition(local: number, global: number, fields: Field[]) {
    this.u8(0x40 | local); // Definition
    this.u8(0); // reserviert
    this.u8(0); // little endian
    this.u16(global);
    this.u8(fields.length);
    for (const f of fields) { this.u8(f.num); this.u8(f.size); this.u8(f.base); }
  }
  data(local: number, fields: Field[]) {
    this.u8(local);
    for (const f of fields) {
      if (f.base === 0x07) { asciiBytes(String(f.value), f.size).forEach((b) => this.u8(b)); continue; }
      const v = Number(f.value);
      if (f.size === 1) this.u8(v);
      else if (f.size === 2) this.u16(v);
      else this.u32(v);
    }
  }
  message(local: number, global: number, fields: Field[]) {
    this.definition(local, global, fields);
    this.data(local, fields);
  }
}

const FIT_EPOCH = Date.UTC(1989, 11, 31, 0, 0, 0) / 1000;

export function buildWorkoutFit(w: FitWorkout): Uint8Array {
  const wr = new Writer();
  const now = Math.floor(Date.now() / 1000) - FIT_EPOCH;

  wr.message(0, 0, [
    { num: 0, size: 1, base: 0x00, value: 5 }, // type = workout
    { num: 1, size: 2, base: 0x84, value: 255 }, // manufacturer = development
    { num: 2, size: 2, base: 0x84, value: 0 }, // product
    { num: 3, size: 4, base: 0x8c, value: (now ^ 0x5a5a5a) >>> 0 }, // serial
    { num: 4, size: 4, base: 0x86, value: now }, // time_created
  ]);

  // Schritte aufbauen (Wiederholungen als Repeat-Schritt hinter dem Block)
  type RawStep = Field[];
  const steps: RawStep[] = [];
  const stepFields = (idx: number, s: FitStep): Field[] => {
    const hasHr = s.hrLow && s.hrHigh;
    const hasPw = !hasHr && s.powerLow && s.powerHigh;
    return [
      { num: 254, size: 2, base: 0x84, value: idx },
      { num: 0, size: 16, base: 0x07, value: s.name || '' },
      { num: 1, size: 1, base: 0x00, value: s.minutes > 0 ? 0 : 5 }, // time | open
      { num: 2, size: 4, base: 0x86, value: s.minutes > 0 ? Math.round(s.minutes * 60 * 1000) : 0 },
      { num: 3, size: 1, base: 0x00, value: hasHr ? 1 : hasPw ? 4 : 2 }, // heart_rate | power | open
      { num: 4, size: 4, base: 0x86, value: 0 },
      { num: 5, size: 4, base: 0x86, value: hasHr ? Math.round(s.hrLow!) + 100 : hasPw ? Math.round(s.powerLow!) + 1000 : 0 },
      { num: 6, size: 4, base: 0x86, value: hasHr ? Math.round(s.hrHigh!) + 100 : hasPw ? Math.round(s.powerHigh!) + 1000 : 0 },
      { num: 7, size: 1, base: 0x00, value: INTENSITY[s.intensity] ?? 0 },
      { num: 8, size: 32, base: 0x07, value: s.notes || '' },
    ];
  };
  for (const b of w.blocks) {
    const first = steps.length;
    for (const s of b.steps) steps.push(stepFields(steps.length, s));
    if (b.repeat > 1 && b.steps.length) {
      steps.push([
        { num: 254, size: 2, base: 0x84, value: steps.length },
        { num: 0, size: 16, base: 0x07, value: '' },
        { num: 1, size: 1, base: 0x00, value: 6 }, // repeat_until_steps_cmplt
        { num: 2, size: 4, base: 0x86, value: first },
        { num: 3, size: 1, base: 0x00, value: 2 },
        { num: 4, size: 4, base: 0x86, value: b.repeat },
        { num: 5, size: 4, base: 0x86, value: 0 },
        { num: 6, size: 4, base: 0x86, value: 0 },
        { num: 7, size: 1, base: 0x00, value: 0 },
        { num: 8, size: 32, base: 0x07, value: '' },
      ]);
    }
  }

  wr.message(1, 26, [
    { num: 4, size: 1, base: 0x00, value: SPORT[w.sport] ?? 0 },
    { num: 6, size: 2, base: 0x84, value: steps.length },
    { num: 8, size: 24, base: 0x07, value: w.name || 'Coach Kira' },
  ]);
  if (steps.length) {
    wr.definition(2, 27, steps[0]);
    for (const s of steps) wr.data(2, s);
  }

  const body = Uint8Array.from(wr.bytes);
  const header = new Uint8Array(14);
  const dv = new DataView(header.buffer);
  header[0] = 14;
  header[1] = 0x20; // Protokoll 2.0
  dv.setUint16(2, 2132, true); // Profil 21.32
  dv.setUint32(4, body.length, true);
  header.set([0x2e, 0x46, 0x49, 0x54], 8); // ".FIT"
  dv.setUint16(12, crc16(header, 0, 12), true);

  const out = new Uint8Array(14 + body.length + 2);
  out.set(header, 0);
  out.set(body, 14);
  const c = crc16(out, 0, 14 + body.length);
  out[out.length - 2] = c & 0xff;
  out[out.length - 1] = (c >> 8) & 0xff;
  return out;
}

export function downloadFit(w: FitWorkout, filename: string) {
  const bytes = buildWorkoutFit(w);
  const blob = new Blob([bytes as unknown as BlobPart], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.endsWith('.fit') ? filename : filename + '.fit';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
