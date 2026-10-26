/**
 * Cloud-Aufruf für strukturierte Aufgaben (z. B. Wochenplan) über den vorhandenen Chat-Weg
 * (Apps Script mode=chat → Perplexity). Das Apps Script reicht nur weiter – kein Sheet-Lesen/-Schreiben.
 * Grenzen des Chat-Wegs: je Nachricht 4000 Zeichen, Kontext 15000 Zeichen, 50 Anfragen/Tag.
 */
import { sendChat, type ChatMsg } from './chatContext';
import type { LlmMsg } from './localLLM';

export async function cloudJson(msgs: LlmMsg[]): Promise<string> {
  const system = msgs.filter((m) => m.role === 'system').map((m) => m.content).join('\n\n');
  const rest = msgs.filter((m) => m.role !== 'system');
  const first = rest[0]?.role === 'user' ? rest.shift()! : null;
  const auftrag = [system, first?.content || ''].join('\n\n').slice(0, 14500);
  const chat: ChatMsg[] = [
    { role: 'user', content: 'Erstelle den Plan genau nach dem AUFTRAG in den Cockpit-Daten. Antworte nur mit dem JSON, ohne Text davor oder danach.' },
    ...rest.map((m) => ({ role: m.role as 'user' | 'assistant', content: m.content.slice(0, 3900) })),
  ];
  const r = await sendChat(chat, { AUFTRAG: auftrag } as any, false);
  return String(r.answer || '').replace(/\[\d+\]/g, '');
}
