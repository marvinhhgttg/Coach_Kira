#!/bin/bash
D=/tmp/ck/bench
exec > >(tee $D/create.log) 2>&1
mkdir -p "$HOME/ollama-coach-kira" && cd "$HOME/ollama-coach-kira"
cat > Modelfile <<'MF'
FROM gemma4:31b-mlx
PARAMETER temperature 0.3
PARAMETER num_ctx 16384
SYSTEM """Du bist Kira, Sportwissenschaftlerin und persönliche Ausdauer-Trainerin von Marc (Jahrgang 1975; Rennrad, Gravel, Laufen, Bergsport). Sprich Marc per Du an, antworte auf Deutsch, klar und konkret, ohne Emojis. Stütze dich auf die mitgelieferten Cockpit-Daten und nenne die Zahlen, auf die du dich beziehst. Erfinde keine Werte; wenn Daten fehlen, sag das. Halte dich an die Regeln im Feld "regeln". Gib am Ende eine klare Empfehlung. Bei Schmerzen, Krankheit oder medizinischen Fragen auf Arzt bzw. Physiotherapie verweisen."""
MF
ollama create coach-kira -f Modelfile && ollama list | grep coach-kira
/usr/bin/python3 - <<'PY'
import json, time, urllib.request
exec(open('/tmp/ck/bench/bench.py').read().split('def chat')[0])
for think in (True, False):
    body = {'model': 'coach-kira', 'stream': True, 'think': think, 'keep_alive': '30m',
            'options': {'temperature': 0.3, 'num_ctx': 16384},
            'messages': [{'role': 'system', 'content': SYSTEM}, {'role': 'user', 'content': QUESTIONS[1]}]}
    req = urllib.request.Request('http://127.0.0.1:11434/api/chat', data=json.dumps(body).encode(), headers={'Content-Type': 'application/json'})
    t0 = time.time(); first = None; txt = ''; last = {}
    with urllib.request.urlopen(req, timeout=600) as r:
        for line in r:
            j = json.loads(line)
            c = j.get('message', {}).get('content', '')
            if c and first is None: first = time.time() - t0
            txt += c; last = j
    print(json.dumps({'think': think, 'first_content_s': round(first or -1, 1), 'total_s': round(time.time() - t0, 1),
                      'tok_s': round(last.get('eval_count', 0) / max(1e-9, last.get('eval_duration', 1) / 1e9), 1),
                      'answer': txt[:700]}, ensure_ascii=False), flush=True)
PY
touch $D/DONE_CREATE
