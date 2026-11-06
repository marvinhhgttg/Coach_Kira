import json, time, urllib.request
D = '/tmp/ck/bench'
ctx = json.load(open(f'{D}/context.json'))
MODELS = ['gemma4:e4b-mlx', 'gemma4:31b-it-qat', 'gpt-oss:120b']
QUESTIONS = [
  'Warum lautet der Tagesentscheid so?',
  'Kann ich morgen hart trainieren?',
  'Wie sieht meine Belastung diese Woche aus und was empfiehlst du für die nächsten Tage?',
]
SYSTEM = '\n'.join([
  'Du bist Kira, Sportwissenschaftlerin und persönliche Ausdauer-Trainerin von Marc (Jahrgang 1975; Rennrad, Gravel, Laufen, Bergsport).',
  'Sprich Marc per Du an, antworte auf Deutsch, klar und konkret, ohne Emojis.',
  'Stütze dich zuerst auf die mitgelieferten Cockpit-Daten (JSON unten). Nenne die Zahlen, auf die du dich beziehst.',
  'Erfinde keine Werte. Wenn Daten fehlen oder unklar sind, sag das offen.',
  'Halte dich an die Regeln aus dem Feld "regeln" (Tagesentscheid, Zielbereiche).',
  'Gib am Ende eine klare Empfehlung (z. B. Einheit, Dauer, Zone oder Pause).',
  'Bei Schmerzen, Krankheit oder medizinischen Fragen: auf Arzt bzw. Physiotherapie verweisen.',
  '', 'COCKPIT-DATEN:', json.dumps(ctx, ensure_ascii=False)])

def chat(model, q, keep):
    body = {'model': model, 'stream': False, 'keep_alive': keep,
            'options': {'temperature': 0.3, 'num_ctx': 16384},
            'messages': [{'role': 'system', 'content': SYSTEM}, {'role': 'user', 'content': q}]}
    req = urllib.request.Request('http://127.0.0.1:11434/api/chat', data=json.dumps(body).encode(), headers={'Content-Type': 'application/json'})
    t0 = time.time()
    with urllib.request.urlopen(req, timeout=900) as r:
        j = json.loads(r.read())
    wall = time.time() - t0
    ns = 1e9
    return {
        'model': model, 'question': q, 'wall_s': round(wall, 1),
        'load_s': round(j.get('load_duration', 0) / ns, 1),
        'prompt_tokens': j.get('prompt_eval_count'), 'prompt_s': round(j.get('prompt_eval_duration', 0) / ns, 1),
        'gen_tokens': j.get('eval_count'), 'gen_s': round(j.get('eval_duration', 0) / ns, 1),
        'tok_per_s': round(j.get('eval_count', 0) / max(1e-9, j.get('eval_duration', 1) / ns), 1),
        'answer': j.get('message', {}).get('content', ''),
        'thinking_chars': len(j.get('message', {}).get('thinking', '') or ''),
    }

res = []
for m in MODELS:
    for k, q in enumerate(QUESTIONS):
        try:
            r = chat(m, q, '0s' if k == len(QUESTIONS) - 1 else '10m')
        except Exception as e:
            r = {'model': m, 'question': q, 'error': str(e)}
        res.append(r)
        print(m, '|', q[:30], '|', {k2: r.get(k2) for k2 in ('wall_s', 'load_s', 'prompt_s', 'gen_tokens', 'tok_per_s', 'error')}, flush=True)
        json.dump(res, open(f'{D}/results.json', 'w'), ensure_ascii=False, indent=1)
