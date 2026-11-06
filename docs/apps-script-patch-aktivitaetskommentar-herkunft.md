# Apps-Script-Patch: Herkunft beim Aktivitätskommentar

## Warum

Bisher speichert `ckDash_saveActivityReview_` jeden Aktivitätskommentar fest mit der Herkunft `local`. Schreibt künftig der Cloud-Ersatz einen Kommentar, weil der Mac 3 × 5 Minuten nicht geliefert hat, stünde im Log fälschlich „lokal“. Nach dem Patch übernimmt das Script die mitgeschickte Herkunft. Der Mac schickt nichts mit, also bleibt es dort bei `local`. Nur der Cockpit-Server schickt `cloud`.

Briefings (`saveBriefing`) und Extras (`saveExtra`) können das schon, dort ist nichts zu tun.

## Datei

`KiraGeminiSupervisor.gs`, Funktion `ckDash_saveActivityReview_`

## Änderung 1

Streiche:

```js
    const model = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
```

Ersetze durch:

```js
    const model = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
    const engine = String(b.engine || 'local') === 'cloud' ? 'cloud' : 'local';
```

(Die Zeile kommt in der Datei mehrfach vor. Gemeint ist nur die in `ckDash_saveActivityReview_`, direkt nach `if (!text) throw new Error('Kein Kommentar-Text übermittelt.');`.)

## Änderung 2

Streiche:

```js
      JSON.stringify({ engine: 'local', model: model, savedAt: ckDash_now_() })
```

Ersetze durch:

```js
      JSON.stringify({ engine: engine, model: model, savedAt: ckDash_now_() })
```

## Danach

Speichern, dann **Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Neue Version → Bereitstellen**. Die Web-App-URL bleibt gleich.
