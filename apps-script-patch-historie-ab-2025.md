# Apps-Script-Patch: Historie ab 26.08.2025

Datei: `KiraGeminiSupervisor.gs` · Funktion `ckDash_parseDays_` (nur von `chartData` genutzt)

## Was und warum

`chartData` liefert höchstens 365 Tage. Deine Timeline reicht aber bis zum 26.08.2025 zurück, das sind heute rund 430 Tage. Dadurch fehlen im Cockpit überall, wo „Alles“ gewählt ist, die ersten gut zwei Monate (26.08.–30.10.2025). Das betrifft unter anderem die neue Auswertung „Votum befolgt“, die Langzeit-Diagramme in Analyse sowie Jahresrückblick und Zeitraumvergleich.

Die Änderung hebt nur die Obergrenze an. Kürzere Zeitfenster (7, 28, 90 Tage …) bleiben genau wie bisher. Bei „Alles“ kommen rund 65 Zeilen mehr; die Antwortzeit steigt dadurch kaum.

Ohne KI, keine neuen Schlüssel, keine Trigger-Änderung.

## Streiche

```javascript
  if (m) {
    return Math.max(1, Math.min(365, Number(m[1])));
  }
```

## Ersetze

```javascript
  if (m) {
    // Historie ab 26.08.2025: Obergrenze 365 → 1000 Tage (nur für „Alles“ relevant)
    return Math.max(1, Math.min(1000, Number(m[1])));
  }
```

## Danach

Bereitstellen → Bereitstellungen verwalten → bestehende Web-App-Bereitstellung bearbeiten → Version „Neue Version“ → Bereitstellen (gleiche URL).

Prüfen: Im Cockpit unter Analyse → „Votum befolgt“ sollte der Zeitraum dann am 26.08.25 (bzw. am ersten Tag mit Morgenwerten) beginnen statt am 31.10.25.
