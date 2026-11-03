#!/bin/zsh
# Coach Kira – Analysen-Update (aerobe Entkopplung): kira_garmin.py einspielen und Runden langer Einheiten nachladen.
set -eu
DST="$HOME/coach-kira-runner"
SRC="${0:A:h}"
cp "$DST/kira_garmin.py" "$DST/kira_garmin.vor-entkopplung.py"
cp "$SRC/kira_garmin.py" "$DST/"
echo "Lade Runden der langen Einheiten nach (einige Minuten) …"
/opt/homebrew/bin/python3 "$DST/kira_garmin.py" entkopplung --von 2025-08-26
echo "Fertig. Neue lange Einheiten bekommen die Entkopplung automatisch beim normalen Garmin-Abruf."
