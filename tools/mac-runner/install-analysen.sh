#!/bin/zsh
# Coach Kira – Analysen-Update: Runden (aerobe Entkopplung) und Garmin-Wetter über den Garmin-Login nachladen.
set -eu
DST="$HOME/coach-kira-runner"
SRC="${0:A:h}"
[ -f "$DST/kira_garmin.vor-entkopplung.py" ] || cp "$DST/kira_garmin.py" "$DST/kira_garmin.vor-entkopplung.py"
cp "$SRC/kira_garmin.py" "$SRC/kira_garmin_runden.py" "$DST/"
if [ ! -x "$DST/.venv-garmin/bin/python" ]; then
  echo "Garmin-Login fehlt – erst: zsh /tmp/ck-garmin-workout/install-garmin-workout.sh"; exit 1
fi
"$DST/.venv-garmin/bin/python" "$DST/kira_garmin_workout.py" test
echo "Lade Runden und Wetter nach (etwa 10 Minuten) …"
/opt/homebrew/bin/python3 "$DST/kira_garmin.py" entkopplung --von 2025-08-26
echo "Fertig. Neue Einheiten bekommen Entkopplung und Wetter automatisch beim normalen Garmin-Abruf."
