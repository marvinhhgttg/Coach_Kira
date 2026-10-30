#!/bin/zsh
# Coach Kira – „An Garmin senden“ einrichten (einmalig). Legt eine eigene Python-Umgebung an,
# installiert garminconnect und speichert deinen Garmin-Login in ~/.garminconnect.
set -eu
DST="$HOME/coach-kira-runner"
SRC="${0:A:h}"
cp "$SRC/kira_garmin_workout.py" "$SRC/kira_runner.py" "$DST/"
/opt/homebrew/bin/python3 -m venv "$DST/.venv-garmin"
"$DST/.venv-garmin/bin/pip" install -q --upgrade pip
"$DST/.venv-garmin/bin/pip" install -q --upgrade "garminconnect>=0.3.20" curl_cffi
echo "Jetzt einmal bei Garmin anmelden (E-Mail, Passwort, ggf. Code aus Mail/App):"
"$DST/.venv-garmin/bin/python" "$DST/kira_garmin_workout.py" login
"$DST/.venv-garmin/bin/python" "$DST/kira_garmin_workout.py" test
echo "Fertig. Der Supervisor-Takt (60 s) holt Aufträge aus dem Cockpit ab."
