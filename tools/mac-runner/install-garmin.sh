#!/bin/zsh
set -eu

SRC="/tmp/ck/garmin"
DST="$HOME/coach-kira-runner"

mkdir -p "$DST/logs"
cp "$SRC/kira_garmin.py" "$DST/"

mkdir -p ~/.config/garmin
chmod 700 ~/.config/garmin

if [[ ! -f "$HOME/.config/garmin/connect-personal-api-key" ]]; then
    echo "HINWEIS: Garmin-Schlüssel nicht gefunden. Datei fehlt oder hat falschen Namen?"
    echo "Erwarter Pfad: ~/.config/garmin/connect-personal-api-key"
    exit 2
fi

chmod 600 ~/.config/garmin/connect-personal-api-key

cp "$SRC/de.coachkira.garmin.plist" ~/Library/LaunchAgents/
launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/de.coachkira.garmin.plist 2>/dev/null || true
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/de.coachkira.garmin.plist 2>&1 | tee ~/coach-kira-runner/logs/garmin.install.log

/opt/homebrew/bin/python3 "$DST/kira_garmin.py" test 2>&1 | tee -a ~/coach-kira-runner/logs/garmin.install.log
