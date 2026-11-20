#!/bin/zsh
set -e
LOG=/tmp/ck/runner/install.log
exec > >(tee "$LOG") 2>&1
mkdir -p ~/coach-kira-runner/logs
cp /tmp/ck/runner/kira_runner.py ~/coach-kira-runner/
cp /tmp/ck/runner/de.coachkira.morning.plist ~/Library/LaunchAgents/
plutil -lint ~/Library/LaunchAgents/de.coachkira.morning.plist
launchctl bootout gui/$(id -u)/de.coachkira.morning 2>/dev/null || true
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/de.coachkira.morning.plist
echo "== check"
/opt/homebrew/bin/python3 ~/coach-kira-runner/kira_runner.py check
echo "== Testlauf KI-Analyse"
/opt/homebrew/bin/python3 ~/coach-kira-runner/kira_runner.py run status
echo "== launchd"
launchctl print gui/$(id -u)/de.coachkira.morning | grep -E "state|run interval|last exit" || true
echo FERTIG
