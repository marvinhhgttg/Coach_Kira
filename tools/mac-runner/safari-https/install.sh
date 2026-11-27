#!/bin/zsh
# Coach Kira – https-Weiterleitung für Safari einrichten (ohne sudo)
set -e
LOG=/tmp/ck/safari_install.log
exec > $LOG 2>&1
echo "start $(date)"
[ -x /opt/homebrew/bin/caddy ] || /opt/homebrew/bin/brew install caddy
cp /tmp/ck/saf/Caddyfile ~/coach-kira-runner/Caddyfile
/opt/homebrew/bin/caddy validate --config ~/coach-kira-runner/Caddyfile --adapter caddyfile
cp /tmp/ck/saf/de.coachkira.ollama-https.plist ~/Library/LaunchAgents/
launchctl unload ~/Library/LaunchAgents/de.coachkira.ollama-https.plist 2>/dev/null || true
launchctl load ~/Library/LaunchAgents/de.coachkira.ollama-https.plist
sleep 4
ls -la ~/Library/Application\ Support/Caddy/pki/authorities/local/ || true
curl -sk -o /dev/null -w "https_status=%{http_code}\n" https://localhost:11435/api/tags || true
curl -sk -o /dev/null -w "cors_status=%{http_code}\n" -H "Origin: https://coachkira.pplx.app" https://localhost:11435/api/tags || true
curl -s -I -H "Origin: https://coachkira.pplx.app" http://localhost:11434/api/tags | grep -i -E "^HTTP|access-control" || true
echo "fertig"
