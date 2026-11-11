#!/bin/bash
# Coach Kira: Ollama-Freigabe + Modelltest (lokal auf dem Mac)
set -u
D=/tmp/ck/bench
LOG=$D/run.log
exec > >(tee -a "$LOG") 2>&1
echo "== Start $(date)"
ORIGIN="https://coachkira.pplx.app"

# 1) Freigabe fuer das Cockpit
launchctl setenv OLLAMA_ORIGINS "$ORIGIN"
mkdir -p "$HOME/Library/LaunchAgents"
PL="$HOME/Library/LaunchAgents/de.coachkira.ollama-origins.plist"
cat > "$PL" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>de.coachkira.ollama-origins</string>
  <key>ProgramArguments</key><array>
    <string>/bin/launchctl</string><string>setenv</string><string>OLLAMA_ORIGINS</string><string>$ORIGIN</string>
  </array>
  <key>RunAtLoad</key><true/>
</dict></plist>
PLIST
launchctl unload "$PL" 2>/dev/null; launchctl load "$PL" && echo "LaunchAgent geladen: $PL"
echo "OLLAMA_ORIGINS=$(launchctl getenv OLLAMA_ORIGINS)"

# 2) Ollama neu starten, damit die Freigabe greift
osascript -e 'quit app "Ollama"' 2>/dev/null
sleep 3
pkill -f "ollama serve" 2>/dev/null
sleep 2
open -a Ollama
for i in $(seq 1 30); do curl -s -m 2 http://127.0.0.1:11434/api/version >/dev/null && break; sleep 1; done
if ! curl -s -m 2 http://127.0.0.1:11434/api/version >/dev/null; then
  echo "App startet keinen Server - starte ollama serve direkt"
  OLLAMA_ORIGINS="$ORIGIN" nohup ollama serve > $D/serve.log 2>&1 &
  sleep 5
fi
echo "Version: $(curl -s http://127.0.0.1:11434/api/version)"
echo "CORS-Test:"; curl -s -i -m 5 -H "Origin: $ORIGIN" http://127.0.0.1:11434/api/version | grep -i -E "^HTTP|access-control-allow-origin"
echo "Modelle:"; ollama list

# 3) Modelltest
/usr/bin/python3 $D/bench.py
echo "== Fertig $(date)"
touch $D/DONE
