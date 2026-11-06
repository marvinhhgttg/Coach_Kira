#!/bin/zsh
# Coach Kira: OLLAMA_ORIGINS setzen und Ollama neu starten, falls es ohne Freigabe läuft.
ORIGIN="https://coachkira.pplx.app"
LOG=~/coach-kira-runner/logs/ollama-origins.log
mkdir -p ~/coach-kira-runner/logs
/bin/launchctl setenv OLLAMA_ORIGINS "$ORIGIN"
check() { /usr/bin/curl -s -m 5 -o /dev/null -w "%{http_code}" -H "Origin: $ORIGIN" http://127.0.0.1:11434/api/tags; }
# bis zu 2 Minuten auf Ollama warten (nach dem Login)
for i in {1..24}; do c=$(check); [[ "$c" != "000" ]] && break; sleep 5; done
if [[ "$c" == "403" ]]; then
  echo "$(date '+%F %T') Ollama ohne Freigabe (403) – starte neu" >> $LOG
  /usr/bin/osascript -e 'quit app "Ollama"' 2>/dev/null
  sleep 3
  /usr/bin/pkill -x Ollama 2>/dev/null; sleep 2
  /usr/bin/pkill -9 -x Ollama 2>/dev/null
  /usr/bin/pkill -f "ollama serve" 2>/dev/null; sleep 2
  /usr/bin/open -a Ollama
  for i in {1..24}; do sleep 5; c=$(check); [[ "$c" == "200" ]] && break; done
fi
echo "$(date '+%F %T') Status $c" >> $LOG
