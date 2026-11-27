#!/bin/zsh
# Coach Kira – native Mac-App bauen und installieren (27.11.2026)
# Die Datei CoachKira.app aus build/CoachKira.xcarchive ist die iPhone-App (Plattform iPhoneOS) –
# die lässt sich im Finder nicht öffnen. Die Mac-App ist das Ziel „CoachKiraMac“ (Name „Coach Kira.app“).
# Aufruf:  zsh ~/coach-kira-ios/mac-app.sh [TEAM-ID]   (Standard 8Z2DKM5LZ9)
set -euo pipefail
TEAM="${1:-8Z2DKM5LZ9}"
cd ~/coach-kira-ios
xcodegen
if xcodebuild -project CoachKira.xcodeproj -list | grep -q '^ *CoachKiraMac$'; then ZIEL=(-scheme CoachKiraMac); else ZIEL=(-target CoachKiraMac); fi
echo "Baue Mac-App (Team $TEAM) …"
xcodebuild -project CoachKira.xcodeproj "${ZIEL[@]}" -configuration Release \
  -destination 'platform=macOS' -derivedDataPath build/mac \
  DEVELOPMENT_TEAM="$TEAM" CODE_SIGN_STYLE=Automatic -allowProvisioningUpdates build | tail -15
APP=$(find build/mac/Build/Products/Release -maxdepth 1 -name '*.app' | head -1)
[[ -n "$APP" ]] || { echo "Keine Mac-App gebaut – Ausgabe oben prüfen."; exit 1; }
codesign --verify --deep --strict "$APP" && echo "Signatur ok: $APP"
rm -rf "/Applications/Coach Kira.app"
ditto "$APP" "/Applications/Coach Kira.app"
xattr -dr com.apple.quarantine "/Applications/Coach Kira.app" 2>/dev/null || true
open "/Applications/Coach Kira.app"
echo "Fertig: /Applications/Coach Kira.app"
