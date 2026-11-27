#!/bin/zsh
# Coach Kira – iPhone-App nach TestFlight hochladen (24.11.2026)
# Voraussetzung: Xcode → Einstellungen → Accounts: Apple-ID mit Entwicklerkonto angemeldet,
# App in App Store Connect angelegt (Bundle-ID de.reichpietsch.coachkira).
# Aufruf:  zsh ~/coach-kira-ios/testflight.sh [TEAM-ID]
set -euo pipefail
TEAM="${1:-}"
if [[ ! "$TEAM" =~ ^[A-Z0-9]{10}$ || "$TEAM" == "AB12CD34EF" ]]; then
  echo "Bitte die echte Team-ID angeben (10 Zeichen, developer.apple.com → Membership), also: zsh ~/coach-kira-ios/testflight.sh <DEINE-TEAM-ID>"
  echo "Hinweis: TestFlight geht nur mit bezahlter Mitgliedschaft im Apple Developer Program."
  exit 1
fi
cd ~/coach-kira-ios
BUILD=$(date +%Y%m%d%H%M)
mkdir -p build
# Build-Nummer steht fest in den Info.plist (App + Widgets) → vor jedem Upload hochzählen, sonst lehnt Apple doppelte Nummern ab
for pl in App/Info.plist Widgets/Info.plist; do /usr/libexec/PlistBuddy -c "Set :CFBundleVersion $BUILD" "$pl"; done
echo "Team $TEAM · Build $BUILD"
xcodebuild -project CoachKira.xcodeproj -scheme CoachKira -configuration Release \
  -destination 'generic/platform=iOS' -archivePath build/CoachKira.xcarchive \
  DEVELOPMENT_TEAM="$TEAM" CURRENT_PROJECT_VERSION="$BUILD" CODE_SIGN_STYLE=Automatic \
  -allowProvisioningUpdates archive | tail -20
cat > build/ExportOptions.plist <<PL
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>method</key><string>app-store-connect</string>
  <key>destination</key><string>upload</string>
  <key>teamID</key><string>$TEAM</string>
  <key>signingStyle</key><string>automatic</string>
  <key>uploadSymbols</key><true/>
  <key>testFlightInternalTestingOnly</key><true/>
</dict></plist>
PL
xcodebuild -exportArchive -archivePath build/CoachKira.xcarchive \
  -exportOptionsPlist build/ExportOptions.plist -exportPath build/export -allowProvisioningUpdates | tail -20
echo "Fertig: Build $BUILD ist hochgeladen. In App Store Connect → TestFlight erscheint er nach 5–15 min."
