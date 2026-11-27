# Coach Kira auf der Garmin-Uhr (fēnix 8 Pro)

Glance und App mit dem Tagesvotum aus dem Cockpit (`/api/widget?uhr=1`). Hintergrund alle 15 min, beim Öffnen frisch.

## Bauen (Mac)
```
cd ~/coach-kira-garmin
SDK="$(cat ~/Library/Application\ Support/Garmin/ConnectIQ/current-sdk.cfg 2>/dev/null)"
"$SDK/bin/monkeyc" -d fenix8pro47mm -f monkey.jungle -o bin/CoachKira.prg -y developer_key.der -l 2
```

## Simulator
```
"$SDK/bin/connectiq" &
"$SDK/bin/monkeydo" bin/CoachKira.prg fenix8pro47mm
```

## Auf die Uhr
Uhr per USB, mit OpenMTP `bin/CoachKira.prg` nach `GARMIN/APPS/` kopieren, trennen.

## Koppeln
App auf der Uhr öffnen → START → 6-stelliger Code. Im Cockpit unter Technik › „iPhone-App, Widgets & Garmin-Uhr“ eingeben → „Uhr koppeln“.
MENU (lange UP) in der App: Kopplung lösen. Widerrufen auch im Cockpit.
