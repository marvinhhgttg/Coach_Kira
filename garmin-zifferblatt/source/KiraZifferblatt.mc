// Coach Kira Zifferblatt im fēnix-Stil (22.11.2026, Version 2)
// Uhrzeit (Stunden weiß, Minuten orange), Datum, Akku als Bogen oben, links Trainingsbereitschaft
// als Ring, rechts heutige Summenlast gegen den Soll-Rahmen als Ring, dazu Ausdauerwert, Wetter,
// Wind und Plan. Kira-Werte kommen aus der Coach-Kira-App (private Komplikation „KIRA DATEN“),
// Akku, Wetter und Wind direkt von der Uhr.
import Toybox.Application;
import Toybox.Complications;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;
import Toybox.System;
import Toybox.Time;
import Toybox.Time.Gregorian;
import Toybox.Weather;
import Toybox.WatchUi;

class KiraZifferblattApp extends Application.AppBase {
    var ansicht as KiraZifferblatt? = null;
    function initialize() {
        AppBase.initialize();
    }

    function getInitialView() as [WatchUi.Views] or [WatchUi.Views, WatchUi.InputDelegates] {
        ansicht = new KiraZifferblatt();
        // 25.11.2026: langer Druck öffnet die Coach-Kira-App (über deren Komplikation)
        return [ansicht, new KiraFaceDelegate(ansicht)];
    }

    // 24.11.2026: Einstellungen aus der Connect IQ App sofort übernehmen
    function onSettingsChanged() as Void {
        if (ansicht != null) { (ansicht as KiraZifferblatt).einstLesen(); }
        WatchUi.requestUpdate();
    }
}

const ORANGE = 0xFF9A1F;
const GRAU = 0x8A8A8A;
const SPUR = 0x2A2A2A;
const VOTE_FARBE = {
    "PUSH" => 0x2ECC71, "TRAIN" => 0x2ECC71, "QUAL" => 0x2ECC71, "EASY" => 0xF1C40F, "REST" => 0xE74C3C
};

class KiraZifferblatt extends WatchUi.WatchFace {
    var kiraId as Complications.Id? = null;
    var kira as String? = null;
    var schlaf as Boolean = false;
    var gesucht as Number = -1;
    // Kleinere Schriften als die Systemgrößen: skalierbare Garmin-Schriften (Roboto Condensed)
    var fKopf as Graphics.FontType = Graphics.FONT_XTINY;
    var fLabel as Graphics.FontType = Graphics.FONT_XTINY;
    var fWert as Graphics.FontType = Graphics.FONT_TINY;
    var fRing as Graphics.FontType = Graphics.FONT_SMALL;
    var fDatum as Graphics.FontType = Graphics.FONT_TINY;
    // 24.11.2026: Einstellungen
    var eEcken as Array<Number> = [0, 1, 2, 3];
    var eRingL as Number = 0;
    var eRingR as Number = 10;
    var eVotum as Boolean = true;
    var eSek as Boolean = true;
    var eErhBB as Boolean = true;

    function initialize() {
        WatchFace.initialize();
        fKopf = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 25, Graphics.FONT_XTINY);
        fLabel = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 23, Graphics.FONT_XTINY);
        fWert = schrift(["RobotoCondensedRegular", "RobotoRegular"], 31, Graphics.FONT_TINY);
        fRing = schrift(["RobotoCondensedBold", "RobotoMedium"], 33, Graphics.FONT_SMALL);
        fDatum = schrift(["RobotoCondensedRegular", "RobotoRegular"], 28, Graphics.FONT_TINY);
        try { Complications.registerComplicationChangeCallback(method(:geaendert)); } catch (e) { }
        kiraSuchen();
        einstLesen();
    }

    function einstLesen() as Void {
        eEcken = [KiraWert.zahlE("pOL", 0), KiraWert.zahlE("pOR", 1), KiraWert.zahlE("pUL", 2), KiraWert.zahlE("pUR", 3)];
        eRingL = KiraWert.zahlE("pRingL", 0);
        eRingR = KiraWert.zahlE("pRingR", 10);
        eVotum = KiraWert.an("pVotumRing");
        eSek = KiraWert.an("pSek");
        eErhBB = KiraWert.an("pErhBB");
        KiraSkala.stil = KiraWert.zahlE("pStil", 0);
    }

    function schrift(namen as Array<String>, px as Number, ersatz as Graphics.FontType) as Graphics.FontType {
        try {
            if (Graphics has :getVectorFont) {
                var f = Graphics.getVectorFont({ :face => namen, :size => px });
                if (f != null) { return f; }
            }
        } catch (e) { }
        return ersatz;
    }

    function kiraSuchen() as Void {
        try {
            var it = Complications.getComplications();
            var c = it.next();
            while (c != null) {
                var ll = c.longLabel;
                var vv = c.value;
                var passt = ll != null && ll.toString().toUpper().find("KIRA DATEN") != null;
                if (!passt && vv instanceof String && vv.find("|") != null && c.getType() == Complications.COMPLICATION_TYPE_INVALID) { passt = true; }
                if (passt) {
                    kiraId = c.complicationId;
                    if (kiraId != null) { Complications.subscribeToUpdates(kiraId); }
                    if (c.value != null) { kira = c.value.toString(); }
                    return;
                }
                c = it.next();
            }
        } catch (e) { }
    }

    function geaendert(id as Complications.Id) as Void {
        try {
            if (kiraId != null && id.equals(kiraId)) {
                var v = Complications.getComplication(id).value;
                if (v != null) { kira = v.toString(); }
            }
        } catch (e) { }
    }

    function onExitSleep() as Void {
        schlaf = false;
        if (kiraId == null) { kiraSuchen(); }
        WatchUi.requestUpdate();
    }

    function onEnterSleep() as Void {
        schlaf = true;
        WatchUi.requestUpdate();
    }

    // 0 vote|1 stufe|2 lo|3 hi|4 ist|5 datum|6 ziel|7 zielTage|8 readiness|9 plan|10 sport|11 zone|12 es|13 es28|14 readinessZeit
    function teile() as Array<String> {
        var r = [] as Array<String>;
        var s = kira;
        if (s == null) { return r; }
        var i = s.find("|");
        while (i != null) {
            r.add(s.substring(0, i) as String);
            s = s.substring(i + 1, s.length()) as String;
            i = s.find("|");
        }
        r.add(s);
        return r;
    }

    function feld(t as Array<String>, i as Number) as String {
        return t.size() > i ? t[i] : "";
    }

    function zahl(t as Array<String>, i as Number) as Number? {
        var s = feld(t, i);
        return s.length() == 0 ? null : s.toNumber();
    }

    function heuteIso() as String {
        var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        return g.year.format("%04d") + "-" + (g.month as Number).format("%02d") + "-" + g.day.format("%02d");
    }

    // Ring wie bei Garmin: dunkle Spur, farbiger Anteil im Uhrzeigersinn ab oben
    function ring(dc as Graphics.Dc, x as Number, y as Number, r as Number, anteil as Float, farbe as Number) as Void {
        dc.setPenWidth(8);
        dc.setColor(SPUR, Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(x, y, r);
        if (anteil > 0.01) {
            var a = anteil > 1.0 ? 1.0 : anteil;
            dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
            var ende = 90 - (a * 360).toNumber();
            if (a >= 0.999) { dc.drawCircle(x, y, r); }
            else { dc.drawArc(x, y, r, Graphics.ARC_CLOCKWISE, 90, ende); }
        }
        dc.setPenWidth(1);
    }

    // Kleiner Block: Bezeichnung (orange, klein) über Wert (weiß)
    function block(dc as Graphics.Dc, x as Number, y as Number, label as String, wert as String, wertFarbe as Number) as Void {
        dc.setColor(KiraSkala.lab(), Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y, fLabel, label, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(wertFarbe, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y + dc.getFontHeight(fLabel) - 2, fWert, wert, Graphics.TEXT_JUSTIFY_CENTER);
    }

    function richtung(grad as Number?) as String {
        if (grad == null) { return ""; }
        var r = ["N", "NO", "O", "SO", "S", "SW", "W", "NW"];
        return r[((grad + 22) % 360) / 45];
    }

    // 23.11.2026: Trainingsbereitschaft im Garmin-Muster – 270°-Bogen (7:30 bis 4:30 Uhr), Segmente
    // Schlecht <25 rot · Niedrig 25–49 orange · Mäßig 50–74 grün · Hoch 75–94 blau · Optimal 95+ lila, Punkt am Wert
    function bereitFarbe(v as Number) as Number {
        if (v >= 95) { return 0x6E45E9; }
        if (v >= 75) { return 0x3A78CF; }
        if (v >= 50) { return 0x5DB75B; }
        if (v >= 25) { return 0xE5822F; }
        return 0xD9433A;
    }

    function bogenWinkel(v as Number) as Number {
        var a = 225 - (v * 27) / 10;
        return ((a % 360) + 360) % 360;
    }

    function bereitBogen(dc as Graphics.Dc, x as Number, y as Number, r as Number, pen as Number, v as Number?) as Void {
        KiraSkala.bogen(dc, x, y, r, pen, v, KiraSkala.tr());
    }


    // 22.11.2026: Votum-Ring am Rand. Grundring = Votumsfarbe gedimmt (voller Kreis = Obergrenze des
    // Soll-Rahmens), Füllung = heutige Summenlast in Votumsfarbe, weißer Strich = Untergrenze,
    // über der Obergrenze läuft eine zweite, gleich breite Runde in Orange innen daneben. Kein Ring im Always-on.
    function stufeFarbe(stufe as String, vote as String) as Number {
        if (stufe.equals("blau")) { return 0x3FA0F0; }
        if (stufe.equals("gruen")) { return 0x2ECC71; }
        if (stufe.equals("gelb")) { return 0xF1C40F; }
        if (stufe.equals("orange")) { return 0xEF7D00; }
        if (stufe.equals("rot")) { return 0xE74C3C; }
        return VOTE_FARBE.hasKey(vote) ? VOTE_FARBE.get(vote) as Number : GRAU;
    }

    function dimmen(f as Number) as Number {
        var r = ((f >> 16) & 0xFF) * 42 / 100;
        var g = ((f >> 8) & 0xFF) * 42 / 100;
        var b = (f & 0xFF) * 42 / 100;
        return (r << 16) | (g << 8) | b;
    }

    function votumRing(dc as Graphics.Dc, cx as Number, cy as Number, stufe as String, vote as String, lo as Number?, hi as Number?, ist as Number?, heute as Boolean) as Void {
        var r = cx - 5;
        dc.setPenWidth(6);
        if (!heute || (stufe.length() == 0 && vote.length() == 0)) {
            dc.setColor(0x303030, Graphics.COLOR_TRANSPARENT);
            dc.drawCircle(cx, cy, r);
            dc.setPenWidth(1);
            return;
        }
        var f = stufeFarbe(stufe, vote);
        dc.setColor(KiraSkala.votSpur(f), Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(cx, cy, r);
        var a = (hi != null && hi > 0 && ist != null && ist > 0) ? ist.toFloat() / hi : 0.0;
        if (a > 0.005) {
            dc.setColor(KiraSkala.votFill(f), Graphics.COLOR_TRANSPARENT);
            if (a >= 0.999) { dc.drawCircle(cx, cy, r); }
            else { dc.drawArc(cx, cy, r, Graphics.ARC_CLOCKWISE, 90, 90 - (a * 360).toNumber()); }
        }
        if (a > 1.005) {
            var b = a - 1.0;
            // zweite Runde gleich breit, direkt innen daneben (1 px Abstand)
            dc.setPenWidth(6);
            dc.setColor(0xFF6A00, Graphics.COLOR_TRANSPARENT);
            if (b >= 0.999) { dc.drawCircle(cx, cy, r - 7); }
            else { dc.drawArc(cx, cy, r - 7, Graphics.ARC_CLOCKWISE, 90, 90 - (b * 360).toNumber()); }
        }
        if (lo != null && hi != null && hi > 0 && lo > 0 && lo < hi) {
            var w = Math.PI * 2 * lo / hi;
            var r1 = r + 3;
            var r2 = r - 11;
            dc.setPenWidth(3);
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawLine(cx + (r1 * Math.sin(w)).toNumber(), cy - (r1 * Math.cos(w)).toNumber(), cx + (r2 * Math.sin(w)).toNumber(), cy - (r2 * Math.cos(w)).toNumber());
        }
        dc.setPenWidth(1);
    }


    // 23.11.2026: Live-Werte direkt von der Uhr (Garmin-Komplikationen): Erholungszeit (Minuten), Body Battery
    function nativ(typ as Complications.Type) as Number? {
        try {
            var c = Complications.getComplication(new Complications.Id(typ));
            if (c != null && c.value != null) { return c.value.toNumber(); }
        } catch (e) { }
        return null;
    }

    function nativZeile() as Array<String> {
        var erh = nativ(Complications.COMPLICATION_TYPE_RECOVERY_TIME);
        var bb = nativ(Complications.COMPLICATION_TYPE_BODY_BATTERY);
        var e = erholtAb(erh, true);
        return [e.equals("ERHOLT") || e.equals("-") ? (e.equals("-") ? "ERH -" : e) : "ERH " + e, "BB " + (bb == null ? "-" : bb.toString())];
    }


    // 23.11.2026: Wochen-Ring (dünn, außen um den Tagessummen-Ring): Wochen-Ist gegen Kiras Wochenbudget
    function wochenRing(dc as Graphics.Dc, x as Number, y as Number, r as Number, ist as Number?, bud as Number?) as Void {
        if (bud == null || bud <= 0) { return; }
        // 23.11.2026: Garmin-Muster mit Kiras Wochenbudget (85 / 100 / 115 %)
        KiraSkala.bogen(dc, x, y, r, 3, ist == null ? 0 : ist, KiraSkala.woche(bud));
    }

    // Garmins Erholungszeit (Minuten, live von der Uhr) als Zeitpunkt: heute „17:20“, sonst „MI 07:10“; kurz = ohne Minuten
    function erholtAb(min as Number?, kurz as Boolean) as String {
        if (min == null) { return "-"; }
        if (min <= 0) { return "ERHOLT"; }
        var jetzt = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var i = Gregorian.info(Time.now().add(new Time.Duration(min * 60)), Time.FORMAT_SHORT);
        var wt = ["", "SO", "MO", "DI", "MI", "DO", "FR", "SA"];
        var hm = i.hour.format("%02d") + ":" + i.min.format("%02d");
        if (i.day == jetzt.day) { return hm; }
        return wt[i.day_of_week as Number] + " " + (kurz ? i.hour.format("%02d") : hm);
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        var w = dc.getWidth();
        var h = dc.getHeight();
        var cx = w / 2;
        var cy = h / 2;
        dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_BLACK);
        dc.clear();
        var uhr = System.getClockTime();
        var std = uhr.hour.format("%02d");
        var min = uhr.min.format("%02d");
        // noch nicht gefunden: einmal pro Minute erneut suchen
        if (kiraId == null && gesucht != uhr.min) { gesucht = uhr.min; kiraSuchen(); }
        var t = teile();
        var heute = feld(t, 5).equals(heuteIso());
        var vote = feld(t, 0);

        if (schlaf) {
            // Always-on: gedimmt, wenige Pixel, wandert jede Minute
            var dx = (uhr.min % 5) * 4 - 8;
            var dy = (uhr.min % 3) * 4 - 4;
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx + dx, cy + dy - 60, Graphics.FONT_NUMBER_MEDIUM, std, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            dc.setColor(0xA05E10, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx + dx, cy + dy + 60, Graphics.FONT_NUMBER_MEDIUM, min, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            return;
        }

        if (eVotum) { votumRing(dc, cx, cy, feld(t, 1), vote, heute ? zahl(t, 2) : null, heute ? zahl(t, 3) : null, heute ? zahl(t, 4) : null, heute); }

        // Akku: Strichbogen oben + Prozent / Tage
        var st = System.getSystemStats();
        var akku = st.battery;
        var n = 30;
        for (var i = 0; i < n; i++) {
            var a = Math.toRadians(135 - i * 90.0 / (n - 1));
            var r1 = cx - 19;
            var r2 = cx - 33;
            dc.setColor(i < (akku / 100.0 * n) ? KiraSkala.akku() : SPUR, Graphics.COLOR_TRANSPARENT);
            dc.setPenWidth(4);
            dc.drawLine(cx + (r1 * Math.cos(a)).toNumber(), cy - (r1 * Math.sin(a)).toNumber(),
                        cx + (r2 * Math.cos(a)).toNumber(), cy - (r2 * Math.sin(a)).toNumber());
        }
        dc.setPenWidth(1);
        var tage = (st has :batteryInDays) && st.batteryInDays != null ? st.batteryInDays.toNumber() : null;
        dc.setColor(KiraSkala.lab(), Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.085).toNumber(), fKopf, akku.toNumber().toString() + "%" + (tage != null ? " · " + tage.toString() + "T" : ""), Graphics.TEXT_JUSTIFY_CENTER);

        // Kira-Votum statt Markenname
        var vf = heute && VOTE_FARBE.hasKey(vote) ? VOTE_FARBE.get(vote) as Number : GRAU;
        dc.setColor(KiraSkala.votText(vf), Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.145).toNumber(), fKopf, kira == null ? (kiraId == null ? "KIRA · APP ÖFFNEN" : "KIRA · WARTET") : "KIRA · " + (vote.length() > 0 ? vote : "-") + (heute && feld(t, 17).equals("1") ? "*" : "") + (heute ? "" : " (ALT)"), Graphics.TEXT_JUSTIFY_CENTER);

        // Uhrzeit: Stunden weiß, Minuten orange (wie fēnix)
        dc.setColor(0xE8E8E8, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.35).toNumber(), Graphics.FONT_NUMBER_HOT, std, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(KiraSkala.minuten(), Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.57).toNumber(), Graphics.FONT_NUMBER_HOT, min, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        // 24.11.2026: Sekunden (nur wach) mittig unter den Minuten
        if (eSek) {
            dc.setColor(KiraSkala.sek(), Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, (h * 0.695).toNumber(), fLabel, uhr.sec.format("%02d"), Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Links: Trainingsbereitschaft (Morgenwert aus Coach Kira)
        var rx = (w * 0.17).toNumber();
        var rr = (w * 0.085).toNumber();
        var ry = (h * 0.49).toNumber();
        // 24.11.2026: Ring links frei belegbar (Standard Trainingsbereitschaft)
        var rgL = KiraWert.ring(eRingL, t, heute);
        KiraSkala.bogen(dc, rx, ry, rr, 7, rgL[0] as Number?, rgL[1] as Array);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, ry, fRing, KiraWert.ringText(eRingL, rgL[0] as Number?), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(KiraSkala.lab(), Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, ry + rr + 8, fLabel, rgL[2] as String, Graphics.TEXT_JUSTIFY_CENTER);

        // Rechts: heutige Summenlast gegen Soll
        var lx = w - rx;
        var lo = heute ? zahl(t, 2) : null;
        var hi = heute ? zahl(t, 3) : null;
        var ist = heute ? zahl(t, 4) : null;
        // 23.11.2026: Tageslast als Garmin-Bogen mit Kiras Soll-Rahmen (unter Soll blau · Rahmen grün · drüber orange · > +15 % rot)
        if (eRingR == 10) {
            KiraSkala.bogen(dc, lx, ry, rr, 7, ist, (hi != null && hi > 0) ? KiraSkala.last(lo, hi) : [[0, 100], [KiraSkala.GRAU]]);
            wochenRing(dc, lx, ry, rr + 8, zahl(t, 15), zahl(t, 16));
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(lx, ry, fRing, ist == null ? "-" : ist.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            dc.setColor(KiraSkala.lab(), Graphics.COLOR_TRANSPARENT);
            dc.drawText(lx, ry + rr + (zahl(t, 16) != null ? 14 : 8), fLabel, (hi == null ? "-" : (lo != null ? lo.toString() + "-" : "") + hi.toString()), Graphics.TEXT_JUSTIFY_CENTER);
        } else {
            // 24.11.2026: Ring rechts frei belegbar
            var rgR = KiraWert.ring(eRingR, t, heute);
            KiraSkala.bogen(dc, lx, ry, rr, 7, rgR[0] as Number?, rgR[1] as Array);
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(lx, ry, fRing, KiraWert.ringText(eRingR, rgR[0] as Number?), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            dc.setColor(KiraSkala.lab(), Graphics.COLOR_TRANSPARENT);
            dc.drawText(lx, ry + rr + 8, fLabel, rgR[2] as String, Graphics.TEXT_JUSTIFY_CENTER);
        }

        // 24.11.2026: vier Ecken frei belegbar (Standard: Temperatur, Ausdauerwert, Wind, Plan)
        var yo = (h * 0.23).toNumber();
        var yu = (h * 0.67).toNumber();
        var pos = [[(w * 0.19).toNumber(), yo], [(w * 0.81).toNumber(), yo], [(w * 0.27).toNumber(), yu], [(w * 0.73).toNumber(), yu]];
        for (var i = 0; i < 4; i++) {
            var ek = KiraWert.ecke(eEcken[i], t, heute);
            if ((ek[0] as String).length() == 0) { continue; }
            var lab = ek[0] as String;
            if (eEcken[i] == 3) {
                var sp = feld(t, 10);
                var pl = heute ? zahl(t, 9) : null;
                lab = "PLAN" + (sp.length() > 0 && heute ? " " + sp.toUpper() : "");
                ek = [lab, pl == null ? "-" : (pl == 0 ? "Ruhe" : pl.toString())];
                // 25.11.2026: nach dem Training „IST“ mit der Tagessumme
                var istE = heute ? zahl(t, 4) : null;
                if (istE != null && istE > 0) { ek = ["IST" + (sp.length() > 0 ? " " + sp.toUpper() : ""), istE.toString()]; }
            }
            block(dc, pos[i][0] as Number, pos[i][1] as Number, ek[0] as String, ek[1] as String, Graphics.COLOR_WHITE);
        }

        // Live von der Uhr: Erholungszeit · Body Battery
        if (eErhBB) {
            var nz = nativZeile();
            dc.setColor(0xC8C8C8, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, (h * 0.818).toNumber(), fLabel, nz[0] + " · " + nz[1], Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Datum unten
        var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var wt = ["", "SO", "MO", "DI", "MI", "DO", "FR", "SA"];
        var mo = ["", "JAN", "FEB", "MÄR", "APR", "MAI", "JUN", "JUL", "AUG", "SEP", "OKT", "NOV", "DEZ"];
        dc.setColor(0xD0D0D0, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.875).toNumber(), fDatum, wt[g.day_of_week as Number] + " " + g.day.toString() + ". " + mo[g.month as Number], Graphics.TEXT_JUSTIFY_CENTER);
    }
}

// 25.11.2026: Touch-and-hold auf dem Zifferblatt → Coach-Kira-App öffnen (Einstellung „Langer Druck öffnet die App“)
class KiraFaceDelegate extends WatchUi.WatchFaceDelegate {
    var ansicht as KiraZifferblatt;

    function initialize(v as KiraZifferblatt) {
        WatchFaceDelegate.initialize();
        ansicht = v;
    }

    function onPress(clickEvent as WatchUi.ClickEvent) as Boolean {
        if (!KiraWert.an("pLangDruck")) { return false; }
        var id = ansicht.kiraId;
        if (id == null) { ansicht.kiraSuchen(); id = ansicht.kiraId; }
        if (id == null) { return false; }
        try {
            Complications.exitTo(id);
            return true;
        } catch (e) {
            return false;
        }
    }
}
