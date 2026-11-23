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
    function initialize() {
        AppBase.initialize();
    }

    function getInitialView() as [WatchUi.Views] or [WatchUi.Views, WatchUi.InputDelegates] {
        return [new KiraZifferblatt()];
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

    function initialize() {
        WatchFace.initialize();
        fKopf = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 25, Graphics.FONT_XTINY);
        fLabel = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 23, Graphics.FONT_XTINY);
        fWert = schrift(["RobotoCondensedRegular", "RobotoRegular"], 31, Graphics.FONT_TINY);
        fRing = schrift(["RobotoCondensedBold", "RobotoMedium"], 33, Graphics.FONT_SMALL);
        fDatum = schrift(["RobotoCondensedRegular", "RobotoRegular"], 28, Graphics.FONT_TINY);
        try { Complications.registerComplicationChangeCallback(method(:geaendert)); } catch (e) { }
        kiraSuchen();
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
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y, fLabel, label, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(wertFarbe, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y + dc.getFontHeight(fLabel) - 2, fWert, wert, Graphics.TEXT_JUSTIFY_CENTER);
    }

    function richtung(grad as Number?) as String {
        if (grad == null) { return ""; }
        var r = ["N", "NO", "O", "SO", "S", "SW", "W", "NW"];
        return r[((grad + 22) % 360) / 45];
    }

    function bereitFarbe(v as Number) as Number {
        if (v >= 75) { return 0x2ECC71; }
        if (v >= 50) { return 0x3FB8F0; }
        if (v >= 25) { return ORANGE; }
        return 0xE74C3C;
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
        dc.setColor(dimmen(f), Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(cx, cy, r);
        var a = (hi != null && hi > 0 && ist != null && ist > 0) ? ist.toFloat() / hi : 0.0;
        if (a > 0.005) {
            dc.setColor(f, Graphics.COLOR_TRANSPARENT);
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
        return ["ERH " + (erh == null ? "-" : ((erh + 59) / 60).toString() + "h"), "BB " + (bb == null ? "-" : bb.toString())];
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

        votumRing(dc, cx, cy, feld(t, 1), vote, heute ? zahl(t, 2) : null, heute ? zahl(t, 3) : null, heute ? zahl(t, 4) : null, heute);

        // Akku: Strichbogen oben + Prozent / Tage
        var st = System.getSystemStats();
        var akku = st.battery;
        var n = 30;
        for (var i = 0; i < n; i++) {
            var a = Math.toRadians(135 - i * 90.0 / (n - 1));
            var r1 = cx - 19;
            var r2 = cx - 33;
            dc.setColor(i < (akku / 100.0 * n) ? ORANGE : SPUR, Graphics.COLOR_TRANSPARENT);
            dc.setPenWidth(4);
            dc.drawLine(cx + (r1 * Math.cos(a)).toNumber(), cy - (r1 * Math.sin(a)).toNumber(),
                        cx + (r2 * Math.cos(a)).toNumber(), cy - (r2 * Math.sin(a)).toNumber());
        }
        dc.setPenWidth(1);
        var tage = (st has :batteryInDays) && st.batteryInDays != null ? st.batteryInDays.toNumber() : null;
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.085).toNumber(), fKopf, akku.toNumber().toString() + "%" + (tage != null ? " · " + tage.toString() + "T" : ""), Graphics.TEXT_JUSTIFY_CENTER);

        // Kira-Votum statt Markenname
        var vf = heute && VOTE_FARBE.hasKey(vote) ? VOTE_FARBE.get(vote) as Number : GRAU;
        dc.setColor(vf, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.145).toNumber(), fKopf, kira == null ? (kiraId == null ? "KIRA · APP ÖFFNEN" : "KIRA · WARTET") : "KIRA · " + (vote.length() > 0 ? vote : "-") + (heute ? "" : " (ALT)"), Graphics.TEXT_JUSTIFY_CENTER);

        // Uhrzeit: Stunden weiß, Minuten orange (wie fēnix)
        dc.setColor(0xE8E8E8, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.35).toNumber(), Graphics.FONT_NUMBER_HOT, std, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.57).toNumber(), Graphics.FONT_NUMBER_HOT, min, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);

        // Links: Trainingsbereitschaft (Morgenwert aus Coach Kira)
        var rx = (w * 0.17).toNumber();
        var rr = (w * 0.085).toNumber();
        var ry = (h * 0.49).toNumber();
        var rd = heute ? zahl(t, 8) : null;
        ring(dc, rx, ry, rr, rd == null ? 0.0 : rd / 100.0, rd == null ? GRAU : bereitFarbe(rd));
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, ry, fRing, rd == null ? "-" : rd.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, ry + rr + 8, fLabel, (heute && feld(t, 14).length() > 0 ? "TR " + feld(t, 14) : "BEREIT"), Graphics.TEXT_JUSTIFY_CENTER);

        // Rechts: heutige Summenlast gegen Soll
        var lx = w - rx;
        var lo = heute ? zahl(t, 2) : null;
        var hi = heute ? zahl(t, 3) : null;
        var ist = heute ? zahl(t, 4) : null;
        var lf = GRAU;
        if (hi != null && ist != null) {
            lf = ist > hi * 1.15 ? 0xE74C3C : ist > hi ? 0xE67E22 : (lo != null && ist >= lo) ? 0x2ECC71 : 0x3FB8F0;
        }
        ring(dc, lx, ry, rr, (hi != null && hi > 0 && ist != null) ? ist.toFloat() / hi : 0.0, lf);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(lx, ry, fRing, ist == null ? "-" : ist.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(lx, ry + rr + 8, fLabel, (hi == null ? "-" : (lo != null ? lo.toString() + "-" : "") + hi.toString()), Graphics.TEXT_JUSTIFY_CENTER);

        // Oben links: Wetter · oben rechts: Ausdauerwert
        var temp = "-";
        var wind = "-";
        try {
            var wc = Weather.getCurrentConditions();
            if (wc != null) {
                if (wc.temperature != null) { temp = wc.temperature.toNumber().toString() + "°"; }
                if (wc.windSpeed != null) { wind = richtung(wc.windBearing) + " " + (wc.windSpeed * 3.6).toNumber().toString(); }
            }
        } catch (e) { }
        var yo = (h * 0.23).toNumber();
        block(dc, (w * 0.19).toNumber(), yo, "TEMP", temp, Graphics.COLOR_WHITE);
        var es = zahl(t, 12);
        block(dc, (w * 0.81).toNumber(), yo, "AUSD.", es == null ? "-" : es.toString(), Graphics.COLOR_WHITE);

        // Unten links: Wind (km/h) · unten rechts: Plan
        var yu = (h * 0.67).toNumber();
        block(dc, (w * 0.27).toNumber(), yu, "WIND", wind, Graphics.COLOR_WHITE);
        var pl = heute ? zahl(t, 9) : null;
        var sp = feld(t, 10);
        block(dc, (w * 0.73).toNumber(), yu, "PLAN" + (sp.length() > 0 && heute ? " " + sp.toUpper() : ""), pl == null ? "-" : (pl == 0 ? "Ruhe" : pl.toString()), Graphics.COLOR_WHITE);

        // Live von der Uhr: Erholungszeit · Body Battery
        var nz = nativZeile();
        dc.setColor(0xC8C8C8, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.818).toNumber(), fLabel, nz[0] + " · " + nz[1], Graphics.TEXT_JUSTIFY_CENTER);

        // Datum unten
        var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var wt = ["", "SO", "MO", "DI", "MI", "DO", "FR", "SA"];
        var mo = ["", "JAN", "FEB", "MÄR", "APR", "MAI", "JUN", "JUL", "AUG", "SEP", "OKT", "NOV", "DEZ"];
        dc.setColor(0xD0D0D0, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.875).toNumber(), fDatum, wt[g.day_of_week as Number] + " " + g.day.toString() + ". " + mo[g.month as Number], Graphics.TEXT_JUSTIFY_CENTER);
    }
}
