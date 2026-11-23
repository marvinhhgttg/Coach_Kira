// Coach Kira Zeigerzifferblatt (22.11.2026)
// Dieselben Werte wie das digitale Zifferblatt (Akku, Votum, Temperatur, Ausdauerwert,
// Trainingsbereitschaft, Summenlast/Soll, Wind, Plan, Datum), in der Mitte zwei Diagramme:
// Last der letzten 7 Tage mit Soll-Band (aus Coach Kira) und Form der letzten 6 Wochen (Fitness gegen Ermüdung).
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

class KiraZeigerApp extends Application.AppBase {
    function initialize() {
        AppBase.initialize();
    }

    function getInitialView() as [WatchUi.Views] or [WatchUi.Views, WatchUi.InputDelegates] {
        return [new KiraZeiger()];
    }
}

const ORANGE = 0xFF9A1F;
const GRAU = 0x8A8A8A;
const SPUR = 0x2A2A2A;
const VOTE_FARBE = {
    "PUSH" => 0x2ECC71, "TRAIN" => 0x2ECC71, "QUAL" => 0x2ECC71, "EASY" => 0xF1C40F, "REST" => 0xE74C3C
};

class KiraZeiger extends WatchUi.WatchFace {
    var kiraId as Complications.Id? = null;
    var kira as String? = null;
    var wocheId as Complications.Id? = null;
    var woche as String? = null;
    var formId as Complications.Id? = null;
    var form as String? = null;
    var schlaf as Boolean = false;
    var gesucht as Number = -1;
    var fKopf as Graphics.FontType = Graphics.FONT_XTINY;
    var fLabel as Graphics.FontType = Graphics.FONT_XTINY;
    var fWert as Graphics.FontType = Graphics.FONT_TINY;
    var fRing as Graphics.FontType = Graphics.FONT_TINY;
    var fKlein as Graphics.FontType = Graphics.FONT_XTINY;

    function initialize() {
        WatchFace.initialize();
        fKopf = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 28, Graphics.FONT_XTINY);
        fLabel = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 25, Graphics.FONT_XTINY);
        fWert = schrift(["RobotoCondensedRegular", "RobotoRegular"], 33, Graphics.FONT_TINY);
        fRing = schrift(["RobotoCondensedBold", "RobotoMedium"], 32, Graphics.FONT_TINY);
        fKlein = schrift(["RobotoCondensedRegular", "RobotoRegular"], 22, Graphics.FONT_XTINY);
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
                var lab = ll != null ? ll.toString().toUpper() : "";
                var vv = c.value;
                var id = c.complicationId;
                if (id != null) {
                    if (lab.find("KIRA DATEN") != null || (kiraId == null && vv instanceof String && vv.find("|") != null && c.getType() == Complications.COMPLICATION_TYPE_INVALID)) {
                        kiraId = id; Complications.subscribeToUpdates(id); if (vv != null) { kira = vv.toString(); }
                    } else if (lab.find("KIRA WOCHE") != null) {
                        wocheId = id; Complications.subscribeToUpdates(id); if (vv != null) { woche = vv.toString(); }
                    } else if (lab.find("KIRA FORM") != null) {
                        formId = id; Complications.subscribeToUpdates(id); if (vv != null) { form = vv.toString(); }
                    }
                }
                c = it.next();
            }
        } catch (e) { }
    }

    function geaendert(id as Complications.Id) as Void {
        try {
            var v = Complications.getComplication(id).value;
            if (v == null) { return; }
            if (kiraId != null && id.equals(kiraId)) { kira = v.toString(); }
            else if (wocheId != null && id.equals(wocheId)) { woche = v.toString(); }
            else if (formId != null && id.equals(formId)) { form = v.toString(); }
        } catch (e) { }
    }

    // ein Zeichen (0-9A-Za-z) → 0..61, "_" oder fehlend → null
    function dekod(s as String, i as Number) as Number? {
        if (i >= s.length()) { return null; }
        var z = s.substring(i, i + 1) as String;
        if (z.equals("_")) { return null; }
        return "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".find(z);
    }

    function onExitSleep() as Void {
        schlaf = false;
        WatchUi.requestUpdate();
    }

    function onEnterSleep() as Void {
        schlaf = true;
        WatchUi.requestUpdate();
    }

    function zerlegen(s as String?, sep as String) as Array<String> {
        var r = [] as Array<String>;
        if (s == null) { return r; }
        var i = s.find(sep);
        while (i != null) {
            r.add(s.substring(0, i) as String);
            s = s.substring(i + 1, s.length()) as String;
            i = s.find(sep);
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

    function ring(dc as Graphics.Dc, x as Number, y as Number, r as Number, anteil as Float, farbe as Number) as Void {
        dc.setPenWidth(7);
        dc.setColor(SPUR, Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(x, y, r);
        if (anteil > 0.01) {
            var a = anteil > 1.0 ? 1.0 : anteil;
            dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
            if (a >= 0.999) { dc.drawCircle(x, y, r); }
            else { dc.drawArc(x, y, r, Graphics.ARC_CLOCKWISE, 90, 90 - (a * 360).toNumber()); }
        }
        dc.setPenWidth(1);
    }

    function block(dc as Graphics.Dc, x as Number, y as Number, label as String, wert as String) as Void {
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y, fLabel, label, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y + dc.getFontHeight(fLabel) - 3, fWert, wert, Graphics.TEXT_JUSTIFY_CENTER);
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
        var grenzen = [0, 25, 50, 75, 95, 100];
        var farben = [0xD9433A, 0xE5822F, 0x5DB75B, 0x3A78CF, 0x6E45E9];
        dc.setPenWidth(pen);
        for (var i = 0; i < 5; i++) {
            var a1 = bogenWinkel(grenzen[i]) - (i > 0 ? 3 : 0);
            var a2 = bogenWinkel(grenzen[i + 1]) + (i < 4 ? 3 : 0);
            dc.setColor(v == null ? 0x3A3A3A : farben[i], Graphics.COLOR_TRANSPARENT);
            dc.drawArc(x, y, r, Graphics.ARC_CLOCKWISE, ((a1 % 360) + 360) % 360, ((a2 % 360) + 360) % 360);
        }
        dc.setPenWidth(1);
        if (v != null) {
            var w = v > 100 ? 100 : (v < 0 ? 0 : v);
            var rad = Math.toRadians(225 - w * 2.7);
            var px = x + (r * Math.cos(rad)).toNumber();
            var py = y - (r * Math.sin(rad)).toNumber();
            dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(px, py, pen / 2 + 5);
            dc.setColor(bereitFarbe(w), Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(px, py, pen / 2 + 3);
        }
    }

    function lastFarbe(ist as Number, lo as Number?, hi as Number?) as Number {
        if (ist <= 0) { return 0x505050; }
        if (hi == null || hi <= 0) { return 0xBBBBBB; }
        if (ist > hi * 1.15) { return 0xE74C3C; }
        if (ist > hi) { return 0xE67E22; }
        if (lo != null && ist >= lo) { return 0x2ECC71; }
        return 0x3FB8F0;
    }

    function zeiger(dc as Graphics.Dc, cx as Number, cy as Number, winkel as Float, laenge as Number, rueck as Number, breite as Number, farbe as Number) as Void {
        var s = Math.sin(winkel);
        var c = Math.cos(winkel);
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.setPenWidth(breite);
        dc.drawLine(cx - (rueck * s).toNumber(), cy + (rueck * c).toNumber(), cx + (laenge * s).toNumber(), cy - (laenge * c).toNumber());
        dc.setPenWidth(1);
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


    // 23.11.2026: Wochen-Ring (dünn, außen um den Tagessummen-Ring): Wochen-Ist gegen Kiras Wochenbudget
    function wochenRing(dc as Graphics.Dc, x as Number, y as Number, r as Number, ist as Number?, bud as Number?) as Void {
        if (bud == null || bud <= 0) { return; }
        var a = ist == null ? 0.0 : ist.toFloat() / bud;
        dc.setPenWidth(4);
        dc.setColor(0x123A40, Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(x, y, r);
        if (a > 0.005) {
            dc.setColor(0x00BCD4, Graphics.COLOR_TRANSPARENT);
            if (a >= 0.999) { dc.drawCircle(x, y, r); }
            else { dc.drawArc(x, y, r, Graphics.ARC_CLOCKWISE, 90, 90 - (a * 360).toNumber()); }
        }
        if (a > 1.005) {
            var b = a - 1.0 > 1.0 ? 1.0 : a - 1.0;
            dc.setColor(0xFF6A00, Graphics.COLOR_TRANSPARENT);
            if (b >= 0.999) { dc.drawCircle(x, y, r); }
            else { dc.drawArc(x, y, r, Graphics.ARC_CLOCKWISE, 90, 90 - (b * 360).toNumber()); }
        }
        dc.setPenWidth(1);
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
        var wStd = ((uhr.hour % 12) * 60 + uhr.min) / 720.0 * Math.PI * 2;
        var wMin = (uhr.min * 60 + uhr.sec) / 3600.0 * Math.PI * 2;

        // Ruhemodus (Always-on, AMOLED): nur die beiden gedimmten Zeiger – keine festen Striche,
        // damit kein Pixel dauerhaft an derselben Stelle leuchtet (23.11.2026)
        if (schlaf) {
            zeiger(dc, cx, cy, wStd, (w * 0.25).toNumber(), 0, 5, 0x909090);
            zeiger(dc, cx, cy, wMin, (w * 0.38).toNumber(), 0, 3, 0x909090);
            return;
        }

        // Stundenstriche
        for (var i = 0; i < 12; i++) {
            if (i == 3 || i == 9) { continue; }   // 3 und 9 Uhr: dort sitzen die Ringe
            var a = i * Math.PI / 6;
            var r1 = cx - 18;
            var r2 = (i % 3 == 0) ? cx - 40 : cx - 32;
            dc.setColor(0xE0E0E0, Graphics.COLOR_TRANSPARENT);
            dc.setPenWidth(i % 3 == 0 ? 5 : 3);
            dc.drawLine(cx + (r1 * Math.sin(a)).toNumber(), cy - (r1 * Math.cos(a)).toNumber(), cx + (r2 * Math.sin(a)).toNumber(), cy - (r2 * Math.cos(a)).toNumber());
        }
        dc.setPenWidth(1);

        if ((kiraId == null || wocheId == null || formId == null) && gesucht != uhr.min) { gesucht = uhr.min; kiraSuchen(); }
        var t = zerlegen(kira, "|");
        var heute = feld(t, 5).equals(heuteIso());
        var vote = feld(t, 0);

        votumRing(dc, cx, cy, feld(t, 1), vote, heute ? zahl(t, 2) : null, heute ? zahl(t, 3) : null, heute ? zahl(t, 4) : null, heute);

        // Kopf: Akku · Votum
        var st = System.getSystemStats();
        var tage = (st has :batteryInDays) && st.batteryInDays != null ? st.batteryInDays.toNumber() : null;
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.085).toNumber(), fKopf, st.battery.toNumber().toString() + "%" + (tage != null ? " · " + tage.toString() + "T" : ""), Graphics.TEXT_JUSTIFY_CENTER);
        var vf = heute && VOTE_FARBE.hasKey(vote) ? VOTE_FARBE.get(vote) as Number : GRAU;
        dc.setColor(vf, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.135).toNumber(), fKopf, kira == null ? (kiraId == null ? "KIRA · APP ÖFFNEN" : "KIRA · WARTET") : "KIRA · " + (vote.length() > 0 ? vote : "-") + (heute ? "" : " (ALT)"), Graphics.TEXT_JUSTIFY_CENTER);

        // Seitenwerte (wie digital)
        var temp = "-";
        var wind = "-";
        try {
            var wc = Weather.getCurrentConditions();
            if (wc != null) {
                if (wc.temperature != null) { temp = wc.temperature.toNumber().toString() + "°"; }
                if (wc.windSpeed != null) { wind = richtung(wc.windBearing) + " " + (wc.windSpeed * 3.6).toNumber().toString(); }
            }
        } catch (e) { }
        var es = zahl(t, 12);
        block(dc, (w * 0.22).toNumber(), (h * 0.25).toNumber(), "TEMP", temp);
        block(dc, (w * 0.78).toNumber(), (h * 0.25).toNumber(), "AUSD.", es == null ? "-" : es.toString());
        var pl = heute ? zahl(t, 9) : null;
        var sp = heute ? feld(t, 10) : "";
        block(dc, (w * 0.22).toNumber(), (h * 0.69).toNumber(), "WIND", wind);
        block(dc, (w * 0.78).toNumber(), (h * 0.69).toNumber(), "PLAN", pl == null ? "-" : (pl == 0 ? "Ruhe" : (sp.length() > 0 ? sp + " " : "") + pl.toString()));
        // Live von der Uhr, Mitte unten zwischen WIND und PLAN: Erholungszeit und Body Battery
        var erh = nativ(Complications.COMPLICATION_TYPE_RECOVERY_TIME);
        var bbn = nativ(Complications.COMPLICATION_TYPE_BODY_BATTERY);
        dc.setColor(0xC8C8C8, Graphics.COLOR_TRANSPARENT);
        var eab = erholtAb(erh, false);
        dc.drawText(cx, (h * 0.718).toNumber(), fKlein, eab.equals("ERHOLT") ? eab : "ERH " + eab, Graphics.TEXT_JUSTIFY_CENTER);
        dc.drawText(cx, (h * 0.772).toNumber(), fKlein, "BB " + (bbn == null ? "-" : bbn.toString()), Graphics.TEXT_JUSTIFY_CENTER);

        var rx = (w * 0.15).toNumber();
        var rr = (w * 0.074).toNumber();
        var rd = heute ? zahl(t, 8) : null;
        bereitBogen(dc, rx, cy, rr, 6, rd);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, cy, fRing, rd == null ? "-" : rd.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, cy + rr + 7, fLabel, (heute && feld(t, 14).length() > 0 ? "TR " + feld(t, 14) : "BEREIT"), Graphics.TEXT_JUSTIFY_CENTER);

        var lx = w - rx;
        var lo = heute ? zahl(t, 2) : null;
        var hi = heute ? zahl(t, 3) : null;
        var ist = heute ? zahl(t, 4) : null;
        ring(dc, lx, cy, rr, (hi != null && hi > 0 && ist != null) ? ist.toFloat() / hi : 0.0, ist == null ? GRAU : lastFarbe(ist, lo, hi));
        wochenRing(dc, lx, cy, rr + 7, zahl(t, 15), zahl(t, 16));
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(lx, cy, fRing, ist == null ? "-" : ist.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(lx, cy + rr + (zahl(t, 16) != null ? 13 : 7), fLabel, hi == null ? "SOLL" : (lo != null ? lo.toString() + "-" : "") + hi.toString(), Graphics.TEXT_JUSTIFY_CENTER);

        // Diagramm 1: Last 7 Tage mit Soll-Band (Komplikation „KIRA WOCHE“: IIIIIII;LLLLLLL;HHHHHHH, ESS / 8)
        var wk = zerlegen(woche, ";");
        var wI = wk.size() > 0 ? wk[0] : "";
        var wL = wk.size() > 1 ? wk[1] : "";
        var wH = wk.size() > 2 ? wk[2] : "";
        var x0 = (w * 0.31).toNumber();
        var x1 = (w * 0.69).toNumber();
        var y0 = (h * 0.24).toNumber();
        var y1 = (h * 0.40).toNumber();
        var n = heute ? wI.length() : 0;
        if (n < 2) {
            dc.setColor(GRAU, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, (y0 + y1) / 2, fKlein, "LAST 7 T · APP ÖFFNEN", Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        } else {
            var vs = [] as Array<Number>;
            var los = [] as Array<Number?>;
            var his = [] as Array<Number?>;
            var mx = 1;
            for (var i = 0; i < n; i++) {
                var v = dekod(wI, i);
                var l1 = dekod(wL, i);
                var h1 = dekod(wH, i);
                vs.add(v == null ? 0 : v * 8);
                los.add(l1 == null ? null : l1 * 8);
                his.add(h1 == null ? null : h1 * 8);
            }
            los[n - 1] = lo;
            his[n - 1] = hi;
            if (ist != null) { vs[n - 1] = ist; }
            for (var i = 0; i < n; i++) {
                if (vs[i] > mx) { mx = vs[i]; }
                var hh = his[i];
                if (hh != null && hh > mx) { mx = hh; }
            }
            var bw = (x1 - x0) / n;
            for (var i = 0; i < n; i++) {
                var x = x0 + i * bw;
                var v = vs[i];
                var l1 = los[i];
                var h1 = his[i];
                if (h1 != null && h1 > 0) {
                    var ya = y1 - ((y1 - y0) * h1 / mx);
                    var yb = y1 - ((y1 - y0) * (l1 == null ? 0 : l1) / mx);
                    dc.setColor(0x28343F, Graphics.COLOR_TRANSPARENT);
                    dc.fillRectangle(x + 2, ya, bw - 4, yb - ya > 1 ? yb - ya : 1);
                }
                var bh = v > 0 ? ((y1 - y0) * v / mx) : 2;
                if (bh < 2) { bh = 2; }
                dc.setColor(lastFarbe(v, l1, h1), Graphics.COLOR_TRANSPARENT);
                dc.fillRectangle(x + bw / 4, y1 - bh, bw / 2, bh);
            }
            var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
            var kz = ["S", "M", "D", "M", "D", "F", "S"];
            for (var i = 0; i < n; i++) {
                var tag = ((g.day_of_week as Number) - 1 - (n - 1 - i) + 70) % 7;
                dc.setColor(i == n - 1 ? Graphics.COLOR_WHITE : GRAU, Graphics.COLOR_TRANSPARENT);
                dc.drawText(x0 + i * bw + bw / 2, y1 + 1, fKlein, kz[tag], Graphics.TEXT_JUSTIFY_CENTER);
            }
        }

        // Diagramm 2: Form 6 Wochen (Komplikation „KIRA FORM“: min,max,ctl,atl;CCC…;AAA…)
        var fk = zerlegen(form, ";");
        var kopf = zerlegen(fk.size() > 0 ? fk[0] : null, ",");
        var fC = fk.size() > 1 ? fk[1] : "";
        var fA = fk.size() > 2 ? fk[2] : "";
        var b0 = (h * 0.545).toNumber();
        var b1 = (h * 0.645).toNumber();
        var m = fC.length() < fA.length() ? fC.length() : fA.length();
        if (m >= 2) {
            for (var k = 0; k < 2; k++) {
                var xs = k == 0 ? fA : fC;
                var pax = -1;
                var pay = 0;
                dc.setColor(k == 0 ? 0xE67E22 : 0x2ECC71, Graphics.COLOR_TRANSPARENT);
                dc.setPenWidth(k == 0 ? 2 : 3);
                for (var i = 0; i < m; i++) {
                    var vv = dekod(xs, i);
                    if (vv == null) { continue; }
                    var px = x0 + (x1 - x0) * i / (m - 1);
                    var py = b1 - (b1 - b0) * vv / 61;
                    if (pax >= 0) { dc.drawLine(pax, pay, px, py); }
                    pax = px;
                    pay = py;
                }
            }
            dc.setPenWidth(1);
        }
        var cTxt = kopf.size() > 2 && kopf[2].length() > 0 ? kopf[2] : "-";
        var aTxt = kopf.size() > 3 && kopf[3].length() > 0 ? kopf[3] : "-";
        var yt = b1 + 2;
        dc.setColor(0x2ECC71, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx - 4, yt, fKlein, "FIT " + cTxt, Graphics.TEXT_JUSTIFY_RIGHT);
        dc.setColor(0xE67E22, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx + 4, yt, fKlein, "ERM " + aTxt, Graphics.TEXT_JUSTIFY_LEFT);

        // Datum
        var gd = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var wt = ["", "SO", "MO", "DI", "MI", "DO", "FR", "SA"];
        var mo = ["", "JAN", "FEB", "MÄR", "APR", "MAI", "JUN", "JUL", "AUG", "SEP", "OKT", "NOV", "DEZ"];
        dc.setColor(0xD0D0D0, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.855).toNumber(), fKopf, wt[gd.day_of_week as Number] + " " + gd.day.toString() + ". " + mo[gd.month as Number], Graphics.TEXT_JUSTIFY_CENTER);

        // Zeiger
        zeiger(dc, cx, cy, wStd, (w * 0.25).toNumber(), 14, 9, 0xEDEDED);
        zeiger(dc, cx, cy, wMin, (w * 0.39).toNumber(), 14, 6, 0xEDEDED);
        zeiger(dc, cx, cy, uhr.sec / 60.0 * Math.PI * 2, (w * 0.41).toNumber(), 26, 2, ORANGE);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.fillCircle(cx, cy, 8);
        dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_TRANSPARENT);
        dc.fillCircle(cx, cy, 3);
    }
}
