// Coach Kira Zeigerzifferblatt (22.11.2026)
// Dieselben Werte wie das digitale Zifferblatt (Akku, Votum, Temperatur, Ausdauerwert,
// Trainingsbereitschaft, Summenlast/Soll, Wind, Plan, Datum), in der Mitte zwei Diagramme:
// Last der letzten 7 Tage mit Soll-Band (aus Coach Kira) und Body Battery der letzten 24 h (live).
import Toybox.Application;
import Toybox.Complications;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;
import Toybox.SensorHistory;
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
    var schlaf as Boolean = false;
    var gesucht as Number = -1;
    var bb as Array<Number> = [] as Array<Number>;   // 48 Halbstundenwerte, -1 = keine Daten
    var bbJetzt as Number? = null;
    var bbStand as Number = -100;
    var fKopf as Graphics.FontType = Graphics.FONT_XTINY;
    var fLabel as Graphics.FontType = Graphics.FONT_XTINY;
    var fWert as Graphics.FontType = Graphics.FONT_TINY;
    var fRing as Graphics.FontType = Graphics.FONT_TINY;
    var fKlein as Graphics.FontType = Graphics.FONT_XTINY;

    function initialize() {
        WatchFace.initialize();
        fKopf = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 24, Graphics.FONT_XTINY);
        fLabel = schrift(["RobotoCondensedBold", "RobotoCondensedRegular"], 21, Graphics.FONT_XTINY);
        fWert = schrift(["RobotoCondensedRegular", "RobotoRegular"], 28, Graphics.FONT_TINY);
        fRing = schrift(["RobotoCondensedBold", "RobotoMedium"], 28, Graphics.FONT_TINY);
        fKlein = schrift(["RobotoCondensedRegular", "RobotoRegular"], 18, Graphics.FONT_XTINY);
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
                    if (vv != null) { kira = vv.toString(); }
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

    function bereitFarbe(v as Number) as Number {
        if (v >= 75) { return 0x2ECC71; }
        if (v >= 50) { return 0x3FB8F0; }
        if (v >= 25) { return ORANGE; }
        return 0xE74C3C;
    }

    function lastFarbe(ist as Number, lo as Number?, hi as Number?) as Number {
        if (ist <= 0) { return 0x505050; }
        if (hi == null || hi <= 0) { return 0xBBBBBB; }
        if (ist > hi * 1.15) { return 0xE74C3C; }
        if (ist > hi) { return 0xE67E22; }
        if (lo != null && ist >= lo) { return 0x2ECC71; }
        return 0x3FB8F0;
    }

    // Body Battery der letzten 24 h in 48 Halbstundenwerte (höchstens alle 10 min neu)
    function bbLaden() as Void {
        var jetzt = Time.now().value();
        if (jetzt - bbStand < 600) { return; }
        bbStand = jetzt;
        var r = [] as Array<Number>;
        for (var i = 0; i < 48; i++) { r.add(-1); }
        bbJetzt = null;
        try {
            if ((Toybox has :SensorHistory) && (SensorHistory has :getBodyBatteryHistory)) {
                var it = SensorHistory.getBodyBatteryHistory({ :period => new Time.Duration(24 * 3600), :order => SensorHistory.ORDER_NEWEST_FIRST });
                var s = it.next();
                while (s != null) {
                    if (s.data != null && s.when != null) {
                        var alter = jetzt - s.when.value();
                        var k = 47 - (alter / 1800);
                        if (k >= 0 && k < 48 && r[k] < 0) { r[k] = s.data.toNumber(); }
                        if (bbJetzt == null) { bbJetzt = s.data.toNumber(); }
                    }
                    s = it.next();
                }
            }
        } catch (e) { }
        bb = r;
    }

    function zeiger(dc as Graphics.Dc, cx as Number, cy as Number, winkel as Float, laenge as Number, rueck as Number, breite as Number, farbe as Number) as Void {
        var s = Math.sin(winkel);
        var c = Math.cos(winkel);
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.setPenWidth(breite);
        dc.drawLine(cx - (rueck * s).toNumber(), cy + (rueck * c).toNumber(), cx + (laenge * s).toNumber(), cy - (laenge * c).toNumber());
        dc.setPenWidth(1);
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

        // Stundenstriche
        for (var i = 0; i < 12; i++) {
            var a = i * Math.PI / 6;
            var r1 = cx - 4;
            var r2 = (i % 3 == 0) ? cx - 26 : cx - 18;
            dc.setColor(schlaf ? 0x606060 : 0xE0E0E0, Graphics.COLOR_TRANSPARENT);
            dc.setPenWidth(i % 3 == 0 ? 5 : 3);
            dc.drawLine(cx + (r1 * Math.sin(a)).toNumber(), cy - (r1 * Math.cos(a)).toNumber(), cx + (r2 * Math.sin(a)).toNumber(), cy - (r2 * Math.cos(a)).toNumber());
        }
        dc.setPenWidth(1);

        if (schlaf) {
            zeiger(dc, cx, cy, wStd, (w * 0.25).toNumber(), 0, 5, 0x909090);
            zeiger(dc, cx, cy, wMin, (w * 0.38).toNumber(), 0, 3, 0x909090);
            return;
        }

        if (kiraId == null && gesucht != uhr.min) { gesucht = uhr.min; kiraSuchen(); }
        var t = zerlegen(kira, "|");
        var heute = feld(t, 5).equals(heuteIso());
        var vote = feld(t, 0);

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
        block(dc, (w * 0.78).toNumber(), (h * 0.69).toNumber(), "PLAN" + (sp.length() > 0 ? " " + sp.toUpper() : ""), pl == null ? "-" : (pl == 0 ? "Ruhe" : pl.toString()));

        var rx = (w * 0.15).toNumber();
        var rr = (w * 0.066).toNumber();
        var rd = heute ? zahl(t, 8) : null;
        ring(dc, rx, cy, rr, rd == null ? 0.0 : rd / 100.0, rd == null ? GRAU : bereitFarbe(rd));
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, cy, fRing, rd == null ? "-" : rd.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(rx, cy + rr + 7, fLabel, "BEREIT", Graphics.TEXT_JUSTIFY_CENTER);

        var lx = w - rx;
        var lo = heute ? zahl(t, 2) : null;
        var hi = heute ? zahl(t, 3) : null;
        var ist = heute ? zahl(t, 4) : null;
        ring(dc, lx, cy, rr, (hi != null && hi > 0 && ist != null) ? ist.toFloat() / hi : 0.0, ist == null ? GRAU : lastFarbe(ist, lo, hi));
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(lx, cy, fRing, ist == null ? "-" : ist.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(ORANGE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(lx, cy + rr + 7, fLabel, hi == null ? "SOLL" : (lo != null ? lo.toString() + "-" : "") + hi.toString(), Graphics.TEXT_JUSTIFY_CENTER);

        // Diagramm 1: Last 7 Tage mit Soll-Band
        var sieben = zerlegen(feld(t, 14), ";");
        var istL = zerlegen(sieben.size() > 0 ? sieben[0] : null, ",");
        var loL = zerlegen(sieben.size() > 1 ? sieben[1] : null, ",");
        var hiL = zerlegen(sieben.size() > 2 ? sieben[2] : null, ",");
        var x0 = (w * 0.31).toNumber();
        var x1 = (w * 0.69).toNumber();
        var y0 = (h * 0.24).toNumber();
        var y1 = (h * 0.40).toNumber();
        var n = heute ? istL.size() : 0;
        if (n > 0) {
            var mx = 1;
            for (var i = 0; i < n; i++) {
                var v = istL[i].length() > 0 ? istL[i].toNumber() : 0;
                var v2 = hiL.size() > i && hiL[i].length() > 0 ? hiL[i].toNumber() : 0;
                if (v > mx) { mx = v; }
                if (v2 > mx) { mx = v2; }
            }
            var bw = (x1 - x0) / n;
            for (var i = 0; i < n; i++) {
                var x = x0 + i * bw;
                var v = istL[i].length() > 0 ? istL[i].toNumber() : 0;
                var l1 = loL.size() > i && loL[i].length() > 0 ? loL[i].toNumber() : null;
                var h1 = hiL.size() > i && hiL[i].length() > 0 ? hiL[i].toNumber() : null;
                if (i == n - 1) { l1 = lo; h1 = hi; }
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
            // Wochentage, heute weiß
            var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
            var kz = ["S", "M", "D", "M", "D", "F", "S"];
            for (var i = 0; i < n; i++) {
                var tag = ((g.day_of_week as Number) - 1 - (n - 1 - i) + 70) % 7;
                dc.setColor(i == n - 1 ? Graphics.COLOR_WHITE : GRAU, Graphics.COLOR_TRANSPARENT);
                dc.drawText(x0 + i * bw + bw / 2, y1 + 1, fKlein, kz[tag], Graphics.TEXT_JUSTIFY_CENTER);
            }
        }

        // Diagramm 2: Body Battery 24 h
        bbLaden();
        var b0 = (h * 0.56).toNumber();
        var b1 = (h * 0.68).toNumber();
        var punkte = [] as Array<[Numeric, Numeric]>;
        var lx0 = x0;
        for (var i = 0; i < 48; i++) {
            if (bb[i] >= 0) { punkte.add([x0 + (x1 - x0) * i / 47, b1 - (b1 - b0) * bb[i] / 100]); }
        }
        if (punkte.size() >= 2) {
            var flaeche = [] as Array<[Numeric, Numeric]>;
            flaeche.addAll(punkte);
            flaeche.add([punkte[punkte.size() - 1][0], b1]);
            flaeche.add([punkte[0][0], b1]);
            dc.setColor(0x123048, Graphics.COLOR_TRANSPARENT);
            dc.fillPolygon(flaeche);
            dc.setColor(0x3FB8F0, Graphics.COLOR_TRANSPARENT);
            dc.setPenWidth(2);
            for (var i = 1; i < punkte.size(); i++) {
                dc.drawLine(punkte[i - 1][0], punkte[i - 1][1], punkte[i][0], punkte[i][1]);
            }
            dc.setPenWidth(1);
        }
        dc.setColor(GRAU, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, b1 + 1, fKlein, "BB " + (bbJetzt == null ? "-" : bbJetzt.toString()), Graphics.TEXT_JUSTIFY_CENTER);

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
