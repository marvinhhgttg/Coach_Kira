// Coach Kira Zifferblatt (22.11.2026)
// Uhrzeit, Datum, Tagesvotum, Soll-Rahmen und heutige Last als Bogen am Rand, Ziel-Countdown,
// dazu Body Battery, Puls und Schritte. Die Kira-Daten kommen aus der Coach-Kira-App
// (private Komplikation „KIRA DATEN“, alle 15 min im Hintergrund aktualisiert) – das
// Zifferblatt selbst braucht keine Kopplung und kein Netz.
import Toybox.Application;
import Toybox.Complications;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;
import Toybox.System;
import Toybox.Time;
import Toybox.Time.Gregorian;
import Toybox.WatchUi;

class KiraZifferblattApp extends Application.AppBase {
    function initialize() {
        AppBase.initialize();
    }

    function getInitialView() as [WatchUi.Views] or [WatchUi.Views, WatchUi.InputDelegates] {
        return [new KiraZifferblatt()];
    }
}

const FARBE = {
    "gruen" => 0x2ECC71, "gelb" => 0xF1C40F, "orange" => 0xE67E22, "rot" => 0xE74C3C, "blau" => 0x3498DB
};
const VOTE_FARBE = {
    "PUSH" => 0x2ECC71, "TRAIN" => 0x2ECC71, "QUAL" => 0x2ECC71, "EASY" => 0xF1C40F, "REST" => 0xE74C3C
};

class KiraZifferblatt extends WatchUi.WatchFace {
    var kiraId as Complications.Id? = null;
    var kira as String? = null;
    var bb as Number? = null;
    var hr as Number? = null;
    var schritte as Number? = null;
    var schlaf as Boolean = false;
    var idBB as Complications.Id? = null;
    var idHR as Complications.Id? = null;
    var idSt as Complications.Id? = null;

    function initialize() {
        WatchFace.initialize();
        try {
            idBB = new Complications.Id(Complications.COMPLICATION_TYPE_BODY_BATTERY);
            idHR = new Complications.Id(Complications.COMPLICATION_TYPE_HEART_RATE);
            idSt = new Complications.Id(Complications.COMPLICATION_TYPE_STEPS);
            Complications.registerComplicationChangeCallback(method(:geaendert));
            Complications.subscribeToUpdates(idBB);
            Complications.subscribeToUpdates(idHR);
            Complications.subscribeToUpdates(idSt);
        } catch (e) { }
        kiraSuchen();
    }

    // Die Komplikation der Coach-Kira-App über ihr Label finden
    function kiraSuchen() as Void {
        try {
            var it = Complications.getComplications();
            var c = it.next();
            while (c != null) {
                var ll = c.longLabel;
                if (ll != null && ll.toString().equals("KIRA DATEN")) {
                    kiraId = c.complicationId;
                    if (kiraId != null) { Complications.subscribeToUpdates(kiraId); }
                    wertKira(c);
                    return;
                }
                c = it.next();
            }
        } catch (e) { }
    }

    function wertKira(c as Complications.Complication) as Void {
        var v = c.value;
        if (v != null) { kira = v.toString(); }
    }

    function geaendert(id as Complications.Id) as Void {
        try {
            var c = Complications.getComplication(id);
            var v = c.value;
            if (kiraId != null && id.equals(kiraId)) { wertKira(c); }
            else if (idBB != null && id.equals(idBB)) { bb = v instanceof Number ? v : null; }
            else if (idHR != null && id.equals(idHR)) { hr = v instanceof Number ? v : null; }
            else if (idSt != null && id.equals(idSt)) { schritte = v instanceof Number ? v : null; }
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

    // vote|stufe|lo|hi|ist|datum|ziel|zielTage|readiness
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

    function zahl(t as Array<String>, i as Number) as Number? {
        if (t.size() <= i || t[i].length() == 0) { return null; }
        return t[i].toNumber();
    }

    function heuteIso() as String {
        var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        return g.year.format("%04d") + "-" + (g.month as Number).format("%02d") + "-" + g.day.format("%02d");
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        var w = dc.getWidth();
        var h = dc.getHeight();
        var cx = w / 2;
        var cy = h / 2;
        dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_BLACK);
        dc.clear();
        var uhr = System.getClockTime();
        var zeit = uhr.hour.format("%02d") + ":" + uhr.min.format("%02d");
        var t = teile();
        var vote = t.size() > 0 ? t[0] : "";
        var heute = t.size() > 5 && t[5].equals(heuteIso());
        var vf = heute && VOTE_FARBE.hasKey(vote) ? VOTE_FARBE.get(vote) as Number : 0x777777;

        if (schlaf) {
            // Always-on: wenige, gedimmte Pixel, Position wandert jede Minute (Einbrennschutz)
            var dx = (uhr.min % 5) * 4 - 8;
            var dy = (uhr.min % 3) * 4 - 4;
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx + dx, cy + dy, Graphics.FONT_NUMBER_MEDIUM, zeit, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            dc.setColor(vf, Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(cx + dx, cy + dy - 70, 5);
            return;
        }

        // Bogen am Rand: heutige Last gegen den Soll-Rahmen (Skala bis 1,3 × Obergrenze)
        var lo = zahl(t, 2);
        var hi = zahl(t, 3);
        var ist = zahl(t, 4);
        var r = cx - 12;
        var start = 225;      // unten links
        var bogen = 270;      // bis unten rechts, im Uhrzeigersinn
        dc.setPenWidth(12);
        dc.setColor(0x262626, Graphics.COLOR_TRANSPARENT);
        dc.drawArc(cx, cy, r, Graphics.ARC_CLOCKWISE, start, start - bogen);
        if (heute && hi != null && hi > 0) {
            var max = hi * 1.3;
            var a0 = start - (bogen * (lo == null ? 0 : lo) / max).toNumber();
            var a1 = start - (bogen * hi / max).toNumber();
            dc.setColor(0x1E5631, Graphics.COLOR_TRANSPARENT);
            dc.drawArc(cx, cy, r, Graphics.ARC_CLOCKWISE, a0, a1);
            if (ist != null && ist > 0) {
                var p = ist > max ? max : ist;
                var c = ist > hi * 1.15 ? 0xE74C3C : ist > hi ? 0xE67E22 : (lo != null && ist >= lo) ? 0x2ECC71 : 0x5DADE2;
                dc.setColor(c, Graphics.COLOR_TRANSPARENT);
                var a2 = start - (bogen * p / max).toNumber();
                if (a2 < start) { dc.drawArc(cx, cy, r, Graphics.ARC_CLOCKWISE, start, a2); }
            }
        }
        dc.setPenWidth(1);

        // Votum oben
        var yV = (h * 0.15).toNumber();
        var vt = kira == null ? "KIRA · App öffnen" : (vote.length() > 0 ? vote : "KIRA") + (heute ? "" : " · alt");
        dc.setColor(vf, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, yV, Graphics.FONT_SMALL, vt, Graphics.TEXT_JUSTIFY_CENTER);
        // Ziel
        var zn = t.size() > 6 ? t[6] : "";
        var zt = zahl(t, 7);
        if (zn.length() > 0 && zt != null) {
            dc.setColor(0x9AD0FF, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, yV + dc.getFontHeight(Graphics.FONT_SMALL), Graphics.FONT_XTINY, zn + " in " + zt.toString() + " T", Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Uhrzeit und Datum
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.47).toNumber(), Graphics.FONT_NUMBER_HOT, zeit, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        var g = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var wt = ["", "So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
        var datum = wt[g.day_of_week as Number] + " " + g.day.format("%02d") + "." + (g.month as Number).format("%02d") + ".";
        dc.setColor(0xBBBBBB, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.595).toNumber(), Graphics.FONT_TINY, datum, Graphics.TEXT_JUSTIFY_CENTER);

        // Rahmen und Last
        var yR = (h * 0.68).toNumber();
        if (heute && hi != null) {
            var rt = (lo != null ? lo.toString() + "-" : "") + hi.toString() + " ESS · heute " + (ist == null ? "0" : ist.toString());
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, yR, Graphics.FONT_XTINY, rt, Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Body Battery, Puls, Schritte
        var yW = (h * 0.75).toNumber();
        var werte = [
            ["BB", bb == null ? "-" : bb.toString()],
            ["HF", hr == null ? "-" : hr.toString()],
            ["Schr.", schritte == null ? "-" : (schritte >= 10000 ? (schritte / 1000).toString() + "k" : schritte.toString())]
        ];
        for (var i = 0; i < 3; i++) {
            var x = cx + (i - 1) * (w * 0.18).toNumber();
            dc.setColor(0x777777, Graphics.COLOR_TRANSPARENT);
            dc.drawText(x, yW, Graphics.FONT_XTINY, werte[i][0], Graphics.TEXT_JUSTIFY_CENTER);
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(x, yW + dc.getFontHeight(Graphics.FONT_XTINY) - 6, Graphics.FONT_TINY, werte[i][1], Graphics.TEXT_JUSTIFY_CENTER);
        }
    }
}
