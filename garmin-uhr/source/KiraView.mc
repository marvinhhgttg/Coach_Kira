import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Communications;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;
import Toybox.ActivityMonitor;
import Toybox.SensorHistory;
import Toybox.PersistedContent;
import Toybox.System;
import Toybox.Time;
import Toybox.Timer;
import Toybox.WatchUi;

// Laden und Koppeln im Vordergrund (ein Objekt, Rückrufe über method(:…))
var kiraNetz as KiraNetz? = null;
function netz() as KiraNetz {
    if (kiraNetz == null) { kiraNetz = new KiraNetz(); }
    return kiraNetz as KiraNetz;
}

class KiraNetz {
    var status as String? = null;   // "laedt", "koppelt", Fehlertext
    var code as String? = null;     // Kopplungscode zum Anzeigen
    var anfrageId as String? = null;
    var timer as Timer.Timer? = null;
    var versuche as Number = 0;

    function initialize() {
    }

    function laden() as Void {
        var tok = Storage.getValue("tok");
        if (tok == null) { return; }
        status = "laedt";
        WatchUi.requestUpdate();
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget", { "uhr" => "1", "mehr" => "1" },
            {
                :method => Communications.HTTP_REQUEST_METHOD_GET,
                :headers => { "Authorization" => "Bearer " + tok },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            method(:geladen));
    }

    function geladen(rc as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        try {
            if (rc == 200 && data instanceof Dictionary) {
                var u = data.get("uhr");
                if (u instanceof Dictionary) { Storage.setValue("u", u); data.remove("uhr"); }
                Storage.setValue("d", data);
                Storage.setValue("at", Time.now().value());
                kiraKomplikation(data);
                status = null;
            } else if (rc == 401) {
                Storage.deleteValue("tok");
                status = "Kopplung widerrufen";
            } else {
                status = rc < 0 ? "Kein Netz (" + rc.toString() + ")" : "Server " + rc.toString();
            }
        } catch (e) {
            status = "Fehler: " + e.getErrorMessage();
        }
        WatchUi.requestUpdate();
    }

    function koppeln() as Void {
        status = "koppelt";
        code = null;
        WatchUi.requestUpdate();
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/anfrage", { "geraet" => "fenix 8 Pro" },
            {
                :method => Communications.HTTP_REQUEST_METHOD_POST,
                :headers => { "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            method(:angefragt));
    }

    function angefragt(rc as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        try {
            if (rc == 200 && data instanceof Dictionary && data.get("id") != null) {
                anfrageId = data.get("id").toString();
                code = data.get("code").toString();
                status = null;
                versuche = 0;
                if (timer == null) { timer = new Timer.Timer(); }
                (timer as Timer.Timer).start(method(:abfragen), 4000, true);
            } else {
                status = rc < 0 ? "Kein Netz (" + rc.toString() + ")" : "Kopplung: Server " + rc.toString();
            }
        } catch (e) {
            status = "Fehler: " + e.getErrorMessage();
        }
        WatchUi.requestUpdate();
    }

    function abfragen() as Void {
        versuche += 1;
        if (versuche > 150 || anfrageId == null) {   // 10 min
            stop("Code abgelaufen – START");
            return;
        }
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/status?id=" + anfrageId, null,
            {
                :method => Communications.HTTP_REQUEST_METHOD_GET,
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            method(:abgefragt));
    }

    function abgefragt(rc as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        try {
            if (rc == 200 && data instanceof Dictionary && data.get("token") != null) {
                Storage.setValue("tok", data.get("token").toString());
                stop(null);
                code = null;
                var app = Application.getApp();
                if (app instanceof KiraApp) { app.hintergrundAn(); }
                laden();
            } else if (rc == 404) {
                stop("Code abgelaufen – START");
            }
        } catch (e) {
            stop("Fehler: " + e.getErrorMessage());
        }
    }

    // 22.11.2026: Befinden 1–5 von der Uhr ans Cockpit
    function befinden(wert as Number) as Void {
        var tok = Storage.getValue("tok");
        if (tok == null) { return; }
        status = "sendet Befinden ...";
        WatchUi.requestUpdate();
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/befinden", { "value" => wert },
            {
                :method => Communications.HTTP_REQUEST_METHOD_POST,
                :headers => { "Authorization" => "Bearer " + tok, "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            method(:befindenFertig));
    }

    function befindenFertig(rc as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        if (rc == 200) {
            status = null;
            Storage.setValue("bef", Time.now().value());
            laden();   // Votum mit Befinden neu holen
        } else {
            status = rc < 0 ? "Kein Netz (" + rc.toString() + ")" : "Befinden: Server " + rc.toString();
        }
        WatchUi.requestUpdate();
    }

    function stop(meldung as String?) as Void {
        if (timer != null) { (timer as Timer.Timer).stop(); }
        anfrageId = null;
        status = meldung;
        WatchUi.requestUpdate();
    }
}

class KiraView extends WatchUi.View {
    // 0 Votum, 1 Werte, 2 Woche, 3 Last 14 T, 4 Morgen 7 T, 5 Form-Kurve, 6 Letzte Einheit, 7 Ziel, 8 Jetzt
    var seite as Number = 0;
    const SEITEN = 9;

    function initialize() {
        View.initialize();
    }

    function onShow() as Void {
        if (Storage.getValue("tok") != null) {
            netz().laden();
        }
    }

    function blaettern(n as Number) as Void {
        seite = (seite + n + SEITEN) % SEITEN;
        WatchUi.requestUpdate();
    }

    // Zeile mittig, auf die Kreisbreite bei y gekürzt; gibt die nächste y-Position zurück
    function zeile(dc as Graphics.Dc, y as Number, font as Graphics.FontType, text as String, farbe as Number) as Number {
        var w = dc.getWidth();
        var r = w / 2;
        var fh = dc.getFontHeight(font);
        var mitte = y + fh / 2 - r;
        var halb = r * r - mitte * mitte;
        var breite = halb > 0 ? (2 * Math.sqrt(halb) * 0.86).toNumber() : 0;
        if (breite < 40) { return y + fh; }
        var t = Graphics.fitTextToArea(KiraHilfen.sauber(text), font, breite, fh, true);
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.drawText(r, y, font, t, Graphics.TEXT_JUSTIFY_CENTER);
        return y + fh;
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        var h = dc.getHeight();
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_BLACK);
        dc.clear();
        if (Storage.getValue("tok") == null) {
            zeigeKopplung(dc, h);
            return;
        }
        var d = KiraHilfen.daten();
        if (!(d instanceof Dictionary)) {
            zeile(dc, (h * 0.12).toNumber(), Graphics.FONT_XTINY, "COACH KIRA", 0x9A9A9A);
            var st = netz().status;
            zeile(dc, (h * 0.45).toNumber(), Graphics.FONT_SMALL, st != null && !st.equals("laedt") ? st : "Lädt ...", Graphics.COLOR_WHITE);
            return;
        }
        var u = Storage.getValue("u");
        if (seite == 0) { seiteVotum(dc, d, h); }
        else if (seite == 1) { seiteWerte(dc, d, h); }
        else if (seite == 2) { seiteTage(dc, d, h); }
        else if (seite == 3) { seiteLast(dc, d, u, h); }
        else if (seite == 4) { seiteMorgen(dc, u, h); }
        else if (seite == 5) { seiteKurve(dc, d, u, h); }
        else if (seite == 6) { seiteLetzte(dc, u, h); }
        else if (seite == 7) { seiteZiel(dc, d, h); }
        else { seiteForm(dc, d, h); }
        // Seitenanzeige rechts
        var w = dc.getWidth();
        for (var i = 0; i < SEITEN; i++) {
            dc.setColor(i == seite ? Graphics.COLOR_WHITE : 0x555555, Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(w - 14, h / 2 - (SEITEN - 1) * 7 + i * 14, 3);
        }
    }

    function seiteVotum(dc as Graphics.Dc, d as Dictionary, h as Number) as Void {
        var w = dc.getWidth();
        var heute = KiraHilfen.vonHeute(d);
        var farbe = heute ? KiraHilfen.farbe(KiraHilfen.feld(d, "votum", "stufe")) : 0x888888;
        var y = (h * 0.11).toNumber();
        y = zeile(dc, y, Graphics.FONT_XTINY, heute ? "COACH KIRA" : "COACH KIRA - STAND GESTERN", 0x9A9A9A);
        y += 10;
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.fillCircle(w / 2, y + 12, 12);
        y += 30;
        var label = KiraHilfen.feld(d, "votum", "label");
        y = zeile(dc, y, Graphics.FONT_LARGE, label == null ? "-" : label.toString().toUpper(), farbe);
        var txt = KiraHilfen.feld(d, "votum", "text");
        if (txt != null) { y = zeile(dc, y, Graphics.FONT_SMALL, txt.toString(), Graphics.COLOR_WHITE); }
        var gr = KiraHilfen.feld(d, "votum", "gruende");
        if (gr instanceof Array && gr.size() > 0) {
            y = zeile(dc, y + 4, Graphics.FONT_XTINY, gr[0].toString(), 0xBBBBBB);
        }
        // Soll heute – das Wichtigste nach dem Votum
        var l = d.get("last");
        if (y < (h * 0.62).toNumber()) { y = (h * 0.62).toNumber(); } else { y += 6; }
        if (l instanceof Dictionary) {
            // 22.11.2026: Soll = Rahmen des Kerns (wie Cockpit), Plan darunter
            y = zeile(dc, y, Graphics.FONT_SMALL, "Soll " + KiraHilfen.sollText(l) + " ESS", Graphics.COLOR_WHITE);
            var sp = l.get("sportHeute");
            var zo = l.get("zoneHeute");
            var plan = l.get("planHeute");
            var sz = (plan != null ? "Plan " + KiraHilfen.zahl(plan, 0) : "") + (sp != null && sp.toString().length() > 0 ? " " + sp.toString() : "") + (zo != null && zo.toString().length() > 0 ? " " + zo.toString() : "");
            var lage = l.get("planLage");
            if (lage != null && lage.toString().equals("darueber") && l.get("erledigt") != true) { sz += " zu viel"; }
            if (l.get("erledigt") == true) { sz += (sz.length() > 0 ? " · " : "") + "erledigt"; }
            if (sz.length() > 0) { zeile(dc, y, Graphics.FONT_XTINY, sz, 0xBBBBBB); }
        }
        unten(dc, h, d);
    }

    function seiteWerte(dc as Graphics.Dc, d as Dictionary, h as Number) as Void {
        var w = dc.getWidth();
        var y = (h * 0.11).toNumber();
        y = zeile(dc, y, Graphics.FONT_XTINY, "MORGENWERTE", 0x9A9A9A);
        var m = d.get("morgen");
        y += 6;
        if (m instanceof Dictionary && m.get("eingetragen") != false) {
            var zh = dc.getFontHeight(Graphics.FONT_XTINY) + dc.getFontHeight(Graphics.FONT_SMALL) + 2;
            kachel(dc, w * 0.33, y, "Readiness", KiraHilfen.zahl(m.get("readiness"), 0));
            kachel(dc, w * 0.67, y, "Schlaf", KiraHilfen.zahl(m.get("schlaf"), 0));
            y += zh;
            kachel(dc, w * 0.33, y, "HRV", KiraHilfen.zahl(m.get("hrv"), 0));
            kachel(dc, w * 0.67, y, "Ruhepuls", KiraHilfen.zahl(m.get("rhr"), 0));
            y += zh;
        } else {
            y = zeile(dc, y + 20, Graphics.FONT_SMALL, "Morgenwerte fehlen", 0xE67E22) + 20;
        }
        var z = d.get("ziel");
        if (z instanceof Dictionary && z.get("name") != null) {
            var zt = z.get("name").toString() + " in " + KiraHilfen.zahl(z.get("tage"), 0) + " T";
            if (z.get("tourform") != null) { zt += " · " + KiraHilfen.zahl(z.get("tourform"), 0) + " %"; }
            y = zeile(dc, y, Graphics.FONT_XTINY, zt, 0x9AD0FF);
        }
        var l = d.get("last");
        if (l instanceof Dictionary) {
            var band = l.get("acwrBand");
            y = zeile(dc, y, Graphics.FONT_XTINY, "ACWR " + KiraHilfen.zahl(l.get("acwr"), 2) + (band != null ? " · " + band.toString() : ""), Graphics.COLOR_WHITE);
            var woche = l.get("woche");
            if (woche instanceof Array && woche.size() > 0) { balken(dc, y + 6, woche); }
        }
    }

    // Seite 3: die nächsten 4 Tage aus dem Plan, lange Einheiten aus Kiras Wochenplan
    function seiteTage(dc as Graphics.Dc, d as Dictionary, h as Number) as Void {
        var w = dc.getWidth();
        var y = (h * 0.11).toNumber();
        y = zeile(dc, y, Graphics.FONT_XTINY, "WOCHE", 0x9A9A9A);
        y += 2;
        var xs = d.get("naechste");
        if (!(xs instanceof Array) || xs.size() == 0) {
            zeile(dc, (h * 0.42).toNumber(), Graphics.FONT_SMALL, "Kein Plan", 0xBBBBBB);
            return;
        }
        var fnt = xs.size() > 4 ? Graphics.FONT_XTINY : Graphics.FONT_TINY;
        var fh = dc.getFontHeight(fnt);
        var lang = "";
        for (var i = 0; i < xs.size() && i < 7; i++) {
            var x = xs[i];
            if (!(x instanceof Dictionary)) { continue; }
            var dt = x.get("d") != null ? x.get("d").toString() : "";
            var ess = x.get("ess");
            var sp = KiraHilfen.sportDe(x.get("sport"));
            var zo = x.get("zone") != null ? x.get("zone").toString() : "";
            var ruhe = ess == null || ess == 0 || sp.equals("Ruhe");
            var lh = x.get("langH");
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.2, y, fnt, KiraHilfen.wochentag(dt).substring(0, 2), Graphics.TEXT_JUSTIFY_LEFT);
            dc.setColor(ruhe ? 0x777777 : Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.48, y, fnt, ruhe ? "-" : KiraHilfen.zahl(ess, 0), Graphics.TEXT_JUSTIFY_RIGHT);
            dc.setColor(lh != null ? 0x9AD0FF : (ruhe ? 0x777777 : 0xBBBBBB), Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.53, y, fnt, ruhe ? "Ruhe" : sp + (zo.length() > 0 && !zo.equals("Off") ? " " + zo : ""), Graphics.TEXT_JUSTIFY_LEFT);
            if (lh != null && lang.length() == 0) { lang = "Lang " + KiraHilfen.wochentag(dt) + " " + KiraHilfen.stunden(lh); }
            y += fh + (fnt == Graphics.FONT_XTINY ? 0 : 4);
        }
        y += 4;
        zeile(dc, y, Graphics.FONT_XTINY, lang.length() > 0 ? lang : "ESS · Sport · Zone", lang.length() > 0 ? 0x9AD0FF : 0x777777);
    }

    // Seite 4: Endurance Score, Fitness, Garmin-Tag
    function seiteForm(dc as Graphics.Dc, d as Dictionary, h as Number) as Void {
        var w = dc.getWidth();
        var y = (h * 0.11).toNumber();
        y = zeile(dc, y, Graphics.FONT_XTINY, "FORM JETZT", 0x9A9A9A);
        var f = d.get("form");
        if (!(f instanceof Dictionary)) {
            zeile(dc, (h * 0.42).toNumber(), Graphics.FONT_SMALL, "Keine Daten", 0xBBBBBB);
            return;
        }
        y = zeile(dc, y + 4, Graphics.FONT_XTINY, "Endurance Score", 0x9A9A9A);
        y = zeile(dc, y, Graphics.FONT_MEDIUM, KiraHilfen.zahl(f.get("es"), 0), Graphics.COLOR_WHITE);
        var e28 = f.get("es28");
        if (e28 != null) {
            var plus = e28 >= 0 ? "+" : "";
            y = zeile(dc, y, Graphics.FONT_XTINY, plus + KiraHilfen.zahl(e28, 0) + " in 28 Tagen", e28 >= 0 ? 0x2ECC71 : 0xE67E22);
        }
        y += 6;
        var zh = dc.getFontHeight(Graphics.FONT_XTINY) + dc.getFontHeight(Graphics.FONT_SMALL) + 2;
        kachel(dc, w * 0.27, y, "CTL", KiraHilfen.zahl(f.get("ctl"), 0));
        kachel(dc, w * 0.50, y, "ATL", KiraHilfen.zahl(f.get("atl"), 0));
        kachel(dc, w * 0.73, y, "Mono", KiraHilfen.zahl(f.get("monotonie"), 1));
        y += zh + 4;
        // 22.11.2026: live von der Uhr statt vom Mac
        var bb = sensorJetzt(:bb);
        var st = sensorJetzt(:stress);
        var sch = null;
        try { sch = ActivityMonitor.getInfo().steps; } catch (e) { }
        kachel(dc, w * 0.27, y, "Body Batt.", KiraHilfen.zahl(bb, 0));
        kachel(dc, w * 0.50, y, "Stress", KiraHilfen.zahl(st, 0));
        kachel(dc, w * 0.73, y, "Schritte", sch != null ? (sch >= 10000 ? (sch / 1000).toString() + "k" : sch.toString()) : "-");
        zeile(dc, y + zh + 2, Graphics.FONT_XTINY, "live von der Uhr", 0x777777);
    }

    // Letzter Wert aus der Sensor-Historie der Uhr
    function sensorJetzt(art as Symbol) {
        try {
            var it = null;
            if (art == :bb && SensorHistory has :getBodyBatteryHistory) { it = SensorHistory.getBodyBatteryHistory({ :period => 1 }); }
            else if (art == :stress && SensorHistory has :getStressHistory) { it = SensorHistory.getStressHistory({ :period => 1 }); }
            if (it == null) { return null; }
            var s = it.next();
            return s != null && s.data != null ? s.data.toNumber() : null;
        } catch (e) {
            return null;
        }
    }

    function arr(u, a as String, b as String) as Array? {
        if (!(u instanceof Dictionary)) { return null; }
        var x = u.get(a);
        if (!(x instanceof Dictionary)) { return null; }
        var y = x.get(b);
        return y instanceof Array ? y : null;
    }

    function ohneDaten(dc as Graphics.Dc, h as Number, titel as String) as Void {
        zeile(dc, (h * 0.11).toNumber(), Graphics.FONT_XTINY, titel, 0x9A9A9A);
        zeile(dc, (h * 0.42).toNumber(), Graphics.FONT_SMALL, "Keine Daten", 0xBBBBBB);
        zeile(dc, (h * 0.55).toNumber(), Graphics.FONT_XTINY, "START = neu laden", 0x777777);
    }

    // Seite 4: Last der letzten 14 Tage mit Soll-Rahmen als Band
    function seiteLast(dc as Graphics.Dc, d as Dictionary, u, h as Number) as Void {
        var ist = arr(u, "l14", "ist");
        var lo = arr(u, "l14", "lo");
        var hi = arr(u, "l14", "hi");
        if (ist == null || lo == null || hi == null || ist.size() == 0) { ohneDaten(dc, h, "LAST 14 TAGE"); return; }
        var w = dc.getWidth();
        zeile(dc, (h * 0.11).toNumber(), Graphics.FONT_XTINY, "LAST 14 TAGE", 0x9A9A9A);
        var n = ist.size();
        var x0 = (w * 0.18).toNumber();
        var x1 = (w * 0.82).toNumber();
        var y0 = (h * 0.25).toNumber();
        var y1 = (h * 0.60).toNumber();
        var mx = 1;
        for (var i = 0; i < n; i++) {
            if (ist[i] > mx) { mx = ist[i]; }
            if (hi[i] > mx) { mx = hi[i]; }
        }
        var bw = (x1 - x0) / n;
        var drin = 0;
        var trainiert = 0;
        var summe = 0;
        for (var i = 0; i < n; i++) {
            var x = x0 + i * bw;
            // Band
            if (hi[i] > 0) {
                var yb1 = y1 - ((y1 - y0) * hi[i] / mx);
                var yb0 = y1 - ((y1 - y0) * (lo[i] > 0 ? lo[i] : 0) / mx);
                dc.setColor(0x33414F, Graphics.COLOR_TRANSPARENT);
                dc.fillRectangle(x + 1, yb1, bw - 2, yb0 - yb1 + 1);
            }
            var v = ist[i];
            summe += v;
            var farbe = 0x555555;
            if (v > 0) {
                trainiert += 1;
                if (hi[i] > 0 && v > hi[i]) { farbe = 0xE67E22; }
                else if (lo[i] >= 0 && v >= lo[i]) { farbe = 0x2ECC71; drin += 1; }
                else { farbe = 0x5DADE2; }
            }
            var bh = v > 0 ? ((y1 - y0) * v / mx) : 3;
            if (bh < 3) { bh = 3; }
            dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x + bw / 4, y1 - bh, bw / 2, bh);
        }
        // heute markieren
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.fillRectangle(x0 + (n - 1) * bw + bw / 4, y1 + 4, bw / 2, 3);
        var y = y1 + 12;
        var l = d.get("last");
        if (l instanceof Dictionary) { y = zeile(dc, y, Graphics.FONT_TINY, "Soll " + KiraHilfen.sollText(l), Graphics.COLOR_WHITE); }
        y = zeile(dc, y, Graphics.FONT_XTINY, "Rahmen " + drin.toString() + "/" + trainiert.toString() + " · Ø " + (summe / n).toString(), 0xBBBBBB);
        zeile(dc, y, Graphics.FONT_XTINY, "grau = Rahmen", 0x777777);
    }

    // Kleine Linie mit Ampelpunkten; v < 0 = fehlt
    function linie(dc as Graphics.Dc, x0 as Number, x1 as Number, y as Number, hoch as Number, xs as Array, stufen as String?, farbe as Number) as Void {
        var lo = 99999;
        var hi = -99999;
        for (var i = 0; i < xs.size(); i++) {
            if (xs[i] < 0) { continue; }
            if (xs[i] < lo) { lo = xs[i]; }
            if (xs[i] > hi) { hi = xs[i]; }
        }
        if (hi < lo) { return; }
        if (hi == lo) { hi = lo + 1; }
        var n = xs.size();
        var dx = n > 1 ? (x1 - x0).toFloat() / (n - 1) : 0.0;
        var px = -1;
        var py = -1;
        dc.setPenWidth(2);
        for (var i = 0; i < n; i++) {
            if (xs[i] < 0) { continue; }
            var x = (x0 + i * dx).toNumber();
            var yy = (y + hoch - (hoch * (xs[i] - lo).toFloat() / (hi - lo))).toNumber();
            if (px >= 0) {
                dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
                dc.drawLine(px, py, x, yy);
            }
            px = x;
            py = yy;
        }
        dc.setPenWidth(1);
        if (stufen != null) {
            for (var i = 0; i < n; i++) {
                if (xs[i] < 0) { continue; }
                var x = (x0 + i * dx).toNumber();
                var yy = (y + hoch - (hoch * (xs[i] - lo).toFloat() / (hi - lo))).toNumber();
                var c = i < stufen.length() ? stufen.substring(i, i + 1) : "x";
                dc.setColor(c.equals("g") ? 0x2ECC71 : c.equals("y") ? 0xF1C40F : c.equals("o") ? 0xE67E22 : c.equals("r") ? 0xE74C3C : c.equals("b") ? 0x3498DB : 0x888888, Graphics.COLOR_TRANSPARENT);
                dc.fillCircle(x, yy, i == n - 1 ? 6 : 4);
            }
        }
    }

    // Seite 5: Morgenwerte der letzten 7 Tage
    function seiteMorgen(dc as Graphics.Dc, u, h as Number) as Void {
        if (!(u instanceof Dictionary) || !(u.get("m7") instanceof Dictionary)) { ohneDaten(dc, h, "MORGEN 7 TAGE"); return; }
        var m = u.get("m7") as Dictionary;
        var w = dc.getWidth();
        zeile(dc, (h * 0.11).toNumber(), Graphics.FONT_XTINY, "MORGEN 7 TAGE", 0x9A9A9A);
        var reihen = [["Ready", "rd", "rdS"], ["HRV", "hrv", "hrvS"], ["Puls", "rhr", "rhrS"], ["Schlaf", "sl", "slS"]];
        var y = (h * 0.22).toNumber();
        var zh = (h * 0.135).toNumber();
        for (var i = 0; i < reihen.size(); i++) {
            var xs = m.get(reihen[i][1]);
            var st = m.get(reihen[i][2]);
            if (!(xs instanceof Array)) { continue; }
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.13, y + zh / 2, Graphics.FONT_XTINY, reihen[i][0], Graphics.TEXT_JUSTIFY_LEFT | Graphics.TEXT_JUSTIFY_VCENTER);
            linie(dc, (w * 0.37).toNumber(), (w * 0.70).toNumber(), y + 10, zh - 20, xs, st instanceof String ? st : null, 0x777777);
            var letzt = xs.size() > 0 ? xs[xs.size() - 1] : -1;
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.88, y + zh / 2, Graphics.FONT_SMALL, letzt < 0 ? "-" : letzt.toString(), Graphics.TEXT_JUSTIFY_RIGHT | Graphics.TEXT_JUSTIFY_VCENTER);
            y += zh;
        }
        zeile(dc, y + 4, Graphics.FONT_XTINY, "rechts = heute", 0x777777);
    }

    // Seite 6: Fitness (CTL) und Ermüdung (ATL) 6 Wochen, Ausdauerwert 90 Tage
    function seiteKurve(dc as Graphics.Dc, d as Dictionary, u, h as Number) as Void {
        var ctl = arr(u, "form", "ctl");
        var atl = arr(u, "form", "atl");
        var es = arr(u, "form", "es");
        if (ctl == null || atl == null) { ohneDaten(dc, h, "FORM-KURVE"); return; }
        var w = dc.getWidth();
        zeile(dc, (h * 0.11).toNumber(), Graphics.FONT_XTINY, "FORM 6 WOCHEN", 0x9A9A9A);
        // gemeinsame Skala für CTL und ATL
        var both = [];
        both.addAll(ctl);
        both.addAll(atl);
        var lo = 99999;
        var hi = -1;
        for (var i = 0; i < both.size(); i++) { if (both[i] >= 0 && both[i] < lo) { lo = both[i]; } if (both[i] > hi) { hi = both[i]; } }
        var x0 = (w * 0.20).toNumber();
        var x1 = (w * 0.80).toNumber();
        var y0 = (h * 0.22).toNumber();
        var hoch = (h * 0.20).toNumber();
        // ATL zuerst (dünn), CTL darüber
        var atlS = skaliert(atl, lo, hi);
        var ctlS = skaliert(ctl, lo, hi);
        linieFest(dc, x0, x1, y0, hoch, atlS, 0xE67E22, 1);
        linieFest(dc, x0, x1, y0, hoch, ctlS, 0x2ECC71, 3);
        var y = y0 + hoch + 6;
        var c = letzter(ctl);
        var a = letzter(atl);
        dc.setColor(0x2ECC71, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w * 0.30, y, Graphics.FONT_XTINY, "CTL " + c.toString(), Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(0xE67E22, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w * 0.70, y, Graphics.FONT_XTINY, "ATL " + a.toString(), Graphics.TEXT_JUSTIFY_CENTER);
        y += dc.getFontHeight(Graphics.FONT_XTINY) + 6;
        if (es != null && es.size() > 1) {
            y = zeile(dc, y, Graphics.FONT_XTINY, "AUSDAUERWERT 90 TAGE", 0x9A9A9A);
            var eh = (h * 0.13).toNumber();
            linie(dc, x0 + 20, x1 - 20, y + 2, eh, es, null, 0x5DADE2);
            y += eh + 8;
            var e1 = letzter(es);
            var e0 = -1;
            for (var i = 0; i < es.size(); i++) { if (es[i] > 0) { e0 = es[i]; break; } }
            var diff = e1 > 0 && e0 > 0 ? e1 - e0 : 0;
            zeile(dc, y, Graphics.FONT_SMALL, e1.toString() + "  " + (diff >= 0 ? "+" : "") + diff.toString(), diff >= 0 ? 0x2ECC71 : 0xE67E22);
        }
    }

    function skaliert(xs as Array, lo as Number, hi as Number) as Array {
        var out = [];
        var sp = hi > lo ? hi - lo : 1;
        for (var i = 0; i < xs.size(); i++) { out.add(xs[i] < 0 ? -1.0 : (xs[i] - lo).toFloat() / sp); }
        return out;
    }

    function linieFest(dc as Graphics.Dc, x0 as Number, x1 as Number, y as Number, hoch as Number, xs as Array, farbe as Number, pen as Number) as Void {
        var n = xs.size();
        var dx = n > 1 ? (x1 - x0).toFloat() / (n - 1) : 0.0;
        var px = -1;
        var py = -1;
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.setPenWidth(pen);
        for (var i = 0; i < n; i++) {
            if (xs[i] < 0) { continue; }
            var x = (x0 + i * dx).toNumber();
            var yy = (y + hoch - hoch * xs[i]).toNumber();
            if (px >= 0) { dc.drawLine(px, py, x, yy); }
            px = x;
            py = yy;
        }
        dc.setPenWidth(1);
    }

    function letzter(xs as Array) as Number {
        for (var i = xs.size() - 1; i >= 0; i--) { if (xs[i] >= 0) { return xs[i]; } }
        return 0;
    }

    // Seite 7: letzte Einheit mit Zonen und Vergleich zum Rahmen
    function seiteLetzte(dc as Graphics.Dc, u, h as Number) as Void {
        if (!(u instanceof Dictionary) || !(u.get("letzte") instanceof Dictionary)) { ohneDaten(dc, h, "LETZTE EINHEIT"); return; }
        var a = u.get("letzte") as Dictionary;
        var w = dc.getWidth();
        var y = zeile(dc, (h * 0.11).toNumber(), Graphics.FONT_XTINY, "LETZTE EINHEIT", 0x9A9A9A);
        var dt = a.get("d") != null ? a.get("d").toString() : "";
        y = zeile(dc, y, Graphics.FONT_XTINY, KiraHilfen.wochentag(dt).substring(0, 2) + " · " + KiraHilfen.zahl(a.get("min"), 0) + " min" + (a.get("km") != null ? " · " + KiraHilfen.zahl(a.get("km"), 1) + " km" : ""), 0xBBBBBB);
        y = zeile(dc, y, Graphics.FONT_XTINY, a.get("name") != null ? a.get("name").toString() : "", Graphics.COLOR_WHITE);
        y = zeile(dc, y + 2, Graphics.FONT_MEDIUM, KiraHilfen.zahl(a.get("ess"), 0) + " ESS", Graphics.COLOR_WHITE);
        var lage = a.get("lage");
        var lf = 0xBBBBBB;
        var lt = "";
        if (lage != null && a.get("hi") != null) {
            var ls = lage.toString();
            lf = ls.equals("darueber") ? 0xE67E22 : ls.equals("im Rahmen") ? 0x2ECC71 : 0x5DADE2;
            lt = (ls.equals("darueber") ? "über" : ls.equals("im Rahmen") ? "im" : "unter") + " Rahmen " + KiraHilfen.zahl(a.get("lo"), 0) + "-" + KiraHilfen.zahl(a.get("hi"), 0);
        }
        if (lt.length() > 0) { y = zeile(dc, y, Graphics.FONT_XTINY, lt, lf); }
        var te = a.get("te");
        if (te != null) { y = zeile(dc, y, Graphics.FONT_XTINY, "TE " + KiraHilfen.zahl(te, 1) + "/" + KiraHilfen.zahl(a.get("tean"), 1), 0xBBBBBB); }
        // Zonenbalken
        var z = a.get("z");
        if (z instanceof Array && z.size() >= 5) {
            var summe = 0;
            for (var i = 0; i < 5; i++) { summe += z[i]; }
            if (summe > 0) {
                var farben = [0x9AA5B1, 0x3498DB, 0x2ECC71, 0xE67E22, 0xE74C3C];
                var x0 = (w * 0.18).toNumber();
                var bw = (w * 0.64).toNumber();
                var x = x0;
                y += 6;
                for (var i = 0; i < 5; i++) {
                    var b = bw * z[i] / summe;
                    dc.setColor(farben[i], Graphics.COLOR_TRANSPARENT);
                    dc.fillRectangle(x, y, b, 14);
                    x += b;
                }
                y += 18;
                var t = "";
                for (var i = 0; i < 5; i++) { if (z[i] > 0) { t += (t.length() > 0 ? " " : "") + "Z" + (i + 1).toString() + " " + z[i].toString() + "'"; } }
                zeile(dc, y, Graphics.FONT_XTINY, t, 0x9A9A9A);
            }
        }
    }

    // Seite 8: Ziel-Countdown als Ring
    function seiteZiel(dc as Graphics.Dc, d as Dictionary, h as Number) as Void {
        var z = d.get("ziel");
        if (!(z instanceof Dictionary) || z.get("name") == null) { ohneDaten(dc, h, "ZIEL"); return; }
        var w = dc.getWidth();
        zeile(dc, (h * 0.11).toNumber(), Graphics.FONT_XTINY, "ZIEL", 0x9A9A9A);
        var p = z.get("tourform");
        var pz = p != null ? p.toNumber() : 0;
        if (pz > 100) { pz = 100; }
        var cx = w / 2;
        var cy = (h * 0.43).toNumber();
        var r = (h * 0.19).toNumber();
        dc.setPenWidth(14);
        dc.setColor(0x333333, Graphics.COLOR_TRANSPARENT);
        dc.drawCircle(cx, cy, r);
        var farbe = pz >= 80 ? 0x2ECC71 : pz >= 60 ? 0xF1C40F : 0xE67E22;
        if (pz > 0) {
            dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
            dc.drawArc(cx, cy, r, Graphics.ARC_CLOCKWISE, 90, 90 - (pz * 360 / 100 >= 360 ? 359 : pz * 360 / 100));
        }
        dc.setPenWidth(1);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, cy, Graphics.FONT_SMALL, p != null ? pz.toString() + "%" : "-", Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        var y = cy + r + 14;
        y = zeile(dc, y, Graphics.FONT_SMALL, z.get("name").toString(), Graphics.COLOR_WHITE);
        y = zeile(dc, y, Graphics.FONT_XTINY, "in " + KiraHilfen.zahl(z.get("tage"), 0) + " Tagen", 0x9AD0FF);
        // nächste lange Einheit aus der Woche
        var xs = d.get("naechste");
        if (xs instanceof Array) {
            for (var i = 0; i < xs.size(); i++) {
                var x = xs[i];
                if (x instanceof Dictionary && x.get("langH") != null) {
                    zeile(dc, y, Graphics.FONT_XTINY, "Lang " + KiraHilfen.wochentag(x.get("d").toString()) + " " + KiraHilfen.stunden(x.get("langH")), 0xBBBBBB);
                    break;
                }
            }
        }
    }

    function kachel(dc as Graphics.Dc, x, y as Number, titel as String, v as String) as Void {
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y, Graphics.FONT_XTINY, titel, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y + dc.getFontHeight(Graphics.FONT_XTINY), Graphics.FONT_SMALL, v, Graphics.TEXT_JUSTIFY_CENTER);
    }

    // 7 Balken der letzten Tage (Ist-ESS), heute rechts
    function balken(dc as Graphics.Dc, y as Number, woche as Array) as Number {
        var w = dc.getWidth();
        var hoch = 36;
        var mx = 1;
        for (var i = 0; i < woche.size(); i++) {
            var e = woche[i];
            if (e instanceof Dictionary && e.get("ist") != null && e.get("ist") > mx) { mx = e.get("ist"); }
        }
        var n = woche.size();
        var breite = (w * 0.5 / n).toNumber();
        var x0 = (w * 0.25).toNumber();
        for (var i = 0; i < n; i++) {
            var e = woche[i];
            var v = e instanceof Dictionary && e.get("ist") != null ? e.get("ist") : 0;
            var bh = v > 0 ? (hoch * v / mx).toNumber() : 2;
            if (bh < 2) { bh = 2; }
            dc.setColor(i == n - 1 ? 0x2ECC71 : 0x777777, Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x0 + i * breite + 3, y + hoch - bh, breite - 6, bh);
        }
        return y + hoch;
    }

    function unten(dc as Graphics.Dc, h as Number, d as Dictionary) as Void {
        var st = netz().status;
        var t = st == null ? "Stand " + KiraHilfen.uhrzeit(Storage.getValue("at")) : st.equals("laedt") ? "aktualisiert ..." : st;
        var m = d.get("morgen");
        if (st == null && m instanceof Dictionary && m.get("befinden") == null && KiraHilfen.vonHeute(d)) { t = "Befinden fehlt - MENU"; }
        zeile(dc, (h * 0.86).toNumber(), Graphics.FONT_XTINY, t, st != null && !st.equals("laedt") ? 0xE67E22 : 0x777777);
    }

    function zeigeKopplung(dc as Graphics.Dc, h as Number) as Void {
        var n = netz();
        zeile(dc, (h * 0.12).toNumber(), Graphics.FONT_XTINY, "COACH KIRA", 0x9A9A9A);
        if (n.code != null) {
            var y = zeile(dc, (h * 0.24).toNumber(), Graphics.FONT_SMALL, "Code im Cockpit eingeben", Graphics.COLOR_WHITE);
            y = zeile(dc, y, Graphics.FONT_XTINY, "Heute > Technik", 0x9A9A9A);
            zeile(dc, y + 16, Graphics.FONT_NUMBER_MEDIUM, n.code as String, 0x2ECC71);
            zeile(dc, (h * 0.74).toNumber(), Graphics.FONT_XTINY, "wartet auf Bestätigung ...", 0x777777);
        } else {
            var st = n.status;
            var y = zeile(dc, (h * 0.36).toNumber(), Graphics.FONT_SMALL, "Mit Cockpit koppeln", Graphics.COLOR_WHITE);
            zeile(dc, y + 4, Graphics.FONT_XTINY, st != null && st.equals("koppelt") ? "fordert Code an ..." : "START drücken", 0x9A9A9A);
            if (st != null && !st.equals("koppelt")) { zeile(dc, (h * 0.62).toNumber(), Graphics.FONT_XTINY, st, 0xE67E22); }
        }
    }
}

class KiraDelegate extends WatchUi.BehaviorDelegate {
    var view as KiraView;

    function initialize(v as KiraView) {
        BehaviorDelegate.initialize();
        view = v;
    }

    function onNextPage() as Boolean {
        view.blaettern(1);
        return true;
    }

    function onPreviousPage() as Boolean {
        view.blaettern(-1);
        return true;
    }

    // START: koppeln bzw. neu laden
    function onSelect() as Boolean {
        if (Storage.getValue("tok") == null) {
            netz().koppeln();
        } else {
            netz().laden();
        }
        return true;
    }

    // MENU (lange UP): Befinden, neu laden, Kopplung lösen
    function onMenu() as Boolean {
        if (Storage.getValue("tok") == null) { return false; }
        kiraMenuZeigen();
        return true;
    }
}

class KiraLoesen extends WatchUi.ConfirmationDelegate {
    function initialize() {
        ConfirmationDelegate.initialize();
    }

    function onResponse(r) as Boolean {
        if (r == WatchUi.CONFIRM_YES) {
            Storage.deleteValue("tok");
            Storage.deleteValue("d");
            WatchUi.requestUpdate();
        }
        return true;
    }
}
