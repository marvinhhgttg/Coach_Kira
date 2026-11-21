import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Communications;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;
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
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget", { "uhr" => "1" },
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
                Storage.setValue("d", data);
                Storage.setValue("at", Time.now().value());
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
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/status", { "id" => anfrageId },
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

    function stop(meldung as String?) as Void {
        if (timer != null) { (timer as Timer.Timer).stop(); }
        anfrageId = null;
        status = meldung;
        WatchUi.requestUpdate();
    }
}

class KiraView extends WatchUi.View {
    var seite as Number = 0;   // 0 Votum, 1 Werte, 2 Nächste Tage, 3 Form
    const SEITEN = 4;

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
        if (seite == 0) { seiteVotum(dc, d, h); }
        else if (seite == 1) { seiteWerte(dc, d, h); }
        else if (seite == 2) { seiteTage(dc, d, h); }
        else { seiteForm(dc, d, h); }
        // Seitenanzeige rechts
        var w = dc.getWidth();
        for (var i = 0; i < SEITEN; i++) {
            dc.setColor(i == seite ? Graphics.COLOR_WHITE : 0x555555, Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(w - 16, h / 2 - 27 + i * 18, 4);
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
            y = zeile(dc, y, Graphics.FONT_SMALL, "Heute " + KiraHilfen.zahl(l.get("sollHeute"), 0) + " ESS", Graphics.COLOR_WHITE);
            var sp = l.get("sportHeute");
            var zo = l.get("zoneHeute");
            var sz = (sp != null ? sp.toString() : "") + (zo != null ? " " + zo.toString() : "");
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
        y = zeile(dc, y, Graphics.FONT_XTINY, "NÄCHSTE TAGE", 0x9A9A9A);
        y += 8;
        var xs = d.get("naechste");
        if (!(xs instanceof Array) || xs.size() == 0) {
            zeile(dc, (h * 0.42).toNumber(), Graphics.FONT_SMALL, "Kein Plan", 0xBBBBBB);
            return;
        }
        var fh = dc.getFontHeight(Graphics.FONT_TINY);
        var lang = "";
        for (var i = 0; i < xs.size() && i < 4; i++) {
            var x = xs[i];
            if (!(x instanceof Dictionary)) { continue; }
            var dt = x.get("d") != null ? x.get("d").toString() : "";
            var ess = x.get("ess");
            var sp = KiraHilfen.sportDe(x.get("sport"));
            var zo = x.get("zone") != null ? x.get("zone").toString() : "";
            var ruhe = ess == null || ess == 0 || sp.equals("Ruhe");
            var lh = x.get("langH");
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.17, y, Graphics.FONT_TINY, KiraHilfen.wochentag(dt).substring(0, 2), Graphics.TEXT_JUSTIFY_LEFT);
            dc.setColor(ruhe ? 0x777777 : Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.48, y, Graphics.FONT_TINY, ruhe ? "-" : KiraHilfen.zahl(ess, 0), Graphics.TEXT_JUSTIFY_RIGHT);
            dc.setColor(lh != null ? 0x9AD0FF : (ruhe ? 0x777777 : 0xBBBBBB), Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.53, y, Graphics.FONT_TINY, ruhe ? "Ruhe" : sp + (zo.length() > 0 && !zo.equals("Off") ? " " + zo : ""), Graphics.TEXT_JUSTIFY_LEFT);
            if (lh != null && lang.length() == 0) { lang = "Lang " + KiraHilfen.wochentag(dt) + " " + KiraHilfen.stunden(lh); }
            y += fh + 4;
        }
        y += 6;
        zeile(dc, y, Graphics.FONT_XTINY, lang.length() > 0 ? lang : "ESS · Sport · Zone", lang.length() > 0 ? 0x9AD0FF : 0x777777);
    }

    // Seite 4: Endurance Score, Fitness, Garmin-Tag
    function seiteForm(dc as Graphics.Dc, d as Dictionary, h as Number) as Void {
        var w = dc.getWidth();
        var y = (h * 0.11).toNumber();
        y = zeile(dc, y, Graphics.FONT_XTINY, "FORM", 0x9A9A9A);
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
        var t = f.get("tag");
        if (t instanceof Dictionary) {
            var bb = t.get("bbAufwachen") != null ? t.get("bbAufwachen") : t.get("bbMax");
            y = zeile(dc, y, Graphics.FONT_XTINY, "BB " + KiraHilfen.zahl(bb, 0) + " · Stress " + KiraHilfen.zahl(t.get("stress"), 0), 0xBBBBBB);
            if (t.get("schritte") != null) { zeile(dc, y, Graphics.FONT_XTINY, KiraHilfen.zahl(t.get("schritte"), 0) + " Schritte", 0x777777); }
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

    // MENU (lange UP): Kopplung lösen
    function onMenu() as Boolean {
        if (Storage.getValue("tok") == null) { return false; }
        WatchUi.pushView(new WatchUi.Confirmation("Kopplung lösen?"), new KiraLoesen(), WatchUi.SLIDE_IMMEDIATE);
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
