import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Communications;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.PersistedContent;
import Toybox.System;
import Toybox.Time;
import Toybox.Timer;
import Toybox.WatchUi;

// Laden und Koppeln im Vordergrund
class KiraNetz {
    static var status as String? = null;   // "laedt", "koppelt", Fehlertext
    static var code as String? = null;     // Kopplungscode zum Anzeigen
    static var anfrageId as String? = null;
    static var timer as Timer.Timer? = null;
    static var versuche as Number = 0;

    static function laden() as Void {
        var tok = Storage.getValue("tok");
        if (tok == null) { return; }
        status = "laedt";
        WatchUi.requestUpdate();
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget", { "uhr" => "1", "frisch" => "0" },
            {
                :method => Communications.HTTP_REQUEST_METHOD_GET,
                :headers => { "Authorization" => "Bearer " + tok },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            new Lang.Method(KiraNetz, :geladen));
    }

    static function geladen(code as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        if (code == 200 && data instanceof Dictionary) {
            Storage.setValue("d", data);
            Storage.setValue("at", Time.now().value());
            Storage.deleteValue("fehler");
            status = null;
        } else if (code == 401) {
            Storage.deleteValue("tok");
            status = "Kopplung im Cockpit widerrufen";
        } else {
            status = code < 0 ? "Kein Handy/Netz (" + code + ")" : "Server " + code;
        }
        WatchUi.requestUpdate();
    }

    static function koppeln() as Void {
        status = "koppelt";
        code = null;
        WatchUi.requestUpdate();
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/anfrage", { "geraet" => "fenix 8 Pro" },
            {
                :method => Communications.HTTP_REQUEST_METHOD_POST,
                :headers => { "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            new Lang.Method(KiraNetz, :angefragt));
    }

    static function angefragt(rc as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        if (rc == 200 && data instanceof Dictionary && data.get("id") != null) {
            anfrageId = data.get("id");
            code = data.get("code");
            versuche = 0;
            if (timer == null) { timer = new Timer.Timer(); }
            timer.start(new Lang.Method(KiraNetz, :abfragen), 4000, true);
        } else {
            status = rc < 0 ? "Kein Handy/Netz (" + rc + ")" : "Kopplung: Server " + rc;
        }
        WatchUi.requestUpdate();
    }

    static function abfragen() as Void {
        versuche += 1;
        if (versuche > 150 || anfrageId == null) {   // 10 min
            stop("Code abgelaufen – START für neuen Code");
            return;
        }
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/status", { "id" => anfrageId },
            {
                :method => Communications.HTTP_REQUEST_METHOD_GET,
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            new Lang.Method(KiraNetz, :abgefragt));
    }

    static function abgefragt(rc as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        if (rc == 200 && data instanceof Dictionary && data.get("token") != null) {
            Storage.setValue("tok", data.get("token"));
            stop(null);
            code = null;
            var app = Application.getApp() as KiraApp;
            app.hintergrundAn();
            laden();
        } else if (rc == 404) {
            stop("Code abgelaufen – START für neuen Code");
        }
    }

    static function stop(meldung as String?) as Void {
        if (timer != null) { timer.stop(); }
        anfrageId = null;
        status = meldung;
        WatchUi.requestUpdate();
    }
}

class KiraView extends WatchUi.View {
    function initialize() {
        View.initialize();
    }

    function onShow() as Void {
        if (Storage.getValue("tok") != null) {
            KiraNetz.laden();
        }
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        var w = dc.getWidth();
        var h = dc.getHeight();
        var cx = w / 2;
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_BLACK);
        dc.clear();

        if (Storage.getValue("tok") == null) {
            zeigeKopplung(dc, w, h);
            return;
        }
        var d = KiraHilfen.daten();
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h * 0.08, Graphics.FONT_XTINY, "COACH KIRA", Graphics.TEXT_JUSTIFY_CENTER);
        if (!(d instanceof Dictionary)) {
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h / 2, Graphics.FONT_SMALL, KiraNetz.status != null && !KiraNetz.status.equals("laedt") ? KiraNetz.status : "Lädt …",
                Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            return;
        }
        var heute = KiraHilfen.vonHeute(d);
        var farbe = heute ? KiraHilfen.farbe(KiraHilfen.feld(d, "votum", "stufe")) : 0x888888;

        // Votum
        var label = KiraHilfen.feld(d, "votum", "label");
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.fillCircle(cx, h * 0.205, 13);
        dc.drawText(cx, h * 0.255, Graphics.FONT_LARGE, label == null ? "–" : label.toUpper(), Graphics.TEXT_JUSTIFY_CENTER);
        var txt = KiraHilfen.feld(d, "votum", "text");
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h * 0.37, Graphics.FONT_XTINY, (txt == null ? "" : txt) + (heute ? "" : " · Stand gestern"), Graphics.TEXT_JUSTIFY_CENTER);
        var gr = KiraHilfen.feld(d, "votum", "gruende");
        if (gr instanceof Array && gr.size() > 0) {
            var g = Graphics.fitTextToArea(gr[0].toString(), Graphics.FONT_XTINY, w * 0.74, h * 0.11, true);
            dc.setColor(0xBBBBBB, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h * 0.43, Graphics.FONT_XTINY, g, Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Morgenwerte
        var m = d.get("morgen");
        var y = h * 0.56;
        if (m instanceof Dictionary) {
            if (m.get("eingetragen") == false) {
                dc.setColor(0xE67E22, Graphics.COLOR_TRANSPARENT);
                dc.drawText(cx, y, Graphics.FONT_XTINY, "Morgenwerte fehlen", Graphics.TEXT_JUSTIFY_CENTER);
            } else {
                wert(dc, w * 0.22, y, "Ready", KiraHilfen.zahl(m.get("readiness"), 0));
                wert(dc, w * 0.41, y, "Schlaf", KiraHilfen.zahl(m.get("schlaf"), 0));
                wert(dc, w * 0.59, y, "HRV", KiraHilfen.zahl(m.get("hrv"), 0));
                wert(dc, w * 0.78, y, "RHR", KiraHilfen.zahl(m.get("rhr"), 0));
            }
        }

        // Last heute
        var l = d.get("last");
        if (l instanceof Dictionary) {
            var soll = "Soll " + KiraHilfen.zahl(l.get("sollHeute"), 0) + " ESS";
            var sp = l.get("sportHeute");
            var zo = l.get("zoneHeute");
            if (sp != null) { soll += " · " + sp; }
            if (zo != null) { soll += " " + zo; }
            if (l.get("erledigt") == true) { soll += " ✓"; }
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h * 0.70, Graphics.FONT_TINY, soll, Graphics.TEXT_JUSTIFY_CENTER);
            var band = l.get("acwrBand");
            dc.setColor(0xBBBBBB, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h * 0.775, Graphics.FONT_XTINY, "ACWR " + KiraHilfen.zahl(l.get("acwr"), 2) + (band != null ? " (" + band + ")" : ""), Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Ziel
        var z = d.get("ziel");
        if (z instanceof Dictionary && z.get("name") != null) {
            var zt = z.get("name") + " in " + KiraHilfen.zahl(z.get("tage"), 0) + " T";
            if (z.get("tourform") != null) { zt += " · Form " + KiraHilfen.zahl(z.get("tourform"), 0) + " %"; }
            dc.setColor(0x9AD0FF, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h * 0.83, Graphics.FONT_XTINY, zt, Graphics.TEXT_JUSTIFY_CENTER);
        }

        // Stand
        var st = KiraNetz.status;
        var unten = st == null ? "Stand " + KiraHilfen.uhrzeit(Storage.getValue("at")) : st.equals("laedt") ? "aktualisiert …" : st;
        dc.setColor(st != null && !st.equals("laedt") ? 0xE67E22 : 0x777777, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h * 0.895, Graphics.FONT_XTINY, unten, Graphics.TEXT_JUSTIFY_CENTER);
    }

    function wert(dc as Graphics.Dc, x, y, titel as String, v as String) as Void {
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y, Graphics.FONT_XTINY, titel, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x, y + 30, Graphics.FONT_SMALL, v, Graphics.TEXT_JUSTIFY_CENTER);
    }

    function zeigeKopplung(dc as Graphics.Dc, w, h) as Void {
        var cx = w / 2;
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h * 0.12, Graphics.FONT_XTINY, "COACH KIRA", Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        if (KiraNetz.code != null) {
            dc.drawText(cx, h * 0.25, Graphics.FONT_SMALL, "Code im Cockpit eingeben", Graphics.TEXT_JUSTIFY_CENTER);
            dc.drawText(cx, h * 0.33, Graphics.FONT_XTINY, "Technik › Garmin-Uhr", Graphics.TEXT_JUSTIFY_CENTER);
            dc.setColor(0x2ECC71, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h * 0.45, Graphics.FONT_NUMBER_MEDIUM, KiraNetz.code, Graphics.TEXT_JUSTIFY_CENTER);
            dc.setColor(0x777777, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h * 0.75, Graphics.FONT_XTINY, "wartet auf Bestätigung …", Graphics.TEXT_JUSTIFY_CENTER);
        } else {
            dc.drawText(cx, h * 0.35, Graphics.FONT_SMALL, "Mit Cockpit koppeln", Graphics.TEXT_JUSTIFY_CENTER);
            dc.drawText(cx, h * 0.47, Graphics.FONT_XTINY, KiraNetz.status != null && KiraNetz.status.equals("koppelt") ? "fordert Code an …" : "START drücken", Graphics.TEXT_JUSTIFY_CENTER);
            if (KiraNetz.status != null && !KiraNetz.status.equals("koppelt")) {
                dc.setColor(0xE67E22, Graphics.COLOR_TRANSPARENT);
                dc.drawText(cx, h * 0.62, Graphics.FONT_XTINY, KiraNetz.status, Graphics.TEXT_JUSTIFY_CENTER);
            }
        }
    }
}

class KiraDelegate extends WatchUi.BehaviorDelegate {
    function initialize() {
        BehaviorDelegate.initialize();
    }

    // START: koppeln bzw. neu laden
    function onSelect() as Boolean {
        if (Storage.getValue("tok") == null) {
            KiraNetz.koppeln();
        } else {
            KiraNetz.laden();
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
