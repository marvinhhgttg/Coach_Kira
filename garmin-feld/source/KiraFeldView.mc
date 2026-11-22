import Toybox.Activity;
import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Attention;
import Toybox.Communications;
import Toybox.FitContributor;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.System;
import Toybox.UserProfile;
import Toybox.WatchUi;

const RATEN_START = [20, 45, 80, 130, 190];

class KiraFeldView extends WatchUi.DataField {
    var sek as Array<Float> = [0.0, 0.0, 0.0, 0.0, 0.0];
    var ess as Float = 0.0;
    var zuletzt as Number? = null;
    var zonen as Array<Number>? = null;
    var gruppe as String = "sonst";
    var maxGemeldet as Boolean = false;
    var minGemeldet as Boolean = false;
    var easyGemeldet as Boolean = false;
    var hinweisBis as Number = 0;
    var hinweis as String = "";
    var fitLive as FitContributor.Field? = null;
    var fitSumme as FitContributor.Field? = null;
    var unterwegs as Boolean = false;
    var naechsterAbruf as Number = 0;
    var seit as Number = 0;

    function initialize() {
        DataField.initialize();
        try {
            fitLive = createField("kira_ess", 0, FitContributor.DATA_TYPE_UINT16,
                { :mesgType => FitContributor.MESG_TYPE_RECORD, :units => "ESS" });
            fitSumme = createField("kira_ess_summe", 1, FitContributor.DATA_TYPE_UINT16,
                { :mesgType => FitContributor.MESG_TYPE_SESSION, :units => "ESS" });
        } catch (e) { }
        sportLesen();
    }

    function sportLesen() as Void {
        var sp = null;
        try {
            var p = Activity.getProfileInfo();
            if (p != null) { sp = p.sport; }
        } catch (e) { }
        if (sp == null) { sp = UserProfile.getCurrentSport(); }
        if (sp == Activity.SPORT_CYCLING) { gruppe = "rad"; }
        else if (sp == Activity.SPORT_RUNNING) { gruppe = "lauf"; }
        else if (sp == Activity.SPORT_HIKING || sp == Activity.SPORT_WALKING || sp == Activity.SPORT_MOUNTAINEERING) { gruppe = "berg"; }
        else if (sp == Activity.SPORT_CROSS_COUNTRY_SKIING || sp == Activity.SPORT_ALPINE_SKIING || sp == Activity.SPORT_SNOWBOARDING) { gruppe = "ski"; }
        else { gruppe = "sonst"; }
        try {
            var z = UserProfile.getHeartRateZones(UserProfile.getCurrentSport());
            if (z != null && z.size() >= 6) { zonen = z; }
        } catch (e) { }
    }

    function feld() as Dictionary? {
        var f = Storage.getValue("f");
        return f instanceof Dictionary ? f : null;
    }

    function raten() as Array {
        var f = feld();
        if (f != null) {
            var r = f.get("raten");
            if (r instanceof Dictionary) {
                var g = r.get(gruppe);
                if (g instanceof Array && g.size() >= 5) { return g; }
            }
        }
        return RATEN_START;
    }

    // Rahmen nur, wenn er von heute ist
    function rahmen() as Array? {
        var f = feld();
        if (f == null || f.get("min") == null || f.get("max") == null) { return null; }
        var d = f.get("datum");
        if (d instanceof String && !d.equals(heute())) { return null; }
        return [(f.get("min") as Number).toNumber(), (f.get("max") as Number).toNumber()];
    }

    function heute() as String {
        var g = Toybox.Time.Gregorian.info(Toybox.Time.now(), Toybox.Time.FORMAT_SHORT);
        return g.year.format("%04d") + "-" + (g.month as Number).format("%02d") + "-" + g.day.format("%02d");
    }

    function fehlerText(rc as Number) as String {
        if (rc == -104) { return "Connect-App öffnen"; }
        if (rc < 0) { return "Netz " + rc.toString(); }
        return "Server " + rc.toString();
    }

    // ---- Netz (direkt aus dem Datenfeld) ----
    function abrufFaellig() as Void {
        // hängende Anfrage nach 45 s verwerfen
        if (unterwegs && System.getTimer() - seit > 45000) { unterwegs = false; Storage.setValue("rc", -300); }
        if (unterwegs || System.getTimer() < naechsterAbruf) { return; }
        if (!System.getDeviceSettings().phoneConnected) {
            Storage.setValue("rc", -104);
            naechsterAbruf = System.getTimer() + 20000;
            return;
        }
        unterwegs = true;
        seit = System.getTimer();
        try {
            var tok = Storage.getValue("tok");
            var aid = Storage.getValue("aid");
            if (tok instanceof String) {
                Communications.makeWebRequest(KIRA_BASIS + "/api/widget/feld", null,
                    { :method => Communications.HTTP_REQUEST_METHOD_GET,
                      :headers => { "Authorization" => "Bearer " + tok },
                      :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON },
                    method(:feldFertig));
            } else if (aid instanceof String) {
                Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/status?id=" + aid, null,
                    { :method => Communications.HTTP_REQUEST_METHOD_GET, :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON },
                    method(:statusFertig));
            } else {
                Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/anfrage", { "geraet" => "fenix Datenfeld" },
                    { :method => Communications.HTTP_REQUEST_METHOD_POST,
                      :headers => { "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON },
                      :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON },
                    method(:anfrageFertig));
            }
        } catch (e) {
            unterwegs = false;
            Storage.setValue("rc", -1);
            naechsterAbruf = System.getTimer() + 30000;
        }
    }

    function fertig(rc as Number?, warte as Number) as Void {
        unterwegs = false;
        Storage.setValue("rc", rc);
        naechsterAbruf = System.getTimer() + warte;
        WatchUi.requestUpdate();
    }

    function anfrageFertig(rc as Number, data as Dictionary or String or Null) as Void {
        if (rc == 200 && data instanceof Dictionary && data.get("id") != null) {
            Storage.setValue("aid", data.get("id").toString());
            Storage.setValue("code", data.get("code").toString());
            fertig(null, 10000);
        } else {
            fertig(rc, 30000);
        }
    }

    function statusFertig(rc as Number, data as Dictionary or String or Null) as Void {
        if (rc == 200 && data instanceof Dictionary && data.get("token") != null) {
            Storage.setValue("tok", data.get("token").toString());
            Storage.deleteValue("aid");
            Storage.deleteValue("code");
            fertig(null, 0);                       // sofort Rahmen holen
        } else if (rc == 200) {
            fertig(null, 10000);                   // wartet auf Bestätigung im Cockpit
        } else if (rc == 404) {
            Storage.deleteValue("aid");            // abgelaufen → neuer Code
            Storage.deleteValue("code");
            fertig(null, 0);
        } else {
            fertig(rc, 30000);
        }
    }

    function feldFertig(rc as Number, data as Dictionary or String or Null) as Void {
        if (rc == 200 && data instanceof Dictionary) {
            data.remove("n");
            data.remove("ok");
            Storage.setValue("f", data as Dictionary<String, Application.PropertyValueType>);
            fertig(null, 15 * 60 * 1000);          // Rahmen alle 15 min auffrischen
        } else if (rc == 401) {
            Storage.deleteValue("tok");            // widerrufen → neu koppeln
            Storage.deleteValue("f");
            fertig(null, 0);
        } else {
            fertig(rc, 30000);
        }
    }

    function onTimerReset() as Void {
        sek = [0.0, 0.0, 0.0, 0.0, 0.0];
        ess = 0.0;
        maxGemeldet = false;
        minGemeldet = false;
        easyGemeldet = false;
    }

    function onTimerStart() as Void {
        sportLesen();
    }

    function compute(info as Activity.Info) as Void {
        abrufFaellig();
        var jetzt = System.getTimer();
        var dt = zuletzt == null ? 0 : jetzt - zuletzt;
        zuletzt = jetzt;
        if (dt < 0 || dt > 5000) { dt = 1000; }
        var hr = info.currentHeartRate;
        if (info.timerState != Activity.TIMER_STATE_ON || hr == null || zonen == null) { return; }
        var z = zonen as Array<Number>;
        var k = -1;
        if (hr < z[0]) { k = -1; }
        else if (hr <= z[1]) { k = 0; }
        else if (hr <= z[2]) { k = 1; }
        else if (hr <= z[3]) { k = 2; }
        else if (hr <= z[4]) { k = 3; }
        else { k = 4; }
        if (k < 0) { return; }
        sek[k] += dt / 1000.0;
        var r = raten();
        var s = 0.0;
        for (var i = 0; i < 5; i++) { s += sek[i] / 3600.0 * (r[i] as Number).toFloat(); }
        ess = s;
        var e = ess.toNumber();
        if (fitLive != null) { fitLive.setData(e); }
        if (fitSumme != null) { fitSumme.setData(e); }
        pruefen(e);
    }

    function pruefen(e as Number) as Void {
        var ra = rahmen();
        if (ra != null) {
            if (!minGemeldet && e >= ra[0] && ra[0] > 0) {
                minGemeldet = true;
                melden("Rahmen erreicht", [new Attention.VibeProfile(60, 300)]);
            }
            if (!maxGemeldet && e > ra[1]) {
                maxGemeldet = true;
                melden("Obergrenze " + ra[1].toString(), [new Attention.VibeProfile(100, 500), new Attention.VibeProfile(0, 250),
                    new Attention.VibeProfile(100, 500), new Attention.VibeProfile(0, 250), new Attention.VibeProfile(100, 500)]);
            }
        }
        var f = feld();
        if (!easyGemeldet && f != null && "EASY".equals(f.get("vote")) && rahmen() != null && sek[2] + sek[3] + sek[4] > 300.0) {
            easyGemeldet = true;
            melden("Easy: locker bleiben", [new Attention.VibeProfile(80, 400), new Attention.VibeProfile(0, 200), new Attention.VibeProfile(80, 400)]);
        }
    }

    function melden(text as String, vibe as Array<Attention.VibeProfile>) as Void {
        hinweis = text;
        hinweisBis = System.getTimer() + 15000;
        if (Attention has :vibrate) {
            try { Attention.vibrate(vibe); } catch (e) { }
        }
    }

    function farbe(e as Number, ra as Array?) as Number {
        if (ra == null) { return Graphics.COLOR_WHITE; }
        if (e > ra[1] * 1.15) { return 0xE74C3C; }
        if (e > ra[1]) { return 0xE67E22; }
        if (e >= ra[0]) { return 0x2ECC71; }
        return 0x5DADE2;
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        var bg = getBackgroundColor();
        var fg = bg == Graphics.COLOR_BLACK ? Graphics.COLOR_WHITE : Graphics.COLOR_BLACK;
        var grau = bg == Graphics.COLOR_BLACK ? 0x9A9A9A : 0x555555;
        dc.setColor(fg, bg);
        dc.clear();
        var w = dc.getWidth();
        var h = dc.getHeight();
        var cx = w / 2;
        var e = ess.toNumber();
        var ra = rahmen();
        var unten = "";
        if (Storage.getValue("tok") == null) {
            var c = Storage.getValue("code");
            var rc = Storage.getValue("rc");
            // Code und ggf. Fehler im Wechsel (alle 3 s)
            if (c != null && (!(rc instanceof Number) || (System.getTimer() / 3000) % 2 == 0)) { unten = "Code " + c.toString(); }
            else if (rc instanceof Number) { unten = fehlerText(rc); }
            else { unten = "koppelt ..."; }
        } else if (ra == null) {
            var rc = Storage.getValue("rc");
            unten = rc instanceof Number ? fehlerText(rc) : "Rahmen lädt";
        } else {
            unten = "/ " + ra[0].toString() + "-" + ra[1].toString();
        }
        if (System.getTimer() < hinweisBis) { unten = hinweis; }
        var c = farbe(e, ra);
        if (c == Graphics.COLOR_WHITE) { c = fg; }
        if (h < 90) {
            // kleines Feld: Wert + Rahmen nebeneinander
            dc.setColor(c, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h / 2 - 4, Graphics.FONT_NUMBER_MILD, e.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            dc.setColor(grau, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h - 2 - dc.getFontHeight(Graphics.FONT_XTINY), Graphics.FONT_XTINY, unten, Graphics.TEXT_JUSTIFY_CENTER);
            return;
        }
        dc.setColor(grau, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.06).toNumber(), Graphics.FONT_XTINY, "KIRA ESS", Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(c, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h / 2, h > 160 ? Graphics.FONT_NUMBER_HOT : Graphics.FONT_NUMBER_MEDIUM, e.toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(System.getTimer() < hinweisBis ? c : grau, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h - (h * 0.06).toNumber() - dc.getFontHeight(Graphics.FONT_TINY), Graphics.FONT_TINY, unten, Graphics.TEXT_JUSTIFY_CENTER);
        // Fortschrittsbalken bis zur Obergrenze
        if (ra != null && ra[1] > 0 && h > 120) {
            var bw = (w * 0.5).toNumber();
            var x = cx - bw / 2;
            var y = (h * 0.06).toNumber() + dc.getFontHeight(Graphics.FONT_XTINY) + 4;
            dc.setColor(0x333333, Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x, y, bw, 6);
            var lo = (bw * ra[0] / ra[1]).toNumber();
            dc.setColor(0x1E5631, Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x + lo, y, bw - lo, 6);
            var p = e >= ra[1] ? bw : (bw * e / ra[1]).toNumber();
            dc.setColor(c, Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x, y, p, 6);
        }
    }
}
