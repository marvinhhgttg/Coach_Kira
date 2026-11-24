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
    // 23.11.2026 Live-Coach: Warnung vor der Obergrenze, Zonenwächter mit Wiederholung, ESS-Verlauf je Minute
    var warnGemeldet as Boolean = false;
    var waechterAb as Float = 300.0;
    var verlauf as Array<Float> = [] as Array<Float>;
    var verlaufMin as Number = -1;
    var startGeprueft as Boolean = false;
    var hinweisBis as Number = 0;
    var hinweis as String = "";
    var fitLive as FitContributor.Field? = null;
    var fitSumme as FitContributor.Field? = null;
    var unterwegs as Boolean = false;
    var naechsterAbruf as Number = 0;
    var seit as Number = 0;
    // 24.11.2026: Einstellungen (Connect IQ App → Coach Kira Feld → Einstellungen)
    var zoneJetzt as Number = 0;
    var timerMs as Number = 0;
    var distM as Float = 0.0;
    var eOben as Number = 0;
    var eMitte as Number = 0;
    var eUnten as Number = 0;
    var vibGrenze as Boolean = true;
    var vibZone as Boolean = true;

    function initialize() {
        DataField.initialize();
        try {
            fitLive = createField("kira_ess", 0, FitContributor.DATA_TYPE_UINT16,
                { :mesgType => FitContributor.MESG_TYPE_RECORD, :units => "ESS" });
            fitSumme = createField("kira_ess_summe", 1, FitContributor.DATA_TYPE_UINT16,
                { :mesgType => FitContributor.MESG_TYPE_SESSION, :units => "ESS" });
        } catch (e) { }
        sportLesen();
        einstLesen();
    }

    function einst(k as String, def) {
        try {
            var v = Application.Properties.getValue(k);
            return v == null ? def : v;
        } catch (e) {
            return def;
        }
    }

    function einstLesen() as Void {
        var v = einst("pOben", 0); eOben = v instanceof Number ? v : 0;
        v = einst("pMitte", 0); eMitte = v instanceof Number ? v : 0;
        v = einst("pUnten", 0); eUnten = v instanceof Number ? v : 0;
        vibGrenze = einst("pVibGrenze", true) != false;
        vibZone = einst("pVibZone", true) != false;
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
        warnGemeldet = false;
        waechterAb = 300.0;
        verlauf = [] as Array<Float>;
        verlaufMin = -1;
        startGeprueft = false;
    }

    // ESS heute schon gespeicherter Einheiten (vom Server, Mac meldet alle 10 min)
    function vorher() as Number {
        var f = feld();
        if (f == null) { return 0; }
        var d = f.get("datum");
        if (d instanceof String && !d.equals(heute())) { return 0; }
        var v = f.get("vorher");
        return v instanceof Number ? v : (v instanceof Float ? v.toNumber() : 0);
    }

    // ESS pro Minute über die letzten bis zu 10 Minuten (null, solange zu wenig Verlauf)
    function rate() as Float? {
        var n = verlauf.size();
        if (n < 4) { return null; }
        var a = n > 11 ? verlauf[n - 11] : verlauf[0];
        var m = n > 11 ? 10 : n - 1;
        var r = (ess - a) / m;
        return r > 0.05 ? r : null;
    }

    function onTimerStart() as Void {
        sportLesen();
    }

    function compute(info as Activity.Info) as Void {
        abrufFaellig();
        // 24.11.2026: Aktivitätswerte für frei belegbare Zeilen
        timerMs = info.timerTime != null ? info.timerTime as Number : 0;
        distM = info.elapsedDistance != null ? (info.elapsedDistance as Float).toFloat() : 0.0;
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
        zoneJetzt = k + 1;
        if (k < 0) { return; }
        sek[k] += dt / 1000.0;
        var r = raten();
        var s = 0.0;
        for (var i = 0; i < 5; i++) { s += sek[i] / 3600.0 * (r[i] as Number).toFloat(); }
        ess = s;
        var tm = info.timerTime != null ? (info.timerTime as Number) / 60000 : 0;
        if (tm != verlaufMin) {
            verlaufMin = tm;
            verlauf.add(ess);
            if (verlauf.size() > 12) { verlauf = verlauf.slice(1, null) as Array<Float>; }
        }
        var e = ess.toNumber();
        if (fitLive != null) { fitLive.setData(e); }
        if (fitSumme != null) { fitSumme.setData(e); }
        pruefen(e);
    }

    function pruefen(e as Number) as Void {
        var ra = rahmen();
        var tag = vorher() + e;
        if (ra != null) {
            // schon vor dem Start erreicht: nicht noch einmal melden
            if (!startGeprueft) {
                startGeprueft = true;
                var v0 = vorher();
                if (ra[0] > 0 && v0 >= ra[0]) { minGemeldet = true; }
                if (v0 > ra[1]) { maxGemeldet = true; warnGemeldet = true; }
            }
            if (!minGemeldet && tag >= ra[0] && ra[0] > 0) {
                minGemeldet = true;
                if (vibGrenze) { melden("Untergrenze erreicht", [new Attention.VibeProfile(60, 300)]); }
            }
            var r = rate();
            if (!warnGemeldet && r != null && tag < ra[1] && (ra[1] - tag) / r <= 5.0) {
                warnGemeldet = true;
                if (vibGrenze) { melden("Obergrenze in ~5 min", [new Attention.VibeProfile(80, 300), new Attention.VibeProfile(0, 200), new Attention.VibeProfile(80, 300)]); }
            }
            if (!maxGemeldet && tag > ra[1]) {
                maxGemeldet = true;
                warnGemeldet = true;
                if (vibGrenze) { melden("Obergrenze " + ra[1].toString(), [new Attention.VibeProfile(100, 500), new Attention.VibeProfile(0, 250),
                    new Attention.VibeProfile(100, 500), new Attention.VibeProfile(0, 250), new Attention.VibeProfile(100, 500)]); }
            }
        }
        // Zonenwächter: Easy/Rest → Zeit ab Zone 3, Train → Zeit ab Zone 4; nach 5 min, dann alle 10 min
        var f = feld();
        if (f != null && rahmen() != null) {
            var vote = f.get("vote");
            var easy = "EASY".equals(vote) || "REST".equals(vote);
            var train = "TRAIN".equals(vote);
            if (vibZone && (easy || train)) {
                var t = easy ? sek[2] + sek[3] + sek[4] : sek[3] + sek[4];
                var schwelle = train && waechterAb < 600.0 ? 600.0 : waechterAb;
                if (t > schwelle) {
                    var min = (t / 60).toNumber();
                    waechterAb = schwelle + 600.0;
                    melden((easy ? (("REST".equals(vote)) ? "Ruhetag: " : "Easy: ") + min.toString() + "' ab Z3" : "Train: " + min.toString() + "' ab Z4"),
                        [new Attention.VibeProfile(80, 400), new Attention.VibeProfile(0, 200), new Attention.VibeProfile(80, 400)]);
                }
            }
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
        if (ra[1] <= 0) { return Graphics.COLOR_WHITE; }
        // 23.11.2026: Garmin-Farbmuster mit Kiras Soll-Rahmen
        return KiraSkala.farbe(KiraSkala.last(ra[0], ra[1]), e);
    }

    // Text für frei belegbare Zeilen. Oben: 1 Zone, 2 Soll, 3 Votum, 4 ACWR, 5 Bereitschaft.
    // Unten: 1 Zone, 2 Soll, 3 ESS/min, 4 ACWR, 5 Bereitschaft, 6 ESS Einheit
    // Aktivität: 100 Timer, 101 Distanz, 102 Uhrzeit
    function aktText(art as Number, mitBez as Boolean) as String {
        if (art == 100) {
            var sk = timerMs / 1000;
            var hh = sk / 3600;
            var t = hh > 0 ? hh.toString() + ":" + ((sk / 60) % 60).format("%02d") + ":" + (sk % 60).format("%02d")
                           : (sk / 60).toString() + ":" + (sk % 60).format("%02d");
            return mitBez ? "Zeit " + t : t;
        }
        if (art == 101) {
            var km = distM / 1000.0;
            var t = km < 10.0 ? km.format("%.2f") : km.format("%.1f");
            var i = t.find(".");
            if (i != null && mitBez) { t = t.substring(0, i) + "," + t.substring(i + 1, t.length()); }   // Zahlenschrift ohne Komma
            return mitBez ? t + " km" : t;
        }
        var c = System.getClockTime();
        return c.hour.format("%02d") + ":" + c.min.format("%02d");
    }

    function zeilenText(art as Number, einheit as Number, e as Number, ra as Array?, istOben as Boolean) as String {
        if (istOben && art >= 6 && art <= 8) { return aktText(art + 94, true); }
        if (!istOben && art >= 7 && art <= 9) { return aktText(art + 93, true); }
        var f = feld();
        if (!istOben && art == 3) {
            var r = rate();
            return r != null ? "ESS/min " + r.format("%.1f") : "ESS/min -";
        }
        if (!istOben && art == 6) { return "Einheit " + einheit.toString(); }
        if (istOben && art == 3) {
            var lb = f != null ? f.get("label") : null;
            return "KIRA " + (lb != null ? lb.toString().toUpper() : "-");
        }
        if (art == 1) {
            var t3 = (sek[2] + sek[3] + sek[4]) / 60;
            return "Zone " + (zoneJetzt > 0 ? zoneJetzt.toString() : "-") + " · Z3+ " + t3.toNumber().toString() + "'";
        }
        if (art == 2) { return ra != null ? "Soll " + ra[0].toString() + "-" + ra[1].toString() : "Soll -"; }
        if (art == 4) {
            var a = f != null ? f.get("acwr") : null;
            return "ACWR " + (a != null ? a.toFloat().format("%.2f") : "-");
        }
        if (art == 5) {
            var b = f != null ? f.get("bereit") : null;
            return "Bereit " + (b != null ? b.toNumber().toString() : "-");
        }
        return "";
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
        var einheit = ess.toNumber();
        var vor = vorher();
        var e = vor + einheit;              // Tagesstand: schon gespeicherte Einheiten + laufende Einheit
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
        } else if (e > ra[1]) {
            unten = "+" + (e - ra[1]).toString() + " über " + ra[1].toString();
        } else {
            // Restbudget bis zur Obergrenze und Zeit bis dort beim aktuellen Tempo
            var r = rate();
            unten = "noch " + (ra[1] - e).toString() + (r != null ? " · ~" + ((ra[1] - e) / r).toNumber().toString() + "'" : " bis " + ra[1].toString());
        }
        if (Storage.getValue("tok") != null && ra != null && eUnten > 0) { unten = zeilenText(eUnten, einheit, e, ra, false); }
        if (System.getTimer() < hinweisBis) { unten = hinweis; }
        var c = farbe(e, ra);
        if (c == Graphics.COLOR_WHITE) { c = fg; }
        // große Zahl nach Einstellung; Farbe bleibt am Tagesstand
        var gross = e.toString();
        if (eMitte == 1) { gross = einheit.toString(); }
        else if (eMitte == 2 && ra != null) { gross = (e > ra[1] ? "+" + (e - ra[1]).toString() : (ra[1] - e).toString()); }
        else if (eMitte >= 3 && eMitte <= 5) { gross = aktText(eMitte + 97, false); c = fg; }
        var oben = vor > 0 ? "KIRA TAG · EINH. " + einheit.toString() : "KIRA ESS";
        if (eMitte == 1) { oben = "KIRA EINHEIT · TAG " + e.toString(); }
        else if (eMitte == 2) { oben = "KIRA BIS OBERGRENZE"; }
        else if (eMitte == 3) { oben = "TIMER · KIRA " + e.toString(); }
        else if (eMitte == 4) { oben = "DISTANZ KM · KIRA " + e.toString(); }
        else if (eMitte == 5) { oben = "UHRZEIT · KIRA " + e.toString(); }
        if (eOben > 0) { oben = zeilenText(eOben, einheit, e, ra, true); }
        if (h < 90) {
            // kleines Feld: Wert + Rahmen nebeneinander
            dc.setColor(c, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h / 2 - 4, Graphics.FONT_NUMBER_MILD, gross, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            dc.setColor(grau, Graphics.COLOR_TRANSPARENT);
            dc.drawText(cx, h - 2 - dc.getFontHeight(Graphics.FONT_XTINY), Graphics.FONT_XTINY, unten, Graphics.TEXT_JUSTIFY_CENTER);
            return;
        }
        dc.setColor(grau, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, (h * 0.06).toNumber(), Graphics.FONT_XTINY, oben, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(c, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h / 2, h > 160 ? Graphics.FONT_NUMBER_HOT : Graphics.FONT_NUMBER_MEDIUM, gross, Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(System.getTimer() < hinweisBis ? c : grau, Graphics.COLOR_TRANSPARENT);
        dc.drawText(cx, h - (h * 0.06).toNumber() - dc.getFontHeight(Graphics.FONT_TINY), Graphics.FONT_TINY, unten, Graphics.TEXT_JUSTIFY_CENTER);
        // Fortschrittsbalken bis zur Obergrenze
        if (ra != null && ra[1] > 0 && h > 120) {
            var bw = (w * 0.5).toNumber();
            var x = cx - bw / 2;
            var y = (h * 0.06).toNumber() + dc.getFontHeight(Graphics.FONT_XTINY) + 4;
            // 23.11.2026: Garmin-Balken (Segmente nach Soll-Rahmen, Punkt am Ist)
            KiraSkala.balken(dc, x, y, bw, 6, e, KiraSkala.last(ra[0], ra[1]), bg);
        }
    }
}
