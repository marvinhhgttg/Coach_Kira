// 24.11.2026: frei belegbare Werte für Ringe und Ecken (Einstellungen in der Connect IQ App)
import Toybox.Application;
import Toybox.Complications;
import Toybox.Lang;
import Toybox.Time;
import Toybox.Time.Gregorian;
import Toybox.Weather;

module KiraWert {
    // Ecken: 0 Temperatur, 1 Ausdauerwert, 2 Wind, 3 Plan, 4 ACWR, 5 Tourform, 6 Body Battery, 7 Schlaf,
    //        8 Herzfrequenz, 9 Schritte, 10 Woche, 11 Erholung, 12 Stress, 13 leer
    function einst(k as String, def) {
        try {
            var v = Application.Properties.getValue(k);
            return v == null ? def : v;
        } catch (e) {
            return def;
        }
    }

    function zahlE(k as String, def as Number) as Number {
        var v = einst(k, def);
        return v instanceof Number ? v : def;
    }

    function an(k as String) as Boolean { return einst(k, true) != false; }

    function nativ(typ as Complications.Type) as Number? {
        try {
            var c = Complications.getComplication(new Complications.Id(typ));
            if (c != null && c.value != null) { return c.value.toNumber(); }
        } catch (e) { }
        return null;
    }

    function feld(t as Array<String>, i as Number) as String { return t.size() > i ? t[i] : ""; }

    function zahl(t as Array<String>, i as Number) as Number? {
        var s = feld(t, i);
        return s.length() == 0 ? null : s.toNumber();
    }

    function richtung(grad as Number?) as String {
        if (grad == null) { return ""; }
        var r = ["N", "NO", "O", "SO", "S", "SW", "W", "NW"];
        return r[((grad + 22) % 360) / 45];
    }

    function acwrText(a as Number?) as String {
        if (a == null) { return "-"; }
        return (a / 100).toString() + "," + (a % 100).format("%02d");
    }

    function erholtAb(min as Number?) as String {
        if (min == null) { return "-"; }
        if (min <= 0) { return "OK"; }
        var jetzt = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var i = Gregorian.info(Time.now().add(new Time.Duration(min * 60)), Time.FORMAT_SHORT);
        var wt = ["", "SO", "MO", "DI", "MI", "DO", "FR", "SA"];
        var hm = i.hour.format("%02d") + ":" + i.min.format("%02d");
        if (i.day == jetzt.day) { return hm; }
        return wt[i.day_of_week as Number] + " " + i.hour.format("%02d");
    }

    // [Bezeichnung, Wert] für eine Ecke
    function ecke(art as Number, t as Array<String>, heute as Boolean) as Array<String> {
        if (art == 0 || art == 2) {
            var temp = "-";
            var wind = "-";
            try {
                var wc = Weather.getCurrentConditions();
                if (wc != null) {
                    if (wc.temperature != null) { temp = wc.temperature.toNumber().toString() + "°"; }
                    if (wc.windSpeed != null) { wind = richtung(wc.windBearing) + " " + (wc.windSpeed * 3.6).toNumber().toString(); }
                }
            } catch (e) { }
            return art == 0 ? ["TEMP", temp] : ["WIND", wind];
        }
        if (art == 1) { var es = zahl(t, 12); return ["AUSD.", es == null ? "-" : es.toString()]; }
        if (art == 3) {
            var pl = heute ? zahl(t, 9) : null;
            var sp = heute ? feld(t, 10) : "";
            return ["PLAN", pl == null ? "-" : (pl == 0 ? "Ruhe" : (sp.length() > 0 ? sp + " " : "") + pl.toString())];
        }
        if (art == 4) { return ["ACWR", acwrText(heute ? zahl(t, 18) : null)]; }
        if (art == 5) { var tf = zahl(t, 19); return ["TOURF.", tf == null ? "-" : tf.toString() + "%"]; }
        if (art == 6) { var b = nativ(Complications.COMPLICATION_TYPE_BODY_BATTERY); return ["BB", b == null ? "-" : b.toString()]; }
        if (art == 7) { var s = nativ(Complications.COMPLICATION_TYPE_SLEEP_SCORE); return ["SCHLAF", s == null ? "-" : s.toString()]; }
        if (art == 8) { var hf = nativ(Complications.COMPLICATION_TYPE_HEART_RATE); return ["HF", hf == null ? "-" : hf.toString()]; }
        if (art == 9) {
            var st = nativ(Complications.COMPLICATION_TYPE_STEPS);
            return ["SCHRITTE", st == null ? "-" : (st >= 10000 ? (st / 1000).toString() + "k" : st.toString())];
        }
        if (art == 10) {
            var wi = zahl(t, 15);
            var wb = zahl(t, 16);
            return ["WOCHE", wi == null ? "-" : wi.toString() + (wb != null ? "/" + wb.toString() : "")];
        }
        if (art == 11) { return ["ERHOL.", erholtAb(nativ(Complications.COMPLICATION_TYPE_RECOVERY_TIME))]; }
        if (art == 12) { var sr = nativ(Complications.COMPLICATION_TYPE_STRESS); return ["STRESS", sr == null ? "-" : sr.toString()]; }
        return ["", ""];
    }

    // Ring: [Wert oder null, Skala, Bezeichnung unten]
    // links: 0 Bereitschaft, 1 Body Battery, 2 Schlaf, 3 Tourform · rechts: 10 Tageslast, 11 Woche, 12 ACWR, 13 Tourform
    function ring(art as Number, t as Array<String>, heute as Boolean) as Array {
        if (art == 1) { return [nativ(Complications.COMPLICATION_TYPE_BODY_BATTERY), KiraSkala.bb(), "BB"]; }
        if (art == 2) { return [nativ(Complications.COMPLICATION_TYPE_SLEEP_SCORE), KiraSkala.schlaf(), "SCHLAF"]; }
        if (art == 3 || art == 13) { return [zahl(t, 19), KiraSkala.tourform(), "TOURFORM"]; }
        if (art == 11) {
            var wb = zahl(t, 16);
            return [zahl(t, 15), (wb != null && wb > 0) ? KiraSkala.woche(wb) : [[0, 100], [KiraSkala.GRAU]], wb == null ? "WOCHE" : "/" + wb.toString()];
        }
        if (art == 12) { return [heute ? zahl(t, 18) : null, KiraSkala.acwr(), "ACWR"]; }
        // 0 Bereitschaft
        return [heute ? zahl(t, 8) : null, KiraSkala.tr(), heute && feld(t, 14).length() > 0 ? "TR " + feld(t, 14) : "BEREIT"];
    }

    // Text in der Ringmitte
    function ringText(art as Number, v as Number?) as String {
        if (v == null) { return "-"; }
        if (art == 12) { return acwrText(v); }
        return v.toString();
    }
}
