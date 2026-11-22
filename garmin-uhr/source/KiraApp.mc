// Coach Kira – Tagesvotum auf der Garmin-Uhr (fēnix 8 Pro), 21.11.2026.
// Holt den Tagesstand aus dem Cockpit (/api/widget?uhr=1, Geräteschlüssel) – im Hintergrund alle 15 min
// und beim Öffnen. Keine KI auf der Uhr: angezeigt wird das fertige Votum des Rechenkerns.
import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Background;
import Toybox.Communications;
import Toybox.Complications;
import Toybox.Notifications;
import Toybox.Lang;
import Toybox.PersistedContent;
import Toybox.System;
import Toybox.Time;
import Toybox.Time.Gregorian;
import Toybox.WatchUi;

(:background :glance)
const KIRA_BASIS = "https://coachkira.pplx.app/port/5000";

(:background :glance)
class KiraApp extends Application.AppBase {
    function initialize() {
        AppBase.initialize();
    }

    function onStart(state as Dictionary?) as Void {
    }

    function onStop(state as Dictionary?) as Void {
    }

    function getInitialView() {
        hintergrundAn();
        var v = new KiraView();
        return [v, new KiraDelegate(v)];
    }

    (:glance)
    function getGlanceView() {
        return [new KiraGlance()];
    }

    function getServiceDelegate() {
        return [new KiraHintergrund()];
    }

    // Ergebnis des Hintergrunddienstes übernehmen
    function onBackgroundData(data) {
        if (data instanceof Dictionary) {
            if (data.hasKey("votum")) {
                Storage.setValue("d", data);
                Storage.setValue("at", Time.now().value());
                Storage.deleteValue("fehler");
                kiraKomplikation(data);
            } else if (data.hasKey("fehler")) {
                Storage.setValue("fehler", data.get("fehler"));
            }
        }
        WatchUi.requestUpdate();
    }

    function hintergrundAn() as Void {
        if (Storage.getValue("tok") == null) {
            return;
        }
        try {
            if (Background.getTemporalEventRegisteredTime() == null) {
                Background.registerForTemporalEvent(new Time.Duration(15 * 60));
            }
        } catch (e) {
        }
    }
}

// 22.11.2026: Komplikation fürs Zifferblatt (Face It): Votum als Kurzlabel, Soll-Rahmen als Wert
(:background :glance)
function kiraKomplikation(d as Dictionary) as Void {
    try {
        if (!(Toybox has :Complications) || !(Complications has :updateComplication)) { return; }
        var v = d.get("votum");
        var l = d.get("last");
        var lab = "KIRA";
        if (v instanceof Dictionary && v.get("vote") != null) {
            var k = v.get("vote").toString();
            lab = k.equals("QUALITY") ? "QUAL" : k;
        }
        var wert = "-";
        if (l instanceof Dictionary) {
            var lo = l.get("sollMin");
            var hi = l.get("sollMax");
            if (hi != null) { wert = (lo != null ? lo.toNumber().toString() + "-" : "") + hi.toNumber().toString(); }
            else if (l.get("sollHeute") != null) { wert = l.get("sollHeute").toNumber().toString(); }
        }
        try {
            Complications.updateComplication(0, { :value => wert, :shortLabel => lab, :unit => "ESS" });
        } catch (e0) { }
        // 22.11.2026: Komplikation 1 (privat) für das Coach-Kira-Zifferblatt:
        // vote|stufe|lo|hi|ist|datum|ziel|zielTage|readiness
        var t = "";
        var st = v instanceof Dictionary && v.get("stufe") != null ? v.get("stufe").toString() : "";
        var lo2 = "", hi2 = "", ist = "";
        if (l instanceof Dictionary) {
            if (l.get("sollMin") != null) { lo2 = l.get("sollMin").toNumber().toString(); }
            if (l.get("sollMax") != null) { hi2 = l.get("sollMax").toNumber().toString(); }
            else if (l.get("sollHeute") != null) { hi2 = l.get("sollHeute").toNumber().toString(); }
            if (l.get("istHeute") != null) { ist = l.get("istHeute").toNumber().toString(); }
        }
        var z = d.get("ziel");
        var zn = "", zt = "";
        if (z instanceof Dictionary) {
            if (z.get("name") != null) { zn = z.get("name").toString(); }
            if (z.get("tage") != null) { zt = z.get("tage").toNumber().toString(); }
        }
        var m = d.get("morgen");
        var rd = m instanceof Dictionary && m.get("readiness") != null ? m.get("readiness").toNumber().toString() : "";
        // ab 22.11. abends zusätzlich: |plan|sport|zone|ausdauerwert|ausdauer28
        var pl = "", sp = "", zo = "", es = "", es28 = "";
        if (l instanceof Dictionary) {
            if (l.get("planHeute") != null) { pl = l.get("planHeute").toNumber().toString(); }
            if (l.get("sportHeute") != null) { sp = l.get("sportHeute").toString(); }
            if (l.get("zoneHeute") != null) { zo = l.get("zoneHeute").toString(); }
        }
        var fo = d.get("form");
        if (fo instanceof Dictionary) {
            if (fo.get("es") != null) { es = fo.get("es").toNumber().toString(); }
            if (fo.get("es28") != null) { es28 = fo.get("es28").toNumber().toString(); }
        }
        t = (lab.equals("KIRA") ? "" : lab) + "|" + st + "|" + lo2 + "|" + hi2 + "|" + ist + "|" + (d.get("datum") != null ? d.get("datum").toString() : "") + "|" + zn + "|" + zt + "|" + rd
            + "|" + pl + "|" + sp + "|" + zo + "|" + es + "|" + es28
            + "|" + sieben(d, l);
        Complications.updateComplication(1, { :value => t });
    } catch (e) {
    }
}

// 7 Tage Last (rollierend, heute zuletzt) mit Soll-Rahmen je Tag fürs Zeigerzifferblatt:
// "ist,ist,…;lo,lo,…;hi,hi,…" – Rahmen aus der letzten Vollabfrage der App (nur wenn von heute)
(:background :glance)
function sieben(d as Dictionary, l) as String {
    var ist = [] as Array;
    var lo = [] as Array;
    var hi = [] as Array;
    try {
        var u = Storage.getValue("u");
        var ud = Storage.getValue("uDatum");
        var l14 = u instanceof Dictionary ? u.get("l14") : null;
        if (l14 instanceof Dictionary && ud != null && d.get("datum") != null && ud.toString().equals(d.get("datum").toString())) {
            var a = l14.get("ist");
            var b = l14.get("lo");
            var c = l14.get("hi");
            if (a instanceof Array && b instanceof Array && c instanceof Array) {
                var n = a.size();
                for (var i = (n > 7 ? n - 7 : 0); i < n; i++) { ist.add(a[i]); lo.add(b[i]); hi.add(c[i]); }
            }
        }
        if (ist.size() == 0 && l instanceof Dictionary && l.get("woche") instanceof Array) {
            var w = l.get("woche") as Array;
            for (var i = 0; i < w.size(); i++) {
                var x = w[i];
                ist.add(x instanceof Dictionary && x.get("ist") != null ? x.get("ist") : 0);
            }
        }
        // heutigen Wert immer frisch
        if (ist.size() > 0 && l instanceof Dictionary && l.get("istHeute") != null) { ist[ist.size() - 1] = l.get("istHeute"); }
    } catch (e) { }
    return liste(ist) + ";" + liste(lo) + ";" + liste(hi);
}

(:background :glance)
function liste(xs as Array) as String {
    var s = "";
    for (var i = 0; i < xs.size(); i++) {
        var v = xs[i];
        s += (i > 0 ? "," : "") + (v == null ? "" : v.toNumber().toString());
    }
    return s;
}

// Hintergrund: höchstens alle 15 min, kleine Antwort
(:background)
class KiraHintergrund extends System.ServiceDelegate {
    function initialize() {
        ServiceDelegate.initialize();
    }

    function onTemporalEvent() as Void {
        var tok = Storage.getValue("tok");
        if (tok == null) {
            Background.exit(null);
            return;
        }
        Communications.makeWebRequest(
            KIRA_BASIS + "/api/widget?uhr=1",
            null,
            {
                :method => Communications.HTTP_REQUEST_METHOD_GET,
                :headers => { "Authorization" => "Bearer " + tok },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            method(:fertig)
        );
    }

    function fertig(code as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        if (code == 200 && data instanceof Dictionary) {
            kiraKomplikation(data);
            melden(data);
            Background.exit(data);
        } else {
            Background.exit({ "fehler" => code });
        }
    }

    // 22.11.2026: Morgenhinweis – einmal am Tag, sobald das Votum für heute mit Morgenwerten steht
    function melden(d as Dictionary) as Void {
        try {
            var datum = d.get("datum");
            var v = d.get("votum");
            var m = d.get("morgen");
            if (datum == null || !(v instanceof Dictionary) || v.get("vote") == null) { return; }
            if (m instanceof Dictionary && m.get("eingetragen") == false) { return; }
            var gemeldet = Storage.getValue("gemeldet");
            if (gemeldet != null && gemeldet.toString().equals(datum.toString())) { return; }
            var jetzt = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
            var heute = jetzt.year.format("%04d") + "-" + jetzt.month.format("%02d") + "-" + jetzt.day.format("%02d");
            if (!heute.equals(datum.toString()) || jetzt.hour < 5 || jetzt.hour >= 13) { return; }
            var l = d.get("last");
            var soll = "";
            if (l instanceof Dictionary && l.get("sollMax") != null) {
                soll = "Soll " + (l.get("sollMin") != null ? l.get("sollMin").toNumber().toString() + "-" : "bis ") + l.get("sollMax").toNumber().toString() + " ESS";
            }
            var titel = "Kira: " + v.get("label").toString();
            var gr = v.get("gruende");
            var text = gr instanceof Array && gr.size() > 0 ? gr[0].toString() : "";
            Storage.setValue("gemeldet", datum.toString());
            if (Toybox has :Notifications && Notifications has :showNotification) {
                Notifications.showNotification(titel, soll, { :body => text, :dismissPrevious => true });
            } else if (Background has :requestApplicationWake) {
                Background.requestApplicationWake(titel + " · " + soll);
            }
        } catch (e) {
        }
    }
}
