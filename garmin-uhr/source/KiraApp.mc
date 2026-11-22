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
        Complications.updateComplication(0, { :value => wert, :shortLabel => lab, :units => "ESS" });
    } catch (e) {
    }
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
            KIRA_BASIS + "/api/widget",
            { "uhr" => "1" },
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
