// Coach Kira – Tagesvotum auf der Garmin-Uhr (fēnix 8 Pro), 21.11.2026.
// Holt den Tagesstand aus dem Cockpit (/api/widget?uhr=1, Geräteschlüssel) – im Hintergrund alle 15 min
// und beim Öffnen. Keine KI auf der Uhr: angezeigt wird das fertige Votum des Rechenkerns.
import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Background;
import Toybox.Communications;
import Toybox.Lang;
import Toybox.PersistedContent;
import Toybox.System;
import Toybox.Time;
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
            Background.exit(data);
        } else {
            Background.exit({ "fehler" => code });
        }
    }
}
