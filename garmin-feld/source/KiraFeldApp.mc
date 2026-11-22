// Coach Kira Datenfeld „Kira ESS“ (22.11.2026)
// Zählt während der Aktivität die Sekunden je Herzfrequenzzone und rechnet sie mit deinen
// persönlichen Raten (ESS je Stunde in Zone 1–5, aus deiner Garmin-Historie) in ESS um.
// Vergleich live gegen den Soll-Rahmen des Tages. Kopplung und Daten laufen im Hintergrund
// (alle 5 min, solange das Datenfeld geladen ist).
import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Background;
import Toybox.Lang;
import Toybox.System;
import Toybox.Time;
import Toybox.WatchUi;

(:background)
const KIRA_BASIS = "https://coachkira.pplx.app/port/5000";

(:background)
class KiraFeldApp extends Application.AppBase {
    var ansicht as KiraFeldView? = null;

    function initialize() {
        AppBase.initialize();
    }

    function getInitialView() as [WatchUi.Views] or [WatchUi.Views, WatchUi.InputDelegates] {
        // Sofort abrufen statt erst nach 5 min: einmaliger Termin (Moment), frühestens 5 min nach dem letzten Lauf
        var naechst = Time.now();
        var zuletzt = Background.getLastTemporalEventTime();
        if (zuletzt != null) {
            var frueh = zuletzt.add(new Time.Duration(5 * 60 + 2));
            if (frueh.greaterThan(naechst)) { naechst = frueh; }
        }
        planen(naechst);
        ansicht = new KiraFeldView();
        return [ansicht];
    }

    function planen(m as Time.Moment) as Void {
        try {
            Background.registerForTemporalEvent(m);
            Storage.setValue("nx", m.value());
        } catch (e) {
            try { Background.registerForTemporalEvent(new Time.Duration(5 * 60)); } catch (e2) { }
        }
    }

    function getServiceDelegate() as [System.ServiceDelegate] {
        return [new KiraFeldDienst()];
    }

    function onBackgroundData(data as Application.PersistableType) as Void {
        Storage.setValue("vz", Time.now().value());
        // nächsten Abruf in 5 min (Kopplung läuft, bis ein Schlüssel da ist; danach Rahmen aktuell halten)
        planen(Time.now().add(new Time.Duration(5 * 60 + 2)));
        if (!(data instanceof Dictionary)) { return; }
        if (data.get("token") != null) {
            Storage.setValue("tok", data.get("token"));
            Storage.deleteValue("aid");
            Storage.deleteValue("code");
        }
        if (data.get("id") != null) {
            Storage.setValue("aid", data.get("id"));
            Storage.setValue("code", data.get("code"));
        }
        if (data.get("neu") != null) {
            Storage.deleteValue("aid");
            Storage.deleteValue("code");
        }
        if (data.get("weg") != null) {
            Storage.deleteValue("tok");
            Storage.deleteValue("f");
        }
        var f = data.get("feld");
        if (f instanceof Dictionary) { Storage.setValue("f", f); }
        var rc = data.get("rc");
        Storage.setValue("rc", rc instanceof Number ? rc : null);
        WatchUi.requestUpdate();
    }
}

(:background)
class KiraFeldDienst extends System.ServiceDelegate {
    var tokNeu as String? = null;

    function initialize() {
        ServiceDelegate.initialize();
    }

    function onTemporalEvent() as Void {
        var tok = Storage.getValue("tok");
        if (tok instanceof String) {
            feldHolen(tok);
            return;
        }
        var aid = Storage.getValue("aid");
        if (aid instanceof String) {
            Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/status", { "id" => aid },
                { :method => Communications.HTTP_REQUEST_METHOD_GET, :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON },
                method(:statusFertig));
        } else {
            Communications.makeWebRequest(KIRA_BASIS + "/api/widget/uhr/anfrage", { "geraet" => "fenix Datenfeld" },
                { :method => Communications.HTTP_REQUEST_METHOD_POST,
                  :headers => { "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON },
                  :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON },
                method(:anfrageFertig));
        }
    }

    function feldHolen(tok as String) as Void {
        Communications.makeWebRequest(KIRA_BASIS + "/api/widget/feld", null,
            { :method => Communications.HTTP_REQUEST_METHOD_GET,
              :headers => { "Authorization" => "Bearer " + tok },
              :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON },
            method(:feldFertig));
    }

    function anfrageFertig(rc as Number, data as Dictionary or String or Null) as Void {
        if (rc == 200 && data instanceof Dictionary && data.get("id") != null) {
            Background.exit({ "id" => data.get("id").toString(), "code" => data.get("code").toString() });
        } else {
            Background.exit({ "rc" => rc });
        }
    }

    function statusFertig(rc as Number, data as Dictionary or String or Null) as Void {
        if (rc == 200 && data instanceof Dictionary && data.get("token") != null) {
            tokNeu = data.get("token").toString();
            feldHolen(tokNeu);
        } else if (rc == 200) {
            Background.exit({ "rc" => null });
        } else if (rc == 404) {
            Background.exit({ "neu" => true });
        } else {
            Background.exit({ "rc" => rc });
        }
    }

    function feldFertig(rc as Number, data as Dictionary or String or Null) as Void {
        var r = {} as Dictionary<String, Application.PropertyValueType>;
        if (tokNeu != null) { r.put("token", tokNeu); }
        if (rc == 200 && data instanceof Dictionary) {
            data.remove("n");
            data.remove("ok");
            r.put("feld", data as Dictionary<String, Application.PropertyValueType>);
        } else if (rc == 401 && tokNeu == null) {
            r.put("weg", true);
        } else {
            r.put("rc", rc);
        }
        Background.exit(r);
    }
}
