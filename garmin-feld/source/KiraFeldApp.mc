// Coach Kira Datenfeld „Kira ESS“ (22.11.2026)
// Zählt während der Aktivität die Sekunden je Herzfrequenzzone und rechnet sie mit deinen
// persönlichen Raten (ESS je Stunde in Zone 1–5, aus deiner Garmin-Historie) in ESS um.
// Vergleich live gegen den Soll-Rahmen des Tages.
// 22.11.2026 abends: Kopplung und Abruf direkt aus dem Datenfeld (ab System 7 erlaubt) statt per
// Hintergrunddienst – kein 5-Minuten-Takt mehr, Wiederholung alle 20–30 s bis es klappt.
import Toybox.Application;
import Toybox.Lang;
import Toybox.WatchUi;

const KIRA_BASIS = "https://coachkira.pplx.app/port/5000";

class KiraFeldApp extends Application.AppBase {
    var ansicht as KiraFeldView? = null;
    function initialize() {
        AppBase.initialize();
    }

    function getInitialView() as [WatchUi.Views] or [WatchUi.Views, WatchUi.InputDelegates] {
        ansicht = new KiraFeldView();
        return [ansicht];
    }

    // 24.11.2026: Einstellungen sofort übernehmen
    function onSettingsChanged() as Void {
        if (ansicht != null) { (ansicht as KiraFeldView).einstLesen(); }
        WatchUi.requestUpdate();
    }
}
