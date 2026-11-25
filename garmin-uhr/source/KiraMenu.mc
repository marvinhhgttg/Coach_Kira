// 22.11.2026: Menü (MENU bzw. UP lang): Befinden eintragen, neu laden, Kopplung lösen
import Toybox.Application.Storage;
import Toybox.Lang;
import Toybox.WatchUi;

function kiraMenuZeigen() as Void {
    var m = new WatchUi.Menu2({ :title => "Coach Kira" });
    var d = Storage.getValue("d");
    if (d instanceof Dictionary && d.get("vorschlag") instanceof Dictionary) {
        var v = d.get("vorschlag") as Dictionary;
        m.addItem(new WatchUi.MenuItem("Plan-Vorschlag", v.get("text") != null ? v.get("text").toString() : "", :plan, null));
    }
    m.addItem(new WatchUi.MenuItem("Befinden eintragen", "1 bis 5 für heute", :befinden, null));
    m.addItem(new WatchUi.MenuItem("Tour-Rechner", "ESS für eine geplante Tour", :tour, null));
    m.addItem(new WatchUi.MenuItem("Plan ändern", "ESS, Ruhetag, Tage tauschen", :planneu, null));
    m.addItem(new WatchUi.MenuItem("Neu laden", null, :laden, null));
    m.addItem(new WatchUi.MenuItem("Einstellungen", "Seiten, Diagramme, Hinweise", :einst, null));
    m.addItem(new WatchUi.MenuItem("Kopplung lösen", null, :loesen, null));
    WatchUi.pushView(m, new KiraMenuDelegate(), WatchUi.SLIDE_UP);
}

class KiraMenuDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() {
        Menu2InputDelegate.initialize();
    }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId();
        if (id == :befinden) {
            var m = new WatchUi.Menu2({ :title => "Befinden heute" });
            m.addItem(new WatchUi.MenuItem("5", "voll trainierbar", 5, null));
            m.addItem(new WatchUi.MenuItem("4", "gut", 4, null));
            m.addItem(new WatchUi.MenuItem("3", "mittel", 3, null));
            m.addItem(new WatchUi.MenuItem("2", "eher nicht", 2, null));
            m.addItem(new WatchUi.MenuItem("1", "gar nicht", 1, null));
            WatchUi.switchToView(m, new KiraBefindenDelegate(), WatchUi.SLIDE_LEFT);
        } else if (id == :plan) {
            var p = new WatchUi.Menu2({ :title => "Plan-Vorschlag" });
            p.addItem(new WatchUi.MenuItem("Übernehmen", "Lasten anpassen", :ja, null));
            p.addItem(new WatchUi.MenuItem("Verwerfen", "Plan bleibt", :nein, null));
            WatchUi.switchToView(p, new KiraPlanDelegate(), WatchUi.SLIDE_LEFT);
        } else if (id == :tour) {
            WatchUi.popView(WatchUi.SLIDE_IMMEDIATE);
            tourMenuZeigen(true);
        } else if (id == :planneu) {
            WatchUi.popView(WatchUi.SLIDE_IMMEDIATE);
            planMenuZeigen();
        } else if (id == :einst) {
            kiraEinstZeigen();
        } else if (id == :laden) {
            WatchUi.popView(WatchUi.SLIDE_DOWN);
            netz().laden();
        } else if (id == :loesen) {
            WatchUi.switchToView(new WatchUi.Confirmation("Kopplung lösen?"), new KiraLoesen(), WatchUi.SLIDE_IMMEDIATE);
        }
    }
}

class KiraBefindenDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() {
        Menu2InputDelegate.initialize();
    }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var v = item.getId();
        WatchUi.popView(WatchUi.SLIDE_DOWN);
        if (v instanceof Number) { netz().befinden(v); }
    }
}

class KiraPlanDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() {
        Menu2InputDelegate.initialize();
    }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId();
        WatchUi.popView(WatchUi.SLIDE_DOWN);
        netz().vorschlag(id == :ja);
    }
}

// 24.11.2026: Einstellungen direkt an der Uhr (gleiche Werte wie in der Connect IQ App)
function kiraEinstZeigen() as Void {
    var m = new WatchUi.Menu2({ :title => "Einstellungen" });
    m.addItem(new WatchUi.MenuItem("Seiten", "ein/aus", :seiten, null));
    m.addItem(new WatchUi.MenuItem("Startseite", KiraEinst.SEITEN_NAMEN[KiraEinst.zahl("pStart", 0) % KiraEinst.N], :start, null));
    m.addItem(new WatchUi.MenuItem("Last-Diagramm", KiraEinst.zahl("pLast", 14).toString() + " Tage", "pLast", null));
    m.addItem(new WatchUi.MenuItem("Morgenwerte", KiraEinst.zahl("pMorgen", 7).toString() + " Tage", "pMorgen", null));
    m.addItem(new WatchUi.MenuItem("Form-Kurve", KiraEinst.zahl("pForm", 42) == 42 ? "6 Wochen" : "90 Tage", "pForm", null));
    m.addItem(new WatchUi.ToggleMenuItem("Plan-Prognose", "ATL/CTL 7 Tage", "pProg", KiraEinst.an("pProg"), null));
    m.addItem(new WatchUi.MenuItem("Hinweise", "ein/aus", :hinweise, null));
    WatchUi.pushView(m, new KiraEinstDelegate(), WatchUi.SLIDE_LEFT);
}

class KiraEinstDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() {
        Menu2InputDelegate.initialize();
    }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId();
        if (id == :seiten) {
            var m = new WatchUi.Menu2({ :title => "Seiten" });
            for (var i = 0; i < KiraEinst.N; i++) {
                var k = "pS" + i.toString();
                m.addItem(new WatchUi.ToggleMenuItem(KiraEinst.SEITEN_NAMEN[i], null, k, KiraEinst.an(k), null));
            }
            WatchUi.pushView(m, new KiraToggleDelegate(), WatchUi.SLIDE_LEFT);
        } else if (id == :hinweise) {
            var m = new WatchUi.Menu2({ :title => "Hinweise" });
            var hs = [["pHVotum", "Morgenvotum"], ["pHFazit", "Fazit nach Einheit"], ["pHBefinden", "Befinden"], ["pHLicht", "Licht aus"],
                      ["pHWarn", "Warnungen"], ["pHPlan", "Plan-Vorschlag"], ["pHErfolg", "Erfolgsmeldungen"]];
            for (var i = 0; i < hs.size(); i++) {
                m.addItem(new WatchUi.ToggleMenuItem(hs[i][1], null, hs[i][0], KiraEinst.an(hs[i][0]), null));
            }
            WatchUi.pushView(m, new KiraToggleDelegate(), WatchUi.SLIDE_LEFT);
        } else if (id == :start) {
            var m = new WatchUi.Menu2({ :title => "Startseite" });
            for (var i = 0; i < KiraEinst.N; i++) { m.addItem(new WatchUi.MenuItem(KiraEinst.SEITEN_NAMEN[i], null, i, null)); }
            WatchUi.pushView(m, new KiraStartDelegate(item), WatchUi.SLIDE_LEFT);
        } else if (id instanceof String && id.equals("pProg")) {
            KiraEinst.setzen("pProg", (item as WatchUi.ToggleMenuItem).isEnabled());
        } else if (id instanceof String) {
            // Zeitraum weiterschalten
            var werte = id.equals("pLast") ? [7, 14, 28] : id.equals("pMorgen") ? [7, 14] : [42, 90];
            var akt = KiraEinst.zahl(id, werte[1 % werte.size()]);
            var i = werte.indexOf(akt);
            var neu = werte[(i + 1) % werte.size()];
            KiraEinst.setzen(id, neu);
            item.setSubLabel(id.equals("pForm") ? (neu == 42 ? "6 Wochen" : "90 Tage") : neu.toString() + " Tage");
        }
        WatchUi.requestUpdate();
    }
}

class KiraToggleDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() {
        Menu2InputDelegate.initialize();
    }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId();
        if (id instanceof String && item instanceof WatchUi.ToggleMenuItem) {
            KiraEinst.setzen(id, item.isEnabled());
        }
        WatchUi.requestUpdate();
    }
}

class KiraStartDelegate extends WatchUi.Menu2InputDelegate {
    var eltern as WatchUi.MenuItem;

    function initialize(e as WatchUi.MenuItem) {
        Menu2InputDelegate.initialize();
        eltern = e;
    }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var v = item.getId();
        if (v instanceof Number) {
            KiraEinst.setzen("pStart", v);
            KiraEinst.setzen("pS" + v.toString(), true);
            eltern.setSubLabel(KiraEinst.SEITEN_NAMEN[v]);
        }
        WatchUi.popView(WatchUi.SLIDE_RIGHT);
    }
}
