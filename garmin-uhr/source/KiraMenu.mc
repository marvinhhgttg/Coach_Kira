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
    m.addItem(new WatchUi.MenuItem("Neu laden", null, :laden, null));
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
