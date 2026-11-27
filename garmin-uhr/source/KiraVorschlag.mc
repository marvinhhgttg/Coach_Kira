// 27.11.2026: Garmin-Tagesvorschlag eintragen (Typ → Minuten → Zone). Garmin liefert die Vorschläge der Uhr
// über keine Schnittstelle aus – Kira zeigt sie als Zweitmeinung zum Votum (ändert das Votum nicht).
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.WatchUi;

var gvTyp as String = "erholung";
var gvMin as Number = 30;

function gvTypZeigen() as Void {
    var m = new WatchUi.Menu2({ :title => "Garmin schlägt vor" });
    var t = [["erholung", "Erholung"], ["basis", "Basis"], ["lang", "Lang"], ["tempo", "Tempo"], ["schwelle", "Schwelle"],
             ["vo2max", "VO2max"], ["anaerob", "Anaerob"], ["ruhe", "Ruhe"], ["loeschen", "Löschen"]];
    for (var i = 0; i < t.size(); i++) { m.addItem(new WatchUi.MenuItem(t[i][1], null, t[i][0], null)); }
    WatchUi.switchToView(m, new GvTypDelegate(), WatchUi.SLIDE_LEFT);
}

class GvTypDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() { Menu2InputDelegate.initialize(); }
    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId() as String;
        gvTyp = id;
        if (id.equals("ruhe") || id.equals("loeschen")) {
            WatchUi.popView(WatchUi.SLIDE_DOWN);
            netz().garminVorschlag({ "typ" => id, "min" => 0 });
            return;
        }
        var f = new GvZahlen(5, 180);
        WatchUi.switchToView(new WatchUi.Picker({ :title => new WatchUi.Text({ :text => "Minuten", :locX => WatchUi.LAYOUT_HALIGN_CENTER, :locY => WatchUi.LAYOUT_VALIGN_BOTTOM, :color => Graphics.COLOR_WHITE }),
            :pattern => [f], :defaults => [gvMin - 5] }), new GvMinDelegate(), WatchUi.SLIDE_LEFT);
    }
}

class GvZahlen extends WatchUi.PickerFactory {
    var lo as Number;
    var hi as Number;
    function initialize(a as Number, b as Number) { PickerFactory.initialize(); lo = a; hi = b; }
    function getSize() as Number { return hi - lo + 1; }
    function getValue(i as Number) as Object? { return lo + i; }
    function getDrawable(i as Number, sel as Boolean) as WatchUi.Drawable? {
        return new WatchUi.Text({ :text => (lo + i).toString(), :color => Graphics.COLOR_WHITE, :font => Graphics.FONT_NUMBER_MEDIUM,
            :locX => WatchUi.LAYOUT_HALIGN_CENTER, :locY => WatchUi.LAYOUT_VALIGN_CENTER });
    }
}

class GvMinDelegate extends WatchUi.PickerDelegate {
    function initialize() { PickerDelegate.initialize(); }
    function onCancel() as Boolean { WatchUi.popView(WatchUi.SLIDE_DOWN); return true; }
    function onAccept(values as Array) as Boolean {
        var v = values[0];
        if (v instanceof Number) { gvMin = v; }
        var m = new WatchUi.Menu2({ :title => "Zone" });
        m.addItem(new WatchUi.MenuItem("Standard", "je nach Typ", "", null));
        var z = ["Z1", "Z2", "Z3", "Z4", "Z5"];
        for (var i = 0; i < z.size(); i++) { m.addItem(new WatchUi.MenuItem(z[i], null, z[i], null)); }
        WatchUi.switchToView(m, new GvZoneDelegate(), WatchUi.SLIDE_LEFT);
        return true;
    }
}

class GvZoneDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() { Menu2InputDelegate.initialize(); }
    function onSelect(item as WatchUi.MenuItem) as Void {
        var z = item.getId() as String;
        WatchUi.popView(WatchUi.SLIDE_DOWN);
        var b = { "typ" => gvTyp, "min" => gvMin } as Dictionary;
        if (z.length() > 0) { b.put("zone", z); }
        netz().garminVorschlag(b);
    }
}
