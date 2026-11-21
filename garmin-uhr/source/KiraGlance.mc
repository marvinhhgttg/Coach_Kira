import Toybox.Graphics;
import Toybox.Lang;
import Toybox.WatchUi;

// Glance: Punkt in Votumfarbe, Votum, Soll heute
(:glance)
class KiraGlance extends WatchUi.GlanceView {
    function initialize() {
        GlanceView.initialize();
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        dc.setColor(Graphics.COLOR_TRANSPARENT, Graphics.COLOR_TRANSPARENT);
        dc.clear();
        var d = KiraHilfen.daten();
        var h = dc.getHeight();
        if (!(d instanceof Dictionary)) {
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(0, h / 2, Graphics.FONT_GLANCE, "COACH KIRA · koppeln", Graphics.TEXT_JUSTIFY_LEFT | Graphics.TEXT_JUSTIFY_VCENTER);
            return;
        }
        var heute = KiraHilfen.vonHeute(d);
        var stufe = KiraHilfen.feld(d, "votum", "stufe");
        var label = KiraHilfen.feld(d, "votum", "label");
        var r = h / 6;
        dc.setColor(heute ? KiraHilfen.farbe(stufe) : 0x888888, Graphics.COLOR_TRANSPARENT);
        dc.fillCircle(r + 2, h / 3, r);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(2 * r + 12, h / 3, Graphics.FONT_GLANCE, "KIRA · " + (label == null ? "–" : label.toUpper()) + (heute ? "" : " (gestern)"),
            Graphics.TEXT_JUSTIFY_LEFT | Graphics.TEXT_JUSTIFY_VCENTER);
        var soll = KiraHilfen.feld(d, "last", "sollHeute");
        var sport = KiraHilfen.feld(d, "last", "sportHeute");
        var zone = KiraHilfen.feld(d, "last", "zoneHeute");
        var zeile = "Soll " + KiraHilfen.zahl(soll, 0) + " ESS" + (sport != null ? " · " + sport : "") + (zone != null ? " " + zone : "");
        dc.setColor(0xBBBBBB, Graphics.COLOR_TRANSPARENT);
        dc.drawText(2, h * 3 / 4, Graphics.FONT_GLANCE, zeile, Graphics.TEXT_JUSTIFY_LEFT | Graphics.TEXT_JUSTIFY_VCENTER);
    }
}
