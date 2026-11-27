// 25.11.2026: Tour-Rechner und Plan ändern direkt an der Uhr
import Toybox.Application.Storage;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.WatchUi;

var kiraTour as Array = [];
var kiraTourErg as Dictionary? = null;
var kiraTourFehler as String? = null;

const TOUR_SPORT = ["gravel", "rad", "lauf", "berg", "abfahrt"];
const TOUR_NAME = ["Gravel/MTB bergauf", "Rennrad", "Lauf", "Bergtour zu Fuß", "Abfahrt Rad"];

function tourSportName(s as String) as String {
    var i = TOUR_SPORT.indexOf(s);
    return i >= 0 ? TOUR_NAME[i] : s;
}

function kmText(v) as String {
    return KiraX.z(v.toFloat());
}

// ---------- Zahleneingabe: UP +, DOWN -, MENU Schrittweite, START übernehmen ----------
class KiraZahlView extends WatchUi.View {
    var titel as String;
    var einheit as String;
    var wert as Float;
    var schritte as Array;
    var si as Number = 0;
    var lo as Float;
    var hi as Float;
    var ziel;   // Objekt mit fertig(wert as Float)

    function initialize(t as String, e as String, w as Float, st as Array, mn as Float, mx as Float, zz) {
        View.initialize();
        titel = t; einheit = e; wert = w; schritte = st; lo = mn; hi = mx; ziel = zz;
    }

    function schritt() as Float { return schritte[si].toFloat(); }

    function aendern(r as Number) as Void {
        wert = wert + r * schritt();
        if (wert < lo) { wert = lo; }
        if (wert > hi) { wert = hi; }
        WatchUi.requestUpdate();
    }

    function text() as String {
        return schritt() < 1 ? KiraX.z(wert) : wert.toNumber().toString();
    }

    function onUpdate(dc as Graphics.Dc) as Void {
        var w = dc.getWidth(), h = dc.getHeight();
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_BLACK);
        dc.clear();
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.16, Graphics.FONT_XTINY, KiraHilfen.sauber(titel), Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.30, Graphics.FONT_SMALL, "UP +", Graphics.TEXT_JUSTIFY_CENTER);
        dc.drawText(w / 2, h * 0.47, Graphics.FONT_NUMBER_MEDIUM, text(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.58, Graphics.FONT_XTINY, einheit, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.64, Graphics.FONT_SMALL, "DOWN -", Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(0x777777, Graphics.COLOR_TRANSPARENT);
        var st = schritt();
        dc.drawText(w / 2, h * 0.78, Graphics.FONT_XTINY, "Schritt " + (st < 1 ? KiraX.z(st) : st.toNumber().toString()) + " · MENU ändert", Graphics.TEXT_JUSTIFY_CENTER);
        dc.drawText(w / 2, h * 0.84, Graphics.FONT_XTINY, "START übernehmen", Graphics.TEXT_JUSTIFY_CENTER);
    }
}

class KiraZahlDelegate extends WatchUi.BehaviorDelegate {
    var v as KiraZahlView;
    function initialize(vv as KiraZahlView) { BehaviorDelegate.initialize(); v = vv; }
    function onPreviousPage() as Boolean { v.aendern(1); return true; }
    function onNextPage() as Boolean { v.aendern(-1); return true; }
    function onMenu() as Boolean { v.si = (v.si + 1) % v.schritte.size(); WatchUi.requestUpdate(); return true; }
    function onSelect() as Boolean { v.ziel.fertig(v.wert); return true; }
}

function zahlZeigen(t as String, e as String, w as Float, st as Array, mn as Float, mx as Float, zz) as Void {
    var v = new KiraZahlView(t, e, w, st, mn, mx, zz);
    WatchUi.switchToView(v, new KiraZahlDelegate(v), WatchUi.SLIDE_LEFT);
}

// ---------- Tour-Rechner ----------
function tourMenuZeigen(neu as Boolean) as Void {
    var m = new WatchUi.Menu2({ :title => "Tour-Rechner" });
    for (var i = 0; i < kiraTour.size(); i++) {
        var t = kiraTour[i] as Dictionary;
        var sub = kmText(t.get("km")) + " km" + (t.get("s").equals("abfahrt") ? "" : " · " + t.get("hm").toString() + " hm");
        m.addItem(new WatchUi.MenuItem((i + 1).toString() + ". " + tourSportName(t.get("s") as String), sub + " (entfernen)", i, null));
    }
    m.addItem(new WatchUi.MenuItem("+ Abschnitt", "Sport, km, hm", :neu, null));
    if (kiraTour.size() > 0) {
        m.addItem(new WatchUi.MenuItem("Berechnen", "aus deinen eigenen Touren", :rechnen, null));
        m.addItem(new WatchUi.MenuItem("Alles leeren", null, :leer, null));
    }
    if (neu) { WatchUi.pushView(m, new KiraTourDelegate(), WatchUi.SLIDE_LEFT); }
    else { WatchUi.switchToView(m, new KiraTourDelegate(), WatchUi.SLIDE_RIGHT); }
}

class KiraTourDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() { Menu2InputDelegate.initialize(); }

    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId();
        if (id == :neu) {
            var m = new WatchUi.Menu2({ :title => "Abschnitt" });
            for (var i = 0; i < TOUR_SPORT.size(); i++) { m.addItem(new WatchUi.MenuItem(TOUR_NAME[i], null, TOUR_SPORT[i], null)); }
            WatchUi.switchToView(m, new KiraTourSportDelegate(), WatchUi.SLIDE_LEFT);
        } else if (id == :rechnen) {
            netz().tour(kiraTour);
            WatchUi.switchToView(new KiraTourErgView(), new KiraTourErgDelegate(), WatchUi.SLIDE_LEFT);
        } else if (id == :leer) {
            kiraTour = [];
            tourMenuZeigen(false);
        } else if (id instanceof Number) {
            kiraTour.remove(kiraTour[id]);
            tourMenuZeigen(false);
        }
    }
}

class KiraTourSportDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() { Menu2InputDelegate.initialize(); }
    function onSelect(item as WatchUi.MenuItem) as Void {
        var s = item.getId() as String;
        zahlZeigen(tourSportName(s) + ": Strecke", "km", s.equals("berg") ? 8.0 : 15.0, [1, 0.1, 5], 0.0, 300.0, new KiraTourKm(s));
    }
    function onBack() as Void { tourMenuZeigen(false); }
}

class KiraTourKm {
    var s as String;
    function initialize(ss as String) { s = ss; }
    function fertig(km as Float) as Void {
        if (s.equals("abfahrt")) {
            kiraTour.add({ "s" => s, "km" => km, "hm" => 0 });
            tourMenuZeigen(false);
        } else {
            zahlZeigen(tourSportName(s) + ": Höhenmeter", "hm bergauf", s.equals("berg") ? 700.0 : 500.0, [50, 10, 100], 0.0, 5000.0, new KiraTourHm(s, km));
        }
    }
}

class KiraTourHm {
    var s as String;
    var km as Float;
    function initialize(ss as String, k as Float) { s = ss; km = k; }
    function fertig(hm as Float) as Void {
        kiraTour.add({ "s" => s, "km" => km, "hm" => hm.toNumber() });
        tourMenuZeigen(false);
    }
}

class KiraTourErgView extends WatchUi.View {
    function initialize() { View.initialize(); }

    function onUpdate(dc as Graphics.Dc) as Void {
        var w = dc.getWidth(), h = dc.getHeight();
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_BLACK);
        dc.clear();
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.10, Graphics.FONT_XTINY, "TOUR-RECHNER", Graphics.TEXT_JUSTIFY_CENTER);
        var e = kiraTourErg;
        if (e == null) {
            var t = kiraTourFehler != null ? kiraTourFehler as String : "rechnet ...";
            dc.setColor(kiraTourFehler != null ? KiraSkala.ORANGE : Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w / 2, h / 2, Graphics.FONT_SMALL, KiraHilfen.sauber(t), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            return;
        }
        dc.setColor(KiraSkala.GRUEN, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.17, Graphics.FONT_NUMBER_MEDIUM, e.get("ess").toString(), Graphics.TEXT_JUSTIFY_CENTER);
        var y = (h * 0.17).toNumber() + dc.getFontHeight(Graphics.FONT_NUMBER_MEDIUM);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, y, Graphics.FONT_XTINY, "ESS (" + e.get("lo").toString() + "-" + e.get("hi").toString() + ")", Graphics.TEXT_JUSTIFY_CENTER);
        y += dc.getFontHeight(Graphics.FONT_XTINY) + 8;
        var ts = e.get("teile");
        if (ts instanceof Array) {
            for (var i = 0; i < ts.size() && i < 5; i++) {
                var t = ts[i] as Dictionary;
                y = KiraX.lr(dc, y, tourSportName(t.get("s").toString()) + " " + kmText(t.get("km")) + " km", t.get("ess").toString(), Graphics.COLOR_WHITE);
                var q = t.get("q");
                if (q != null) {
                    dc.setColor(0x777777, Graphics.COLOR_TRANSPARENT);
                    dc.drawText(w * 0.17, y, Graphics.FONT_XTINY, KiraHilfen.sauber(q.toString()), Graphics.TEXT_JUSTIFY_LEFT);
                    y += dc.getFontHeight(Graphics.FONT_XTINY) + 2;
                }
            }
        }
        var u = KiraX.teil("rest");
        if (u != null) {
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w / 2, h * 0.84, Graphics.FONT_XTINY, "Heutiger Rahmen bis " + u.get("max").toString(), Graphics.TEXT_JUSTIFY_CENTER);
        }
    }
}

class KiraTourErgDelegate extends WatchUi.BehaviorDelegate {
    function initialize() { BehaviorDelegate.initialize(); }
    function onBack() as Boolean { tourMenuZeigen(false); return true; }
    function onSelect() as Boolean { tourMenuZeigen(false); return true; }
}

// ---------- Plan ändern ----------
function planMenuZeigen() as Void {
    var pl = KiraX.liste("pl");
    var m = new WatchUi.Menu2({ :title => "Plan ändern" });
    var n = 0;
    if (pl != null) {
        for (var i = 0; i < pl.size(); i++) {
            var p = pl[i] as Dictionary;
            n += 1;
            var sub = p.get("lock") == true ? "gesperrt" : (p.get("s") != null ? p.get("s").toString() : "") + (p.get("z") != null && !p.get("z").toString().equals("Off") ? " " + p.get("z").toString() : "");
            m.addItem(new WatchUi.MenuItem(p.get("t").toString() + "  " + p.get("ess").toString() + " ESS", KiraHilfen.sauber(sub), i, null));
        }
    }
    if (n == 0) { m.addItem(new WatchUi.MenuItem("Kein Plan geladen", "START auf der Hauptseite", :nix, null)); }
    WatchUi.pushView(m, new KiraPlanTagDelegate(), WatchUi.SLIDE_LEFT);
}

class KiraPlanTagDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() { Menu2InputDelegate.initialize(); }
    function onSelect(item as WatchUi.MenuItem) as Void {
        var i = item.getId();
        var pl = KiraX.liste("pl");
        if (!(i instanceof Number) || pl == null) { WatchUi.popView(WatchUi.SLIDE_RIGHT); return; }
        var p = pl[i] as Dictionary;
        if (p.get("lock") == true) {
            if (WatchUi has :showToast) { WatchUi.showToast("Tag ist gesperrt", null); }
            return;
        }
        var m = new WatchUi.Menu2({ :title => p.get("t").toString() });
        m.addItem(new WatchUi.MenuItem("ESS ändern", "jetzt " + p.get("ess").toString(), :ess, null));
        m.addItem(new WatchUi.MenuItem("Ruhetag", "0 ESS", :ruhe, null));
        m.addItem(new WatchUi.MenuItem("Tauschen mit ...", "anderer Tag", :tausch, null));
        WatchUi.switchToView(m, new KiraPlanArtDelegate(p), WatchUi.SLIDE_LEFT);
    }
}

class KiraPlanArtDelegate extends WatchUi.Menu2InputDelegate {
    var p as Dictionary;
    function initialize(pp as Dictionary) { Menu2InputDelegate.initialize(); p = pp; }
    function onSelect(item as WatchUi.MenuItem) as Void {
        var id = item.getId();
        if (id == :ess) {
            zahlZeigen(p.get("t").toString() + ": neue Last", "ESS", p.get("ess").toFloat(), [10, 5, 50], 0.0, 400.0, new KiraPlanEss(p.get("d").toString()));
        } else if (id == :ruhe) {
            WatchUi.popView(WatchUi.SLIDE_DOWN);
            netz().planAendern({ "art" => "ruhe", "d" => p.get("d") });
        } else if (id == :tausch) {
            var pl = KiraX.liste("pl");
            var m = new WatchUi.Menu2({ :title => "Tauschen mit" });
            if (pl != null) {
                for (var i = 0; i < pl.size(); i++) {
                    var q = pl[i] as Dictionary;
                    if (q.get("lock") == true || q.get("d").toString().equals(p.get("d").toString())) { continue; }
                    m.addItem(new WatchUi.MenuItem(q.get("t").toString() + "  " + q.get("ess").toString(), q.get("s") != null ? q.get("s").toString() : null, q.get("d").toString(), null));
                }
            }
            WatchUi.switchToView(m, new KiraPlanTauschDelegate(p.get("d").toString()), WatchUi.SLIDE_LEFT);
        }
    }
}

class KiraPlanTauschDelegate extends WatchUi.Menu2InputDelegate {
    var d as String;
    function initialize(dd as String) { Menu2InputDelegate.initialize(); d = dd; }
    function onSelect(item as WatchUi.MenuItem) as Void {
        WatchUi.popView(WatchUi.SLIDE_DOWN);
        netz().planAendern({ "art" => "tausch", "d" => d, "d2" => item.getId() });
    }
}

class KiraPlanEss {
    var d as String;
    function initialize(dd as String) { d = dd; }
    function fertig(ess as Float) as Void {
        WatchUi.popView(WatchUi.SLIDE_DOWN);
        netz().planAendern({ "art" => "ess", "d" => d, "ess" => ess.toNumber() });
    }
}
