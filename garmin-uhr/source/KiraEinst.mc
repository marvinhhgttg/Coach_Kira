// 24.11.2026: Einstellungen (Connect IQ App → Coach Kira → Einstellungen, oder Uhr-Menü)
import Toybox.Application;
import Toybox.Lang;
import Toybox.Time;
import Toybox.Time.Gregorian;

(:background :glance)
module KiraEinst {
    const SEITEN_NAMEN = ["Votum", "Werte", "Nächste Tage", "Last", "Morgenwerte", "Form-Kurve", "Letzte Einheit", "Ziel", "Form & Ausdauer", "Erholung & Woche",
        "Vorausblick", "Heute noch", "Zonen 7/28", "Erholungstrend", "Wochenbilanz", "Bereitschaft morgen", "Ziel-Countdown", "Treffsicherheit"];
    const N = 18;

    function wert(k as String, def) {
        try {
            var v = Application.Properties.getValue(k);
            return v == null ? def : v;
        } catch (e) {
            return def;
        }
    }

    function an(k as String) as Boolean {
        return wert(k, true) != false;
    }

    function zahl(k as String, def as Number) as Number {
        var v = wert(k, def);
        return v instanceof Number ? v : def;
    }

    function setzen(k as String, v) as Void {
        try { Application.Properties.setValue(k, v); } catch (e) { }
    }

    // aktive Seiten in fester Reihenfolge; mindestens Votum
    function seiten() as Array {
        var xs = [];
        for (var i = 0; i < N; i++) { if (an("pS" + i.toString())) { xs.add(i); } }
        if (xs.size() == 0) { xs.add(0); }
        return xs;
    }

    function ruhig() as Boolean {
        var von = zahl("pRuheVon", 23);
        var bis = zahl("pRuheBis", 6);
        if (von == bis) { return false; }
        var h = Gregorian.info(Time.now(), Time.FORMAT_SHORT).hour;
        return von < bis ? (h >= von && h < bis) : (h >= von || h < bis);
    }

    // Schlüssel der Hinweisart zu einer Server-id
    function hinweisSchluessel(id as String) as String {
        if (id.find("plan-") == 0) { return "pHPlan"; }
        if (id.find("ok-") == 0 || id.find("wo-") == 0) { return "pHErfolg"; }
        if (id.find("-befinden") != null) { return "pHBefinden"; }
        if (id.find("-licht") != null) { return "pHLicht"; }
        if (id.find("-acwr") != null || id.find("-rhr") != null) { return "pHWarn"; }
        return "pHVotum";
    }

    // letzte n Einträge einer Liste
    function ende(xs as Array, n as Number) as Array {
        return xs.size() > n ? xs.slice(xs.size() - n, null) : xs;
    }
}
