import Toybox.Application.Storage;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Time;
import Toybox.Time.Gregorian;

(:glance)
module KiraHilfen {
    function farbe(stufe) as Number {
        if (stufe == null) { return 0x888888; }
        if (stufe.equals("gruen")) { return 0x2ECC71; }
        if (stufe.equals("gelb")) { return 0xF1C40F; }
        if (stufe.equals("orange")) { return 0xE67E22; }
        if (stufe.equals("rot")) { return 0xE74C3C; }
        return 0x888888;
    }

    // Zahl als Text, Dezimalkomma; null → "–"
    function zahl(v, stellen as Number) as String {
        if (v == null) { return "-"; }
        var s;
        if (v instanceof Float || v instanceof Double) {
            s = v.format("%." + stellen + "f");
        } else if (v instanceof Number || v instanceof Long) {
            s = v.toString();
        } else {
            return v.toString();
        }
        var i = s.find(".");
        if (i != null) {
            s = s.substring(0, i) + "," + s.substring(i + 1, s.length());
        }
        return s;
    }

    // Zeichen, die die Uhrschrift nicht kennt, ersetzen
    function sauber(t as String) as String {
        var ersatz = [["→", ">"], ["✓", ""], ["–", "-"], ["—", "-"], ["…", "..."], ["≥", ">="], ["≤", "<="], ["×", "x"]];
        for (var i = 0; i < ersatz.size(); i++) {
            var a = ersatz[i][0];
            var b = ersatz[i][1];
            var p = t.find(a);
            while (p != null) {
                t = t.substring(0, p) + b + t.substring(p + a.length(), t.length());
                p = t.find(a);
            }
        }
        return t;
    }

    function sportDe(s) as String {
        if (s == null) { return ""; }
        var t = s.toString();
        if (t.equals("Bike")) { return "Rad"; }
        if (t.equals("Run")) { return "Lauf"; }
        if (t.equals("Hike")) { return "Wandern"; }
        if (t.equals("Ski")) { return "Skitour"; }
        if (t.equals("Walk")) { return "Gehen"; }
        if (t.equals("Off") || t.length() == 0) { return "Ruhe"; }
        return t;
    }

    // "2026-11-22" → "So 22."
    function wochentag(iso as String) as String {
        if (iso.length() < 10) { return iso; }
        var m = Gregorian.moment({ :year => iso.substring(0, 4).toNumber(), :month => iso.substring(5, 7).toNumber(), :day => iso.substring(8, 10).toNumber(), :hour => 12 });
        var i = Gregorian.info(m, Time.FORMAT_SHORT);
        var wt = ["", "So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
        return wt[i.day_of_week as Number] + " " + iso.substring(8, 10) + ".";
    }

    function stunden(v) as String {
        if (v == null) { return ""; }
        var m = (v.toFloat() * 60).toNumber();
        return (m / 60).toString() + ":" + (m % 60).format("%02d") + " h";
    }

    function feld(d, a as String, b as String) {
        if (!(d instanceof Dictionary)) { return null; }
        var x = d.get(a);
        if (!(x instanceof Dictionary)) { return null; }
        return x.get(b);
    }

    function uhrzeit(sek) as String {
        if (sek == null) { return "-"; }
        var i = Gregorian.info(new Time.Moment(sek), Time.FORMAT_SHORT);
        return i.hour.format("%02d") + ":" + i.min.format("%02d");
    }

    function daten() {
        return Storage.getValue("d");
    }

    // Stand von heute? (datum im Format YYYY-MM-DD)
    function vonHeute(d) as Boolean {
        if (!(d instanceof Dictionary)) { return false; }
        var dt = d.get("datum");
        if (dt == null) { return false; }
        var i = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        var h = i.year.format("%04d") + "-" + i.month.format("%02d") + "-" + i.day.format("%02d");
        return h.equals(dt);
    }
}
