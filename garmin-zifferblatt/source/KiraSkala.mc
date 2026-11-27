import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;

// 23.11.2026: Garmin-Farbmuster für alle Ringe, Bögen und Balken – jeweils mit individuellen Grenzwerten.
// Rot · Orange · Grün · Blau · Lila wie in Garmin Connect. Segmente plus Punkt am Wert.
module KiraSkala {
    const ROT = 0xD9433A;
    const ORANGE = 0xE5822F;
    const GRUEN = 0x5DB75B;
    const BLAU = 0x3A78CF;
    const LILA = 0x6E45E9;
    const GRAU = 0x3A3A3A;

    // 25.11.2026: Farbstil (Einstellung pStil): 0 Bunt · 1 Dezent · 2 Grau mit Akzent · 3 Monochrom
    var stil as Number = 0;

    function dim(f as Number, p as Number) as Number {
        return ((((f >> 16) & 0xFF) * p / 100) << 16) | ((((f >> 8) & 0xFF) * p / 100) << 8) | ((f & 0xFF) * p / 100);
    }

    function wahl(xs as Array<Number>) as Number { return xs[stil < 0 || stil > 3 ? 0 : stil]; }

    // 26.11.2026: Farbstile klar getrennt – alle Ringe, Bögen, Balken, Punkte und Texte laufen hierüber
    //   0 Bunt: Garmin-Farben · 1 Dezent: gedimmte Farben, orange Minuten · 2 Grau mit Akzent: graue Skalen,
    //   nur Wert-Punkt und Votum farbig · 3 Monochrom: alles grau/weiß, nur Warnungen (rot/orange) farbig
    function lab() as Number { return wahl([0xFF9A1F, 0x9A9A9A, 0x8A8A8A, 0x8C8C8C]); }
    function akku() as Number { return wahl([0xFF9A1F, 0x7A7A7A, 0x6E6E6E, 0x6E6E6E]); }
    function minuten() as Number { return wahl([0xFF9A1F, 0xFF9A1F, 0xE8E8E8, 0xE8E8E8]); }
    function sek() as Number { return wahl([0xFF9A1F, 0x9A9A9A, 0x9A9A9A, 0x8C8C8C]); }
    function mitte() as Number { return wahl([0xFF9A1F, 0xFF9A1F, 0xB0B0B0, 0xB0B0B0]); }
    function votText(f as Number) as Number { return stil == 3 ? 0xB0B0B0 : (stil == 1 ? dim(f, 85) : f); }
    // Skalen-Segmente
    function seg(f as Number) as Number { return stil == 1 ? dim(f, 55) : (stil == 2 ? 0x464646 : (stil == 3 ? 0x3C3C3C : f)); }
    // Punkt am Wert: in Monochrom nur bei rot/orange farbig
    function punkt(f as Number) as Number { return stil == 3 && f != ROT && f != ORANGE ? 0xE6E6E6 : f; }
    // Votum-Ring: Füllung, Grundring und zweite Runde über der Obergrenze
    function votFill(f as Number) as Number { return stil == 1 ? dim(f, 70) : (stil == 3 ? 0x9A9A9A : f); }
    function votSpur(f as Number) as Number { return stil == 0 ? dim(f, 42) : (stil == 1 ? dim(f, 28) : 0x2E2E2E); }
    function ueber() as Number { return wahl([0xFF6A00, 0xB04A00, 0xFF6A00, 0x707070]); }

    // Skala = [grenzen (n+1 Werte, aufsteigend), farben (n)]
    function farbe(sk as Array, v as Numeric?) as Number {
        if (v == null) { return 0x888888; }
        var g = sk[0] as Array;
        var f = sk[1] as Array;
        for (var i = f.size() - 1; i >= 0; i--) {
            if (v >= g[i]) { return f[i] as Number; }
        }
        return f[0] as Number;
    }

    function tr() as Array { return [[0, 25, 50, 75, 95, 100], [ROT, ORANGE, GRUEN, BLAU, LILA]]; }
    function bb() as Array { return [[0, 26, 51, 76, 100], [ROT, ORANGE, GRUEN, BLAU]]; }
    function schlaf() as Array { return [[0, 60, 80, 90, 100], [ROT, ORANGE, GRUEN, BLAU]]; }
    // 24.11.2026: ACWR × 100 – unter 0,8 blau · bis 1,3 grün · bis 1,5 orange · darüber rot
    function acwr() as Array { return [[50, 80, 130, 150, 200], [BLAU, GRUEN, ORANGE, ROT]]; }
    function tourform() as Array { return [[0, 60, 80, 100], [ORANGE, GRUEN, BLAU]]; }

    // Tageslast gegen Kiras Soll-Rahmen lo–hi: unter Soll blau · im Rahmen grün · über Soll orange · > +15 % rot
    function last(lo as Number?, hi as Number) as Array {
        var o = hi * 115 / 100;
        var m = hi * 13 / 10;
        if (o <= hi) { o = hi + 1; }
        if (m <= o) { m = o + 1; }
        if (lo != null && lo > 0 && lo < hi) { return [[0, lo, hi, o, m], [BLAU, GRUEN, ORANGE, ROT]]; }
        return [[0, hi, o, m], [GRUEN, ORANGE, ROT]];
    }

    // Wochenlast gegen Kiras Wochenbudget: < 85 % blau · bis 100 % grün · bis 115 % orange · darüber rot
    function woche(bud as Number) as Array {
        return [[0, bud * 85 / 100, bud, bud * 115 / 100, bud * 13 / 10], [BLAU, GRUEN, ORANGE, ROT]];
    }

    function winkel(sk as Array, v as Numeric) as Number {
        var g = sk[0] as Array;
        var lo = g[0] as Number;
        var hi = g[g.size() - 1] as Number;
        var x = v < lo ? lo : (v > hi ? hi : v);
        return (225 - (x - lo).toFloat() * 270.0 / (hi - lo)).toNumber();
    }

    function norm(a as Number) as Number { return ((a % 360) + 360) % 360; }

    // 270°-Bogen (7:30 bis 4:30 Uhr), Segmente mit kleiner Lücke, Punkt am Wert (schwarz umrandet)
    function bogen(dc as Graphics.Dc, x as Number, y as Number, r as Number, pen as Number, v as Numeric?, sk as Array) as Void {
        var g = sk[0] as Array;
        var f = sk[1] as Array;
        dc.setPenWidth(pen);
        for (var i = 0; i < f.size(); i++) {
            var a1 = winkel(sk, g[i] as Number) - (i > 0 ? 2 : 0);
            var a2 = winkel(sk, g[i + 1] as Number) + (i < f.size() - 1 ? 2 : 0);
            if (a1 - a2 < 2) { continue; }
            dc.setColor(v == null ? GRAU : seg(f[i] as Number), Graphics.COLOR_TRANSPARENT);
            dc.drawArc(x, y, r, Graphics.ARC_CLOCKWISE, norm(a1), norm(a2));
        }
        dc.setPenWidth(1);
        if (v != null) {
            var rad = Math.toRadians(winkel(sk, v));
            var px = x + (r * Math.cos(rad)).toNumber();
            var py = y - (r * Math.sin(rad)).toNumber();
            dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(px, py, pen / 2 + 5);
            dc.setColor(punkt(farbe(sk, v)), Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(px, py, pen / 2 + 3);
        }
    }

    // Waagrechter Balken: Segmente mit 2 px Lücke, Punkt am Wert
    function balken(dc as Graphics.Dc, x as Number, y as Number, bw as Number, hoch as Number, v as Numeric?, sk as Array, rand as Number) as Void {
        var g = sk[0] as Array;
        var f = sk[1] as Array;
        var lo = g[0] as Number;
        var hi = g[g.size() - 1] as Number;
        for (var i = 0; i < f.size(); i++) {
            var x1 = x + (bw * ((g[i] as Number) - lo) / (hi - lo));
            var x2 = x + (bw * ((g[i + 1] as Number) - lo) / (hi - lo)) - (i < f.size() - 1 ? 2 : 0);
            if (x2 <= x1) { continue; }
            dc.setColor(v == null ? GRAU : seg(f[i] as Number), Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x1, y, x2 - x1, hoch);
        }
        if (v != null) {
            var vv = v < lo ? lo : (v > hi ? hi : v);
            var px = x + (bw * (vv - lo).toFloat() / (hi - lo)).toNumber();
            var cy = y + hoch / 2;
            dc.setColor(rand, Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(px, cy, hoch / 2 + 5);
            dc.setColor(punkt(farbe(sk, v)), Graphics.COLOR_TRANSPARENT);
            dc.fillCircle(px, cy, hoch / 2 + 3);
        }
    }
}
