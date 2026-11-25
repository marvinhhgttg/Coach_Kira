// 25.11.2026: Zusatzseiten 10–17 (Vorausblick, Heute noch, Zonen, Erholungstrend, Wochenbilanz,
// Bereitschaft morgen, Ziel-Countdown, Treffsicherheit). Daten: Storage "x" (Server: uhr.x)
import Toybox.Application.Storage;
import Toybox.Graphics;
import Toybox.Lang;
import Toybox.Math;

module KiraX {
    const ZF = [0x9A9A9A, 0x3A78CF, 0x5DB75B, 0xE5822F, 0xD9433A];

    function teil(k as String) as Dictionary? {
        var x = Storage.getValue("x");
        if (!(x instanceof Dictionary)) { return null; }
        var t = x.get(k);
        return t instanceof Dictionary ? t : null;
    }

    function liste(k as String) as Array? {
        var x = Storage.getValue("x");
        if (!(x instanceof Dictionary)) { return null; }
        var t = x.get(k);
        return t instanceof Array ? t : null;
    }

    function z(v) as String {
        if (v == null) { return "-"; }
        if (v instanceof Float || v instanceof Double) {
            var s = (Math.round(v * 10) / 10.0).format("%.1f");
            var p = s.find(".");
            return p == null ? s : s.substring(0, p) + "," + s.substring(p + 1, s.length());
        }
        return v.toString();
    }

    function z2(v) as String {
        if (v == null) { return "-"; }
        var s = v.toFloat().format("%.2f");
        var p = s.find(".");
        return p == null ? s : s.substring(0, p) + "," + s.substring(p + 1, s.length());
    }

    function vz(v) as String {
        if (v == null) { return "-"; }
        return (v > 0 ? "+" : "") + v.toString();
    }

    function farbeStufe(s) as Number {
        if (s == null) { return Graphics.COLOR_WHITE; }
        var t = s.toString();
        return t.equals("gruen") ? KiraSkala.GRUEN : t.equals("gelb") ? 0xE8C547 : t.equals("orange") ? KiraSkala.ORANGE : t.equals("rot") ? KiraSkala.ROT : Graphics.COLOR_WHITE;
    }

    // Absatz über mehrere Zeilen, mittig; gibt nächste y-Position zurück
    function absatz(dc as Graphics.Dc, y as Number, hoehe as Number, breite as Number, font as Graphics.FontType, text as String, farbe as Number) as Number {
        var t = Graphics.fitTextToArea(KiraHilfen.sauber(text), font, breite, hoehe, true);
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.drawText(dc.getWidth() / 2, y, font, t, Graphics.TEXT_JUSTIFY_CENTER);
        return y + hoehe;
    }

    // Zeile links-rechts (Beschriftung grau, Wert weiß)
    function lr(dc as Graphics.Dc, y as Number, l as String, r as String, fr as Number) as Number {
        var w = dc.getWidth();
        var fh = dc.getFontHeight(Graphics.FONT_XTINY);
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w * 0.18, y, Graphics.FONT_XTINY, KiraHilfen.sauber(l), Graphics.TEXT_JUSTIFY_LEFT);
        dc.setColor(fr, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w * 0.82, y, Graphics.FONT_XTINY, KiraHilfen.sauber(r), Graphics.TEXT_JUSTIFY_RIGHT);
        return y + fh + 2;
    }

    function polylinie(dc as Graphics.Dc, x0 as Number, x1 as Number, y as Number, hoch as Number, xs as Array, lo as Float, hi as Float, farbe as Number) as Void {
        var n = xs.size();
        if (n < 2 || hi <= lo) { return; }
        dc.setColor(farbe, Graphics.COLOR_TRANSPARENT);
        dc.setPenWidth(3);
        var px = -1, py = -1;
        for (var i = 0; i < n; i++) {
            var v = xs[i];
            if (v == null || v < 0) { px = -1; continue; }
            var x = x0 + ((x1 - x0) * i / (n - 1)).toNumber();
            var yy = y + hoch - ((hoch * (v - lo) / (hi - lo))).toNumber();
            if (yy < y) { yy = y; }
            if (yy > y + hoch) { yy = y + hoch; }
            if (px >= 0) { dc.drawLine(px, py, x, yy); }
            px = x; py = yy;
        }
        dc.setPenWidth(1);
    }

    function grenzen(xs as Array) as Array {
        var lo = 99999.0, hi = -99999.0;
        for (var i = 0; i < xs.size(); i++) {
            var v = xs[i];
            if (v == null || v < 0) { continue; }
            if (v < lo) { lo = v.toFloat(); }
            if (v > hi) { hi = v.toFloat(); }
        }
        if (hi < lo) { return [0.0, 1.0]; }
        if (hi - lo < 4) { hi += 2; lo -= 2; }
        return [lo, hi];
    }

    function letzter(xs as Array) {
        for (var i = xs.size() - 1; i >= 0; i--) { if (xs[i] != null && xs[i] >= 0) { return xs[i]; } }
        return null;
    }

    // 10 Vorausblick
    function vorausblick(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var b = teil("vb");
        if (b == null) { v.ohneDaten(dc, h, "VORAUSBLICK"); return; }
        var w = dc.getWidth();
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "VORAUSBLICK", 0x9A9A9A);
        var st = b.get("stufe");
        var f = farbeStufe(st);
        var ac = b.get("acwr");
        y = v.zeile(dc, y, Graphics.FONT_SMALL, st != null ? st.toString().toUpper() : "-", f);
        if (ac instanceof Dictionary) {
            y = v.zeile(dc, y, Graphics.FONT_XTINY, "ACWR " + z2(ac.get("w")) + " · " + ac.get("d").toString(), Graphics.COLOR_WHITE);
        }
        var p = b.get("prog");
        if (p != null) { y = absatz(dc, y + 4, (h * 0.33).toNumber(), (w * 0.84).toNumber(), Graphics.FONT_XTINY, p.toString(), 0xDDDDDD); }
        var tf = b.get("tf");
        if (tf instanceof Dictionary) { v.zeile(dc, y + 2, Graphics.FONT_XTINY, "Tour " + tf.get("tag").toString() + " (" + z(tf.get("s")) + ")", KiraSkala.GRUEN); }
    }

    // 11 Was geht heute noch
    function heuteNoch(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var r = teil("rest");
        if (r == null) { v.ohneDaten(dc, h, "HEUTE NOCH"); return; }
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "HEUTE NOCH", 0x9A9A9A);
        var ess = r.get("ess");
        if (ess == null || ess <= 0) {
            y = v.zeile(dc, y + 10, Graphics.FONT_SMALL, "Rahmen erreicht", KiraSkala.ORANGE);
            y = v.zeile(dc, y + 4, Graphics.FONT_XTINY, z(r.get("ist")) + " ESS · Soll " + z(r.get("min")) + "-" + z(r.get("max")), Graphics.COLOR_WHITE);
            y = v.zeile(dc, y + 8, Graphics.FONT_XTINY, "Nur noch locker bewegen", 0x9A9A9A);
            v.zeile(dc, y, Graphics.FONT_XTINY, "oder Mobility/Dehnen", 0x9A9A9A);
            return;
        }
        y = v.zeile(dc, y, Graphics.FONT_NUMBER_MEDIUM, ess.toString(), KiraSkala.GRUEN);
        y = v.zeile(dc, y, Graphics.FONT_XTINY, "ESS bis Obergrenze " + z(r.get("max")), Graphics.COLOR_WHITE);
        var hs = r.get("halte");
        if (hs != null && hs > 0) { y = v.zeile(dc, y, Graphics.FONT_XTINY, hs.toString() + " bis Haltewert", 0x9A9A9A); }
        y += 6;
        var os = r.get("opts");
        if (os instanceof Array) {
            for (var i = 0; i < os.size() && i < 4; i++) {
                var o = os[i];
                if (!(o instanceof Dictionary)) { continue; }
                y = lr(dc, y, o.get("s").toString(), "ca. " + o.get("min").toString() + " min", Graphics.COLOR_WHITE);
            }
        }
    }

    function zonenBalken(dc as Graphics.Dc, y as Number, ps as Array, titel as String) as Number {
        var w = dc.getWidth();
        var x0 = (w * 0.16).toNumber(), x1 = (w * 0.84).toNumber();
        dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x0, y, Graphics.FONT_XTINY, titel, Graphics.TEXT_JUSTIFY_LEFT);
        y += dc.getFontHeight(Graphics.FONT_XTINY) + 2;
        var x = x0;
        var bh = 22;
        for (var i = 0; i < 5 && i < ps.size(); i++) {
            var bw = ((x1 - x0) * ps[i] / 100.0).toNumber();
            if (bw <= 0) { continue; }
            dc.setColor(ZF[i], Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x, y, bw, bh);
            if (bw > 34) {
                dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_TRANSPARENT);
                dc.drawText(x + bw / 2, y + bh / 2, Graphics.FONT_XTINY, ps[i].toString(), Graphics.TEXT_JUSTIFY_CENTER | Graphics.TEXT_JUSTIFY_VCENTER);
            }
            x += bw;
        }
        return y + bh + 6;
    }

    // 12 Zonenverteilung
    function zonen(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var zz = teil("zonen");
        if (zz == null) { v.ohneDaten(dc, h, "ZONEN"); return; }
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "ZONEN %", 0x9A9A9A);
        y = zonenBalken(dc, y + 8, zz.get("p7") as Array, "7 Tage · " + z(zz.get("h7")) + " h");
        y = zonenBalken(dc, y + 4, zz.get("p28") as Array, "28 Tage · " + z(zz.get("h28")) + " h");
        y += 4;
        var g7 = zz.get("grau7");
        y = lr(dc, y, "Z3 grau", z(g7) + " / " + z(zz.get("grau28")) + " %", g7 != null && g7 > 25 ? KiraSkala.ORANGE : Graphics.COLOR_WHITE);
        y = lr(dc, y, "Z4+Z5", z(zz.get("hart7")) + " / " + z(zz.get("hart28")) + " %", Graphics.COLOR_WHITE);
        var t = g7 != null && g7 > 25 ? "Viel Z3: mehr Z2" : "Verteilung passt";
        v.zeile(dc, y + 4, Graphics.FONT_XTINY, t, 0x9A9A9A);
    }

    // 13 Erholungstrend 14 Tage
    function erholung(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var e = teil("erh");
        if (e == null) { v.ohneDaten(dc, h, "ERHOLUNGSTREND"); return; }
        var w = dc.getWidth();
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "ERHOLUNG 14 T", 0x9A9A9A);
        y += 6;
        var reihen = [["HRV", "hrv", 0xB58CFF], ["Schlaf", "sl", 0x3A78CF], ["Score", "sc", 0x5DB75B]];
        var zh = (h * 0.15).toNumber();
        var x0 = (w * 0.36).toNumber(), x1 = (w * 0.70).toNumber();
        for (var i = 0; i < reihen.size(); i++) {
            var xs = e.get(reihen[i][1]);
            if (!(xs instanceof Array)) { continue; }
            var gr = grenzen(xs);
            if (i == 0 && e.get("lo") != null && e.get("hi") != null) {
                var lo = e.get("lo").toFloat(), hi = e.get("hi").toFloat();
                if (lo < gr[0]) { gr[0] = lo - 2; }
                if (hi > gr[1]) { gr[1] = hi + 2; }
                var yl = y + 6 + zh - 12 - ((zh - 12) * (lo - gr[0]) / (gr[1] - gr[0])).toNumber();
                var yh = y + 6 + zh - 12 - ((zh - 12) * (hi - gr[0]) / (gr[1] - gr[0])).toNumber();
                dc.setColor(0x2A2A2A, Graphics.COLOR_TRANSPARENT);
                dc.fillRectangle(x0, yh, x1 - x0, yl - yh);
            }
            polylinie(dc, x0, x1, y + 6, zh - 12, xs, gr[0], gr[1], reihen[i][2]);
            dc.setColor(0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.13, y + zh / 2, Graphics.FONT_XTINY, reihen[i][0], Graphics.TEXT_JUSTIFY_LEFT | Graphics.TEXT_JUSTIFY_VCENTER);
            var l = letzter(xs);
            var t = l == null ? "-" : i == 1 ? z(l / 10.0) : l.toString();
            dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w * 0.88, y + zh / 2, Graphics.FONT_SMALL, t, Graphics.TEXT_JUSTIFY_RIGHT | Graphics.TEXT_JUSTIFY_VCENTER);
            y += zh;
        }
        var bb = e.get("bb");
        var bl = bb instanceof Array ? letzter(bb) : null;
        var sch = e.get("schuld");
        var t2 = "Schuld " + z(sch) + " h" + (bl != null ? " · BB " + bl.toString() : "");
        v.zeile(dc, y + 4, Graphics.FONT_XTINY, t2, sch != null && sch > 3 ? KiraSkala.ORANGE : 0x9A9A9A);
    }

    // 14 Wochenbilanz Mo–So
    function woche(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var b = teil("wb");
        if (b == null || !(b.get("tage") instanceof Array)) { v.ohneDaten(dc, h, "WOCHENBILANZ"); return; }
        var w = dc.getWidth();
        var tage = b.get("tage") as Array;
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "WOCHE", 0x9A9A9A);
        var mx = 1, sIst = 0, sPlan = 0;
        for (var i = 0; i < tage.size(); i++) {
            var t = tage[i] as Dictionary;
            var a = t.get("ist"), p = t.get("plan");
            if (a != null && a > mx) { mx = a; }
            if (p != null && p > mx) { mx = p; }
            if (a != null && a > 0) { sIst += a; }
            if (p != null && p > 0 && (a == null || a < 0)) { sPlan += p; }
        }
        var hoch = (h * 0.22).toNumber();
        var x0 = (w * 0.16).toNumber();
        var bw = ((w * 0.68) / 7).toNumber();
        y += 6;
        for (var i = 0; i < tage.size(); i++) {
            var t = tage[i] as Dictionary;
            var a = t.get("ist"), p = t.get("plan");
            var x = x0 + i * bw + 4;
            if (p != null && p > 0) {
                var ph = (hoch * p / mx).toNumber();
                dc.setColor(0x666666, Graphics.COLOR_TRANSPARENT);
                dc.setPenWidth(2);
                dc.drawRectangle(x, y + hoch - ph, bw - 8, ph);
                dc.setPenWidth(1);
            }
            if (a != null && a > 0) {
                var ah = (hoch * a / mx).toNumber();
                dc.setColor(t.get("heute") == true ? KiraSkala.GRUEN : 0xBBBBBB, Graphics.COLOR_TRANSPARENT);
                dc.fillRectangle(x + 3, y + hoch - ah, bw - 14, ah);
            }
            dc.setColor(t.get("heute") == true ? Graphics.COLOR_WHITE : 0x9A9A9A, Graphics.COLOR_TRANSPARENT);
            dc.drawText(x + (bw - 8) / 2, y + hoch + 2, Graphics.FONT_XTINY, t.get("t").toString(), Graphics.TEXT_JUSTIFY_CENTER);
        }
        y += hoch + dc.getFontHeight(Graphics.FONT_XTINY) + 6;
        y = lr(dc, y, "Ist/Plan", sIst.toString() + " / " + sPlan.toString(), Graphics.COLOR_WHITE);
        var mo = b.get("mono");
        y = lr(dc, y, "Monotonie", z(mo), mo != null && mo > 2.0 ? KiraSkala.ORANGE : Graphics.COLOR_WHITE);
        var en = b.get("entl");
        lr(dc, y, "Entlastung", en != null ? en.toString() : "keine", en == null ? KiraSkala.ORANGE : Graphics.COLOR_WHITE);
    }

    // 15 Bereitschaft morgen
    function morgen(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var b = teil("bm");
        if (b == null) { v.ohneDaten(dc, h, "MORGEN"); return; }
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "MORGEN", 0x9A9A9A);
        var p = b.get("prog");
        y = v.zeile(dc, y, Graphics.FONT_NUMBER_MEDIUM, p != null ? p.toString() : "-", p != null && p >= 60 ? KiraSkala.GRUEN : p != null && p >= 40 ? 0xE8C547 : KiraSkala.ORANGE);
        var t = "heute " + z(b.get("heute")) + (b.get("progTag") != null ? " > " + b.get("progTag").toString() : "") + (b.get("mae") != null ? " (±" + z(b.get("mae")) + ")" : "");
        y = v.zeile(dc, y, Graphics.FONT_XTINY, t, Graphics.COLOR_WHITE);
        y += 6;
        y = v.zeile(dc, y, Graphics.FONT_XTINY, "Erfahrung nach Last:", 0x9A9A9A);
        var sz = b.get("szen");
        if (sz instanceof Array) {
            for (var i = 0; i < sz.size(); i++) {
                var s = sz[i] as Dictionary;
                var d1 = s.get("d1");
                var lab = s.get("lab").toString();
                if (lab.equals("Obergrenze")) { lab = "Max"; }
                y = lr(dc, y, lab + " " + s.get("ess").toString(), s.get("rd").toString() + " (" + vz(d1) + ")", d1 != null && d1 < -10 ? KiraSkala.ORANGE : Graphics.COLOR_WHITE);
            }
        }
    }

    // 16 Ziel-Countdown
    function ziel(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var b = teil("zc");
        if (b == null) { v.ohneDaten(dc, h, "ZIEL"); return; }
        var w = dc.getWidth();
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "ZIEL", 0x9A9A9A);
        y = v.zeile(dc, y, Graphics.FONT_SMALL, b.get("name").toString(), Graphics.COLOR_WHITE);
        y = v.zeile(dc, y, Graphics.FONT_NUMBER_MILD, z(b.get("tage")) + " T", KiraSkala.GRUEN);
        y = v.zeile(dc, y, Graphics.FONT_XTINY, b.get("datum") != null ? b.get("datum").toString() : "", 0x9A9A9A);
        var ist = b.get("ist"), prog = b.get("prog");
        if (ist instanceof Array && ist.size() > 1) {
            var alle = [];
            alle.addAll(ist);
            if (prog instanceof Array) { alle.addAll(prog); }
            var gr = grenzen(alle);
            var x0 = (w * 0.18).toNumber(), x1 = (w * 0.82).toNumber();
            var n = alle.size();
            var hoch = (h * 0.14).toNumber();
            var xm = x0 + ((x1 - x0) * (ist.size() - 1) / (n - 1)).toNumber();
            polylinie(dc, x0, xm, y + 6, hoch, ist, gr[0], gr[1], Graphics.COLOR_WHITE);
            if (prog instanceof Array && prog.size() > 0) {
                var pz = [ist[ist.size() - 1]];
                pz.addAll(prog);
                polylinie(dc, xm, x1, y + 6, hoch, pz, gr[0], gr[1], KiraSkala.BLAU);
            }
            y += hoch + 10;
        }
        var ce = b.get("ctlEnde");
        var t = "CTL jetzt " + z(ist instanceof Array && ist.size() > 0 ? ist[ist.size() - 1] : null) + " > Ziel " + z(ce);
        y = v.zeile(dc, y, Graphics.FONT_XTINY, t, Graphics.COLOR_WHITE);
        if (b.get("tf") != null) { v.zeile(dc, y, Graphics.FONT_XTINY, "Tourform " + z(b.get("tf")) + " %", 0x9A9A9A); }
    }

    // 17 Treffsicherheit 28 Tage
    function treffer(v as KiraView, dc as Graphics.Dc, h as Number) as Void {
        var b = teil("tr");
        if (b == null || b.get("n") == null || b.get("n") == 0) { v.ohneDaten(dc, h, "TREFFSICHERHEIT"); return; }
        var w = dc.getWidth();
        var y = v.zeile(dc, (h * 0.10).toNumber(), Graphics.FONT_XTINY, "TREFFER 28 T", 0x9A9A9A);
        var pd = b.get("pDrin");
        y = v.zeile(dc, y, Graphics.FONT_NUMBER_MEDIUM, z(pd) + "%", pd != null && pd >= 60 ? KiraSkala.GRUEN : 0xE8C547);
        y = v.zeile(dc, y, Graphics.FONT_XTINY, "Tage im Rahmen (" + z(b.get("n")) + ")", Graphics.COLOR_WHITE);
        // Stapelbalken drin/über/unter
        var n = b.get("n");
        var x0 = (w * 0.18).toNumber(), x1 = (w * 0.82).toNumber();
        var teile = [[b.get("unter"), KiraSkala.BLAU], [b.get("drin"), KiraSkala.GRUEN], [b.get("ueber"), KiraSkala.ORANGE]];
        var x = x0;
        y += 6;
        for (var i = 0; i < 3; i++) {
            var bw = ((x1 - x0) * teile[i][0] / n.toFloat()).toNumber();
            dc.setColor(teile[i][1], Graphics.COLOR_TRANSPARENT);
            dc.fillRectangle(x, y, bw, 14);
            x += bw;
        }
        y += 22;
        y = v.zeile(dc, y, Graphics.FONT_XTINY, "Readiness danach:", 0x9A9A9A);
        y = lr(dc, y, "drin (" + z(b.get("drin")) + ")", vz(b.get("dDrin")), Graphics.COLOR_WHITE);
        y = lr(dc, y, "darüber (" + z(b.get("ueber")) + ")", vz(b.get("dUeber")), KiraSkala.ORANGE);
        lr(dc, y, "darunter (" + z(b.get("unter")) + ")", vz(b.get("dUnter")), KiraSkala.BLAU);
    }
}
