// Coach Kira – Tagesvotum auf der Garmin-Uhr (fēnix 8 Pro), 21.11.2026.
// Holt den Tagesstand aus dem Cockpit (/api/widget?uhr=1, Geräteschlüssel) – im Hintergrund alle 15 min
// und beim Öffnen. Keine KI auf der Uhr: angezeigt wird das fertige Votum des Rechenkerns.
import Toybox.ActivityMonitor;
import Toybox.Application;
import Toybox.Application.Storage;
import Toybox.Background;
import Toybox.Communications;
import Toybox.Complications;
import Toybox.Notifications;
import Toybox.Lang;
import Toybox.PersistedContent;
import Toybox.System;
import Toybox.Time;
import Toybox.Time.Gregorian;
import Toybox.WatchUi;

(:background :glance)
const KIRA_BASIS = "https://coachkira.pplx.app/port/5000";

(:background :glance)
class KiraApp extends Application.AppBase {
    function initialize() {
        AppBase.initialize();
    }

    function onStart(state as Dictionary?) as Void {
    }

    function onStop(state as Dictionary?) as Void {
    }

    function getInitialView() {
        hintergrundAn();
        var v = new KiraView();
        return [v, new KiraDelegate(v)];
    }

    (:glance)
    function getGlanceView() {
        return [new KiraGlance()];
    }

    function getServiceDelegate() {
        return [new KiraHintergrund()];
    }

    // Ergebnis des Hintergrunddienstes übernehmen
    function onBackgroundData(data) {
        if (data instanceof Dictionary) {
            if (data.hasKey("votum")) {
                Storage.setValue("d", data);
                Storage.setValue("at", Time.now().value());
                Storage.deleteValue("fehler");
                kiraKomplikation(data);
                takt(data);
            } else if (data.hasKey("fehler")) {
                Storage.setValue("fehler", data.get("fehler"));
            }
        }
        WatchUi.requestUpdate();
    }

    // 24.11.2026: Stufe 3 – morgens (5–10 Uhr), bei vorläufigem Votum und bis 2 h nach einer neuen Einheit alle 5 min, sonst alle 15 min
    function takt(d as Dictionary) as Void {
        try {
            var schnell = false;
            var h = Gregorian.info(Time.now(), Time.FORMAT_SHORT).hour;
            if (h >= 5 && h < 10) { schnell = true; }
            var v = d.get("votum");
            if (v instanceof Dictionary && v.get("vorlaeufig") == true) { schnell = true; }
            var l = d.get("last");
            var ist = l instanceof Dictionary && l.get("istHeute") != null ? l.get("istHeute").toNumber() : 0;
            var alt = Storage.getValue("istAlt");
            if (alt == null || ist != alt) { Storage.setValue("istAlt", ist); if (alt != null && ist > alt) { Storage.setValue("istNeuAt", Time.now().value()); } }
            var neu = Storage.getValue("istNeuAt");
            if (neu != null && Time.now().value() - neu < 2 * 3600) { schnell = true; }
            var soll = schnell ? 5 * 60 : 15 * 60;
            if (Storage.getValue("takt") != soll) {
                Background.registerForTemporalEvent(new Time.Duration(soll));
                Storage.setValue("takt", soll);
            }
        } catch (e) { }
    }

    function hintergrundAn() as Void {
        if (Storage.getValue("tok") == null) {
            return;
        }
        try {
            if (Background.getTemporalEventRegisteredTime() == null) {
                Background.registerForTemporalEvent(new Time.Duration(15 * 60));
            }
        } catch (e) {
        }
    }
}

// 22.11.2026: Komplikation fürs Zifferblatt (Face It): Votum als Kurzlabel, Soll-Rahmen als Wert
(:background :glance)
function kiraKomplikation(d as Dictionary) as Void {
    try {
        if (!(Toybox has :Complications) || !(Complications has :updateComplication)) { return; }
        var v = d.get("votum");
        var l = d.get("last");
        var lab = "KIRA";
        if (v instanceof Dictionary && v.get("vote") != null) {
            var k = v.get("vote").toString();
            lab = k.equals("QUALITY") ? "QUAL" : k;
        }
        var wert = "-";
        if (l instanceof Dictionary) {
            var lo = l.get("sollMin");
            var hi = l.get("sollMax");
            if (hi != null) { wert = (lo != null ? lo.toNumber().toString() + "-" : "") + hi.toNumber().toString(); }
            else if (l.get("sollHeute") != null) { wert = l.get("sollHeute").toNumber().toString(); }
        }
        try {
            Complications.updateComplication(0, { :value => wert, :shortLabel => lab, :unit => "ESS" });
        } catch (e0) { }
        // 22.11.2026: Komplikation 1 (privat) für das Coach-Kira-Zifferblatt:
        // vote|stufe|lo|hi|ist|datum|ziel|zielTage|readiness
        var t = "";
        var st = v instanceof Dictionary && v.get("stufe") != null ? v.get("stufe").toString() : "";
        var lo2 = "", hi2 = "", ist = "";
        if (l instanceof Dictionary) {
            if (l.get("sollMin") != null) { lo2 = l.get("sollMin").toNumber().toString(); }
            if (l.get("sollMax") != null) { hi2 = l.get("sollMax").toNumber().toString(); }
            else if (l.get("sollHeute") != null) { hi2 = l.get("sollHeute").toNumber().toString(); }
            if (l.get("istHeute") != null) { ist = l.get("istHeute").toNumber().toString(); }
        }
        var z = d.get("ziel");
        var zn = "", zt = "";
        if (z instanceof Dictionary) {
            if (z.get("name") != null) { zn = z.get("name").toString(); }
            if (z.get("tage") != null) { zt = z.get("tage").toNumber().toString(); }
        }
        var m = d.get("morgen");
        var rd = m instanceof Dictionary && m.get("readiness") != null ? m.get("readiness").toNumber().toString() : "";
        // ab 22.11. abends zusätzlich: |plan|sport|zone|ausdauerwert|ausdauer28
        var pl = "", sp = "", zo = "", es = "", es28 = "";
        if (l instanceof Dictionary) {
            if (l.get("planHeute") != null) { pl = l.get("planHeute").toNumber().toString(); }
            if (l.get("sportHeute") != null) { sp = l.get("sportHeute").toString(); }
            if (l.get("zoneHeute") != null) { zo = l.get("zoneHeute").toString(); }
        }
        var fo = d.get("form");
        if (fo instanceof Dictionary) {
            if (fo.get("es") != null) { es = fo.get("es").toNumber().toString(); }
            if (fo.get("es28") != null) { es28 = fo.get("es28").toNumber().toString(); }
        }
        // 23.11.2026: |14 Uhrzeit des Garmin-Bereitschaftswerts (z. B. 05:50), leer wenn unbekannt
        var rz = m instanceof Dictionary && m.get("readinessZeit") != null ? m.get("readinessZeit").toString() : "";
        // 23.11.2026: |15 Wochen-Ist |16 Wochenbudget (Wochen-Ring auf den Zifferblättern)
        var wi = "", wb = "";
        var wo = d.get("woche");
        if (wo instanceof Dictionary) {
            if (wo.get("ist") != null) { wi = wo.get("ist").toNumber().toString(); }
            if (wo.get("budget") != null) { wb = wo.get("budget").toNumber().toString(); }
        }
        t = (lab.equals("KIRA") ? "" : lab) + "|" + st + "|" + lo2 + "|" + hi2 + "|" + ist + "|" + (d.get("datum") != null ? d.get("datum").toString() : "") + "|" + zn + "|" + zt + "|" + rd
            + "|" + pl + "|" + sp + "|" + zo + "|" + es + "|" + es28 + "|" + rz + "|" + wi + "|" + wb
            + "|" + (v instanceof Dictionary && v.get("vorlaeufig") == true ? "1" : "");
        Complications.updateComplication(1, { :value => t });
        // Lange Texte lehnt die Uhr ab – Diagrammdaten daher kompakt und getrennt (je 1 Zeichen pro Wert)
        try { Complications.updateComplication(2, { :value => sieben(d, l) }); } catch (e2) { }
        try { Complications.updateComplication(3, { :value => formKurve(d) }); } catch (e3) { }
    } catch (e) {
    }
}

// 7 Tage Last (rollierend, heute zuletzt) mit Soll-Rahmen je Tag fürs Zeigerzifferblatt:
// kompakt "IIIIIII;LLLLLLL;HHHHHHH" (ESS / 8) – Rahmen aus der letzten Vollabfrage der App (nur wenn von heute)
(:background :glance)
function sieben(d as Dictionary, l) as String {
    var ist = [] as Array;
    var lo = [] as Array;
    var hi = [] as Array;
    try {
        var u = Storage.getValue("u");
        var ud = Storage.getValue("uDatum");
        var l14 = u instanceof Dictionary ? u.get("l14") : null;
        if (l14 instanceof Dictionary && ud != null && d.get("datum") != null && ud.toString().equals(d.get("datum").toString())) {
            var a = l14.get("ist");
            var b = l14.get("lo");
            var c = l14.get("hi");
            if (a instanceof Array && b instanceof Array && c instanceof Array) {
                var n = a.size();
                for (var i = (n > 7 ? n - 7 : 0); i < n; i++) { ist.add(a[i]); lo.add(b[i]); hi.add(c[i]); }
            }
        }
        if (ist.size() == 0 && l instanceof Dictionary && l.get("woche") instanceof Array) {
            var w = l.get("woche") as Array;
            for (var i = 0; i < w.size(); i++) {
                var x = w[i];
                ist.add(x instanceof Dictionary && x.get("ist") != null ? x.get("ist") : 0);
            }
        }
        // heutigen Wert immer frisch
        if (ist.size() > 0 && l instanceof Dictionary && l.get("istHeute") != null) { ist[ist.size() - 1] = l.get("istHeute"); }
    } catch (e) { }
    return kodiert(ist, 8) + ";" + kodiert(lo, 8) + ";" + kodiert(hi, 8);
}

// Werte als je 1 Zeichen (0-9A-Za-z = 0..61), Wert / teiler, fehlend = "_"
(:background :glance)
const ZEICHEN = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";

(:background :glance)
function kodiert(xs as Array, teiler as Number) as String {
    var s = "";
    for (var i = 0; i < xs.size(); i++) {
        var v = xs[i];
        if (v == null || v.toNumber() < 0) { s += "_"; continue; }
        var k = (v.toFloat() / teiler + 0.5).toNumber();
        if (k > 61) { k = 61; }
        s += ZEICHEN.substring(k, k + 1);
    }
    return s;
}

// 23.11.2026: Form fürs Zeigerzifferblatt – 7 Tage zurück + 7 Tage Plan-Prognose (Kiras Formel, vom Server)
// "v2,max,ctlHeute,atlHeute,acwrMax,acwrTag,planSumme;CCCCCCCccccccc;AAAAAAAaaaaaaa;PPPPPPP"
// Kurven auf 0..61 zwischen 0 und max skaliert, Plan-ESS / 8; Prognose fehlt = "_"
(:background :glance)
function formKurve(d as Dictionary) as String {
    var c = [] as Array;
    var a = [] as Array;
    var pe = [] as Array;
    var pa = [] as Array;
    var pc = [] as Array;
    var am = 0.0;
    var at = -1;
    try {
        var u = Storage.getValue("u");
        var ud = Storage.getValue("uDatum");
        var vonHeute = ud != null && d.get("datum") != null && ud.toString().equals(d.get("datum").toString());
        var f = u instanceof Dictionary ? u.get("form") : null;
        if (f instanceof Dictionary && f.get("ctl") instanceof Array && f.get("atl") instanceof Array) {
            var xc = f.get("ctl") as Array;
            var xa = f.get("atl") as Array;
            var n = xc.size() < xa.size() ? xc.size() : xa.size();
            for (var i = (n > 7 ? n - 7 : 0); i < n; i++) { c.add(xc[i]); a.add(xa[i]); }
        }
        var fo = d.get("form");
        if (fo instanceof Dictionary && fo.get("ctl") != null && fo.get("atl") != null) {
            if (c.size() == 0) { c.add(fo.get("ctl")); a.add(fo.get("atl")); }
            else { c[c.size() - 1] = fo.get("ctl"); a[a.size() - 1] = fo.get("atl"); }
        }
        var p7 = u instanceof Dictionary ? u.get("plan7") : null;
        if (vonHeute && p7 instanceof Dictionary && p7.get("ess") instanceof Array && p7.get("atl") instanceof Array && p7.get("ctl") instanceof Array) {
            pe = p7.get("ess") as Array;
            pa = p7.get("atl") as Array;
            pc = p7.get("ctl") as Array;
            var ac = p7.get("acwr");
            if (ac instanceof Array) {
                for (var i = 0; i < ac.size(); i++) { if (ac[i] != null && ac[i].toFloat() > am) { am = ac[i].toFloat(); at = i; } }
            }
        }
    } catch (e) { }
    var mx = 0.0;
    var alle = [c, a, pa, pc];
    for (var k = 0; k < 4; k++) {
        var xs = alle[k] as Array;
        for (var i = 0; i < xs.size(); i++) { var v = xs[i]; if (v != null && v.toFloat() > mx) { mx = v.toFloat(); } }
    }
    if (c.size() == 0 || mx <= 0) { return ""; }
    mx = mx * 1.08;
    var sc = "";
    var sa = "";
    for (var i = 0; i < c.size(); i++) { sc += zeichen(c[i], 0.0, mx); sa += zeichen(a[i], 0.0, mx); }
    for (var i = 0; i < pc.size(); i++) { sc += zeichen(pc[i], 0.0, mx); }
    for (var i = 0; i < pa.size(); i++) { sa += zeichen(pa[i], 0.0, mx); }
    var sum = 0;
    for (var i = 0; i < pe.size(); i++) { if (pe[i] != null) { sum += pe[i].toNumber(); } }
    var cl = c[c.size() - 1];
    var al = a[a.size() - 1];
    return "v2," + mx.toNumber().toString() + "," + (cl == null ? "" : cl.toNumber().toString()) + "," + (al == null ? "" : al.toNumber().toString())
        + "," + (at >= 0 ? (am * 100 + 0.5).toNumber().toString() : "") + "," + (at >= 0 ? at.toString() : "") + "," + (pe.size() > 0 ? sum.toString() : "")
        + ";" + sc + ";" + sa + ";" + kodiert(pe, 8);
}

(:background :glance)
function zeichen(v, mn as Float, mx as Float) as String {
    if (v == null || v.toFloat() < 0) { return "_"; }
    var k = ((v.toFloat() - mn) / (mx - mn) * 61 + 0.5).toNumber();
    if (k < 0) { k = 0; }
    if (k > 61) { k = 61; }
    return ZEICHEN.substring(k, k + 1);
}

(:background :glance)
function liste(xs as Array) as String {
    var s = "";
    for (var i = 0; i < xs.size(); i++) {
        var v = xs[i];
        s += (i > 0 ? "," : "") + (v == null ? "" : v.toNumber().toString());
    }
    return s;
}

// 23.11.2026: Garmins Erholungszeit live von der Uhr (Minuten) und daraus „erholt ab“
(:background :glance)
function erholMinuten() as Number? {
    // Uhr-Apps dürfen keine Komplikationen lesen – Erholungszeit (Stunden) aus ActivityMonitor
    try {
        var i = ActivityMonitor.getInfo();
        if (i has :timeToRecovery && i.timeToRecovery != null) { return (i.timeToRecovery as Number) * 60; }
    } catch (e) { }
    return null;
}

(:background :glance)
function erholtAb(min as Number) as String {
    if (min <= 0) { return "jetzt"; }
    var jetzt = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
    var i = Gregorian.info(Time.now().add(new Time.Duration(min * 60)), Time.FORMAT_SHORT);
    var wt = ["", "So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
    var tag = i.day == jetzt.day ? "heute" : wt[i.day_of_week as Number];
    // Erholungszeit kommt in ganzen Stunden – daher nur auf die Stunde genau
    return tag + " ~" + i.hour.format("%02d") + " Uhr";
}

// Nächster geplanter Tag mit Intensität (Zone 3 und höher) gegen die Erholungszeit: [Text, passt?] oder null
(:background :glance)
function intensitaetCheck(d as Dictionary, min as Number?) as Array? {
    var n = d.get("naechste");
    if (!(n instanceof Array)) { return null; }
    for (var k = 0; k < n.size(); k++) {
        var x = n[k];
        if (!(x instanceof Dictionary)) { continue; }
        var z = x.get("zone") != null ? x.get("zone").toString() : "";
        var e = x.get("ess");
        if (e == null || e.toNumber() <= 0) { continue; }
        if (z.find("Z3") == null && z.find("Z4") == null && z.find("Z5") == null) { continue; }
        var iso = x.get("d").toString();
        var m = Gregorian.moment({ :year => iso.substring(0, 4).toNumber(), :month => iso.substring(5, 7).toNumber(), :day => iso.substring(8, 10).toNumber(), :hour => 8 });
        // Gregorian.moment rechnet in UTC – 8 Uhr UTC ≈ 9 Uhr Berlin, als Trainingsbeginn genügt das
        var wt = ["", "So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
        var info = Gregorian.info(m, Time.FORMAT_SHORT);
        var name = wt[info.day_of_week as Number] + " " + (x.get("sport") != null ? x.get("sport").toString() : "") + " " + z;
        var passt = min == null || Time.now().add(new Time.Duration(min * 60)).compare(m) <= 0;
        return [name, passt];
    }
    return null;
}

// Hintergrund: höchstens alle 15 min, kleine Antwort
(:background)
class KiraHintergrund extends System.ServiceDelegate {
    function initialize() {
        ServiceDelegate.initialize();
    }

    function onTemporalEvent() as Void {
        var tok = Storage.getValue("tok");
        if (tok == null) {
            Background.exit(null);
            return;
        }
        Communications.makeWebRequest(
            KIRA_BASIS + "/api/widget?uhr=1",
            null,
            {
                :method => Communications.HTTP_REQUEST_METHOD_GET,
                :headers => { "Authorization" => "Bearer " + tok },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON
            },
            method(:fertig)
        );
    }

    function fertig(code as Number, data as Null or ByteArray or Dictionary or String or PersistedContent.Iterator) as Void {
        if (code == 200 && data instanceof Dictionary) {
            kiraKomplikation(data);
            melden(data);
            fazitMelden(data);
            hinweiseMelden(data);
            Background.exit(data);
        } else {
            Background.exit({ "fehler" => code });
        }
    }

    // 23.11.2026: Kiras Fazit nach dem Speichern – einmal je neuer Einheit (Server erkennt sie ca. 10 min nach dem Abgleich)
    function fazitMelden(d as Dictionary) as Void {
        try {
            var f = d.get("fazit");
            if (!(f instanceof Dictionary) || f.get("id") == null) { return; }
            var id = f.get("id").toString();
            var alt = Storage.getValue("fazitId");
            if (alt != null && alt.toString().equals(id)) { return; }
            Storage.setValue("fazitId", id);
            var jetzt = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
            if (jetzt.hour < 6 || jetzt.hour >= 23) { return; }
            var text = f.get("text") != null ? f.get("text").toString() : "";
            var min = erholMinuten();
            if (min != null) { text += " Erholt ab " + erholtAb(min) + "."; }
            var ic = intensitaetCheck(d, min);
            if (ic != null && !(ic[1] as Boolean)) { text += " " + (ic[0] as String) + " kommt zu früh."; }
            var titel = f.get("titel") != null ? f.get("titel").toString() : "Kira";
            var unter = f.get("unter") != null ? f.get("unter").toString() : "";
            if (Toybox has :Notifications && Notifications has :showNotification) {
                Notifications.showNotification(titel, unter, { :body => text, :dismissPrevious => false });
            } else if (Background has :requestApplicationWake) {
                Background.requestApplicationWake(titel + " · " + unter);
            }
        } catch (e) {
        }
    }

    // 24.11.2026: Hinweise vom Server (Befinden, Licht aus, ACWR/Ruhepuls, Plan-Vorschlag, vorläufiges Votum) – je id einmal
    function hinweiseMelden(d as Dictionary) as Void {
        try {
            var xs = d.get("hinweise");
            if (!(xs instanceof Array) || xs.size() == 0) { return; }
            var gesehen = Storage.getValue("hinw");
            var liste = gesehen instanceof Array ? gesehen : [];
            for (var i = 0; i < xs.size(); i++) {
                var h = xs[i];
                if (!(h instanceof Dictionary) || h.get("id") == null) { continue; }
                var id = h.get("id").toString();
                var neu = true;
                for (var k = 0; k < liste.size(); k++) { if (liste[k].toString().equals(id)) { neu = false; break; } }
                if (!neu) { continue; }
                liste.add(id);
                while (liste.size() > 12) { liste = liste.slice(1, null); }
                Storage.setValue("hinw", liste);
                var titel = h.get("titel") != null ? h.get("titel").toString() : "Kira";
                var text = h.get("text") != null ? h.get("text").toString() : "";
                if (Toybox has :Notifications && Notifications has :showNotification) {
                    Notifications.showNotification(titel, "", { :body => text, :dismissPrevious => false });
                } else if (Background has :requestApplicationWake) {
                    Background.requestApplicationWake(titel);
                }
                return;   // höchstens ein Hinweis je Abruf
            }
        } catch (e) {
        }
    }

    // 22.11.2026: Morgenhinweis – einmal am Tag, sobald das Votum für heute mit Morgenwerten steht
    function melden(d as Dictionary) as Void {
        try {
            var datum = d.get("datum");
            var v = d.get("votum");
            var m = d.get("morgen");
            if (datum == null || !(v instanceof Dictionary) || v.get("vote") == null) { return; }
            if (m instanceof Dictionary && m.get("eingetragen") == false) { return; }
            var gemeldet = Storage.getValue("gemeldet");
            if (gemeldet != null && gemeldet.toString().equals(datum.toString())) { return; }
            var jetzt = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
            var heute = jetzt.year.format("%04d") + "-" + jetzt.month.format("%02d") + "-" + jetzt.day.format("%02d");
            if (!heute.equals(datum.toString()) || jetzt.hour < 5 || jetzt.hour >= 13) { return; }
            var l = d.get("last");
            var soll = "";
            if (l instanceof Dictionary && l.get("sollMax") != null) {
                soll = "Soll " + (l.get("sollMin") != null ? l.get("sollMin").toNumber().toString() + "-" : "bis ") + l.get("sollMax").toNumber().toString() + " ESS";
            }
            var titel = "Kira: " + v.get("label").toString();
            var gr = v.get("gruende");
            var text = gr instanceof Array && gr.size() > 0 ? gr[0].toString() : "";
            Storage.setValue("gemeldet", datum.toString());
            if (Toybox has :Notifications && Notifications has :showNotification) {
                Notifications.showNotification(titel, soll, { :body => text, :dismissPrevious => true });
            } else if (Background has :requestApplicationWake) {
                Background.requestApplicationWake(titel + " · " + soll);
            }
        } catch (e) {
        }
    }
}
