import SwiftUI
import WidgetKit

/// Native Heute-Seite (optionale Etappe, 20.11.2026).
/// Zeigt den Tagesstand aus /api/widget (dieselbe Logik wie die Heute-Seite im Cockpit) und
/// funktioniert auch ohne Netz: dann mit dem zuletzt gespeicherten Stand (App Group, wie die Widgets).
struct HeuteNativView: View {
    @Environment(\.dismiss) private var dismiss
    var offline: Bool = false

    @State private var daten: WidgetDaten? = KiraStore.letzter()
    @State private var fehler: String?
    @State private var laeuft = false

    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(alignment: .leading, spacing: 14) {
                    hinweis
                    if let d = daten {
                        votumKarte(d)
                        morgenKarte(d)
                        lastKarte(d)
                        if let z = d.ziel { zielKarte(z) }
                        if let t = d.texte { texteKarte(t) }
                    } else {
                        Text("Noch kein Tagesstand gespeichert. Sobald die App einmal online war, steht er hier auch ohne Netz.")
                            .font(.callout)
                            .foregroundStyle(.secondary)
                            .padding(.top, 24)
                    }
                }
                .padding(16)
            }
            .background(KiraStil.hintergrund.ignoresSafeArea())
            .refreshable { await neu() }
            .navigationTitle("Heute")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .cancellationAction) {
                    Button(laeuft ? "Lädt …" : "Aktualisieren") { Task { await neu() } }
                        .disabled(laeuft)
                }
                ToolbarItem(placement: .confirmationAction) {
                    Button("Cockpit") { dismiss() }
                }
            }
            .task { if !offline { await neu() } }
        }
        .preferredColorScheme(.dark)
    }

    // MARK: - Bausteine

    @ViewBuilder private var hinweis: some View {
        let standText = "Stand \(KiraStil.datumKurz(daten?.datum)) \(KiraStil.uhrzeit(daten?.stand))"
        if let f = fehler {
            Label("\(f) Angezeigt wird der letzte gespeicherte Stand (\(standText)).", systemImage: "wifi.slash")
                .font(.footnote)
                .foregroundStyle(KiraStil.orange)
        } else if offline {
            Label("Kein Netz – letzter gespeicherter Stand (\(standText)).", systemImage: "wifi.slash")
                .font(.footnote)
                .foregroundStyle(KiraStil.orange)
        } else if daten != nil {
            Text(standText).font(.footnote).foregroundStyle(.secondary)
        }
    }

    private func karte<Inhalt: View>(_ titel: String, @ViewBuilder _ inhalt: () -> Inhalt) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(titel.uppercased()).font(.caption2.weight(.semibold)).foregroundStyle(.white.opacity(0.55))
            inhalt()
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(12)
        .background(Color.white.opacity(0.05), in: RoundedRectangle(cornerRadius: 12))
    }

    private func votumKarte(_ d: WidgetDaten) -> some View {
        karte("Tagesvotum") {
            HStack(spacing: 10) {
                Circle().fill(KiraStil.farbe(d.votum?.stufe)).frame(width: 18, height: 18)
                VStack(alignment: .leading, spacing: 2) {
                    Text(d.votum?.label ?? "–").font(.title2.weight(.bold)).foregroundStyle(.white)
                    Text(d.votum?.text ?? "").font(.subheadline).foregroundStyle(.white.opacity(0.75))
                }
            }
            ForEach(Array((d.votum?.gruende ?? []).enumerated()), id: \.offset) { _, g in
                Text("· \(g)").font(.footnote).foregroundStyle(.white.opacity(0.7))
            }
        }
    }

    private func wert(_ titel: String, _ text: String, farbe: Color = .white) -> some View {
        VStack(alignment: .leading, spacing: 2) {
            Text(titel).font(.caption2).foregroundStyle(.white.opacity(0.55))
            Text(text).font(.headline.monospacedDigit()).foregroundStyle(farbe)
        }
        .frame(maxWidth: .infinity, alignment: .leading)
    }

    private func morgenKarte(_ d: WidgetDaten) -> some View {
        let m = d.morgen
        return karte("Morgenwerte") {
            if m?.eingetragen == false {
                Text("Morgenwerte noch nicht eingetragen.").font(.footnote).foregroundStyle(KiraStil.orange)
            }
            HStack {
                wert("Readiness", KiraStil.zahl(m?.readiness), farbe: KiraStil.trFarbe(m?.readiness))
                wert("Schlaf", "\(KiraStil.zahl(m?.schlaf)) · \(KiraStil.zahl(m?.schlafH, stellen: 1)) h", farbe: KiraStil.schlafFarbe(m?.schlaf))
            }
            HStack {
                wert("HRV", "\(KiraStil.zahl(m?.hrv)) ms", farbe: KiraStil.stufeHrv(m?.hrvStufe))
                wert("Ruhepuls", rhrText(m), farbe: KiraStil.stufeG(m?.rhrStufe))
                wert("Befinden", KiraStil.zahl(m?.befinden))
            }
        }
    }

    private func rhrText(_ m: WidgetDaten.Morgen?) -> String {
        guard let r = m?.rhr else { return "–" }
        guard let dlt = m?.rhrDelta else { return KiraStil.zahl(r) }
        let vz = dlt > 0 ? "+" : ""
        return "\(KiraStil.zahl(r)) (\(vz)\(KiraStil.zahl(dlt)))"
    }

    private func lastKarte(_ d: WidgetDaten) -> some View {
        let l = d.last
        let planSport = [l?.sportHeute, l?.zoneHeute].compactMap { $0 }.filter { !$0.isEmpty }.joined(separator: " · ")
        let plan = [l?.planKurz ?? "", planSport].filter { !$0.isEmpty }.joined(separator: " · ") + (l?.planLage == "darueber" && l?.erledigt != true ? " · über dem Rahmen" : "")
        return karte("Last") {
            HStack {
                wert("Soll heute", "\(l?.soll ?? "–") ESS")
                wert(l?.erledigt == true ? "Ist (erledigt)" : "Ist", "\(KiraStil.zahl(l?.istHeute)) ESS", farbe: (l?.istHeute ?? 0) > 0 ? KiraStil.lastFarbe(l?.istHeute, lo: l?.sollMin, hi: l?.sollMax) : .white)
                wert("ACWR", KiraStil.zahl(l?.acwr, stellen: 2), farbe: acwrFarbe(l?.acwrBand))
            }
            if !plan.isEmpty {
                Text(plan).font(.footnote).foregroundStyle(.white.opacity(0.7))
            }
            if let w = l?.woche, !w.isEmpty { wocheBalken(w) }
        }
    }

    private func acwrFarbe(_ band: String?) -> Color {
        switch band {
        case "hoch": return KiraStil.gRot
        case "erhoeht": return KiraStil.gOrange
        case "niedrig": return KiraStil.gBlau
        case "korridor": return KiraStil.gGruen
        default: return .white
        }
    }

    private func wocheBalken(_ w: [WidgetDaten.Tag]) -> some View {
        let maxi = max(1, w.map { $0.ist ?? 0 }.max() ?? 1)
        return VStack(alignment: .leading, spacing: 4) {
            Text("Letzte 7 Tage (Ist-ESS)").font(.caption2).foregroundStyle(.white.opacity(0.55))
            HStack(alignment: .bottom, spacing: 6) {
                ForEach(Array(w.enumerated()), id: \.offset) { _, t in
                    VStack(spacing: 3) {
                        RoundedRectangle(cornerRadius: 3)
                            .fill(KiraStil.blau)
                            .frame(height: max(3, CGFloat((t.ist ?? 0) / maxi) * 56))
                        Text(KiraStil.datumKurz(t.d).prefix(2)).font(.system(size: 9)).foregroundStyle(.white.opacity(0.5))
                    }
                    .frame(maxWidth: .infinity)
                }
            }
            .frame(height: 72, alignment: .bottom)
        }
    }

    private func zielKarte(_ z: WidgetDaten.Ziel) -> some View {
        karte("Ziel") {
            HStack {
                wert(z.name, "in \(z.tage) Tagen")
                wert("Tourform", z.tourform == nil ? "–" : "\(KiraStil.zahl(z.tourform)) %", farbe: KiraStil.tourformFarbe(z.tourform))
            }
        }
    }

    private func texteKarte(_ t: WidgetDaten.Texte) -> some View {
        karte("Kira-Texte") {
            Text(texteZeile(t)).font(.footnote).foregroundStyle(.white.opacity(0.8))
            ForEach(Array((t.offen ?? []).enumerated()), id: \.offset) { _, o in
                Text("· \(o.label): \(o.phase)").font(.footnote).foregroundStyle(.white.opacity(0.65))
            }
            if let m = t.mac {
                Text(m.erreichbar == true ? (m.ollama == false ? "Mac erreichbar, Ollama nicht bereit" : "Mac bereit") : "Mac nicht erreichbar")
                    .font(.caption2).foregroundStyle(.white.opacity(0.55))
            }
        }
    }

    private func texteZeile(_ t: WidgetDaten.Texte) -> String {
        let c = t.cloud ?? 0
        let basis = "\(t.fertig) von \(t.gesamt) fertig"
        return c > 0 ? basis + " · \(c) aus der Cloud" : basis
    }

    // MARK: - Laden

    private func neu() async {
        laeuft = true
        defer { laeuft = false }
        do {
            daten = try await KiraAPI.laden()
            fehler = nil
            WidgetCenter.shared.reloadAllTimelines()
        } catch {
            daten = KiraStore.letzter()
            fehler = (error as? KiraFehler)?.errorDescription ?? "Keine Verbindung."
        }
    }
}
