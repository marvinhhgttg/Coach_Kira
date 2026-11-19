import WidgetKit
import SwiftUI
import AppIntents

// MARK: - Zeitleiste

struct KiraEintrag: TimelineEntry {
    let date: Date
    let daten: WidgetDaten?
    let fehler: String?
}

struct KiraProvider: TimelineProvider {
    func placeholder(in context: Context) -> KiraEintrag {
        KiraEintrag(date: Date(), daten: .beispiel, fehler: nil)
    }

    func getSnapshot(in context: Context, completion: @escaping (KiraEintrag) -> Void) {
        if context.isPreview {
            completion(KiraEintrag(date: Date(), daten: KiraStore.letzter() ?? .beispiel, fehler: nil))
            return
        }
        Task { completion(await Self.eintrag()) }
    }

    func getTimeline(in context: Context, completion: @escaping (Timeline<KiraEintrag>) -> Void) {
        Task {
            let e = await Self.eintrag()
            completion(Timeline(entries: [e], policy: .after(Self.naechsteAktualisierung())))
        }
    }

    static func eintrag() async -> KiraEintrag {
        do {
            return KiraEintrag(date: Date(), daten: try await KiraAPI.laden(), fehler: nil)
        } catch {
            return KiraEintrag(date: Date(), daten: KiraStore.letzter(), fehler: error.localizedDescription)
        }
    }

    /// Morgens (6–10 Uhr) alle 15 min, sonst alle 30 min – iOS entscheidet am Ende selbst über das Budget.
    static func naechsteAktualisierung() -> Date {
        let h = Calendar.current.component(.hour, from: Date())
        let minuten = (6...10).contains(h) ? 15 : 30
        return Date().addingTimeInterval(Double(minuten) * 60)
    }
}

// MARK: - Knöpfe (nur Mac-Aufträge, nie Cloud)

struct GarminAbrufenIntent: AppIntent {
    static var title: LocalizedStringResource = "Garmin abrufen"
    static var description = IntentDescription("Gibt dem Mac den Auftrag, Garmin Connect sofort abzurufen.")

    func perform() async throws -> some IntentResult {
        try await KiraAPI.aktion("garmin")
        return .result()
    }
}

struct BriefingLokalIntent: AppIntent {
    static var title: LocalizedStringResource = "KI-Analyse lokal"
    static var description = IntentDescription("Gibt dem Mac den Auftrag, die KI-Analyse lokal zu erzeugen.")

    func perform() async throws -> some IntentResult {
        try await KiraAPI.aktion("briefing_lokal")
        return .result()
    }
}

// MARK: - Ansichten

struct KiraWidgetView: View {
    @Environment(\.widgetFamily) private var family
    let eintrag: KiraEintrag

    var body: some View {
        Group {
            if let d = eintrag.daten {
                switch family {
                case .systemSmall: KleinView(d: d)
                case .systemMedium: MittelView(d: d)
                case .systemLarge, .systemExtraLarge: GrossView(d: d)
                case .accessoryCircular: RundView(d: d)
                case .accessoryRectangular: RechteckView(d: d)
                case .accessoryInline: InlineView(d: d)
                @unknown default: KleinView(d: d)
                }
            } else {
                LeerView(fehler: eintrag.fehler)
            }
        }
        .containerBackground(for: .widget) { KiraStil.hintergrund }
    }
}

private struct LeerView: View {
    let fehler: String?
    var body: some View {
        VStack(alignment: .leading, spacing: 4) {
            Text("Coach Kira").font(.caption.bold()).foregroundStyle(.white)
            Text(fehler ?? "Noch keine Daten.")
                .font(.caption2)
                .foregroundStyle(.white.opacity(0.7))
            Text("App öffnen und mit dem Cockpit koppeln.")
                .font(.caption2)
                .foregroundStyle(.white.opacity(0.5))
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

private struct Kopf: View {
    let d: WidgetDaten
    var body: some View {
        let farbe = KiraStil.farbe(d.votum?.stufe)
        HStack(alignment: .firstTextBaseline, spacing: 6) {
            Circle().fill(farbe).frame(width: 9, height: 9)
            Text(KiraStil.farbname(d.votum?.stufe)).font(.headline.weight(.heavy)).foregroundStyle(farbe)
            Text("· \(d.votum?.text ?? "–")").font(.subheadline).foregroundStyle(.white).lineLimit(1)
            Spacer(minLength: 4)
            Text(KiraStil.datumKurz(d.datum)).font(.caption2).foregroundStyle(.white.opacity(0.6))
        }
    }
}

private struct Wert: View {
    let titel: String
    let wert: String
    var farbe: Color = .white
    var body: some View {
        VStack(alignment: .leading, spacing: 1) {
            Text(titel).font(.caption2).foregroundStyle(.white.opacity(0.55))
            Text(wert).font(.subheadline.weight(.semibold).monospacedDigit()).foregroundStyle(farbe)
        }
    }
}

private func hrvText(_ m: WidgetDaten.Morgen?) -> String {
    guard let m else { return "–" }
    switch m.hrvStufe {
    case "rot": return "unter Band"
    case "orange": return "unausgegl."
    case "gelb": return "knapp"
    case "gruen", "blau": return "ausgeglichen"
    default: return m.hrv.map { KiraStil.zahl($0) + " ms" } ?? "–"
    }
}

private func essZeile(_ l: WidgetDaten.Last?) -> String {
    guard let l else { return "ESS heute: –" }
    let soll = KiraStil.zahl(l.sollHeute)
    if l.erledigt == true { return "ESS heute: Soll \(soll) · Ist \(KiraStil.zahl(l.istHeute))" }
    let sport = [l.sportHeute, l.zoneHeute].compactMap { $0 }.filter { !$0.isEmpty }.joined(separator: " ")
    return "ESS heute: Soll \(soll)" + (sport.isEmpty ? "" : " · \(sport)")
}

private func texteZeile(_ t: WidgetDaten.Texte?) -> String {
    guard let t, t.gesamt > 0 else { return "Kira-Texte: noch nichts offen" }
    if let o = t.offen?.first {
        let p = o.phase == "lokal" ? "lokal in Arbeit" : o.phase == "cloud" ? "Cloud-Ersatz" : "wartet auf Mac"
        return "Kira-Texte: \(t.fertig)/\(t.gesamt) · \(o.label) \(p)"
    }
    let wo = (t.cloud ?? 0) > 0 ? "davon \(t.cloud ?? 0) Cloud" : "lokal"
    return "Kira-Texte: \(t.fertig)/\(t.gesamt) fertig (\(wo))"
}

private struct KleinView: View {
    let d: WidgetDaten
    var body: some View {
        let farbe = KiraStil.farbe(d.votum?.stufe)
        VStack(alignment: .leading, spacing: 3) {
            HStack {
                Text("HEUTE").font(.caption2.bold()).foregroundStyle(.white.opacity(0.6))
                Spacer()
                Circle().fill(farbe).frame(width: 10, height: 10)
            }
            Text(KiraStil.farbname(d.votum?.stufe)).font(.title.weight(.heavy)).foregroundStyle(farbe).minimumScaleFactor(0.7)
            Text(d.votum?.text ?? "–").font(.caption).foregroundStyle(.white).lineLimit(2)
            Spacer(minLength: 0)
            Text("Readiness \(KiraStil.zahl(d.morgen?.readiness))").font(.caption.weight(.semibold).monospacedDigit()).foregroundStyle(.white)
            Text("Soll \(KiraStil.zahl(d.last?.sollHeute)) ESS · \(KiraStil.uhrzeit(d.stand))").font(.caption2).foregroundStyle(.white.opacity(0.55))
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

private struct MittelView: View {
    let d: WidgetDaten
    var body: some View {
        VStack(alignment: .leading, spacing: 6) {
            Kopf(d: d)
            HStack(spacing: 14) {
                Wert(titel: "Readiness", wert: KiraStil.zahl(d.morgen?.readiness))
                Wert(titel: "Schlaf", wert: KiraStil.zahl(d.morgen?.schlaf))
                Wert(titel: "HRV", wert: hrvText(d.morgen))
                Wert(titel: "RHR", wert: KiraStil.zahl(d.morgen?.rhr))
            }
            Text(essZeile(d.last)).font(.caption.monospacedDigit()).foregroundStyle(.white)
            Link(destination: URL(string: "coachkira://eingabe")!) {
                Text(texteZeile(d.texte)).font(.caption2).foregroundStyle(.white.opacity(0.6)).lineLimit(1)
            }
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

private struct Balken: View {
    let woche: [WidgetDaten.Tag]
    var body: some View {
        let werte = woche.map { $0.ist ?? 0 }
        let maxW = max(werte.max() ?? 1, 1)
        HStack(alignment: .bottom, spacing: 4) {
            ForEach(Array(werte.enumerated()), id: \.offset) { i, w in
                RoundedRectangle(cornerRadius: 2)
                    .fill(i == werte.count - 1 ? KiraStil.blau : Color.white.opacity(0.45))
                    .frame(width: 12, height: max(3, CGFloat(w / maxW) * 34))
            }
        }
        .frame(height: 36, alignment: .bottom)
    }
}

private struct GrossView: View {
    let d: WidgetDaten
    var body: some View {
        VStack(alignment: .leading, spacing: 8) {
            Kopf(d: d)
            HStack(spacing: 14) {
                Wert(titel: "Readiness", wert: KiraStil.zahl(d.morgen?.readiness))
                Wert(titel: "Schlaf", wert: KiraStil.zahl(d.morgen?.schlaf))
                Wert(titel: "HRV", wert: hrvText(d.morgen))
                Wert(titel: "RHR", wert: KiraStil.zahl(d.morgen?.rhr))
            }
            Text(essZeile(d.last)).font(.caption.monospacedDigit()).foregroundStyle(.white)
            HStack(alignment: .bottom) {
                VStack(alignment: .leading, spacing: 2) {
                    Text("ACWR \(KiraStil.zahl(d.last?.acwr, stellen: 2))")
                        .font(.caption.weight(.semibold).monospacedDigit()).foregroundStyle(.white)
                    Text(bandText(d.last?.acwrBand)).font(.caption2).foregroundStyle(.white.opacity(0.6))
                }
                Spacer()
                VStack(alignment: .trailing, spacing: 2) {
                    Balken(woche: d.last?.woche ?? [])
                    Text("Last 7 Tage (Ist)").font(.caption2).foregroundStyle(.white.opacity(0.5))
                }
            }
            if let z = d.ziel {
                Link(destination: URL(string: "coachkira://ziele")!) {
                    VStack(alignment: .leading, spacing: 1) {
                        Text("Nächstes Ziel: \(z.name)").font(.caption.weight(.semibold)).foregroundStyle(.white).lineLimit(1)
                        Text("in \(z.tage) Tagen" + (z.tourform.map { " · Tourform \(KiraStil.zahl($0)) %" } ?? ""))
                            .font(.caption2).foregroundStyle(.white.opacity(0.6))
                    }
                }
            }
            Link(destination: URL(string: "coachkira://eingabe")!) {
                Text(texteZeile(d.texte)).font(.caption2).foregroundStyle(.white.opacity(0.6)).lineLimit(1)
            }
            Spacer(minLength: 0)
            HStack(spacing: 8) {
                Button(intent: GarminAbrufenIntent()) {
                    Label("Garmin abrufen", systemImage: "arrow.down.circle").font(.caption2.weight(.semibold))
                }
                .tint(KiraStil.blau)
                Button(intent: BriefingLokalIntent()) {
                    Label("Analyse lokal", systemImage: "desktopcomputer").font(.caption2.weight(.semibold))
                }
                .tint(KiraStil.gruen)
                Spacer()
                Text("Stand \(KiraStil.uhrzeit(d.stand))").font(.caption2).foregroundStyle(.white.opacity(0.45))
            }
            .buttonStyle(.bordered)
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }

    private func bandText(_ b: String?) -> String {
        switch b {
        case "korridor": return "im Korridor"
        case "niedrig": return "unter dem Korridor"
        case "erhoeht": return "erhöht – keine Qualität"
        case "hoch": return "hoch – maximal Easy"
        default: return ""
        }
    }
}

private struct RundView: View {
    let d: WidgetDaten
    var body: some View {
        let r = d.morgen?.readiness ?? 0
        Gauge(value: min(max(r, 0), 100), in: 0...100) {
            Text("R")
        } currentValueLabel: {
            Text(KiraStil.zahl(d.morgen?.readiness))
        }
        .gaugeStyle(.accessoryCircular)
        .tint(KiraStil.farbe(d.votum?.stufe))
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

private struct RechteckView: View {
    let d: WidgetDaten
    var body: some View {
        VStack(alignment: .leading, spacing: 1) {
            Text("Kira · \(KiraStil.farbname(d.votum?.stufe))").font(.headline.weight(.heavy))
            Text("Soll \(KiraStil.zahl(d.last?.sollHeute)) ESS · ACWR \(KiraStil.zahl(d.last?.acwr, stellen: 2))").font(.caption)
            Text("Readiness \(KiraStil.zahl(d.morgen?.readiness)) · Schlaf \(KiraStil.zahl(d.morgen?.schlaf))").font(.caption2)
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

private struct InlineView: View {
    let d: WidgetDaten
    var body: some View {
        Text("Kira: \(KiraStil.farbname(d.votum?.stufe)) · \(KiraStil.zahl(d.last?.sollHeute)) ESS")
    }
}

// MARK: - Widget

struct KiraWidget: Widget {
    let kind = "KiraWidget"

    var body: some WidgetConfiguration {
        StaticConfiguration(kind: kind, provider: KiraProvider()) { e in
            KiraWidgetView(eintrag: e)
        }
        .configurationDisplayName("Coach Kira")
        .description("Tagesentscheid, Morgenwerte, Last, Ziel und Kira-Texte.")
        .supportedFamilies([.systemSmall, .systemMedium, .systemLarge, .accessoryCircular, .accessoryRectangular, .accessoryInline])
    }
}

@main
struct KiraWidgetBundle: WidgetBundle {
    var body: some Widget {
        KiraWidget()
    }
}
