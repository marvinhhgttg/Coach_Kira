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
                case .systemLarge: GrossView(d: d)
                case .systemExtraLarge: ExtraGrossView(d: d)
                #if os(iOS)
                case .accessoryCircular: RundView(d: d)
                case .accessoryRectangular: RechteckView(d: d)
                case .accessoryInline: InlineView(d: d)
                #endif
                @unknown default: KleinView(d: d)
                }
            } else {
                LeerView(fehler: eintrag.fehler)
            }
        }
        .containerBackground(for: .widget) { KiraStil.hintergrund }
    }
}

struct LeerView: View {
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

struct Kopf: View {
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

struct Wert: View {
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
    if l.erledigt == true { return "ESS heute: Soll \(l.soll) · Ist \(KiraStil.zahl(l.istHeute))" }
    let sport = [l.sportHeute, l.zoneHeute].compactMap { $0 }.filter { !$0.isEmpty }.joined(separator: " ")
    let plan = l.planKurz.isEmpty ? "" : " · \(l.planKurz)" + (sport.isEmpty ? "" : " \(sport)") + (l.planLage == "darueber" ? " (zu viel)" : "")
    return "ESS heute: Soll \(l.soll)" + plan
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

struct KleinView: View {
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
            Text("Soll \(d.last?.soll ?? "–") ESS · \(KiraStil.uhrzeit(d.stand))").font(.caption2).foregroundStyle(.white.opacity(0.55))
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

struct MittelView: View {
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

struct Balken: View {
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

struct GrossView: View {
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

#if os(iOS)
struct RundView: View {
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

#endif

struct RechteckView: View {
    let d: WidgetDaten
    var body: some View {
        VStack(alignment: .leading, spacing: 1) {
            Text("Kira · \(KiraStil.farbname(d.votum?.stufe))").font(.headline.weight(.heavy))
            Text("Soll \(d.last?.soll ?? "–") ESS · ACWR \(KiraStil.zahl(d.last?.acwr, stellen: 2))").font(.caption)
            Text("Readiness \(KiraStil.zahl(d.morgen?.readiness)) · Schlaf \(KiraStil.zahl(d.morgen?.schlaf))").font(.caption2)
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

struct InlineView: View {
    let d: WidgetDaten
    var body: some View {
        Text("Kira: \(KiraStil.farbname(d.votum?.stufe)) · \(d.last?.soll ?? "–") ESS")
    }
}

// MARK: - Weitere Ansichten

struct ZielKleinView: View {
    let d: WidgetDaten
    var body: some View {
        VStack(alignment: .leading, spacing: 4) {
            Text("ZIELTOUR").font(.caption2.bold()).foregroundStyle(.white.opacity(0.6))
            if let z = d.ziel {
                Text(z.name).font(.subheadline.weight(.semibold)).foregroundStyle(.white).lineLimit(2).minimumScaleFactor(0.8)
                Spacer(minLength: 0)
                HStack(alignment: .firstTextBaseline, spacing: 3) {
                    Text("\(z.tage)").font(.title.weight(.heavy).monospacedDigit()).foregroundStyle(KiraStil.blau)
                    Text(z.tage == 1 ? "Tag" : "Tage").font(.caption).foregroundStyle(.white.opacity(0.7))
                }
                if let t = z.tourform {
                    ProgressView(value: min(max(t, 0), 100), total: 100).tint(tourformFarbe(t))
                    Text("Tourform \(KiraStil.zahl(t)) %").font(.caption2.monospacedDigit()).foregroundStyle(.white.opacity(0.7))
                }
            } else {
                Text("Keine Zieltour geplant").font(.caption).foregroundStyle(.white.opacity(0.7))
                Spacer(minLength: 0)
            }
        }
        .widgetURL(URL(string: "coachkira://ziele"))
    }
}

private func tourformFarbe(_ t: Double) -> Color {
    t >= 85 ? KiraStil.gruen : t >= 60 ? KiraStil.gelb : KiraStil.orange
}

struct ZielRechteckView: View {
    let d: WidgetDaten
    var body: some View {
        VStack(alignment: .leading, spacing: 1) {
            if let z = d.ziel {
                Text(z.name).font(.headline).lineLimit(1)
                Text("in \(z.tage) Tagen").font(.caption)
                Text("Tourform \(KiraStil.zahl(z.tourform)) %").font(.caption2)
            } else {
                Text("Keine Zieltour").font(.headline)
            }
        }
        .widgetURL(URL(string: "coachkira://ziele"))
    }
}

struct MacPunkt: View {
    let t: WidgetDaten.Texte?
    var body: some View {
        let ok = t?.mac?.erreichbar == true && t?.mac?.ollama != false
        HStack(spacing: 4) {
            Circle().fill(ok ? KiraStil.gruen : KiraStil.orange).frame(width: 7, height: 7)
            Text(ok ? "Mac bereit" : "Mac nicht erreichbar").font(.caption2).foregroundStyle(.white.opacity(0.6))
        }
    }
}

private func phaseText(_ p: String) -> String {
    switch p {
    case "lokal": return "lokal in Arbeit"
    case "cloud": return "Cloud-Ersatz läuft"
    default: return "wartet auf Mac"
    }
}

struct TexteKleinView: View {
    let d: WidgetDaten
    var body: some View {
        let t = d.texte
        let offen = t?.offen ?? []
        VStack(alignment: .leading, spacing: 4) {
            Text("KIRA-TEXTE").font(.caption2.bold()).foregroundStyle(.white.opacity(0.6))
            HStack(alignment: .firstTextBaseline, spacing: 2) {
                Text("\(t?.fertig ?? 0)").font(.title.weight(.heavy).monospacedDigit())
                    .foregroundStyle(offen.isEmpty ? KiraStil.gruen : KiraStil.orange)
                Text("/ \(t?.gesamt ?? 0) fertig").font(.caption).foregroundStyle(.white.opacity(0.7))
            }
            if let o = offen.first {
                Text(o.label).font(.caption2.weight(.semibold)).foregroundStyle(.white).lineLimit(1)
                Text(phaseText(o.phase)).font(.caption2).foregroundStyle(.white.opacity(0.6))
            } else if (t?.cloud ?? 0) > 0 {
                Text("davon \(t?.cloud ?? 0) aus der Cloud").font(.caption2).foregroundStyle(KiraStil.blau)
            } else {
                Text("alles lokal").font(.caption2).foregroundStyle(KiraStil.gruen)
            }
            Spacer(minLength: 0)
            MacPunkt(t: t)
        }
        .widgetURL(URL(string: "coachkira://eingabe"))
    }
}

struct TexteMittelView: View {
    let d: WidgetDaten
    var body: some View {
        let t = d.texte
        let offen = t?.offen ?? []
        VStack(alignment: .leading, spacing: 5) {
            HStack {
                Text("Kira-Texte").font(.headline).foregroundStyle(.white)
                Text("\(t?.fertig ?? 0)/\(t?.gesamt ?? 0) fertig").font(.subheadline.monospacedDigit())
                    .foregroundStyle(offen.isEmpty ? KiraStil.gruen : KiraStil.orange)
                Spacer()
                MacPunkt(t: t)
            }
            if offen.isEmpty {
                Text((t?.cloud ?? 0) > 0 ? "Alle Texte da, davon \(t?.cloud ?? 0) aus der Cloud." : (t?.gesamt ?? 0) > 0 ? "Alle Texte da, lokal erzeugt." : "Heute noch nichts offen.")
                    .font(.caption).foregroundStyle(.white.opacity(0.75))
            } else {
                ForEach(Array(offen.prefix(3).enumerated()), id: \.offset) { _, o in
                    HStack {
                        Text(o.label).font(.caption).foregroundStyle(.white).lineLimit(1)
                        Spacer()
                        Text(phaseText(o.phase)).font(.caption2).foregroundStyle(.white.opacity(0.6))
                    }
                }
            }
            Spacer(minLength: 0)
            HStack(spacing: 8) {
                Button(intent: BriefingLokalIntent()) {
                    Label("Analyse lokal", systemImage: "desktopcomputer").font(.caption2.weight(.semibold))
                }
                .tint(KiraStil.gruen)
                Button(intent: GarminAbrufenIntent()) {
                    Label("Garmin abrufen", systemImage: "arrow.down.circle").font(.caption2.weight(.semibold))
                }
                .tint(KiraStil.blau)
                Spacer()
                Text(KiraStil.uhrzeit(d.stand)).font(.caption2).foregroundStyle(.white.opacity(0.45))
            }
            .buttonStyle(.bordered)
        }
        .widgetURL(URL(string: "coachkira://eingabe"))
    }
}

struct BelastungKleinView: View {
    let d: WidgetDaten
    var body: some View {
        let band = d.last?.acwrBand
        let farbe: Color = band == "hoch" ? KiraStil.rot : band == "erhoeht" ? KiraStil.orange : band == "niedrig" ? KiraStil.gelb : KiraStil.gruen
        VStack(alignment: .leading, spacing: 3) {
            Text("BELASTUNG").font(.caption2.bold()).foregroundStyle(.white.opacity(0.6))
            HStack(alignment: .firstTextBaseline, spacing: 4) {
                Text(KiraStil.zahl(d.last?.acwr, stellen: 2)).font(.title2.weight(.heavy).monospacedDigit()).foregroundStyle(farbe)
                Text("ACWR").font(.caption2).foregroundStyle(.white.opacity(0.6))
            }
            Text("Soll \(d.last?.soll ?? "–") · Ist \(KiraStil.zahl(d.last?.istHeute))").font(.caption2.monospacedDigit()).foregroundStyle(.white)
            Spacer(minLength: 0)
            Balken(woche: d.last?.woche ?? [])
        }
        .widgetURL(URL(string: "coachkira://analyse"))
    }
}

/// Gemeinsame Hülle: leere Daten → Hinweis, sonst Inhalt; immer mit Kira-Hintergrund
struct Huelle<Inhalt: View>: View {
    let eintrag: KiraEintrag
    @ViewBuilder let inhalt: (WidgetDaten) -> Inhalt
    var body: some View {
        Group {
            if let d = eintrag.daten { inhalt(d) } else { LeerView(fehler: eintrag.fehler) }
        }
        .containerBackground(for: .widget) { KiraStil.hintergrund }
    }
}

struct TagesentscheidView: View {
    @Environment(\.widgetFamily) private var family
    let e: KiraEintrag
    var body: some View {
        Huelle(eintrag: e) { d in
            #if os(iOS)
            switch family {
            case .accessoryCircular: RundView(d: d)
            case .accessoryInline: InlineView(d: d)
            default: KleinView(d: d)
            }
            #else
            KleinView(d: d)
            #endif
        }
    }
}

struct MorgenLageView: View {
    @Environment(\.widgetFamily) private var family
    let e: KiraEintrag
    var body: some View {
        Huelle(eintrag: e) { d in
            if istRechteck(family) { RechteckView(d: d) } else { MittelView(d: d) }
        }
    }
}

struct ZieltourView: View {
    @Environment(\.widgetFamily) private var family
    let e: KiraEintrag
    var body: some View {
        Huelle(eintrag: e) { d in
            if istRechteck(family) { ZielRechteckView(d: d) } else { ZielKleinView(d: d) }
        }
    }
}

struct TexteView: View {
    @Environment(\.widgetFamily) private var family
    let e: KiraEintrag
    var body: some View {
        Huelle(eintrag: e) { d in
            if family == .systemMedium { TexteMittelView(d: d) } else { TexteKleinView(d: d) }
        }
    }
}

/// Widget-Größen je Plattform: Sperrbildschirm-Varianten gibt es nur auf dem iPhone
enum KiraFamilien {
    #if os(iOS)
    static let alle: [WidgetFamily] = [.systemSmall, .systemMedium, .systemLarge, .systemExtraLarge, .accessoryCircular, .accessoryRectangular, .accessoryInline]
    static let tagesentscheid: [WidgetFamily] = [.systemSmall, .accessoryCircular, .accessoryInline]
    static let morgenLage: [WidgetFamily] = [.systemMedium, .accessoryRectangular]
    static let zieltour: [WidgetFamily] = [.systemSmall, .accessoryRectangular]
    #else
    static let alle: [WidgetFamily] = [.systemSmall, .systemMedium, .systemLarge, .systemExtraLarge]
    static let tagesentscheid: [WidgetFamily] = [.systemSmall]
    static let morgenLage: [WidgetFamily] = [.systemMedium]
    static let zieltour: [WidgetFamily] = [.systemSmall]
    #endif
}

func istRechteck(_ f: WidgetFamily) -> Bool {
    #if os(iOS)
    return f == .accessoryRectangular
    #else
    return false
    #endif
}

// MARK: - Widgets

/// Alles in einem (wie bisher): Klein, Mittel, Groß und Sperrbildschirm
struct KiraWidget: Widget {
    let kind = "KiraWidget"

    var body: some WidgetConfiguration {
        StaticConfiguration(kind: kind, provider: KiraProvider()) { e in
            KiraWidgetView(eintrag: e)
        }
        .configurationDisplayName("Coach Kira")
        .description("Tagesentscheid, Morgenwerte, Last, Ziel und Kira-Texte.")
        .supportedFamilies(KiraFamilien.alle)
    }
}

struct TagesentscheidWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraTagesentscheid", provider: KiraProvider()) { e in TagesentscheidView(e: e) }
            .configurationDisplayName("Tagesentscheid")
            .description("Votum in Garmin-Farbe, Readiness und Soll für heute.")
            .supportedFamilies(KiraFamilien.tagesentscheid)
    }
}

struct MorgenLageWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraMorgenLage", provider: KiraProvider()) { e in MorgenLageView(e: e) }
            .configurationDisplayName("Morgen-Lage")
            .description("Readiness, Schlaf, HRV, Ruhepuls und ESS heute.")
            .supportedFamilies(KiraFamilien.morgenLage)
    }
}

struct WocheZielWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraWocheZiel", provider: KiraProvider()) { e in
            Huelle(eintrag: e) { d in GrossView(d: d) }
        }
        .configurationDisplayName("Woche & Ziel")
        .description("Tagesstand, ACWR, Last der letzten 7 Tage, nächstes Ziel und Knöpfe für den Mac.")
        .supportedFamilies([.systemLarge])
    }
}

/// Extragroßes Widget (iPad, Mac): links der Tagesstand, rechts Zieltour und Kira-Texte.
/// Entwurf vom lokalen Modell coach-kira, in der Cloud geprüft (Vorschau entfernt, URL-Schema ergänzt).
struct ExtraGrossView: View {
    let d: WidgetDaten

    var body: some View {
        HStack(alignment: .top, spacing: 16) {
            VStack(alignment: .leading, spacing: 8) {
                GrossView(d: d)
            }
            .frame(maxWidth: .infinity, alignment: .leading)

            Divider()
                .overlay(Color.white.opacity(0.15))

            VStack(alignment: .leading, spacing: 10) {
                ZielKleinView(d: d)
                    .frame(height: 130)

                Divider()
                    .overlay(Color.white.opacity(0.15))

                TexteKleinView(d: d)
            }
            .frame(width: 250)
        }
        .widgetURL(URL(string: "coachkira://heute"))
    }
}

/// Extragroß: nur iPad und Mac (auf dem iPhone bietet iOS diese Größe nicht an)
struct CockpitXLWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraCockpitXL", provider: KiraProvider()) { e in
            Huelle(eintrag: e) { d in ExtraGrossView(d: d) }
        }
        .configurationDisplayName("Cockpit extragroß")
        .description("Tagesstand, ACWR, Woche, Zieltour und Kira-Texte auf einen Blick (iPad und Mac).")
        .supportedFamilies([.systemExtraLarge])
    }
}

struct ZieltourWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraZieltour", provider: KiraProvider()) { e in ZieltourView(e: e) }
            .configurationDisplayName("Zieltour")
            .description("Countdown zur nächsten Zieltour mit Tourform.")
            .supportedFamilies(KiraFamilien.zieltour)
    }
}

struct TexteWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraTexte", provider: KiraProvider()) { e in TexteView(e: e) }
            .configurationDisplayName("Kira-Texte")
            .description("Stand der KI-Texte nach dem Absenden, lokal oder Cloud-Ersatz, mit Knopf für die lokale Analyse.")
            .supportedFamilies([.systemSmall, .systemMedium])
    }
}

struct BelastungWidget: Widget {
    var body: some WidgetConfiguration {
        StaticConfiguration(kind: "KiraBelastung", provider: KiraProvider()) { e in
            Huelle(eintrag: e) { d in BelastungKleinView(d: d) }
        }
        .configurationDisplayName("Belastung")
        .description("ACWR, Soll und Ist heute, Last der letzten 7 Tage.")
        .supportedFamilies([.systemSmall])
    }
}

@main
struct KiraWidgetBundle: WidgetBundle {
    var body: some Widget {
        TagesentscheidWidget()
        MorgenLageWidget()
        WocheZielWidget()
        CockpitXLWidget()
        ZieltourWidget()
        TexteWidget()
        BelastungWidget()
        KiraWidget()
    }
}
