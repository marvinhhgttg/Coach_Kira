import Foundation
import SwiftUI

/// Gemeinsam für App und Widgets: Adresse, Speicher (App Group), Datenmodell zu /api/widget, Abruf.
enum KiraConfig {
    /// iPhone: group.de.reichpietsch.coachkira · Mac: <Team-ID>.de.reichpietsch.coachkira (aus Info.plist „KiraAppGroup“)
    static let appGroup: String = {
        if let g = Bundle.main.object(forInfoDictionaryKey: "KiraAppGroup") as? String, !g.isEmpty, !g.contains("$(") { return g }
        return "group.de.reichpietsch.coachkira"
    }()
    static let server = URL(string: "https://coachkira.pplx.app/port/5000")!
    static let cockpit = URL(string: "https://coachkira.pplx.app/")!
}

// MARK: - Datenmodell (entspricht der Antwort von /api/widget)

struct WidgetDaten: Codable {
    struct Votum: Codable {
        let stufe: String
        let vote: String?
        let label: String
        let text: String
        let gruende: [String]?
    }
    struct Morgen: Codable {
        let eingetragen: Bool?
        let readiness: Double?
        let schlaf: Double?
        let schlafH: Double?
        let hrv: Double?
        let hrvStufe: String?
        let rhr: Double?
        let rhrDelta: Double?
        let befinden: Double?
    }
    struct Tag: Codable {
        let d: String
        let ist: Double?
    }
    struct Last: Codable {
        let sollHeute: Double?
        let sportHeute: String?
        let zoneHeute: String?
        let istHeute: Double?
        let erledigt: Bool?
        let acwr: Double?
        let acwrBand: String?
        let woche: [Tag]?
    }
    struct Ziel: Codable {
        let name: String
        let datum: String
        let tage: Int
        let tourform: Double?
    }
    struct Offen: Codable {
        let label: String
        let phase: String
        let versuche: Int?
    }
    struct Mac: Codable {
        let erreichbar: Bool?
        let ollama: Bool?
    }
    struct Texte: Codable {
        let fertig: Int
        let gesamt: Int
        let cloud: Int?
        let offen: [Offen]?
        let mac: Mac?
    }
    let ok: Bool
    let stand: String?
    let datum: String?
    let votum: Votum?
    let morgen: Morgen?
    let last: Last?
    let ziel: Ziel?
    let texte: Texte?

    /// Für die Widget-Galerie (Vorschau ohne Daten)
    static let beispiel = WidgetDaten(
        ok: true, stand: nil, datum: nil,
        votum: Votum(stufe: "gruen", vote: "TRAIN", label: "Train", text: "normal trainieren", gruende: nil),
        morgen: Morgen(eingetragen: true, readiness: 72, schlaf: 84, schlafH: 7.4, hrv: 52, hrvStufe: "gruen", rhr: 48, rhrDelta: 0, befinden: 4),
        last: Last(sollHeute: 120, sportHeute: "Run", zoneHeute: "Z2", istHeute: 0, erledigt: false, acwr: 1.08, acwrBand: "korridor",
                   woche: [95.0, 0, 140, 60, 0, 180, 0].enumerated().map { Tag(d: "\($0.offset)", ist: $0.element) }),
        ziel: Ziel(name: "Skitour Allgäu", datum: "2026-12-12", tage: 23, tourform: 86),
        texte: Texte(fertig: 3, gesamt: 3, cloud: 0, offen: [], mac: Mac(erreichbar: true, ollama: true))
    )
}

// MARK: - Speicher (App Group: App und Widget teilen Schlüssel und letzten Stand)

enum KiraStore {
    static var defaults: UserDefaults { UserDefaults(suiteName: KiraConfig.appGroup) ?? .standard }

    static var token: String? {
        get { defaults.string(forKey: "token") }
        set { defaults.set(newValue, forKey: "token") }
    }

    static func speichere(_ data: Data) {
        defaults.set(data, forKey: "letzter")
    }

    static func letzter() -> WidgetDaten? {
        guard let d = defaults.data(forKey: "letzter") else { return nil }
        return try? JSONDecoder().decode(WidgetDaten.self, from: d)
    }
}

// MARK: - Abruf

enum KiraFehler: LocalizedError {
    case nichtGekoppelt
    case server(String)

    var errorDescription: String? {
        switch self {
        case .nichtGekoppelt: return "App ist noch nicht mit dem Cockpit gekoppelt."
        case .server(let s): return s
        }
    }
}

private struct ServerAntwort: Decodable {
    let ok: Bool
    let token: String?
    let text: String?
    let error: String?
}

enum KiraAPI {
    static func laden(frisch: Bool = false) async throws -> WidgetDaten {
        guard let t = KiraStore.token, !t.isEmpty else { throw KiraFehler.nichtGekoppelt }
        var c = URLComponents(url: KiraConfig.server.appendingPathComponent("api/widget"), resolvingAgainstBaseURL: false)!
        if frisch { c.queryItems = [URLQueryItem(name: "frisch", value: "1")] }
        var r = URLRequest(url: c.url!, timeoutInterval: 60)
        r.setValue("Bearer \(t)", forHTTPHeaderField: "Authorization")
        r.cachePolicy = .reloadIgnoringLocalCacheData
        let (data, resp) = try await URLSession.shared.data(for: r)
        let code = (resp as? HTTPURLResponse)?.statusCode ?? 0
        if code == 401 { throw KiraFehler.nichtGekoppelt }
        guard let d = try? JSONDecoder().decode(WidgetDaten.self, from: data), d.ok else {
            let f = try? JSONDecoder().decode(ServerAntwort.self, from: data)
            throw KiraFehler.server(f?.error ?? "Server antwortet nicht (HTTP \(code)).")
        }
        KiraStore.speichere(data)
        return d
    }

    static func verbinden(code: String, geraet: String) async throws {
        var r = URLRequest(url: KiraConfig.server.appendingPathComponent("api/widget/verbinden"), timeoutInterval: 30)
        r.httpMethod = "POST"
        r.setValue("application/json", forHTTPHeaderField: "Content-Type")
        r.httpBody = try JSONSerialization.data(withJSONObject: ["code": code, "geraet": geraet])
        let (data, _) = try await URLSession.shared.data(for: r)
        let a = try JSONDecoder().decode(ServerAntwort.self, from: data)
        guard a.ok, let t = a.token else { throw KiraFehler.server(a.error ?? "Kopplung fehlgeschlagen.") }
        KiraStore.token = t
    }

    @discardableResult
    static func aktion(_ name: String) async throws -> String {
        guard let t = KiraStore.token, !t.isEmpty else { throw KiraFehler.nichtGekoppelt }
        var r = URLRequest(url: KiraConfig.server.appendingPathComponent("api/widget/aktion"), timeoutInterval: 30)
        r.httpMethod = "POST"
        r.setValue("application/json", forHTTPHeaderField: "Content-Type")
        r.setValue("Bearer \(t)", forHTTPHeaderField: "Authorization")
        r.httpBody = try JSONSerialization.data(withJSONObject: ["aktion": name])
        let (data, _) = try await URLSession.shared.data(for: r)
        let a = try JSONDecoder().decode(ServerAntwort.self, from: data)
        guard a.ok else { throw KiraFehler.server(a.error ?? "Aktion fehlgeschlagen.") }
        return a.text ?? "erledigt"
    }
}

// MARK: - Darstellung (Farben wie im Cockpit, Garmin-Spektrum)

enum KiraStil {
    static let hintergrund = Color(red: 0.043, green: 0.071, blue: 0.125)
    static let blau = Color(red: 0.098, green: 0.463, blue: 0.824)
    static let gruen = Color(red: 0.263, green: 0.627, blue: 0.278)
    static let gelb = Color(red: 0.992, green: 0.847, blue: 0.208)
    static let orange = Color(red: 0.984, green: 0.549, blue: 0.0)
    static let rot = Color(red: 0.898, green: 0.224, blue: 0.208)

    static func farbe(_ stufe: String?) -> Color {
        switch stufe {
        case "blau": return blau
        case "gruen": return gruen
        case "gelb": return gelb
        case "orange": return orange
        case "rot": return rot
        default: return .gray
        }
    }

    static func farbname(_ stufe: String?) -> String {
        switch stufe {
        case "blau": return "BLAU"
        case "gruen": return "GRÜN"
        case "gelb": return "GELB"
        case "orange": return "ORANGE"
        case "rot": return "ROT"
        default: return "–"
        }
    }

    static func zahl(_ v: Double?, stellen: Int = 0) -> String {
        guard let v else { return "–" }
        let f = NumberFormatter()
        f.locale = Locale(identifier: "de_DE")
        f.minimumFractionDigits = stellen
        f.maximumFractionDigits = stellen
        return f.string(from: NSNumber(value: v)) ?? "–"
    }

    static func uhrzeit(_ iso: String?) -> String {
        guard let iso else { return "–" }
        let p = ISO8601DateFormatter()
        p.formatOptions = [.withInternetDateTime, .withFractionalSeconds]
        let d = p.date(from: iso) ?? ISO8601DateFormatter().date(from: iso)
        guard let d else { return "–" }
        let f = DateFormatter()
        f.locale = Locale(identifier: "de_DE")
        f.dateFormat = "HH:mm"
        return f.string(from: d)
    }

    static func datumKurz(_ iso: String?) -> String {
        guard let iso, iso.count >= 10 else { return "" }
        let f = DateFormatter()
        f.dateFormat = "yyyy-MM-dd"
        guard let d = f.date(from: String(iso.prefix(10))) else { return "" }
        let g = DateFormatter()
        g.locale = Locale(identifier: "de_DE")
        g.dateFormat = "EE dd.MM."
        return g.string(from: d)
    }
}
