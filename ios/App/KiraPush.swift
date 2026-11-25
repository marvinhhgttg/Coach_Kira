// 25.11.2026: Push-Mitteilungen – Kiras Hinweise sofort auf iPhone und iPad.
// Die App meldet ihre Gerätekennung beim Cockpit an; der Mac verschickt die Hinweise über Apple (APNs).
import SwiftUI
import UIKit
import UserNotifications

enum KiraPush {
    private static let schluessel = "pushToken"

    #if DEBUG
    static let umgebung = "development"
    #else
    static let umgebung = "production"
    #endif

    /// Erlaubnis einholen und bei Apple registrieren (bei jedem Start; Apple liefert dann die aktuelle Kennung)
    static func starten() {
        let c = UNUserNotificationCenter.current()
        c.requestAuthorization(options: [.alert, .sound, .badge]) { ok, _ in
            guard ok else { return }
            DispatchQueue.main.async { UIApplication.shared.registerForRemoteNotifications() }
        }
    }

    static func erhalten(_ token: Data) {
        let hex = token.map { String(format: "%02x", $0) }.joined()
        KiraStore.defaults.set(hex, forKey: schluessel)
        melden(hex)
    }

    /// Nach dem Koppeln die gemerkte Kennung nachreichen
    static func nachKopplung() {
        if let hex = KiraStore.defaults.string(forKey: schluessel) { melden(hex) } else { starten() }
    }

    private static func melden(_ hex: String) {
        guard KiraStore.token != nil else { return }
        Task { try? await KiraAPI.pushMelden(geraeteToken: hex, umgebung: umgebung) }
    }
}

final class KiraAppDelegate: NSObject, UIApplicationDelegate, UNUserNotificationCenterDelegate {
    func application(_ application: UIApplication, didFinishLaunchingWithOptions launchOptions: [UIApplication.LaunchOptionsKey: Any]? = nil) -> Bool {
        UNUserNotificationCenter.current().delegate = self
        KiraPush.starten()
        return true
    }

    func application(_ application: UIApplication, didRegisterForRemoteNotificationsWithDeviceToken deviceToken: Data) {
        KiraPush.erhalten(deviceToken)
    }

    func application(_ application: UIApplication, didFailToRegisterForRemoteNotificationsWithError error: Error) {
        print("Push-Registrierung fehlgeschlagen: \(error.localizedDescription)")
    }

    /// Auch bei geöffneter App als Banner zeigen
    func userNotificationCenter(_ center: UNUserNotificationCenter, willPresent notification: UNNotification) async -> UNNotificationPresentationOptions {
        [.banner, .sound, .list]
    }

    /// Tippen auf die Mitteilung öffnet die Heute-Seite
    func userNotificationCenter(_ center: UNUserNotificationCenter, didReceive response: UNNotificationResponse) async {
        await MainActor.run { KiraSeiten.gehe("heute") }
    }
}
