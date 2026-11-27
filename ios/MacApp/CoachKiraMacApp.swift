import SwiftUI
import AppKit
import WebKit
import WidgetKit

/// Coach Kira für den Mac: Cockpit im eigenen Fenster, Menü mit Tastenkürzeln, Kopplung für die Mac-Widgets.
/// Lokal nutzt das Cockpit hier direkt dein Ollama (https://localhost:11435 über Caddy) – mit Live-Text.
@main
struct CoachKiraMacApp: App {
    @StateObject private var steuerung = MacSteuerung()

    var body: some Scene {
        WindowGroup("Coach Kira") {
            MacContentView()
                .environmentObject(steuerung)
                .frame(minWidth: 960, minHeight: 640)
        }
        .defaultSize(width: 1440, height: 960)
        .commands {
            CommandGroup(replacing: .newItem) {}
            CommandMenu("Cockpit") {
                Button("Heute") { steuerung.gehe("heute") }.keyboardShortcut("1")
                Button("Eingabe") { steuerung.gehe("eingabe") }.keyboardShortcut("2")
                Button("Plan") { steuerung.gehe("plan") }.keyboardShortcut("3")
                Button("Ziele") { steuerung.gehe("ziele") }.keyboardShortcut("4")
                Button("Analyse") { steuerung.gehe("analyse") }.keyboardShortcut("5")
                Button("Berichte") { steuerung.gehe("berichte") }.keyboardShortcut("6")
                Button("Log") { steuerung.gehe("log") }.keyboardShortcut("7")
                Divider()
                Button("Neu laden") { steuerung.neuLaden() }.keyboardShortcut("r")
                Button("Widgets aktualisieren") { WidgetCenter.shared.reloadAllTimelines() }.keyboardShortcut("r", modifiers: [.command, .shift])
                Divider()
                Button("Größer") { steuerung.zoom(+0.1) }.keyboardShortcut("+")
                Button("Kleiner") { steuerung.zoom(-0.1) }.keyboardShortcut("-")
                Button("Standardgröße (175 %)") { steuerung.zoomStandard() }.keyboardShortcut("0")
                Divider()
                Button("Im Browser öffnen") { NSWorkspace.shared.open(steuerung.aktuelleURL ?? KiraConfig.cockpit) }
            }
        }

        Settings {
            KopplungMacView()
                .frame(width: 480)
                .padding()
        }
    }
}

final class MacSteuerung: ObservableObject {
    weak var webView: WKWebView?
    var aktuelleURL: URL? { webView?.url }

    /// Schriftgröße der Cockpit-Seite: Standard 175 %, wird gemerkt
    static let zoomStandardWert: Double = 1.75
    static var gespeicherterZoom: Double {
        let z = UserDefaults.standard.double(forKey: "kiraMacZoom")
        return z > 0 ? z : zoomStandardWert
    }

    func zoom(_ schritt: Double) { setzeZoom((webView.map { Double($0.pageZoom) } ?? Self.gespeicherterZoom) + schritt) }
    func zoomStandard() { setzeZoom(Self.zoomStandardWert) }

    private func setzeZoom(_ z: Double) {
        let wert = min(max((z * 10).rounded() / 10, 0.8), 3.0)
        webView?.pageZoom = wert
        UserDefaults.standard.set(wert, forKey: "kiraMacZoom")
    }

    func neuLaden() {
        webView?.reload()
        WidgetCenter.shared.reloadAllTimelines()
    }

    func gehe(_ ziel: String) {
        guard let u = URL(string: KiraConfig.cockpit.absoluteString + "#" + ziel) else { return }
        if let wv = webView, wv.url?.host == u.host {
            // gleiche Seite: nur Sprungziel wechseln (kein Neuladen)
            wv.evaluateJavaScript("location.hash = '#\(ziel)'", completionHandler: nil)
        } else {
            webView?.load(URLRequest(url: u))
        }
    }
}

struct MacContentView: View {
    @EnvironmentObject private var steuerung: MacSteuerung
    @State private var zeigeKopplung = false

    var body: some View {
        MacWebView(steuerung: steuerung)
            .background(KiraStil.hintergrund)
            .toolbar {
                ToolbarItemGroup(placement: .navigation) {
                    Button { steuerung.webView?.goBack() } label: { Image(systemName: "chevron.left") }
                        .help("Zurück")
                    Button { steuerung.neuLaden() } label: { Image(systemName: "arrow.clockwise") }
                        .help("Neu laden (⌘R)")
                }
                ToolbarItemGroup(placement: .primaryAction) {
                    Button { steuerung.gehe("heute") } label: { Label("Heute", systemImage: "sun.max") }
                    Button { steuerung.gehe("eingabe") } label: { Label("Eingabe", systemImage: "square.and.pencil") }
                    Button { steuerung.gehe("plan") } label: { Label("Plan", systemImage: "calendar") }
                    Button { steuerung.gehe("analyse") } label: { Label("Analyse", systemImage: "chart.xyaxis.line") }
                    Button { zeigeKopplung = true } label: {
                        Label("Widgets", systemImage: KiraStore.token == nil ? "link.badge.plus" : "square.grid.2x2")
                    }
                    .help("Kopplung und Widgets")
                }
            }
            .sheet(isPresented: $zeigeKopplung) {
                VStack(alignment: .trailing) {
                    KopplungMacView()
                    Button("Fertig") { zeigeKopplung = false }.keyboardShortcut(.defaultAction)
                }
                .padding()
                .frame(width: 500)
            }
            .onOpenURL { url in
                let ziel = (url.host ?? url.path).trimmingCharacters(in: CharacterSet(charactersIn: "/")).lowercased()
                steuerung.gehe(ziel.isEmpty ? "heute" : ziel)
            }
            .onAppear {
                if KiraStore.token == nil { zeigeKopplung = true }
            }
    }
}

/// Cockpit im WKWebView (macOS). Externe Links (z. B. Garmin Connect) öffnen im Standardbrowser.
struct MacWebView: NSViewRepresentable {
    let steuerung: MacSteuerung

    func makeCoordinator() -> Coordinator { Coordinator() }

    func makeNSView(context: Context) -> WKWebView {
        let cfg = WKWebViewConfiguration()
        cfg.websiteDataStore = .default()
        // Das Cockpit meldet „abgesendet“ → Widgets etwas später neu laden
        cfg.userContentController.add(context.coordinator, name: "kiraMac")
        let wv = WKWebView(frame: .zero, configuration: cfg)
        wv.uiDelegate = context.coordinator
        wv.navigationDelegate = context.coordinator
        wv.allowsBackForwardNavigationGestures = true
        wv.allowsMagnification = true
        wv.pageZoom = MacSteuerung.gespeicherterZoom
        steuerung.webView = wv
        wv.load(URLRequest(url: KiraConfig.cockpit))
        return wv
    }

    func updateNSView(_ nsView: WKWebView, context: Context) {}

    final class Coordinator: NSObject, WKScriptMessageHandler, WKUIDelegate, WKNavigationDelegate {
        func userContentController(_ ucc: WKUserContentController, didReceive message: WKScriptMessage) {
            DispatchQueue.main.asyncAfter(deadline: .now() + 20) {
                WidgetCenter.shared.reloadAllTimelines()
            }
        }

        // target="_blank" → Standardbrowser
        func webView(_ webView: WKWebView, createWebViewWith configuration: WKWebViewConfiguration,
                     for navigationAction: WKNavigationAction, windowFeatures: WKWindowFeatures) -> WKWebView? {
            if let u = navigationAction.request.url { NSWorkspace.shared.open(u) }
            return nil
        }

        // Links auf andere Seiten als das Cockpit → Standardbrowser
        func webView(_ webView: WKWebView, decidePolicyFor navigationAction: WKNavigationAction,
                     decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
            if navigationAction.navigationType == .linkActivated, let u = navigationAction.request.url,
               let h = u.host, h != KiraConfig.cockpit.host {
                NSWorkspace.shared.open(u)
                decisionHandler(.cancel)
                return
            }
            decisionHandler(.allow)
        }

        // Datei-Downloads und Bestätigungsdialoge des Cockpits (z. B. „Wirklich löschen?“)
        func webView(_ webView: WKWebView, runJavaScriptConfirmPanelWithMessage message: String,
                     initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping (Bool) -> Void) {
            let a = NSAlert()
            a.messageText = message
            a.addButton(withTitle: "OK")
            a.addButton(withTitle: "Abbrechen")
            completionHandler(a.runModal() == .alertFirstButtonReturn)
        }

        func webView(_ webView: WKWebView, runJavaScriptAlertPanelWithMessage message: String,
                     initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping () -> Void) {
            let a = NSAlert()
            a.messageText = message
            a.runModal()
            completionHandler()
        }
    }
}

/// Kopplung mit dem Cockpit für die Mac-Widgets (eigener Schlüssel, unabhängig vom iPhone).
struct KopplungMacView: View {
    @State private var code = ""
    @State private var meldung: String?
    @State private var laeuft = false
    @State private var gekoppelt = KiraStore.token != nil
    @State private var stand: WidgetDaten? = KiraStore.letzter()

    var body: some View {
        Form {
            if !gekoppelt {
                Section {
                    Text("Im Cockpit unter Heute → Technik → „iPhone-App & Widgets“ auf „Kopplungscode erzeugen“ klicken und den Code hier eingeben.")
                        .font(.callout)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                    TextField("6-stelliger Code", text: $code)
                        .font(.title2.monospacedDigit())
                        .onSubmit { Task { await verbinden() } }
                    Button(laeuft ? "Verbinde …" : "Verbinden") { Task { await verbinden() } }
                        .disabled(code.trimmingCharacters(in: .whitespaces).count != 6 || laeuft)
                } header: {
                    Text("Mac-Widgets mit dem Cockpit verbinden").font(.headline)
                }
            } else {
                Section {
                    Label("Mit dem Cockpit gekoppelt", systemImage: "checkmark.seal.fill")
                        .foregroundStyle(KiraStil.gruen)
                    if let s = stand, let v = s.votum {
                        HStack {
                            Circle().fill(KiraStil.farbe(v.stufe)).frame(width: 10, height: 10)
                            Text("\(KiraStil.farbname(v.stufe)) · \(v.text)")
                            Spacer()
                            Text("Stand \(KiraStil.uhrzeit(s.stand))").foregroundStyle(.secondary).font(.footnote)
                        }
                    }
                    Button(laeuft ? "Lädt …" : "Widgets jetzt aktualisieren") { Task { await aktualisieren() } }
                        .disabled(laeuft)
                    Text("Widgets hinzufügen: Rechtsklick auf den Schreibtisch → „Widgets bearbeiten“ → „Coach Kira“. Oder Mitteilungszentrale öffnen (Klick auf Datum/Uhrzeit) → ganz unten „Widgets bearbeiten“.")
                        .font(.footnote)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                    Button("Kopplung lösen", role: .destructive) {
                        KiraStore.token = nil
                        gekoppelt = false
                        WidgetCenter.shared.reloadAllTimelines()
                    }
                } header: {
                    Text("Mac-Widgets").font(.headline)
                }
            }
            if let m = meldung {
                Text(m).font(.footnote)
            }
        }
        .formStyle(.grouped)
    }

    private func verbinden() async {
        laeuft = true
        defer { laeuft = false }
        do {
            try await KiraAPI.verbinden(code: code.trimmingCharacters(in: .whitespaces), geraet: Host.current().localizedName ?? "Mac")
            gekoppelt = true
            meldung = "Verbunden. Die Widgets holen jetzt den Tagesstand."
            stand = try? await KiraAPI.laden(frisch: true)
            WidgetCenter.shared.reloadAllTimelines()
        } catch {
            meldung = error.localizedDescription
        }
    }

    private func aktualisieren() async {
        laeuft = true
        defer { laeuft = false }
        do {
            stand = try await KiraAPI.laden(frisch: true)
            meldung = "Aktualisiert."
            WidgetCenter.shared.reloadAllTimelines()
        } catch {
            meldung = error.localizedDescription
            if let f = error as? KiraFehler, case .nichtGekoppelt = f { gekoppelt = false }
        }
    }
}
