import SwiftUI
import UIKit
import WebKit
import WidgetKit

@main
struct CoachKiraApp: App {
    var body: some Scene {
        WindowGroup {
            ContentView()
                .preferredColorScheme(.dark)
        }
    }
}

/// Hauptansicht: das Cockpit in voller Größe, dazu ein kleiner Knopf für Kopplung und Widgets.
struct ContentView: View {
    @State private var ziel: URL = KiraConfig.cockpit
    @State private var zeigeKopplung = false
    @Environment(\.scenePhase) private var phase

    var body: some View {
        ZStack(alignment: .bottomTrailing) {
            KiraWebView(url: ziel)
                .ignoresSafeArea(edges: .bottom)
                .background(KiraStil.hintergrund)
            Button {
                zeigeKopplung = true
            } label: {
                Image(systemName: KiraStore.token == nil ? "link.badge.plus" : "square.grid.2x2")
                    .font(.system(size: 15, weight: .semibold))
                    .foregroundStyle(.white)
                    .padding(10)
                    .background(.ultraThinMaterial, in: Circle())
            }
            .padding(.trailing, 14)
            .padding(.bottom, 18)
            .accessibilityLabel("Kopplung und Widgets")
        }
        .sheet(isPresented: $zeigeKopplung) { KopplungView() }
        .onOpenURL { url in ziel = KiraLinks.cockpitURL(fuer: url) }
        .onChange(of: phase) { _, neu in
            if neu == .active { WidgetCenter.shared.reloadAllTimelines() }
        }
        .onAppear {
            if KiraStore.token == nil { zeigeKopplung = true }
        }
    }
}

/// Deep Links aus den Widgets: coachkira://heute, coachkira://eingabe, coachkira://ziele …
enum KiraLinks {
    static func cockpitURL(fuer url: URL) -> URL {
        let ziel = (url.host ?? url.path).trimmingCharacters(in: CharacterSet(charactersIn: "/")).lowercased()
        let erlaubt = ["heute", "eingabe", "plan", "ziele", "analyse", "berichte", "log"]
        guard erlaubt.contains(ziel) else { return KiraConfig.cockpit }
        return URL(string: KiraConfig.cockpit.absoluteString + "#" + ziel) ?? KiraConfig.cockpit
    }
}

/// Cockpit im WKWebView. Merkt sich Anmeldung und Einstellungen (localStorage) wie Safari.
struct KiraWebView: UIViewRepresentable {
    let url: URL

    func makeCoordinator() -> Coordinator { Coordinator() }

    func makeUIView(context: Context) -> WKWebView {
        let cfg = WKWebViewConfiguration()
        cfg.websiteDataStore = .default()
        cfg.allowsInlineMediaPlayback = true
        // Das Cockpit meldet „abgesendet“ → Widgets neu laden
        cfg.userContentController.add(context.coordinator, name: "kira")
        let wv = WKWebView(frame: .zero, configuration: cfg)
        wv.isOpaque = false
        wv.backgroundColor = UIColor(red: 0.043, green: 0.071, blue: 0.125, alpha: 1)
        wv.scrollView.backgroundColor = wv.backgroundColor
        wv.allowsBackForwardNavigationGestures = true
        let rc = UIRefreshControl()
        rc.addTarget(context.coordinator, action: #selector(Coordinator.neuLaden(_:)), for: .valueChanged)
        wv.scrollView.refreshControl = rc
        context.coordinator.webView = wv
        context.coordinator.geladen = url
        wv.load(URLRequest(url: url))
        return wv
    }

    func updateUIView(_ wv: WKWebView, context: Context) {
        if context.coordinator.geladen != url {
            context.coordinator.geladen = url
            wv.load(URLRequest(url: url))
        }
    }

    final class Coordinator: NSObject, WKScriptMessageHandler {
        weak var webView: WKWebView?
        var geladen: URL?

        @objc func neuLaden(_ rc: UIRefreshControl) {
            webView?.reload()
            DispatchQueue.main.asyncAfter(deadline: .now() + 1.0) { rc.endRefreshing() }
            WidgetCenter.shared.reloadAllTimelines()
        }

        func userContentController(_ ucc: WKUserContentController, didReceive message: WKScriptMessage) {
            // Nach dem Absenden braucht die Auswertung einen Moment – dann Widgets neu laden
            DispatchQueue.main.asyncAfter(deadline: .now() + 20) {
                WidgetCenter.shared.reloadAllTimelines()
            }
        }
    }
}

/// Kopplung mit dem Cockpit (Code aus Heute → Technik → „iPhone-App & Widgets“).
struct KopplungView: View {
    @Environment(\.dismiss) private var dismiss
    @State private var code = ""
    @State private var meldung: String?
    @State private var laeuft = false
    @State private var gekoppelt = KiraStore.token != nil
    @State private var stand: WidgetDaten? = KiraStore.letzter()

    var body: some View {
        NavigationStack {
            Form {
                if !gekoppelt {
                    Section {
                        Text("Im Cockpit unter Heute → Technik → „iPhone-App & Widgets“ auf „Kopplungscode erzeugen“ tippen und den Code hier eingeben.")
                            .font(.footnote)
                            .foregroundStyle(.secondary)
                        TextField("6-stelliger Code", text: $code)
                            .keyboardType(.numberPad)
                            .font(.title2.monospacedDigit())
                        Button(laeuft ? "Verbinde …" : "Verbinden") {
                            Task { await verbinden() }
                        }
                        .disabled(code.trimmingCharacters(in: .whitespaces).count != 6 || laeuft)
                    } header: {
                        Text("Mit Cockpit verbinden")
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
                        Button(laeuft ? "Lädt …" : "Widgets jetzt aktualisieren") {
                            Task { await aktualisieren() }
                        }
                        .disabled(laeuft)
                    } header: {
                        Text("Widgets")
                    } footer: {
                        Text("Widgets hinzufügen: Homescreen lange drücken → „Bearbeiten“ → „Widget hinzufügen“ → Coach Kira. Für den Sperrbildschirm: Sperrbildschirm lange drücken → „Anpassen“.")
                    }
                    Section {
                        Button("Kopplung lösen", role: .destructive) {
                            KiraStore.token = nil
                            gekoppelt = false
                            WidgetCenter.shared.reloadAllTimelines()
                        }
                    }
                }
                if let m = meldung {
                    Section { Text(m).font(.footnote) }
                }
            }
            .navigationTitle("Coach Kira")
            .toolbar {
                ToolbarItem(placement: .confirmationAction) { Button("Fertig") { dismiss() } }
            }
        }
    }

    private func verbinden() async {
        laeuft = true
        defer { laeuft = false }
        do {
            try await KiraAPI.verbinden(code: code.trimmingCharacters(in: .whitespaces), geraet: UIDevice.current.name)
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
