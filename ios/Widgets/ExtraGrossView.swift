import SwiftUI
import WidgetKit

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
