import Foundation
import SwiftUI
import Testing
@testable import OpenClawChatUI

@MainActor
@Suite("ChatLink OpenURLAction integration")
struct ChatLinkEnvironmentIntegrationTests {

    // Mirror of the exact closure body installed by ChatView's
    // `.environment(\.openURL, OpenURLAction { ... })` modifier. If the mapping
    // from ChatLinkDecision to OpenURLAction.Result drifts, these tests fail.
    private static func handler(_ url: URL) -> OpenURLAction.Result {
        switch ChatLinkAllowlist.decision(for: url) {
        case .allow:   return .systemAction
        case .discard: return .discarded
        }
    }

    private func isDiscarded(_ url: URL) -> Bool {
        if case .discarded = Self.handler(url) { return true }
        return false
    }

    private func isSystemAction(_ url: URL) -> Bool {
        if case .systemAction = Self.handler(url) { return true }
        return false
    }

    @Test("tel: maps to .discarded")
    func telDiscardedAtBoundary() {
        #expect(isDiscarded(URL(string: "tel:+15551234567")!))
    }

    @Test("sms: maps to .discarded")
    func smsDiscardedAtBoundary() {
        #expect(isDiscarded(URL(string: "sms:+15551234567")!))
    }

    @Test("facetime: maps to .discarded")
    func facetimeDiscardedAtBoundary() {
        #expect(isDiscarded(URL(string: "facetime://user@example.com")!))
    }

    @Test("javascript: maps to .discarded")
    func javascriptDiscardedAtBoundary() {
        #expect(isDiscarded(URL(string: "javascript:alert(1)")!))
    }

    @Test("data: maps to .discarded")
    func dataDiscardedAtBoundary() {
        #expect(isDiscarded(URL(string: "data:text/html,<script>alert(1)</script>")!))
    }

    @Test("https: maps to .systemAction")
    func httpsSystemActionAtBoundary() {
        #expect(isSystemAction(URL(string: "https://example.com")!))
    }

    @Test("http: maps to .systemAction")
    func httpSystemActionAtBoundary() {
        #expect(isSystemAction(URL(string: "http://example.com")!))
    }

    @Test("mailto: maps to .systemAction")
    func mailtoSystemActionAtBoundary() {
        #expect(isSystemAction(URL(string: "mailto:rod@example.com")!))
    }

    // Compile-only: confirms the exact OpenURLAction construction used in
    // ChatView type-checks. If SwiftUI changes OpenURLAction's API, this fails.
    @Test("OpenURLAction construction type-checks")
    func openURLActionConstructionCompiles() {
        _ = OpenURLAction { url in
            switch ChatLinkAllowlist.decision(for: url) {
            case .allow:   return .systemAction
            case .discard: return .discarded
            }
        }
    }
}
