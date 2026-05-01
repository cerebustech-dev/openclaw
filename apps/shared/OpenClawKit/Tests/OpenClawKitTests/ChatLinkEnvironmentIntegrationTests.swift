import Foundation
import SwiftUI
import Testing
@testable import OpenClawChatUI

@MainActor
@Suite("ChatLink OpenURLAction integration")
struct ChatLinkEnvironmentIntegrationTests {

    // SwiftUI's `OpenURLAction.Result` does not conform to `Equatable`, so we
    // cannot introspect the result at runtime via `if case .discarded = …`.
    // The URL-scheme behavioural matrix is covered exhaustively in
    // ChatLinkAllowlistTests against the Equatable `ChatLinkDecision` enum.
    // Here we lock the boundary that ChatView depends on at compile time:
    // the exact `OpenURLAction` closure body and its decision-to-result
    // mapping must type-check against the SwiftUI API surface. If SwiftUI
    // renames a case (e.g., `.discarded` → `.cancelled`) or breaks the
    // closure signature, this test will fail at build time.

    @Test("OpenURLAction construction with ChatLinkAllowlist decision compiles")
    func openURLActionConstructionCompiles() {
        _ = OpenURLAction { url in
            switch ChatLinkAllowlist.decision(for: url) {
            case .allow:   return .systemAction
            case .discard: return .discarded
            }
        }
    }
}
