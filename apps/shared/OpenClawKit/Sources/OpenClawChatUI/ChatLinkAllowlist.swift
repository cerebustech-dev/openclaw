import Foundation

enum ChatLinkDecision: Equatable {
    case allow
    case discard
}

enum ChatLinkAllowlist {
    static let allowedSchemes: Set<String> = ["https", "http", "mailto"]

    static func decision(for url: URL) -> ChatLinkDecision {
        guard let scheme = url.scheme?.lowercased(),
              allowedSchemes.contains(scheme) else {
            return .discard
        }
        if scheme == "http" || scheme == "https" {
            guard let host = url.host(), !host.isEmpty else {
                return .discard
            }
        }
        return .allow
    }
}
