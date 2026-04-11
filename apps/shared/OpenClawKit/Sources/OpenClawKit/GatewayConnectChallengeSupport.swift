import Foundation
import OpenClawProtocol

public enum GatewayConnectChallengeSupport {
    public static let minimumNonceLength = 8

    public static func nonce(from payload: [String: OpenClawProtocol.AnyCodable]?) -> String? {
        guard let nonce = payload?["nonce"]?.value as? String else { return nil }
        let trimmed = nonce.trimmingCharacters(in: .whitespacesAndNewlines)
        guard trimmed.count >= minimumNonceLength else { return nil }
        return trimmed
    }

    public static func waitForNonce<E: Error>(
        timeoutSeconds: Double,
        onTimeout: @escaping @Sendable () -> E,
        receiveNonce: @escaping @Sendable () async throws -> String?) async throws -> String
    {
        try await AsyncTimeout.withTimeout(
            seconds: timeoutSeconds,
            onTimeout: onTimeout,
            operation: {
                while true {
                    if let nonce = try await receiveNonce() {
                        return nonce
                    }
                }
            })
    }
}
