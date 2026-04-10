import Foundation
import Testing
@testable import OpenClawKit
import OpenClawProtocol

// MARK: - ConfigurableFakeWebSocketTask

/// Extended WebSocket task fake for security audit tests.
/// Supports configurable nonce, event injection, and full frame capture.
final class ConfigurableFakeWebSocketTask: WebSocketTasking, @unchecked Sendable {
    private let lock = NSLock()
    private let challengeNonce: String?
    private let helloAuth: [String: Any]?
    private let helloError: [String: Any]?
    private var _state: URLSessionTask.State = .suspended
    private var connectRequestId: String?
    private var connectParams: [String: Any]?
    private var receivePhase = 0
    private var pendingReceiveHandler:
        (@Sendable (Result<URLSessionWebSocketTask.Message, Error>) -> Void)?
    private var injectedEvents: [[String: Any]] = []
    private var _sentFrames: [[String: Any]] = []

    /// - Parameters:
    ///   - challengeNonce: The nonce to include in the connect.challenge event.
    ///     `nil` = no challenge sent (timeout testing), `""` = empty nonce, string = valid nonce.
    ///   - helloAuth: Auth payload in the hello-ok response.
    ///   - helloError: If set, the connect response will be an error response.
    init(challengeNonce: String? = "test-nonce-12345678",
         helloAuth: [String: Any]? = nil,
         helloError: [String: Any]? = nil)
    {
        self.challengeNonce = challengeNonce
        self.helloAuth = helloAuth
        self.helloError = helloError
    }

    var state: URLSessionTask.State {
        get { self.lock.withLock { self._state } }
        set { self.lock.withLock { self._state = newValue } }
    }

    func resume() {
        self.state = .running
    }

    func cancel(with closeCode: URLSessionWebSocketTask.CloseCode, reason: Data?) {
        _ = (closeCode, reason)
        self.state = .canceling
        let handler = self.lock.withLock { () -> (@Sendable (Result<URLSessionWebSocketTask.Message, Error>) -> Void)? in
            defer { self.pendingReceiveHandler = nil }
            return self.pendingReceiveHandler
        }
        handler?(Result<URLSessionWebSocketTask.Message, Error>.failure(URLError(.cancelled)))
    }

    func send(_ message: URLSessionWebSocketTask.Message) async throws {
        let data: Data? = switch message {
        case let .data(d): d
        case let .string(s): s.data(using: .utf8)
        @unknown default: nil
        }
        guard let data else { return }
        if let obj = try? JSONSerialization.jsonObject(with: data) as? [String: Any] {
            self.lock.withLock {
                self._sentFrames.append(obj)
            }
            if obj["type"] as? String == "req",
               obj["method"] as? String == "connect",
               let id = obj["id"] as? String
            {
                let params = obj["params"] as? [String: Any]
                self.lock.withLock {
                    self.connectRequestId = id
                    self.connectParams = params
                }
            }
        }
    }

    func latestConnectAuth() -> [String: Any]? {
        self.lock.withLock {
            (self.connectParams?["auth"] as? [String: Any]) ?? [:]
        }
    }

    func latestConnectParams() -> [String: Any]? {
        self.lock.withLock { self.connectParams }
    }

    func allSentFrames() -> [[String: Any]] {
        self.lock.withLock { self._sentFrames }
    }

    func sendPing(pongReceiveHandler: @escaping @Sendable (Error?) -> Void) {
        pongReceiveHandler(nil)
    }

    /// Inject an event frame that will be delivered on the next receive() call after handshake.
    func injectServerEvent(_ frame: [String: Any]) {
        self.lock.withLock {
            self.injectedEvents.append(frame)
        }
    }

    func receive() async throws -> URLSessionWebSocketTask.Message {
        let phase = self.lock.withLock { () -> Int in
            let current = self.receivePhase
            self.receivePhase += 1
            return current
        }
        // Phase 0: send connect.challenge (if nonce configured)
        if phase == 0 {
            if let nonce = self.challengeNonce {
                return .data(Self.connectChallengeData(nonce: nonce))
            } else {
                // No challenge — block until cancelled (for timeout testing)
                while !Task.isCancelled {
                    try await Task.sleep(nanoseconds: 50_000_000) // 50ms
                }
                throw URLError(.cancelled)
            }
        }
        // Phase 1: wait for connect request, then send hello-ok or error
        if phase == 1 {
            for _ in 0..<50 {
                let id = self.lock.withLock { self.connectRequestId }
                if let id {
                    if let error = self.helloError {
                        return .data(Self.connectErrorData(id: id, error: error))
                    }
                    return .data(Self.connectOkData(id: id, auth: self.helloAuth))
                }
                try await Task.sleep(nanoseconds: 1_000_000)
            }
            return .data(Self.connectOkData(id: "connect", auth: self.helloAuth))
        }
        // Phase 2+: deliver injected events, then block
        let event = self.lock.withLock { () -> [String: Any]? in
            let idx = phase - 2
            if idx < self.injectedEvents.count {
                return self.injectedEvents[idx]
            }
            return nil
        }
        if let event {
            return .data((try? JSONSerialization.data(withJSONObject: event)) ?? Data())
        }
        // Block until cancelled or event injected
        while !Task.isCancelled {
            try await Task.sleep(nanoseconds: 50_000_000)
        }
        throw URLError(.cancelled)
    }

    func receive(
        completionHandler: @escaping @Sendable (Result<URLSessionWebSocketTask.Message, Error>) -> Void)
    {
        self.lock.withLock { self.pendingReceiveHandler = completionHandler }
    }

    func emitReceiveFailure() {
        let handler = self.lock.withLock { () -> (@Sendable (Result<URLSessionWebSocketTask.Message, Error>) -> Void)? in
            self._state = .canceling
            defer { self.pendingReceiveHandler = nil }
            return self.pendingReceiveHandler
        }
        handler?(Result<URLSessionWebSocketTask.Message, Error>.failure(URLError(.networkConnectionLost)))
    }

    // MARK: - Frame Builders

    private static func connectChallengeData(nonce: String) -> Data {
        let frame: [String: Any] = [
            "type": "event",
            "event": "connect.challenge",
            "payload": ["nonce": nonce],
        ]
        return (try? JSONSerialization.data(withJSONObject: frame)) ?? Data()
    }

    private static func connectOkData(id: String, auth: [String: Any]? = nil) -> Data {
        var payload: [String: Any] = [
            "type": "hello-ok",
            "protocol": 2,
            "server": [
                "version": "test",
                "connId": "test",
            ],
            "features": [
                "methods": [] as [Any],
                "events": [] as [Any],
            ],
            "snapshot": [
                "presence": [["ts": 1]],
                "health": [:] as [String: Any],
                "stateVersion": [
                    "presence": 0,
                    "health": 0,
                ],
                "uptimeMs": 0,
            ],
            "policy": [
                "maxPayload": 1,
                "maxBufferedBytes": 1,
                "tickIntervalMs": 30_000,
            ],
        ]
        if let auth {
            payload["auth"] = auth
        }
        let frame: [String: Any] = [
            "type": "res",
            "id": id,
            "ok": true,
            "payload": payload,
        ]
        return (try? JSONSerialization.data(withJSONObject: frame)) ?? Data()
    }

    private static func connectErrorData(id: String, error: [String: Any]) -> Data {
        let frame: [String: Any] = [
            "type": "res",
            "id": id,
            "ok": false,
            "error": error,
        ]
        return (try? JSONSerialization.data(withJSONObject: frame)) ?? Data()
    }
}

// MARK: - MultiGenerationFakeSession

/// Tracks WebSocket tasks across reconnect generations.
final class MultiGenerationFakeSession: WebSocketSessioning, @unchecked Sendable {
    private let lock = NSLock()
    private let taskFactory: () -> ConfigurableFakeWebSocketTask
    private var tasks: [ConfigurableFakeWebSocketTask] = []

    init(taskFactory: @escaping () -> ConfigurableFakeWebSocketTask) {
        self.taskFactory = taskFactory
    }

    convenience init(challengeNonce: String? = "test-nonce-12345678",
                     helloAuth: [String: Any]? = nil)
    {
        self.init(taskFactory: {
            ConfigurableFakeWebSocketTask(challengeNonce: challengeNonce, helloAuth: helloAuth)
        })
    }

    func makeWebSocketTask(url: URL) -> WebSocketTaskBox {
        _ = url
        return self.lock.withLock {
            let task = self.taskFactory()
            self.tasks.append(task)
            return WebSocketTaskBox(task: task)
        }
    }

    func taskCount() -> Int {
        self.lock.withLock { self.tasks.count }
    }

    func taskAt(_ index: Int) -> ConfigurableFakeWebSocketTask? {
        self.lock.withLock {
            guard index < self.tasks.count else { return nil }
            return self.tasks[index]
        }
    }

    func latestTask() -> ConfigurableFakeWebSocketTask? {
        self.lock.withLock { self.tasks.last }
    }

    func authForGeneration(_ n: Int) -> [String: Any]? {
        self.lock.withLock {
            guard n < self.tasks.count else { return nil }
            return self.tasks[n].latestConnectAuth()
        }
    }

    func paramsForGeneration(_ n: Int) -> [String: Any]? {
        self.lock.withLock {
            guard n < self.tasks.count else { return nil }
            return self.tasks[n].latestConnectParams()
        }
    }
}

// MARK: - Temp State Dir Helper

/// Creates a temporary directory for OPENCLAW_STATE_DIR and restores on deinit.
final class TempStateDir {
    let url: URL
    private let previousValue: String?

    init() {
        self.url = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try? FileManager.default.createDirectory(at: self.url, withIntermediateDirectories: true)
        self.previousValue = ProcessInfo.processInfo.environment["OPENCLAW_STATE_DIR"]
        setenv("OPENCLAW_STATE_DIR", self.url.path, 1)
    }

    func cleanup() {
        if let prev = self.previousValue {
            setenv("OPENCLAW_STATE_DIR", prev, 1)
        } else {
            unsetenv("OPENCLAW_STATE_DIR")
        }
        try? FileManager.default.removeItem(at: self.url)
    }

    /// Path to the identity directory used by DeviceIdentityStore/DeviceAuthStore.
    var identityDir: URL {
        self.url.appendingPathComponent("identity", isDirectory: true)
    }

    var deviceJsonURL: URL {
        self.identityDir.appendingPathComponent("device.json")
    }

    var deviceAuthJsonURL: URL {
        self.identityDir.appendingPathComponent("device-auth.json")
    }
}

// MARK: - Default Connect Options

func defaultTestConnectOptions(includeDeviceIdentity: Bool = true) -> GatewayConnectOptions {
    GatewayConnectOptions(
        role: "operator",
        scopes: ["operator.read"],
        caps: [],
        commands: [],
        permissions: [:],
        clientId: "openclaw-ios-test",
        clientMode: "ui",
        clientDisplayName: "iOS Test",
        includeDeviceIdentity: includeDeviceIdentity)
}
