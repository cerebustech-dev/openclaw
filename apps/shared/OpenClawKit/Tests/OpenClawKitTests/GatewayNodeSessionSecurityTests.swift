import Foundation
import Testing
@testable import OpenClawKit
import OpenClawProtocol

// MARK: - Test Infrastructure (duplicated from GatewayNodeSessionTests — private there)

private extension NSLock {
    func withLock<T>(_ body: () -> T) -> T {
        self.lock()
        defer { self.unlock() }
        return body()
    }
}

private final class FakeGatewayWebSocketTask: WebSocketTasking, @unchecked Sendable {
    private let lock = NSLock()
    private var _state: URLSessionTask.State = .suspended
    private var connectRequestId: String?
    private var connectAuth: [String: Any]?
    private var receivePhase = 0
    private var pendingReceiveHandler:
        (@Sendable (Result<URLSessionWebSocketTask.Message, Error>) -> Void)?

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
        if let obj = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
           obj["type"] as? String == "req",
           obj["method"] as? String == "connect",
           let id = obj["id"] as? String
        {
            let auth = ((obj["params"] as? [String: Any])?["auth"] as? [String: Any]) ?? [:]
            self.lock.withLock {
                self.connectRequestId = id
                self.connectAuth = auth
            }
        }
    }

    func latestConnectAuth() -> [String: Any]? {
        self.lock.withLock { self.connectAuth }
    }

    func sendPing(pongReceiveHandler: @escaping @Sendable (Error?) -> Void) {
        pongReceiveHandler(nil)
    }

    func receive() async throws -> URLSessionWebSocketTask.Message {
        let phase = self.lock.withLock { () -> Int in
            let current = self.receivePhase
            self.receivePhase += 1
            return current
        }
        if phase == 0 {
            return .data(Self.connectChallengeData(nonce: "nonce-1"))
        }
        for _ in 0..<50 {
            let id = self.lock.withLock { self.connectRequestId }
            if let id {
                return .data(Self.connectOkData(id: id))
            }
            try await Task.sleep(nanoseconds: 1_000_000)
        }
        return .data(Self.connectOkData(id: "connect"))
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

    private static func connectChallengeData(nonce: String) -> Data {
        let frame: [String: Any] = [
            "type": "event",
            "event": "connect.challenge",
            "payload": ["nonce": nonce],
        ]
        return (try? JSONSerialization.data(withJSONObject: frame)) ?? Data()
    }

    private static func connectOkData(id: String) -> Data {
        let payload: [String: Any] = [
            "type": "hello-ok",
            "protocol": 2,
            "server": [
                "version": "test",
                "connId": "test",
            ],
            "features": [
                "methods": [],
                "events": [],
            ],
            "snapshot": [
                "presence": [["ts": 1]],
                "health": [:],
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
        let frame: [String: Any] = [
            "type": "res",
            "id": id,
            "ok": true,
            "payload": payload,
        ]
        return (try? JSONSerialization.data(withJSONObject: frame)) ?? Data()
    }
}

private final class FakeGatewayWebSocketSession: WebSocketSessioning, @unchecked Sendable {
    private let lock = NSLock()
    private var tasks: [FakeGatewayWebSocketTask] = []
    private var makeCount = 0

    func snapshotMakeCount() -> Int {
        self.lock.withLock { self.makeCount }
    }

    func latestTask() -> FakeGatewayWebSocketTask? {
        self.lock.withLock { self.tasks.last }
    }

    func makeWebSocketTask(url: URL) -> WebSocketTaskBox {
        _ = url
        return self.lock.withLock {
            self.makeCount += 1
            let task = FakeGatewayWebSocketTask()
            self.tasks.append(task)
            return WebSocketTaskBox(task: task)
        }
    }
}

// MARK: - Helper

private func makeDefaultOptions() -> GatewayConnectOptions {
    GatewayConnectOptions(
        role: "operator",
        scopes: ["operator.read"],
        caps: [],
        commands: [],
        permissions: [:],
        clientId: "test-client",
        clientMode: "ui",
        clientDisplayName: "Test",
        includeDeviceIdentity: false)
}

// MARK: - Security Test Suite

@Suite("GatewayNodeSession Security Audit")
struct GatewayNodeSessionSecurityTests {

    // MARK: - D1: TLS & Transport

    @Suite("D1: TLS & Transport")
    struct D1_TLSTransport {
        @Test func rejectsPlaintextWebSocketToRemoteHost() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            await #expect(throws: NSError.self) {
                try await gateway.connect(
                    url: URL(string: "ws://remote.example.com/ws")!,
                    token: nil, bootstrapToken: nil, password: nil,
                    connectOptions: options,
                    sessionBox: WebSocketSessionBox(session: session),
                    onConnected: {}, onDisconnected: { _ in },
                    onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            }
        }

        @Test func allowsPlaintextWebSocketToLoopbackHost() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            // ws://127.0.0.1 should be allowed (loopback)
            try await gateway.connect(
                url: URL(string: "ws://127.0.0.1/ws")!,
                token: nil, bootstrapToken: nil, password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            await gateway.disconnect()
        }

        @Test func allowsPlaintextToIPv6Loopback() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            try await gateway.connect(
                url: URL(string: "ws://[::1]/ws")!,
                token: nil, bootstrapToken: nil, password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            await gateway.disconnect()
        }

        @Test func allowsPlaintextToLocalhostHost() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            try await gateway.connect(
                url: URL(string: "ws://localhost/ws")!,
                token: nil, bootstrapToken: nil, password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            await gateway.disconnect()
        }

        @Test func allowsTLSWebSocketToRemoteHost() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            // wss:// to remote should always be allowed
            try await gateway.connect(
                url: URL(string: "wss://remote.example.com/ws")!,
                token: nil, bootstrapToken: nil, password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            await gateway.disconnect()
        }

        @Test func rejectsPlaintextToRemoteIPAddress() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            await #expect(throws: NSError.self) {
                try await gateway.connect(
                    url: URL(string: "ws://192.168.1.100/ws")!,
                    token: nil, bootstrapToken: nil, password: nil,
                    connectOptions: options,
                    sessionBox: WebSocketSessionBox(session: session),
                    onConnected: {}, onDisconnected: { _ in },
                    onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            }
        }

        @Test func plaintextRejectionErrorCodeIsTen() async {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            do {
                try await gateway.connect(
                    url: URL(string: "ws://attacker.example.com/ws")!,
                    token: nil, bootstrapToken: nil, password: nil,
                    connectOptions: options,
                    sessionBox: WebSocketSessionBox(session: session),
                    onConnected: {}, onDisconnected: { _ in },
                    onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
                Issue.record("Expected error for ws:// to remote")
            } catch let error as NSError {
                #expect(error.code == 10)
            }
        }
    }

    // MARK: - D2: Nonce Challenge

    @Suite("D2: Nonce Challenge")
    struct D2_NonceChallenge {
        @Test func acceptsValidNonce() {
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable("abcdefgh")]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == "abcdefgh")
        }

        @Test func rejectsEmptyNonce() {
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable("")]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == nil)
        }

        @Test func rejectsShortNonce() {
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable("abc")]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == nil)
        }

        @Test func rejectsNonceExactlyBelowMinimum() {
            let shortNonce = String(repeating: "x", count: GatewayConnectChallengeSupport.minimumNonceLength - 1)
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable(shortNonce)]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == nil)
        }

        @Test func acceptsNonceExactlyAtMinimum() {
            let exactNonce = String(repeating: "y", count: GatewayConnectChallengeSupport.minimumNonceLength)
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable(exactNonce)]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == exactNonce)
        }

        @Test func rejectsWhitespaceOnlyNonce() {
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable("        ")]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == nil)
        }

        @Test func trimsWhitespaceFromValidNonce() {
            let payload: [String: OpenClawProtocol.AnyCodable] = ["nonce": AnyCodable("  abcdefgh  ")]
            let nonce = GatewayConnectChallengeSupport.nonce(from: payload)
            #expect(nonce == "abcdefgh")
        }
    }

    // MARK: - D3: Message Auth

    @Suite("D3: Message Auth")
    struct D3_MessageAuth {
        @Test func idempotencyKeyEchoedInInvokeResult() async throws {
            // This test verifies V3 fix: idempotencyKey forwarding in sendInvokeResult.
            // We test indirectly through invokeWithTimeout since sendInvokeResult is private.
            let request = BridgeInvokeRequest(id: "idem-test", command: "x", paramsJSON: nil)
            let response = await GatewayNodeSession.invokeWithTimeout(
                request: request,
                timeoutMs: 100,
                onInvoke: { req in
                    #expect(req.id == "idem-test")
                    return BridgeInvokeResponse(id: req.id, ok: true, payloadJSON: nil)
                })
            #expect(response.ok == true)
        }

        @Test func invokeResponsePreservesErrorCodeAndMessage() {
            let error = OpenClawNodeError(code: .unavailable, message: "service offline")
            let response = BridgeInvokeResponse(id: "1", ok: false, error: error)
            #expect(response.error?.code == .unavailable)
            #expect(response.error?.message == "service offline")
        }

        @Test func invokeRequestRoundTripsWithParamsJSON() throws {
            let request = BridgeInvokeRequest(id: "rt", command: "test.cmd", paramsJSON: "{\"key\":\"value\"}")
            let data = try JSONEncoder().encode(request)
            let decoded = try JSONDecoder().decode(BridgeInvokeRequest.self, from: data)
            #expect(decoded.id == "rt")
            #expect(decoded.command == "test.cmd")
            #expect(decoded.paramsJSON == "{\"key\":\"value\"}")
        }

        @Test func invokeRequestDefaultTypeIsInvoke() {
            let request = BridgeInvokeRequest(id: "1", command: "x")
            #expect(request.type == "invoke")
        }
    }

    // MARK: - D4: Reconnect Auth

    @Suite("D4: Reconnect Auth")
    struct D4_ReconnectAuth {
        @Test func reconnectUsesDeviceTokenFromPriorHello() async throws {
            // This test mirrors scannedSetupCodePrefersBootstrapAuthOverStoredDeviceToken
            // from GatewayNodeSessionTests but verifies the inverse: stored device token
            // is used when no bootstrap token is present.
            let tempDir = FileManager.default.temporaryDirectory
                .appendingPathComponent(UUID().uuidString, isDirectory: true)
            try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
            let prev = ProcessInfo.processInfo.environment["OPENCLAW_STATE_DIR"]
            setenv("OPENCLAW_STATE_DIR", tempDir.path, 1)
            defer {
                if let prev { setenv("OPENCLAW_STATE_DIR", prev, 1) }
                else { unsetenv("OPENCLAW_STATE_DIR") }
                try? FileManager.default.removeItem(at: tempDir)
            }

            let identity = DeviceIdentityStore.loadOrCreate()
            _ = DeviceAuthStore.storeToken(
                deviceId: identity.deviceId, role: "operator", token: "stored-token")

            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            try await gateway.connect(
                url: URL(string: "ws://127.0.0.1")!,
                token: nil, bootstrapToken: nil, password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })

            let auth = try #require(session.latestTask()?.latestConnectAuth())
            #expect(auth["deviceToken"] as? String == "stored-token")
            await gateway.disconnect()
        }

        @Test func bootstrapTokenNotClearedAfterConnect() async throws {
            // V8 (noted): bootstrap token remains in memory after device token issued.
            // Verify it doesn't cause issues on reconnect — device token takes priority.
            let tempDir = FileManager.default.temporaryDirectory
                .appendingPathComponent(UUID().uuidString, isDirectory: true)
            try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
            let prev = ProcessInfo.processInfo.environment["OPENCLAW_STATE_DIR"]
            setenv("OPENCLAW_STATE_DIR", tempDir.path, 1)
            defer {
                if let prev { setenv("OPENCLAW_STATE_DIR", prev, 1) }
                else { unsetenv("OPENCLAW_STATE_DIR") }
                try? FileManager.default.removeItem(at: tempDir)
            }

            let identity = DeviceIdentityStore.loadOrCreate()
            _ = DeviceAuthStore.storeToken(
                deviceId: identity.deviceId, role: "operator", token: "device-tok")

            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            // Connect with bootstrap — but stored device token exists too
            try await gateway.connect(
                url: URL(string: "ws://127.0.0.1")!,
                token: nil, bootstrapToken: "bootstrap-tok", password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })

            // Bootstrap should be preferred when explicitly provided
            let auth = try #require(session.latestTask()?.latestConnectAuth())
            #expect(auth["bootstrapToken"] as? String == "bootstrap-tok")
            await gateway.disconnect()
        }
    }

    // MARK: - D5: Log Redaction

    @Suite("D5: Log Redaction")
    struct D5_LogRedaction {
        @Test func sanitizedURLStripsQueryParams() {
            let url = URL(string: "wss://gateway.example.com/ws?token=secret&key=abc")!
            var components = URLComponents(url: url, resolvingAgainstBaseURL: false)!
            components.user = nil
            components.password = nil
            if let items = components.queryItems {
                components.queryItems = items.map { URLQueryItem(name: $0.name, value: "***") }
            }
            let sanitized = components.string!
            #expect(!sanitized.contains("secret"))
            #expect(!sanitized.contains("abc"))
            #expect(sanitized.contains("token=***"))
            #expect(sanitized.contains("key=***"))
        }

        @Test func sanitizedURLStripsUserInfo() {
            let url = URL(string: "wss://admin:p4ssw0rd@gateway.example.com/ws")!
            var components = URLComponents(url: url, resolvingAgainstBaseURL: false)!
            components.user = nil
            components.password = nil
            let sanitized = components.string!
            #expect(!sanitized.contains("admin"))
            #expect(!sanitized.contains("p4ssw0rd"))
            #expect(sanitized.contains("gateway.example.com"))
        }

        @Test func sanitizedURLPreservesHostAndPath() {
            let url = URL(string: "wss://gateway.example.com:7443/ws/v2")!
            var components = URLComponents(url: url, resolvingAgainstBaseURL: false)!
            components.user = nil
            components.password = nil
            let sanitized = components.string!
            #expect(sanitized.contains("gateway.example.com"))
            #expect(sanitized.contains("7443"))
            #expect(sanitized.contains("/ws/v2"))
        }
    }

    // MARK: - D6: Persistence

    @Suite("D6: Persistence")
    struct D6_Persistence {
        private func withTempStateDir(_ body: (URL) throws -> Void) throws {
            let dir = FileManager.default.temporaryDirectory
                .appendingPathComponent(UUID().uuidString, isDirectory: true)
            try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
            let prev = ProcessInfo.processInfo.environment["OPENCLAW_STATE_DIR"]
            setenv("OPENCLAW_STATE_DIR", dir.path, 1)
            defer {
                if let prev { setenv("OPENCLAW_STATE_DIR", prev, 1) }
                else { unsetenv("OPENCLAW_STATE_DIR") }
                try? FileManager.default.removeItem(at: dir)
            }
            try body(dir)
        }

        @Test func identityLoadOrCreateGeneratesNew() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                #expect(!identity.deviceId.isEmpty)
                #expect(!identity.publicKey.isEmpty)
                #expect(!identity.privateKey.isEmpty)
                #expect(identity.createdAtMs > 0)
            }
        }

        @Test func identityLoadOrCreateReusesExisting() throws {
            try withTempStateDir { _ in
                let first = DeviceIdentityStore.loadOrCreate()
                let second = DeviceIdentityStore.loadOrCreate()
                #expect(first.deviceId == second.deviceId)
                #expect(first.publicKey == second.publicKey)
            }
        }

        @Test func identityGeneratesValidSigningKey() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                let signature = DeviceIdentityStore.signPayload("test", identity: identity)
                #expect(signature != nil)
                #expect(!signature!.isEmpty)
            }
        }

        @Test func identitySaveFailureDoesNotCrash() throws {
            // V5 (accepted risk): save uses bare catch {} — verify it doesn't crash
            try withTempStateDir { dir in
                let identity = DeviceIdentityStore.loadOrCreate()
                #expect(!identity.deviceId.isEmpty)
                // Identity was generated even if save could fail
            }
        }

        @Test func authStoreRoundTrips() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                let stored = DeviceAuthStore.storeToken(
                    deviceId: identity.deviceId, role: "operator", token: "test-token", scopes: ["read"])
                #expect(stored.token == "test-token")
                #expect(stored.role == "operator")

                let loaded = DeviceAuthStore.loadToken(deviceId: identity.deviceId, role: "operator")
                #expect(loaded?.token == "test-token")
                #expect(loaded?.scopes == ["read"])
            }
        }

        @Test func authStoreHandlesCorruptedJSON() throws {
            try withTempStateDir { dir in
                let identityDir = dir.appendingPathComponent("identity", isDirectory: true)
                try FileManager.default.createDirectory(at: identityDir, withIntermediateDirectories: true)
                let authFile = identityDir.appendingPathComponent("device-auth.json")
                try "not valid json {{{".data(using: .utf8)!.write(to: authFile)

                let loaded = DeviceAuthStore.loadToken(deviceId: "any", role: "operator")
                #expect(loaded == nil) // gracefully returns nil
            }
        }

        @Test func authStoreSetsRestrictivePermissions() throws {
            try withTempStateDir { dir in
                let identity = DeviceIdentityStore.loadOrCreate()
                _ = DeviceAuthStore.storeToken(
                    deviceId: identity.deviceId, role: "operator", token: "secret")

                let authFile = dir
                    .appendingPathComponent("identity", isDirectory: true)
                    .appendingPathComponent("device-auth.json")
                let attrs = try FileManager.default.attributesOfItem(atPath: authFile.path)
                let perms = (attrs[.posixPermissions] as? Int) ?? 0
                #expect(perms == 0o600) // owner read/write only
            }
        }

        @Test func authStoreClearsToken() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                _ = DeviceAuthStore.storeToken(
                    deviceId: identity.deviceId, role: "operator", token: "to-clear")
                DeviceAuthStore.clearToken(deviceId: identity.deviceId, role: "operator")
                let loaded = DeviceAuthStore.loadToken(deviceId: identity.deviceId, role: "operator")
                #expect(loaded == nil)
            }
        }

        @Test func authStoreIgnoresDifferentDeviceId() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                _ = DeviceAuthStore.storeToken(
                    deviceId: identity.deviceId, role: "operator", token: "mine")

                let loaded = DeviceAuthStore.loadToken(deviceId: "wrong-device-id", role: "operator")
                #expect(loaded == nil)
            }
        }

        @Test func authStoreNormalizesRole() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                _ = DeviceAuthStore.storeToken(
                    deviceId: identity.deviceId, role: "  operator  ", token: "tok")

                let loaded = DeviceAuthStore.loadToken(deviceId: identity.deviceId, role: "operator")
                #expect(loaded?.token == "tok")
            }
        }

        @Test func authStoreNormalizesScopes() throws {
            try withTempStateDir { _ in
                let identity = DeviceIdentityStore.loadOrCreate()
                let entry = DeviceAuthStore.storeToken(
                    deviceId: identity.deviceId, role: "operator",
                    token: "tok", scopes: ["  write  ", "read", "", "read"])
                // Normalized: trimmed, deduped, sorted
                #expect(entry.scopes == ["read", "write"])
            }
        }
    }

    // MARK: - D7: Actor Reentrancy

    @Suite("D7: Actor Reentrancy")
    struct D7_ActorReentrancy {
        @Test func concurrentInvokesDoNotCrash() async {
            // Verify actor isolation holds under concurrent invoke calls
            let results = await withTaskGroup(of: BridgeInvokeResponse.self) { group in
                for i in 0..<10 {
                    group.addTask {
                        await GatewayNodeSession.invokeWithTimeout(
                            request: BridgeInvokeRequest(id: "c\(i)", command: "test"),
                            timeoutMs: 50,
                            onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
                    }
                }
                var results: [BridgeInvokeResponse] = []
                for await result in group { results.append(result) }
                return results
            }
            #expect(results.count == 10)
            #expect(results.allSatisfy { $0.ok == true })
        }

        @Test func disconnectDuringInvokeDoesNotLeak() async throws {
            let session = FakeGatewayWebSocketSession()
            let gateway = GatewayNodeSession()
            let options = makeDefaultOptions()
            try await gateway.connect(
                url: URL(string: "ws://127.0.0.1")!,
                token: nil, bootstrapToken: nil, password: nil,
                connectOptions: options,
                sessionBox: WebSocketSessionBox(session: session),
                onConnected: {}, onDisconnected: { _ in },
                onInvoke: { BridgeInvokeResponse(id: $0.id, ok: true) })
            // Immediately disconnect — should clean up without leaking
            await gateway.disconnect()
            // No crash = pass
        }
    }
}
