import Foundation
import Testing
@testable import OpenClawKit

// MARK: - Domain 6: Persistence Failure Modes

@Suite(.serialized)
struct D6_DeviceIdentityPersistenceTests {

    @Test
    func identityLoadHandlesCorruptedJSON() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        // Write garbage bytes to device.json
        try? FileManager.default.createDirectory(
            at: stateDir.identityDir, withIntermediateDirectories: true)
        try? Data([0xFF, 0xFE, 0x00, 0x01, 0xDE, 0xAD]).write(to: stateDir.deviceJsonURL)

        // loadOrCreate should handle corruption and generate fresh identity
        let identity = DeviceIdentityStore.loadOrCreate()
        #expect(!identity.deviceId.isEmpty)
        #expect(!identity.publicKey.isEmpty)
        #expect(!identity.privateKey.isEmpty)
    }

    @Test
    func identityLoadHandlesEmptyFile() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        try? FileManager.default.createDirectory(
            at: stateDir.identityDir, withIntermediateDirectories: true)
        try? Data().write(to: stateDir.deviceJsonURL)

        let identity = DeviceIdentityStore.loadOrCreate()
        #expect(!identity.deviceId.isEmpty)
    }

    @Test
    func identityLoadHandlesPartialJSON() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        try? FileManager.default.createDirectory(
            at: stateDir.identityDir, withIntermediateDirectories: true)
        let truncated = #"{"deviceId":"abc","publicKey":"#
        try? truncated.data(using: .utf8)?.write(to: stateDir.deviceJsonURL)

        let identity = DeviceIdentityStore.loadOrCreate()
        #expect(!identity.deviceId.isEmpty)
        // Should have generated a new identity, not returned the partial one
        #expect(identity.deviceId != "abc")
    }

    @Test
    func identityLoadHandlesEmptyFields() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        try? FileManager.default.createDirectory(
            at: stateDir.identityDir, withIntermediateDirectories: true)
        let emptyFields = #"{"deviceId":"","publicKey":"","privateKey":"","createdAtMs":0}"#
        try? emptyFields.data(using: .utf8)?.write(to: stateDir.deviceJsonURL)

        let identity = DeviceIdentityStore.loadOrCreate()
        // Should regenerate since fields are empty
        #expect(!identity.deviceId.isEmpty)
        #expect(!identity.publicKey.isEmpty)
    }

    @Test
    func identityRegenerationAfterCorruptionYieldsNewDeviceId() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        // First: create a valid identity
        let original = DeviceIdentityStore.loadOrCreate()
        let originalId = original.deviceId

        // Corrupt the file
        try? Data([0xFF, 0xFE]).write(to: stateDir.deviceJsonURL)

        // Reload — should generate new identity with different deviceId
        let regenerated = DeviceIdentityStore.loadOrCreate()
        #expect(regenerated.deviceId != originalId,
                "Regenerated identity should have a different deviceId (pairing is lost)")
    }

    @Test
    func identitySignPayloadProducesValidSignature() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        let identity = DeviceIdentityStore.loadOrCreate()
        let payload = "test-payload-data"
        let signature = DeviceIdentityStore.signPayload(payload, identity: identity)

        #expect(signature != nil)
        #expect(!signature!.isEmpty)
        // Signature should be base64url encoded (no +, /, or = characters)
        #expect(!signature!.contains("+"))
        #expect(!signature!.contains("/"))
        #expect(!signature!.contains("="))
    }
}

@Suite(.serialized)
struct D6_DeviceAuthStorePersistenceTests {

    @Test
    func authStoreHandlesCorruptedJSON() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        try? FileManager.default.createDirectory(
            at: stateDir.identityDir, withIntermediateDirectories: true)
        try? Data([0xDE, 0xAD, 0xBE, 0xEF]).write(to: stateDir.deviceAuthJsonURL)

        let token = DeviceAuthStore.loadToken(deviceId: "any-id", role: "operator")
        #expect(token == nil)
    }

    @Test
    func authStoreHandlesMismatchedDeviceId() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        _ = DeviceAuthStore.storeToken(
            deviceId: "device-A",
            role: "operator",
            token: "token-for-A")

        let result = DeviceAuthStore.loadToken(deviceId: "device-B", role: "operator")
        #expect(result == nil)
    }

    @Test
    func authStoreHandlesVersionMismatch() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        try? FileManager.default.createDirectory(
            at: stateDir.identityDir, withIntermediateDirectories: true)
        let futureVersion = #"{"version":99,"deviceId":"test","tokens":{}}"#
        try? futureVersion.data(using: .utf8)?.write(to: stateDir.deviceAuthJsonURL)

        let token = DeviceAuthStore.loadToken(deviceId: "test", role: "operator")
        #expect(token == nil)
    }

    @Test
    func authStoreSetsRestrictivePermissions() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        _ = DeviceAuthStore.storeToken(
            deviceId: "test-device",
            role: "operator",
            token: "test-token")

        let attrs = try? FileManager.default.attributesOfItem(atPath: stateDir.deviceAuthJsonURL.path)
        if let permissions = attrs?[.posixPermissions] as? Int {
            #expect(permissions == 0o600, "Auth store file should have 0600 permissions")
        }
        // On platforms where posixPermissions is not available, this is a no-op
    }

    @Test
    func authStoreScopesNormalized() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        let entry = DeviceAuthStore.storeToken(
            deviceId: "test-device",
            role: "  operator  ",
            token: "test-token",
            scopes: ["  operator.write  ", "operator.read", "", "operator.read"])

        // Role should be trimmed
        #expect(entry.role == "operator")
        // Scopes should be trimmed, deduplicated, and sorted
        #expect(entry.scopes == ["operator.read", "operator.write"])
    }

    @Test
    func authStoreClearTokenRemovesEntry() {
        let stateDir = TempStateDir()
        defer { stateDir.cleanup() }

        _ = DeviceAuthStore.storeToken(
            deviceId: "test-device",
            role: "operator",
            token: "test-token")

        #expect(DeviceAuthStore.loadToken(deviceId: "test-device", role: "operator") != nil)

        DeviceAuthStore.clearToken(deviceId: "test-device", role: "operator")

        #expect(DeviceAuthStore.loadToken(deviceId: "test-device", role: "operator") == nil)
    }
}
