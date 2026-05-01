import Foundation
import Testing
@testable import OpenClawChatUI

@Suite("ChatLinkAllowlist")
struct ChatLinkAllowlistTests {

    @Test("https with host is allowed")
    func httpsAllowed() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "https://example.com")!) == .allow)
    }

    @Test("http with host is allowed")
    func httpAllowed() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "http://example.com")!) == .allow)
    }

    @Test("mailto is allowed")
    func mailtoAllowed() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "mailto:rod@example.com")!) == .allow)
    }

    @Test("uppercase HTTPS is allowed (case-insensitive scheme)")
    func uppercaseHttpsAllowed() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "HTTPS://example.com")!) == .allow)
    }

    @Test("mixed-case MailTo is allowed")
    func mixedCaseMailtoAllowed() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "MailTo:rod@example.com")!) == .allow)
    }

    @Test("https with empty host is discarded", arguments: [
        "https://",
        "http://",
    ])
    func emptyHostDiscarded(urlString: String) {
        let url = URL(string: urlString)!
        #expect(ChatLinkAllowlist.decision(for: url) == .discard)
    }

    @Test("tel: is discarded", arguments: [
        "tel:+15551234567",
        "tel:5551234567",
    ])
    func telDiscarded(urlString: String) {
        #expect(ChatLinkAllowlist.decision(for: URL(string: urlString)!) == .discard)
    }

    @Test("sms: is discarded")
    func smsDiscarded() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "sms:+15551234567")!) == .discard)
    }

    @Test("facetime schemes are discarded", arguments: [
        "facetime://user@example.com",
        "facetime-audio://user@example.com",
    ])
    func facetimeDiscarded(urlString: String) {
        #expect(ChatLinkAllowlist.decision(for: URL(string: urlString)!) == .discard)
    }

    @Test("javascript: is discarded")
    func javascriptDiscarded() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "javascript:alert(1)")!) == .discard)
    }

    @Test("data URL is discarded")
    func dataDiscarded() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "data:text/html,<script>alert(1)</script>")!) == .discard)
    }

    @Test("file URL is discarded")
    func fileDiscarded() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "file:///etc/passwd")!) == .discard)
    }

    @Test("custom app schemes are discarded", arguments: [
        "myapp://action",
        "shortcuts://run-shortcut?name=foo",
        "itms-services://?action=download-manifest",
    ])
    func customSchemeDiscarded(urlString: String) {
        #expect(ChatLinkAllowlist.decision(for: URL(string: urlString)!) == .discard)
    }

    @Test("ftp: is discarded")
    func ftpDiscarded() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "ftp://example.com/x")!) == .discard)
    }

    @Test("URL with no scheme is discarded")
    func noSchemeDiscarded() {
        #expect(ChatLinkAllowlist.decision(for: URL(string: "//example.com")!) == .discard)
    }
}
