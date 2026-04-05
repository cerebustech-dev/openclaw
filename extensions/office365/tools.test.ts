import { describe, expect, it, vi, beforeEach } from "vitest";
import type { GraphClient } from "./src/graph-client.js";
import { createEmailListTool } from "./src/tools/email-list.js";
import { createEmailReadTool } from "./src/tools/email-read.js";
import { createEmailSendTool } from "./src/tools/email-send.js";
import { createEmailReplyTool } from "./src/tools/email-reply.js";
import { createEmailSearchTool } from "./src/tools/email-search.js";
import { GraphApiError } from "./src/types.js";
import { formatEventSummary, validateDateRange, checkConflicts } from "./src/tools/_calendar-shared.js";
import { createCalendarListTool } from "./src/tools/calendar-list.js";
import { createCalendarUpdateTool } from "./src/tools/calendar-update.js";
import { createCalendarDeleteTool } from "./src/tools/calendar-delete.js";
import { createCalendarCreateTool } from "./src/tools/calendar-create.js";
import { createEmailAttachmentReadTool } from "./src/tools/email-attachment-read.js";
import type { GraphEvent } from "./src/types.js";

// ── Mock Graph client ───────────────────────────────────────────────────────

function createMockGraphClient(): GraphClient & {
  _fetchJsonMock: ReturnType<typeof vi.fn>;
  _fetchMock: ReturnType<typeof vi.fn>;
} {
  const fetchJsonMock = vi.fn();
  const fetchMock = vi.fn().mockResolvedValue(new Response("", { status: 200 }));
  return {
    fetchJson: fetchJsonMock,
    fetch: fetchMock,
    setCredential: vi.fn(),
    _fetchJsonMock: fetchJsonMock,
    _fetchMock: fetchMock,
  };
}

function parseResult(result: { content: Array<{ text?: string }> }) {
  return JSON.parse(result.content[0].text ?? "{}");
}

// ── email_list tests ────────────────────────────────────────────────────────

describe("email_list", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createEmailListTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createEmailListTool({ graphClient: client });
  });

  it("with defaults calls /me/messages with correct query params", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", {});

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages",
      expect.objectContaining({
        $top: "10",
        $orderby: "receivedDateTime desc",
        $count: "true",
        $select: expect.stringContaining("id,subject,from"),
      }),
      expect.any(Object),
    );
  });

  it("with folder param uses /me/mailFolders/{folder}/messages", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", { folder: "SentItems" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/mailFolders/SentItems/messages",
      expect.any(Object),
      expect.any(Object),
    );
  });

  it("maps friendly folder names to Graph identifiers", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", { folder: "sent" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/mailFolders/SentItems/messages",
      expect.any(Object),
      expect.any(Object),
    );
  });

  it("returns actionable error for unknown folder", async () => {
    const result = await tool.execute("id", { folder: "bogus" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("Unknown folder");
    expect(parsed.error.message).toContain("SentItems");
  });

  it("clamps top to 1-50", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", { top: 999 });
    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ $top: "50" }),
      expect.any(Object),
    );

    await tool.execute("id", { top: -5 });
    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ $top: "1" }),
      expect.any(Object),
    );
  });

  it("wraps search in quotes and adds ConsistencyLevel header", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", { search: "hello world" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ $search: '"hello world"' }),
      expect.objectContaining({ ConsistencyLevel: "eventual" }),
    );
  });

  it("rejects combined search + filter", async () => {
    const result = await tool.execute("id", {
      search: "test",
      filter: "isRead eq false",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("Cannot combine");
  });

  it("warns about ignored orderBy when search is used", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    const result = await tool.execute("id", {
      search: "test",
      orderBy: "subject asc",
    });
    const parsed = parseResult(result);

    expect(parsed.data.warnings).toBeDefined();
    expect(parsed.data.warnings[0]).toContain("ignored");
  });

  it("formats message summaries correctly", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [
        {
          id: "msg-1",
          subject: "Test Subject",
          from: { emailAddress: { name: "Sender Name", address: "sender@test.com" } },
          toRecipients: [
            { emailAddress: { address: "to@test.com" } },
          ],
          receivedDateTime: "2026-04-02T10:00:00Z",
          isRead: false,
          hasAttachments: true,
          bodyPreview: "Preview text...",
          importance: "high",
          flag: { flagStatus: "flagged" },
        },
      ],
      "@odata.count": 1,
    });

    const result = await tool.execute("id", {});
    const parsed = parseResult(result);

    expect(parsed.data.messages).toHaveLength(1);
    const msg = parsed.data.messages[0];
    expect(msg.id).toBe("msg-1");
    expect(msg.subject).toBe("Test Subject");
    expect(msg.from).toBe("Sender Name <sender@test.com>");
    expect(msg.to).toEqual(["to@test.com"]);
    expect(msg.isRead).toBe(false);
    expect(msg.hasAttachments).toBe(true);
    expect(msg.importance).toBe("high");
    expect(msg.flagStatus).toBe("flagged");
  });

  it("handles non-ASCII subjects correctly", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [
        {
          id: "msg-1",
          subject: "日本語のメール — Ünïcödë",
          from: { emailAddress: { address: "test@test.com" } },
          toRecipients: [],
        },
      ],
    });

    const result = await tool.execute("id", {});
    const parsed = parseResult(result);
    expect(parsed.data.messages[0].subject).toBe("日本語のメール — Ünïcödë");
  });

  it("returns pagination info", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{ id: "msg-1", subject: "Test" }],
      "@odata.count": 42,
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
    });

    const result = await tool.execute("id", { top: 10 });
    const parsed = parseResult(result);

    expect(parsed.data.totalCount).toBe(42);
    expect(parsed.data.hasMore).toBe(true);
    expect(parsed.data.nextSkip).toBe(10);
  });

  // ── Issue 7: search backslash escaping ─────────────────────────────────

  it("escapes backslashes in search before quotes", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", { search: 'path\\to\\"file' });

    const call = client._fetchJsonMock.mock.calls[0];
    // Backslashes escaped first, then quotes
    expect(call[1].$search).toBe('"path\\\\to\\\\\\"file"');
  });

  it("handles search ending with backslash without producing malformed OData", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", { search: "trailing\\" });

    const call = client._fetchJsonMock.mock.calls[0];
    // A trailing \ must be escaped to \\ so the query is: "trailing\\"
    expect(call[1].$search).toBe('"trailing\\\\"');
  });

  // ── Issue 8: folder ID validation ──────────────────────────────────────

  it("rejects folder strings with path separators", async () => {
    const result = await tool.execute("id", { folder: "../../etc/passwd" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
  });

  it("rejects folder strings with spaces", async () => {
    const result = await tool.execute("id", { folder: "some folder with spaces and more" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
  });

  it("accepts valid Graph API folder IDs (alphanumeric + special chars)", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    // Typical Graph folder ID format
    await tool.execute("id", { folder: "AAMkADQ0MTg3MDMyLTY5ZTItNGI3ZS04OTI3LWRjNjI=" });

    expect(client._fetchJsonMock).toHaveBeenCalled();
  });

  // ── Issue 12: error message sanitization ───────────────────────────────

  it("sanitizes error messages from non-GraphApiError exceptions", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new Error("GET https://graph.microsoft.com/v1.0/me/messages?token=secret123 failed"),
    );

    const result = await tool.execute("id", {});
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    // Should NOT leak the raw URL with token
    expect(parsed.error.message).not.toContain("secret123");
    expect(parsed.error.message).not.toContain("graph.microsoft.com");
  });

  it("uses GraphApiError message directly (already sanitized)", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Graph API /me/messages failed (403). Check that the Azure app has the required scopes.", "permission", 403),
    );

    const result = await tool.execute("id", {});
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("permission");
    expect(parsed.error.message).toContain("required scopes");
  });
});

// ── email_read tests ────────────────────────────────────────────────────────

describe("email_read", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createEmailReadTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createEmailReadTool({ graphClient: client });
  });

  it("fetches full message by ID", async () => {
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      subject: "Test",
      from: { emailAddress: { name: "Sender", address: "s@test.com" } },
      toRecipients: [{ emailAddress: { address: "to@test.com" } }],
      ccRecipients: [],
      bccRecipients: [],
      replyTo: [],
      body: { contentType: "text", content: "Hello world" },
      isRead: true,
      hasAttachments: false,
      importance: "normal",
      flag: { flagStatus: "notFlagged" },
      conversationId: "conv-1",
    });

    const result = await tool.execute("id", { messageId: "msg-1" });
    const parsed = parseResult(result);

    expect(parsed.data.id).toBe("msg-1");
    expect(parsed.data.subject).toBe("Test");
    expect(parsed.data.from).toBe("Sender <s@test.com>");
    expect(parsed.data.body).toBe("Hello world");
    expect(parsed.data.bodyTruncated).toBe(false);
    expect(parsed.data.conversationId).toBe("conv-1");
  });

  it("fetches attachment metadata when hasAttachments is true", async () => {
    client._fetchJsonMock
      .mockResolvedValueOnce({
        id: "msg-1",
        hasAttachments: true,
        body: { contentType: "text", content: "Body" },
      })
      .mockResolvedValueOnce({
        value: [
          { id: "att-1", name: "file.pdf", contentType: "application/pdf", size: 12345 },
        ],
      });

    const result = await tool.execute("id", { messageId: "msg-1" });
    const parsed = parseResult(result);

    expect(parsed.data.attachments).toHaveLength(1);
    expect(parsed.data.attachments[0].name).toBe("file.pdf");

    // Verify attachment call used the filter
    const attachCall = client._fetchJsonMock.mock.calls[1];
    expect(attachCall[1]).toEqual(
      expect.objectContaining({ $filter: "isInline eq false" }),
    );
  });

  it("sends PATCH when markAsRead is true", async () => {
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      isRead: false,
      hasAttachments: false,
      body: { contentType: "text", content: "Body" },
    });

    await tool.execute("id", { messageId: "msg-1", markAsRead: true });

    expect(client._fetchMock).toHaveBeenCalledWith(
      "/me/messages/msg-1",
      expect.objectContaining({
        method: "PATCH",
        body: JSON.stringify({ isRead: true }),
      }),
    );
  });

  it("does not PATCH when already read", async () => {
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      isRead: true,
      hasAttachments: false,
      body: { contentType: "text", content: "Body" },
    });

    await tool.execute("id", { messageId: "msg-1", markAsRead: true });

    expect(client._fetchMock).not.toHaveBeenCalled();
  });

  it("truncates body over 100KB", async () => {
    const largeBody = "A".repeat(200 * 1024);
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      hasAttachments: false,
      body: { contentType: "text", content: largeBody },
    });

    const result = await tool.execute("id", { messageId: "msg-1" });
    const parsed = parseResult(result);

    expect(parsed.data.bodyTruncated).toBe(true);
    expect(parsed.data.body).toContain("[... truncated at 100KB]");
    // Verify it's actually shorter
    expect(parsed.data.body.length).toBeLessThan(largeBody.length);
  });

  it("sets Prefer header for bodyFormat", async () => {
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      hasAttachments: false,
      body: { contentType: "html", content: "<p>Hello</p>" },
    });

    await tool.execute("id", { messageId: "msg-1", bodyFormat: "html" });

    const call = client._fetchJsonMock.mock.calls[0];
    expect(call[2]).toEqual(
      expect.objectContaining({ Prefer: 'outlook.body-content-type="html"' }),
    );
  });

  it("returns error for missing messageId", async () => {
    const result = await tool.execute("id", { messageId: "" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
  });

  // ── Issue 9: HTML body truncation safety ───────────────────────────────

  it("marks truncated HTML body as unsafe", async () => {
    const largeHtml = "<div>" + "A".repeat(200 * 1024) + "</div>";
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      hasAttachments: false,
      body: { contentType: "html", content: largeHtml },
    });

    const result = await tool.execute("id", { messageId: "msg-1", bodyFormat: "html" });
    const parsed = parseResult(result);

    expect(parsed.data.bodyTruncated).toBe(true);
    expect(parsed.data.bodyUnsafeHtml).toBe(true);
    expect(parsed.data.body).toContain("truncated");
  });

  it("does not mark non-truncated HTML as unsafe", async () => {
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      hasAttachments: false,
      body: { contentType: "html", content: "<p>Short HTML</p>" },
    });

    const result = await tool.execute("id", { messageId: "msg-1", bodyFormat: "html" });
    const parsed = parseResult(result);

    expect(parsed.data.bodyTruncated).toBe(false);
    expect(parsed.data.bodyUnsafeHtml).toBeUndefined();
  });

  it("does not mark truncated text body as unsafe html", async () => {
    const largeText = "A".repeat(200 * 1024);
    client._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      hasAttachments: false,
      body: { contentType: "text", content: largeText },
    });

    const result = await tool.execute("id", { messageId: "msg-1" });
    const parsed = parseResult(result);

    expect(parsed.data.bodyTruncated).toBe(true);
    expect(parsed.data.bodyUnsafeHtml).toBeUndefined();
  });

  // ── Issue 12: error message sanitization ───────────────────────────────

  it("sanitizes error messages from non-GraphApiError exceptions", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new Error("GET https://graph.microsoft.com/v1.0/me/messages/abc?token=secret123 failed"),
    );

    const result = await tool.execute("id", { messageId: "abc" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    // Should NOT leak the raw URL with token
    expect(parsed.error.message).not.toContain("secret123");
    expect(parsed.error.message).not.toContain("graph.microsoft.com");
  });

  it("uses GraphApiError message directly (already sanitized)", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Graph API /me/messages/abc failed (404). The requested resource was not found.", "not_found", 404),
    );

    const result = await tool.execute("id", { messageId: "abc" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("not_found");
    expect(parsed.error.message).toContain("not found");
  });
});

// ── email_send tests ───────────────────────────────────────────────────────

describe("email_send", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createEmailSendTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createEmailSendTool({ graphClient: client });
  });

  it("sends email with required fields", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const result = await tool.execute("id", {
      to: ["alice@example.com"],
      subject: "Hello",
      body: "<p>Hi Alice</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.data).toBeDefined();
    expect(parsed.error).toBeUndefined();
    expect(parsed.data.sent).toBe(true);

    expect(client._fetchMock).toHaveBeenCalledWith(
      "/me/sendMail",
      expect.objectContaining({
        method: "POST",
        body: expect.any(String),
      }),
    );

    // Verify the POST body structure
    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.subject).toBe("Hello");
    expect(callBody.message.body).toEqual({
      contentType: "HTML",
      content: "<p>Hi Alice</p>",
    });
    expect(callBody.message.toRecipients).toEqual([
      { emailAddress: { address: "alice@example.com" } },
    ]);
  });

  it("sends email with cc and bcc", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      to: ["alice@example.com"],
      subject: "Test",
      body: "<p>Body</p>",
      cc: ["bob@example.com"],
      bcc: ["charlie@example.com"],
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.ccRecipients).toEqual([
      { emailAddress: { address: "bob@example.com" } },
    ]);
    expect(callBody.message.bccRecipients).toEqual([
      { emailAddress: { address: "charlie@example.com" } },
    ]);
  });

  it("handles multiple to recipients", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      to: ["alice@example.com", "bob@example.com"],
      subject: "Group",
      body: "<p>Hey all</p>",
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.toRecipients).toHaveLength(2);
    expect(callBody.message.toRecipients[1].emailAddress.address).toBe("bob@example.com");
  });

  it("accepts a single string for to (not just array)", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const result = await tool.execute("id", {
      to: "alice@example.com",
      subject: "Solo",
      body: "<p>Hi</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.data.sent).toBe(true);

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.toRecipients).toEqual([
      { emailAddress: { address: "alice@example.com" } },
    ]);
  });

  it("returns error for missing to", async () => {
    const result = await tool.execute("id", {
      subject: "Hello",
      body: "<p>Hi</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("to");
  });

  it("returns error for empty to array", async () => {
    const result = await tool.execute("id", {
      to: [],
      subject: "Hello",
      body: "<p>Hi</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
  });

  it("returns error for missing subject", async () => {
    const result = await tool.execute("id", {
      to: ["alice@example.com"],
      body: "<p>Hi</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("subject");
  });

  it("returns error for missing body", async () => {
    const result = await tool.execute("id", {
      to: ["alice@example.com"],
      subject: "Hello",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("body");
  });

  it("omits cc/bcc from request when not provided", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      to: ["alice@example.com"],
      subject: "Test",
      body: "<p>Body</p>",
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.ccRecipients).toBeUndefined();
    expect(callBody.message.bccRecipients).toBeUndefined();
  });

  it("handles GraphApiError", async () => {
    client._fetchMock.mockRejectedValue(
      new GraphApiError("Graph API /me/sendMail failed (403). Check that the Azure app has the required scopes.", "permission", 403),
    );

    const result = await tool.execute("id", {
      to: ["alice@example.com"],
      subject: "Hello",
      body: "<p>Hi</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("permission");
    expect(parsed.error.message).toContain("required scopes");
  });

  it("sanitizes error messages from non-GraphApiError exceptions", async () => {
    client._fetchMock.mockRejectedValue(
      new Error("POST https://graph.microsoft.com/v1.0/me/sendMail?token=secret123 failed"),
    );

    const result = await tool.execute("id", {
      to: ["alice@example.com"],
      subject: "Hello",
      body: "<p>Hi</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.message).not.toContain("secret123");
    expect(parsed.error.message).not.toContain("graph.microsoft.com");
  });

  // ── Round 3: Send with attachments ────────────────────────────────────────

  it("sends email with single attachment", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const result = await tool.execute("id", {
      to: "alice@example.com",
      subject: "Report",
      body: "<p>See attached</p>",
      attachments: [{
        name: "report.pdf",
        contentType: "application/pdf",
        contentBytes: "SGVsbG8=",
      }],
    });
    const parsed = parseResult(result);

    expect(parsed.data.sent).toBe(true);

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toHaveLength(1);
    expect(callBody.message.attachments[0]["@odata.type"]).toBe("#microsoft.graph.fileAttachment");
    expect(callBody.message.attachments[0].name).toBe("report.pdf");
    expect(callBody.message.attachments[0].contentType).toBe("application/pdf");
    expect(callBody.message.attachments[0].contentBytes).toBe("SGVsbG8=");
  });

  it("sends email with multiple attachments", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const attachments = [
      { name: "a.txt", contentType: "text/plain", contentBytes: "YQ==" },
      { name: "b.txt", contentType: "text/plain", contentBytes: "Yg==" },
      { name: "c.txt", contentType: "text/plain", contentBytes: "Yw==" },
    ];

    await tool.execute("id", {
      to: "alice@example.com",
      subject: "Files",
      body: "<p>3 files</p>",
      attachments,
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toHaveLength(3);
  });

  it("omits attachments from request when not provided", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      to: "alice@example.com",
      subject: "No attach",
      body: "<p>Hello</p>",
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toBeUndefined();
  });

  it("omits attachments from request when empty array", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      to: "alice@example.com",
      subject: "Empty",
      body: "<p>Hello</p>",
      attachments: [],
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toBeUndefined();
  });

  // ── Round 4: Send attachment edge cases ────────────────────────────────────

  it("returns error when attachment exceeds 3MB (decoded)", async () => {
    // 4MB base64 ≈ 3MB decoded. Create something clearly over 3MB decoded.
    const oversized = "A".repeat(4_200_000); // ~3.15MB decoded
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "Big",
      body: "<p>Too big</p>",
      attachments: [{ name: "big.bin", contentType: "application/octet-stream", contentBytes: oversized }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("3MB");
  });

  it("allows attachment exactly at 3MB boundary", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));
    // 3MB = 3145728 bytes. In base64: ceil(3145728 / 3) * 4 = 4194304 chars
    const exactly3mb = "A".repeat(4_194_304);

    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "Exact",
      body: "<p>Exactly 3MB</p>",
      attachments: [{ name: "exact.bin", contentType: "application/octet-stream", contentBytes: exactly3mb }],
    }));

    expect(result.data.sent).toBe(true);
  });

  it("returns error when attachment missing name", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "No name",
      body: "<p>Missing</p>",
      attachments: [{ contentType: "text/plain", contentBytes: "dGVzdA==" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("name");
  });

  it("returns error when attachment missing contentBytes", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "No bytes",
      body: "<p>Missing</p>",
      attachments: [{ name: "file.txt", contentType: "text/plain" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("contentBytes");
  });

  it("returns error when attachment missing contentType", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "No type",
      body: "<p>Missing</p>",
      attachments: [{ name: "file.txt", contentBytes: "dGVzdA==" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("contentType");
  });

  it("returns error for whitespace-only name", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "WS name",
      body: "<p>Bad</p>",
      attachments: [{ name: "   ", contentType: "text/plain", contentBytes: "dGVzdA==" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("name");
  });

  it("returns error for whitespace-only contentType", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "WS type",
      body: "<p>Bad</p>",
      attachments: [{ name: "file.txt", contentType: "  ", contentBytes: "dGVzdA==" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("contentType");
  });

  it("returns error for whitespace-only contentBytes", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "WS bytes",
      body: "<p>Bad</p>",
      attachments: [{ name: "file.txt", contentType: "text/plain", contentBytes: "   " }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("contentBytes");
  });

  it("returns error when too many attachments (>10)", async () => {
    const attachments = Array.from({ length: 11 }, (_, i) => ({
      name: `file${i}.txt`,
      contentType: "text/plain",
      contentBytes: "dGVzdA==",
    }));

    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "Too many",
      body: "<p>Overflow</p>",
      attachments,
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("10");
  });

  it("allows exactly 10 attachments", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const attachments = Array.from({ length: 10 }, (_, i) => ({
      name: `file${i}.txt`,
      contentType: "text/plain",
      contentBytes: "dGVzdA==",
    }));

    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "Max",
      body: "<p>10 files</p>",
      attachments,
    }));

    expect(result.data.sent).toBe(true);
    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toHaveLength(10);
  });

  it("sends with attachments + cc + bcc combined", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      to: "alice@example.com",
      subject: "Combined",
      body: "<p>All fields</p>",
      cc: ["bob@example.com"],
      bcc: ["charlie@example.com"],
      attachments: [{ name: "doc.pdf", contentType: "application/pdf", contentBytes: "AAAA" }],
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toHaveLength(1);
    expect(callBody.message.ccRecipients).toHaveLength(1);
    expect(callBody.message.bccRecipients).toHaveLength(1);
  });

  it("validates all attachments before sending", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "Partial",
      body: "<p>First good, second bad</p>",
      attachments: [
        { name: "good.txt", contentType: "text/plain", contentBytes: "dGVzdA==" },
        { contentType: "text/plain", contentBytes: "dGVzdA==" }, // missing name
      ],
    }));

    expect(result.error.category).toBe("user_input");
    expect(client._fetchMock).not.toHaveBeenCalled();
  });

  it("returns error for invalid base64 payload", async () => {
    const result = parseResult(await tool.execute("id", {
      to: "alice@example.com",
      subject: "Bad b64",
      body: "<p>Invalid</p>",
      attachments: [{ name: "file.txt", contentType: "text/plain", contentBytes: "not!valid!base64!!!" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("base64");
  });
});

// ── email_reply tests ──────────────────────────────────────────────────────

describe("email_reply", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createEmailReplyTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createEmailReplyTool({ graphClient: client });
  });

  it("replies to a message with body", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const result = await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Thanks for the info</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.data).toBeDefined();
    expect(parsed.error).toBeUndefined();
    expect(parsed.data.replied).toBe(true);

    expect(client._fetchMock).toHaveBeenCalledWith(
      "/me/messages/msg-1/reply",
      expect.objectContaining({
        method: "POST",
        body: expect.any(String),
      }),
    );

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.body).toEqual({
      contentType: "HTML",
      content: "<p>Thanks for the info</p>",
    });
  });

  it("uses /replyAll when replyAll is true", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Noted</p>",
      replyAll: true,
    });

    expect(client._fetchMock).toHaveBeenCalledWith(
      "/me/messages/msg-1/replyAll",
      expect.objectContaining({ method: "POST" }),
    );
  });

  it("uses /reply by default (replyAll false)", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Ok</p>",
    });

    expect(client._fetchMock).toHaveBeenCalledWith(
      "/me/messages/msg-1/reply",
      expect.objectContaining({ method: "POST" }),
    );
  });

  it("encodes messageId in path", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      messageId: "AAMkAD+special/chars=",
      body: "<p>Reply</p>",
    });

    const calledPath = client._fetchMock.mock.calls[0][0];
    expect(calledPath).toBe(`/me/messages/${encodeURIComponent("AAMkAD+special/chars=")}/reply`);
  });

  it("returns error for missing messageId", async () => {
    const result = await tool.execute("id", {
      body: "<p>Reply</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("messageId");
  });

  it("returns error for empty messageId", async () => {
    const result = await tool.execute("id", {
      messageId: "",
      body: "<p>Reply</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
  });

  it("returns error for missing body", async () => {
    const result = await tool.execute("id", {
      messageId: "msg-1",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("body");
  });

  it("handles GraphApiError", async () => {
    client._fetchMock.mockRejectedValue(
      new GraphApiError("Graph API /me/messages/msg-1/reply failed (404). The requested resource was not found.", "not_found", 404),
    );

    const result = await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Reply</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("not_found");
    expect(parsed.error.message).toContain("not found");
  });

  it("sanitizes error messages from non-GraphApiError exceptions", async () => {
    client._fetchMock.mockRejectedValue(
      new Error("POST https://graph.microsoft.com/v1.0/me/messages/msg-1/reply?token=secret123 failed"),
    );

    const result = await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Reply</p>",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.message).not.toContain("secret123");
    expect(parsed.error.message).not.toContain("graph.microsoft.com");
  });

  // ── Round 5: Reply with attachments ────────────────────────────────────────

  it("replies with single attachment", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const result = parseResult(await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>See attached</p>",
      attachments: [{
        name: "notes.pdf",
        contentType: "application/pdf",
        contentBytes: "SGVsbG8=",
      }],
    }));

    expect(result.data.replied).toBe(true);

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toHaveLength(1);
    expect(callBody.message.attachments[0]["@odata.type"]).toBe("#microsoft.graph.fileAttachment");
    expect(callBody.message.attachments[0].name).toBe("notes.pdf");
    expect(client._fetchMock.mock.calls[0][0]).toContain("/reply");
  });

  it("replyAll with attachments", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>All see this</p>",
      replyAll: true,
      attachments: [{ name: "file.txt", contentType: "text/plain", contentBytes: "dGVzdA==" }],
    });

    const calledPath = client._fetchMock.mock.calls[0][0] as string;
    expect(calledPath).toContain("/replyAll");
    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toHaveLength(1);
  });

  it("omits attachments when not provided (legacy format preserved)", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>No attachments</p>",
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toBeUndefined();
    expect(callBody.message.body).toEqual({ contentType: "HTML", content: "<p>No attachments</p>" });
  });

  it("omits attachments when empty array", async () => {
    client._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Empty array</p>",
      attachments: [],
    });

    const callBody = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(callBody.message.attachments).toBeUndefined();
  });

  it("returns error when reply attachment exceeds 3MB", async () => {
    const oversized = "A".repeat(4_200_000);
    const result = parseResult(await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Too big</p>",
      attachments: [{ name: "big.bin", contentType: "application/octet-stream", contentBytes: oversized }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("3MB");
    expect(client._fetchMock).not.toHaveBeenCalled();
  });

  it("returns error when reply attachment missing name", async () => {
    const result = parseResult(await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Missing name</p>",
      attachments: [{ contentType: "text/plain", contentBytes: "dGVzdA==" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("name");
  });

  it("returns error when too many reply attachments", async () => {
    const attachments = Array.from({ length: 11 }, (_, i) => ({
      name: `file${i}.txt`,
      contentType: "text/plain",
      contentBytes: "dGVzdA==",
    }));

    const result = parseResult(await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Too many</p>",
      attachments,
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("10");
  });

  it("returns error for invalid base64 in reply attachment", async () => {
    const result = parseResult(await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Bad base64</p>",
      attachments: [{ name: "file.txt", contentType: "text/plain", contentBytes: "not!valid!base64!!!" }],
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("base64");
  });
});

// ── Phase 3: Multi-account routing + policy enforcement ───────────────────

// ── email_search tests ─────────────────────────────────────────────────────

describe("email_search", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createEmailSearchTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createEmailSearchTool({ graphClient: client });
  });

  // ── Cycle 1: free-text query ──────────────────────────────────────────────

  it("calls /me/messages with $search and ConsistencyLevel header", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "quarterly report" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages",
      expect.objectContaining({
        $search: expect.stringContaining("quarterly report"),
        $count: "true",
      }),
      expect.objectContaining({ ConsistencyLevel: "eventual" }),
    );
  });

  it("returns formatted message summaries", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [
        {
          id: "msg-1",
          subject: "Test Subject",
          from: { emailAddress: { name: "Sender", address: "sender@test.com" } },
          toRecipients: [{ emailAddress: { address: "to@test.com" } }],
          receivedDateTime: "2026-04-02T10:00:00Z",
          isRead: false,
          hasAttachments: true,
          bodyPreview: "Preview text",
          importance: "high",
          flag: { flagStatus: "flagged" },
        },
      ],
      "@odata.count": 42,
      "@odata.nextLink": "https://graph.microsoft.com/next",
    });

    const result = await tool.execute("id", { query: "test" });
    const parsed = parseResult(result);

    expect(parsed.data.messages).toHaveLength(1);
    expect(parsed.data.messages[0]).toEqual({
      id: "msg-1",
      subject: "Test Subject",
      from: "Sender <sender@test.com>",
      to: ["to@test.com"],
      receivedDateTime: "2026-04-02T10:00:00Z",
      isRead: false,
      hasAttachments: true,
      bodyPreview: "Preview text",
      importance: "high",
      flagStatus: "flagged",
    });
    expect(parsed.data.totalCount).toBe(42);
    expect(parsed.data.hasMore).toBe(true);
  });

  // ── Cycle 2: structured KQL fields ─────────────────────────────────────────

  it("from field generates KQL from: clause", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { from: "alice@test.com" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages",
      expect.objectContaining({
        $search: expect.stringContaining("from:alice@test.com"),
      }),
      expect.objectContaining({ ConsistencyLevel: "eventual" }),
    );
  });

  it("subject with spaces is quoted in KQL", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { subject: "quarterly report" });

    // KQL: subject:"quarterly report" → OData wrapping escapes inner quotes
    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    expect(searchArg).toContain("subject:");
    expect(searchArg).toContain("quarterly report");
  });

  it("to field generates KQL to: clause", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { to: "bob@test.com" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages",
      expect.objectContaining({
        $search: expect.stringContaining("to:bob@test.com"),
      }),
      expect.any(Object),
    );
  });

  it("combines from, subject, and query into one KQL string", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { from: "alice@test.com", subject: "report", query: "financials" });

    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    expect(searchArg).toContain("from:alice@test.com");
    expect(searchArg).toContain("subject:report");
    expect(searchArg).toContain("financials");
  });

  // ── Cycle 3: hasAttachments ─────────────────────────────────────────────────

  it("hasAttachments true adds KQL clause", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "invoice", hasAttachments: true });

    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    expect(searchArg).toContain("hasAttachments:true");
    expect(searchArg).toContain("invoice");
  });

  it("hasAttachments false adds KQL clause", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { from: "test@test.com", hasAttachments: false });

    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    expect(searchArg).toContain("hasAttachments:false");
  });

  // ── Cycle 4: date range ───────────────────────────────────────────────────

  it("dateFrom with query uses KQL received>= instead of $filter", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "report", dateFrom: "2026-01-01" });

    const q = client._fetchJsonMock.mock.calls[0][1];
    // Dates go into KQL when combined with text search (Graph blocks $search + $filter)
    expect(q.$search).toContain("received>=2026-01-01");
    expect(q.$search).toContain("report");
    expect(q.$filter).toBeUndefined();
  });

  it("dateTo with query uses KQL received<= with inclusive date", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "report", dateTo: "2026-03-31" });

    const q = client._fetchJsonMock.mock.calls[0][1];
    expect(q.$search).toContain("received<=2026-03-31");
    expect(q.$filter).toBeUndefined();
  });

  it("dateFrom + dateTo + from combines into KQL with no $filter", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { from: "alice@test.com", dateFrom: "2026-01-01", dateTo: "2026-03-31" });

    const q = client._fetchJsonMock.mock.calls[0][1];
    expect(q.$search).toContain("from:alice@test.com");
    expect(q.$search).toContain("received>=2026-01-01");
    expect(q.$search).toContain("received<=2026-03-31");
    expect(q.$filter).toBeUndefined();
  });

  it("date-only search omits $search and ConsistencyLevel", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { dateFrom: "2026-01-01" });

    const query = client._fetchJsonMock.mock.calls[0][1];
    expect(query.$filter).toContain("receivedDateTime ge");
    expect(query.$search).toBeUndefined();
    expect(query.$orderby).toBe("receivedDateTime desc");

    // ConsistencyLevel should NOT be sent for filter-only
    const headers = client._fetchJsonMock.mock.calls[0][2];
    expect(headers?.ConsistencyLevel).toBeUndefined();
  });

  it("rejects invalid dateFrom format", async () => {
    const result = await tool.execute("id", { query: "test", dateFrom: "not-a-date" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("dateFrom");
  });

  it("rejects dateFrom after dateTo", async () => {
    const result = await tool.execute("id", { query: "test", dateFrom: "2026-06-01", dateTo: "2026-01-01" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("before");
  });

  // ── Cycle 5: pagination and ordering ──────────────────────────────────────

  it("defaults to $top 10", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "test" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ $top: "10" }),
      expect.any(Object),
    );
  });

  it("clamps top to 1-50", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "test", top: 999 });
    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ $top: "50" }),
      expect.any(Object),
    );

    client._fetchJsonMock.mockClear();
    await tool.execute("id", { query: "test", top: -5 });
    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ $top: "1" }),
      expect.any(Object),
    );
  });

  it("returns pagination info", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [],
      "@odata.count": 42,
      "@odata.nextLink": "https://graph.microsoft.com/next",
    });

    const result = await tool.execute("id", { query: "test" });
    const parsed = parseResult(result);

    expect(parsed.data.totalCount).toBe(42);
    expect(parsed.data.hasMore).toBe(true);
    expect(parsed.data.nextSkip).toBe(10);
  });

  it("filter-only search includes $orderby", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { dateFrom: "2026-01-01" });

    const query = client._fetchJsonMock.mock.calls[0][1];
    expect(query.$orderby).toBe("receivedDateTime desc");
  });

  // ── Cycle 6: error handling and multi-account ─────────────────────────────

  it("preserves GraphApiError category", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Insufficient permissions", "permission", 403),
    );

    const result = await tool.execute("id", { query: "test" });
    const parsed = parseResult(result);

    expect(parsed.error.category).toBe("permission");
    expect(parsed.error.message).toBe("Insufficient permissions");
  });

  it("sanitizes non-GraphApiError messages", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new Error("GET https://graph.microsoft.com/v1.0/me/messages?token=secret123"),
    );

    const result = await tool.execute("id", { query: "test" });
    const parsed = parseResult(result);

    expect(parsed.error.category).toBe("transient");
    expect(parsed.error.message).not.toContain("secret123");
    expect(parsed.error.message).not.toContain("graph.microsoft.com");
  });

  it("routes via resolveClient when provided", async () => {
    const altClient = createMockGraphClient();
    altClient._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    const resolveClient = vi.fn().mockReturnValue(altClient);
    const tool2 = createEmailSearchTool({ graphClient: client, resolveClient });

    await tool2.execute("id", { query: "test", account: "rod" });

    expect(resolveClient).toHaveBeenCalledWith("email_search", "rod");
    expect(altClient._fetchJsonMock).toHaveBeenCalled();
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("returns policy denial error from resolveClient", async () => {
    const resolveClient = vi.fn().mockImplementation(() => {
      throw new GraphApiError(
        "Tool email_search is not permitted for account 'openclaw'. Allowed accounts: ['rod'].",
        "user_input",
        403,
      );
    });
    const tool2 = createEmailSearchTool({ graphClient: client, resolveClient });

    const result = await tool2.execute("id", { query: "test", account: "openclaw" });
    const parsed = parseResult(result);

    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("not permitted");
  });

  // ── Cycle 7: KQL escaping edge cases ──────────────────────────────────────

  it("escapes internal quotes in subject", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { subject: 'the "big" deal' });

    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    // KQL quotes the value, OData escapes the whole thing
    expect(searchArg).toContain("subject:");
    expect(searchArg).toContain("big");
    expect(searchArg).toContain("deal");
  });

  it("double-escapes backslashes for OData layer", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "test\\path" });

    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    // Backslash in query → \\ in OData
    expect(searchArg).toContain("\\\\");
  });

  it("bare email in from is not unnecessarily quoted", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { from: "alice@test.com" });

    const searchArg = client._fetchJsonMock.mock.calls[0][1].$search;
    // Should be from:alice@test.com (no extra quotes around the email)
    expect(searchArg).toContain("from:alice@test.com");
  });

  // ── Validation ────────────────────────────────────────────────────────────

  it("rejects empty params with user_input error", async () => {
    const result = await tool.execute("id", {});
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("at least one search");
  });
});

describe("multi-account tool routing", () => {
  let rodClient: ReturnType<typeof createMockGraphClient>;
  let openclawClient: ReturnType<typeof createMockGraphClient>;

  function makeResolveClient(
    clients: Map<string, GraphClient>,
    policy: Map<string, string[]>, // toolName → permitted account IDs
  ) {
    return (toolName: string, accountId?: string): GraphClient => {
      const targetId = accountId ?? "rod"; // default
      if (!clients.has(targetId)) {
        throw new GraphApiError(
          `Unknown account '${targetId}'. Available accounts: ${[...clients.keys()].join(", ")}.`,
          "user_input",
          400,
        );
      }
      const permitted = policy.get(toolName);
      if (permitted && !permitted.includes(targetId)) {
        throw new GraphApiError(
          `Tool ${toolName} is not permitted for account '${targetId}'. Allowed accounts: [${permitted.map((a) => `'${a}'`).join(", ")}].`,
          "user_input",
          403,
        );
      }
      return clients.get(targetId)!;
    };
  }

  beforeEach(() => {
    rodClient = createMockGraphClient();
    openclawClient = createMockGraphClient();
  });

  it("email_list uses resolveClient when provided", async () => {
    rodClient._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["email_list", ["rod"]], ["email_send", ["openclaw"]]]);
    const tool = createEmailListTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    await tool.execute("id", {});

    expect(rodClient._fetchJsonMock).toHaveBeenCalled();
    expect(openclawClient._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("email_send uses resolveClient to route to openclaw", async () => {
    openclawClient._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["email_list", ["rod"]], ["email_send", ["openclaw"]]]);
    const tool = createEmailSendTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    await tool.execute("id", {
      to: "test@example.com",
      subject: "Test",
      body: "<p>Hi</p>",
      account: "openclaw",
    });

    expect(openclawClient._fetchMock).toHaveBeenCalled();
    expect(rodClient._fetchMock).not.toHaveBeenCalled();
  });

  it("policy denial: email_send with account=rod is blocked", async () => {
    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["email_send", ["openclaw"]]]);
    const tool = createEmailSendTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    const result = await tool.execute("id", {
      to: "test@example.com",
      subject: "Test",
      body: "<p>Hi</p>",
      account: "rod",
    });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("not permitted");
    expect(parsed.error.message).toContain("openclaw");
  });

  it("policy denial: unknown account returns actionable error", async () => {
    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["email_list", ["rod"]]]);
    const tool = createEmailListTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    const result = await tool.execute("id", { account: "nonexistent" });
    const parsed = parseResult(result);

    expect(parsed.error).toBeDefined();
    expect(parsed.error.category).toBe("user_input");
    expect(parsed.error.message).toContain("Unknown account");
    expect(parsed.error.message).toContain("rod");
    expect(parsed.error.message).toContain("openclaw");
  });

  it("email_read uses resolveClient with explicit account override", async () => {
    openclawClient._fetchJsonMock.mockResolvedValue({
      id: "msg-1",
      hasAttachments: false,
      body: { contentType: "text", content: "Body" },
    });

    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["email_read", ["rod", "openclaw"]]]); // permitted for both
    const tool = createEmailReadTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    await tool.execute("id", { messageId: "msg-1", account: "openclaw" });

    expect(openclawClient._fetchJsonMock).toHaveBeenCalled();
    expect(rodClient._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("email_reply uses resolveClient", async () => {
    openclawClient._fetchMock.mockResolvedValue(new Response("", { status: 202 }));

    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["email_reply", ["openclaw"]]]);
    const tool = createEmailReplyTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    await tool.execute("id", {
      messageId: "msg-1",
      body: "<p>Thanks</p>",
      account: "openclaw",
    });

    expect(openclawClient._fetchMock).toHaveBeenCalled();
    expect(rodClient._fetchMock).not.toHaveBeenCalled();
  });

  it("legacy fallback: tools use deps.graphClient when resolveClient is absent", async () => {
    rodClient._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    const tool = createEmailListTool({ graphClient: rodClient });

    await tool.execute("id", {});

    expect(rodClient._fetchJsonMock).toHaveBeenCalled();
  });

  it("account param is ignored gracefully when resolveClient is absent", async () => {
    rodClient._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    const tool = createEmailListTool({ graphClient: rodClient });

    await tool.execute("id", { account: "openclaw" });

    expect(rodClient._fetchJsonMock).toHaveBeenCalled();
  });

  it("calendar_list routes to rod via resolveClient", async () => {
    rodClient._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["calendar_list", ["rod"]]]);
    const tool = createCalendarListTool({
      graphClient: openclawClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    await tool.execute("id", { account: "rod" });

    expect(rodClient._fetchJsonMock).toHaveBeenCalled();
    expect(openclawClient._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("calendar_create policy denial for wrong account", async () => {
    const clients = new Map<string, GraphClient>([["rod", rodClient], ["openclaw", openclawClient]]);
    const policy = new Map([["calendar_create", ["rod"]]]);
    const tool = createCalendarCreateTool({
      graphClient: rodClient,
      resolveClient: makeResolveClient(clients, policy),
    });

    const result = parseResult(await tool.execute("id", {
      account: "openclaw",
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("not permitted");
  });
});

// ── formatEventSummary tests ───────────────────────────────────────────────

describe("formatEventSummary", () => {
  it("formats a full event with all fields", () => {
    const event: GraphEvent = {
      id: "evt-1",
      subject: "Team Standup",
      bodyPreview: "Daily sync meeting",
      start: { dateTime: "2026-04-04T09:00:00", timeZone: "America/Detroit" },
      end: { dateTime: "2026-04-04T09:30:00", timeZone: "America/Detroit" },
      location: { displayName: "Conference Room A" },
      organizer: { emailAddress: { name: "Rod", address: "rod@test.com" } },
      attendees: [
        {
          emailAddress: { name: "Alice", address: "alice@test.com" },
          type: "required",
          status: { response: "accepted", time: "2026-04-03T12:00:00Z" },
        },
      ],
      isAllDay: false,
      isCancelled: false,
      isOnlineMeeting: true,
      onlineMeetingUrl: "https://teams.microsoft.com/meet/123",
      importance: "high",
      showAs: "busy",
      webLink: "https://outlook.office.com/calendar/item/123",
    };

    const result = formatEventSummary(event);

    expect(result).toEqual({
      id: "evt-1",
      subject: "Team Standup",
      start: { dateTime: "2026-04-04T09:00:00", timeZone: "America/Detroit" },
      end: { dateTime: "2026-04-04T09:30:00", timeZone: "America/Detroit" },
      location: "Conference Room A",
      organizer: "Rod <rod@test.com>",
      attendees: [
        { email: "alice@test.com", name: "Alice", type: "required", response: "accepted" },
      ],
      isAllDay: false,
      isCancelled: false,
      isOnlineMeeting: true,
      onlineMeetingUrl: "https://teams.microsoft.com/meet/123",
      importance: "high",
      showAs: "busy",
      bodyPreview: "Daily sync meeting",
      webLink: "https://outlook.office.com/calendar/item/123",
    });
  });

  it("handles missing optional fields with defaults", () => {
    const event: GraphEvent = { id: "evt-minimal" };

    const result = formatEventSummary(event);

    expect(result.id).toBe("evt-minimal");
    expect(result.subject).toBe("(no subject)");
    expect(result.start).toBeNull();
    expect(result.end).toBeNull();
    expect(result.location).toBe("");
    expect(result.organizer).toBe("(unknown)");
    expect(result.attendees).toEqual([]);
    expect(result.isAllDay).toBe(false);
    expect(result.isCancelled).toBe(false);
    expect(result.isOnlineMeeting).toBe(false);
    expect(result.onlineMeetingUrl).toBeNull();
    expect(result.importance).toBe("normal");
    expect(result.showAs).toBe("busy");
    expect(result.bodyPreview).toBe("");
    expect(result.webLink).toBeNull();
  });

  it("formats multiple attendees with type and response status", () => {
    const event: GraphEvent = {
      id: "evt-2",
      attendees: [
        {
          emailAddress: { name: "Alice", address: "alice@test.com" },
          type: "required",
          status: { response: "accepted", time: "2026-04-03T12:00:00Z" },
        },
        {
          emailAddress: { address: "bob@test.com" },
          type: "optional",
          status: { response: "tentativelyAccepted", time: "2026-04-03T13:00:00Z" },
        },
        {
          emailAddress: { name: "Room 101", address: "room101@test.com" },
          type: "resource",
        },
      ],
    };

    const result = formatEventSummary(event);

    expect(result.attendees).toEqual([
      { email: "alice@test.com", name: "Alice", type: "required", response: "accepted" },
      { email: "bob@test.com", name: "", type: "optional", response: "tentativelyAccepted" },
      { email: "room101@test.com", name: "Room 101", type: "resource", response: "none" },
    ]);
  });

  it("formats organizer with name+email and address-only", () => {
    const withName: GraphEvent = {
      id: "evt-3",
      organizer: { emailAddress: { name: "Rod", address: "rod@test.com" } },
    };
    expect(formatEventSummary(withName).organizer).toBe("Rod <rod@test.com>");

    const addressOnly: GraphEvent = {
      id: "evt-4",
      organizer: { emailAddress: { address: "rod@test.com" } },
    };
    expect(formatEventSummary(addressOnly).organizer).toBe("rod@test.com");
  });

  it("carries through bodyPreview", () => {
    const event: GraphEvent = {
      id: "evt-5",
      bodyPreview: "Please review the attached proposal before our meeting.",
    };

    expect(formatEventSummary(event).bodyPreview).toBe(
      "Please review the attached proposal before our meeting.",
    );
  });

  it("formats cancelled event", () => {
    const event: GraphEvent = {
      id: "evt-6",
      subject: "Cancelled Meeting",
      isCancelled: true,
    };

    const result = formatEventSummary(event);
    expect(result.isCancelled).toBe(true);
    expect(result.subject).toBe("Cancelled Meeting");
  });

  it("passes through online meeting URL when isOnlineMeeting is true", () => {
    const event: GraphEvent = {
      id: "evt-7",
      isOnlineMeeting: true,
      onlineMeetingUrl: "https://teams.microsoft.com/meet/456",
    };

    const result = formatEventSummary(event);
    expect(result.isOnlineMeeting).toBe(true);
    expect(result.onlineMeetingUrl).toBe("https://teams.microsoft.com/meet/456");
  });
});

// ── validateDateRange tests ────────────────────────────────────────────────

describe("validateDateRange", () => {
  it("returns null when both dates are missing (no date filter)", () => {
    expect(validateDateRange(undefined, undefined)).toBeNull();
  });

  it("returns error when only startDateTime provided", () => {
    const result = validateDateRange("2026-04-04T09:00:00", undefined);
    expect(result).not.toBeNull();
    expect(result!.error).toContain("Both startDateTime and endDateTime");
    expect(result!.error).toContain("only startDateTime");
  });

  it("returns error when only endDateTime provided", () => {
    const result = validateDateRange(undefined, "2026-04-04T17:00:00");
    expect(result).not.toBeNull();
    expect(result!.error).toContain("Both startDateTime and endDateTime");
    expect(result!.error).toContain("only endDateTime");
  });

  it("returns error for invalid startDateTime", () => {
    const result = validateDateRange("not-a-date", "2026-04-04T17:00:00");
    expect(result).not.toBeNull();
    expect(result!.error).toContain("Invalid startDateTime");
    expect(result!.error).toContain("ISO 8601");
  });

  it("returns error for invalid endDateTime", () => {
    const result = validateDateRange("2026-04-04T09:00:00", "garbage");
    expect(result).not.toBeNull();
    expect(result!.error).toContain("Invalid endDateTime");
  });

  it("returns error when startDateTime >= endDateTime", () => {
    const result = validateDateRange("2026-04-04T17:00:00", "2026-04-04T09:00:00");
    expect(result).not.toBeNull();
    expect(result!.error).toContain("must be before");
  });

  it("returns error when startDateTime equals endDateTime", () => {
    const result = validateDateRange("2026-04-04T09:00:00", "2026-04-04T09:00:00");
    expect(result).not.toBeNull();
    expect(result!.error).toContain("must be before");
  });

  it("returns null for valid date range", () => {
    expect(validateDateRange("2026-04-04T09:00:00", "2026-04-04T17:00:00")).toBeNull();
  });
});

// ── checkConflicts tests ───────────────────────────────────────────────────

describe("checkConflicts", () => {
  let client: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    client = createMockGraphClient();
  });

  const START = "2026-04-05T09:00:00";
  const END = "2026-04-05T10:00:00";
  const TZ = "America/Detroit";

  function makeEvent(overrides: Partial<GraphEvent> & { id: string }): GraphEvent {
    return {
      subject: "Test Event",
      start: { dateTime: "2026-04-05T09:00:00", timeZone: TZ },
      end: { dateTime: "2026-04-05T09:30:00", timeZone: TZ },
      showAs: "busy",
      isCancelled: false,
      ...overrides,
    };
  }

  it("returns empty result when no overlapping events", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    const result = await checkConflicts(client, START, END, TZ);

    expect(result.hasConflicts).toBe(false);
    expect(result.conflicts).toHaveLength(0);
    expect(result.scanIncomplete).toBe(false);

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/calendarView",
      expect.objectContaining({ startDateTime: START, endDateTime: END }),
      expect.objectContaining({ Prefer: `outlook.timezone="${TZ}"` }),
    );
  });

  it("detects event with showAs=busy as conflict", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: "busy" })],
    });

    const result = await checkConflicts(client, START, END, TZ);

    expect(result.hasConflicts).toBe(true);
    expect(result.conflicts).toHaveLength(1);
    expect(result.conflicts[0].id).toBe("evt-1");
    expect(result.conflicts[0].showAs).toBe("busy");
    expect(result.conflicts[0].subject).toBe("Test Event");
    expect(result.conflicts[0].start).toBeDefined();
  });

  it("detects event with showAs=oof as conflict", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: "oof" })],
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.hasConflicts).toBe(true);
    expect(result.conflicts[0].showAs).toBe("oof");
  });

  it("detects event with showAs=workingElsewhere as conflict", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: "workingElsewhere" })],
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.hasConflicts).toBe(true);
    expect(result.conflicts[0].showAs).toBe("workingElsewhere");
  });

  it("ignores event with showAs=free", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: "free" })],
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.hasConflicts).toBe(false);
    expect(result.conflicts).toHaveLength(0);
  });

  it("ignores event with showAs=tentative", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: "tentative" })],
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.hasConflicts).toBe(false);
  });

  it("ignores cancelled event with showAs=busy", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: "busy", isCancelled: true })],
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.hasConflicts).toBe(false);
  });

  it("treats missing showAs as non-conflict", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [makeEvent({ id: "evt-1", showAs: undefined })],
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.hasConflicts).toBe(false);
  });

  it("excludeEventId filters out the specified event", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [
        makeEvent({ id: "evt-self", showAs: "busy" }),
        makeEvent({ id: "evt-other", showAs: "busy", subject: "Other Meeting" }),
      ],
    });

    const result = await checkConflicts(client, START, END, TZ, "evt-self");
    expect(result.hasConflicts).toBe(true);
    expect(result.conflicts).toHaveLength(1);
    expect(result.conflicts[0].id).toBe("evt-other");
  });

  it("sets scanIncomplete when @odata.nextLink present", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?$skip=50",
    });

    const result = await checkConflicts(client, START, END, TZ);
    expect(result.scanIncomplete).toBe(true);
  });
});

// ── calendar_list tests ────────────────────────────────────────────────────

describe("calendar_list", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createCalendarListTool>;

  const SAMPLE_EVENT: GraphEvent = {
    id: "evt-1",
    subject: "Team Standup",
    bodyPreview: "Daily sync",
    start: { dateTime: "2026-04-04T09:00:00", timeZone: "America/Detroit" },
    end: { dateTime: "2026-04-04T09:30:00", timeZone: "America/Detroit" },
    location: { displayName: "Room A" },
    organizer: { emailAddress: { name: "Rod", address: "rod@test.com" } },
    attendees: [],
    isAllDay: false,
    importance: "normal",
    showAs: "busy",
  };

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createCalendarListTool({ graphClient: client });
  });

  it("with defaults calls /me/events with correct query params", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", {});

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/events",
      expect.objectContaining({
        $top: "10",
        $orderby: "start/dateTime asc",
        $count: "true",
      }),
      expect.any(Object),
    );
  });

  it("default orderBy is start/dateTime asc", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", {});

    const query = client._fetchJsonMock.mock.calls[0][1];
    expect(query.$orderby).toBe("start/dateTime asc");
  });

  it("formats event summaries via formatEventSummary", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [SAMPLE_EVENT],
      "@odata.count": 1,
    });

    const result = parseResult(await tool.execute("id", {}));

    expect(result.data.events).toHaveLength(1);
    expect(result.data.events[0].id).toBe("evt-1");
    expect(result.data.events[0].subject).toBe("Team Standup");
    expect(result.data.events[0].organizer).toBe("Rod <rod@test.com>");
  });

  it("returns pagination info", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [SAMPLE_EVENT],
      "@odata.count": 42,
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/events?$skip=10",
    });

    const result = parseResult(await tool.execute("id", {}));

    expect(result.data.totalCount).toBe(42);
    expect(result.data.hasMore).toBe(true);
    expect(result.data.nextSkip).toBe(10);
  });

  it("uses /me/calendarView when both startDateTime and endDateTime provided", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", {
      startDateTime: "2026-04-04T00:00:00",
      endDateTime: "2026-04-04T23:59:59",
    });

    const [path, query] = client._fetchJsonMock.mock.calls[0];
    expect(path).toBe("/me/calendarView");
    expect(query.startDateTime).toBe("2026-04-04T00:00:00");
    expect(query.endDateTime).toBe("2026-04-04T23:59:59");
  });

  it("returns error when only startDateTime provided", async () => {
    const result = parseResult(await tool.execute("id", {
      startDateTime: "2026-04-04T09:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("Both startDateTime and endDateTime");
  });

  it("returns error when only endDateTime provided", async () => {
    const result = parseResult(await tool.execute("id", {
      endDateTime: "2026-04-04T17:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("Both startDateTime and endDateTime");
  });

  it("returns error when startDateTime >= endDateTime", async () => {
    const result = parseResult(await tool.execute("id", {
      startDateTime: "2026-04-04T17:00:00",
      endDateTime: "2026-04-04T09:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("must be before");
  });

  it("returns error for invalid date format", async () => {
    const result = parseResult(await tool.execute("id", {
      startDateTime: "not-a-date",
      endDateTime: "2026-04-04T17:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("Invalid startDateTime");
  });

  it("uses default timezone America/Detroit in Prefer header", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", {});

    const extraHeaders = client._fetchJsonMock.mock.calls[0][2];
    expect(extraHeaders.Prefer).toBe('outlook.timezone="America/Detroit"');
  });

  it("uses custom timezone when provided", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { timeZone: "America/Los_Angeles" });

    const extraHeaders = client._fetchJsonMock.mock.calls[0][2];
    expect(extraHeaders.Prefer).toBe('outlook.timezone="America/Los_Angeles"');
  });

  it("clamps top to 1-50", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { top: 999 });
    expect(client._fetchJsonMock.mock.calls[0][1].$top).toBe("50");

    client._fetchJsonMock.mockClear();

    await tool.execute("id", { top: -5 });
    expect(client._fetchJsonMock.mock.calls[0][1].$top).toBe("1");
  });

  it("passes filter and orderBy params", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", {
      filter: "importance eq 'high'",
      orderBy: "start/dateTime desc",
    });

    const query = client._fetchJsonMock.mock.calls[0][1];
    expect(query.$filter).toBe("importance eq 'high'");
    expect(query.$orderby).toBe("start/dateTime desc");
  });

  it("preserves GraphApiError category", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Forbidden", "permission", 403),
    );

    const result = parseResult(await tool.execute("id", {}));

    expect(result.error.category).toBe("permission");
    expect(result.error.message).toBe("Forbidden");
  });

  it("sanitizes non-GraphApiError messages", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new Error("token=abc123 leaked at https://internal.url"),
    );

    const result = parseResult(await tool.execute("id", {}));

    expect(result.error.category).toBe("transient");
    expect(result.error.message).not.toContain("abc123");
    expect(result.error.message).not.toContain("internal.url");
  });

  it("routes via resolveClient when provided", async () => {
    const altClient = createMockGraphClient();
    altClient._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    const resolveClient = vi.fn().mockReturnValue(altClient);
    const routedTool = createCalendarListTool({ graphClient: client, resolveClient });

    await routedTool.execute("id", { account: "rod" });

    expect(resolveClient).toHaveBeenCalledWith("calendar_list", "rod");
    expect(altClient._fetchJsonMock).toHaveBeenCalled();
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("returns policy denial error from resolveClient", async () => {
    const resolveClient = vi.fn().mockImplementation(() => {
      throw new GraphApiError(
        "Tool calendar_list is not permitted for account 'openclaw'.",
        "user_input",
        403,
      );
    });
    const routedTool = createCalendarListTool({ graphClient: client, resolveClient });

    const result = parseResult(await routedTool.execute("id", { account: "openclaw" }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("not permitted");
  });
});

// ── calendar_update tests ──────────────────────────────────────────────────

describe("calendar_update", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createCalendarUpdateTool>;

  const UPDATED_EVENT: GraphEvent = {
    id: "evt-update-1",
    subject: "Updated Standup",
    start: { dateTime: "2026-04-04T10:00:00", timeZone: "America/Detroit" },
    end: { dateTime: "2026-04-04T10:30:00", timeZone: "America/Detroit" },
  };

  beforeEach(() => {
    client = createMockGraphClient();
    // Default: no conflicts (conflict check returns empty when both dates provided)
    client._fetchJsonMock.mockResolvedValue({ value: [] });
    client._fetchMock.mockResolvedValue(
      new Response(JSON.stringify(UPDATED_EVENT), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      }),
    );
    tool = createCalendarUpdateTool({ graphClient: client });
  });

  it("updates subject only — PATCH body contains only subject", async () => {
    await tool.execute("id", { eventId: "evt-1", subject: "New Title" });

    const [path, init] = client._fetchMock.mock.calls[0];
    expect(path).toBe("/me/events/evt-1");
    expect(init.method).toBe("PATCH");
    const body = JSON.parse(init.body);
    expect(body.subject).toBe("New Title");
    expect(body.start).toBeUndefined();
    expect(body.end).toBeUndefined();
  });

  it("returns updated event formatted (200 with body)", async () => {
    const result = parseResult(await tool.execute("id", {
      eventId: "evt-update-1",
      subject: "Updated Standup",
    }));

    expect(result.data.updated).toBe(true);
    expect(result.data.event.id).toBe("evt-update-1");
    expect(result.data.event.subject).toBe("Updated Standup");
  });

  it("handles PATCH returning empty response (204/no body)", async () => {
    client._fetchMock.mockResolvedValue(
      new Response(null, { status: 204 }),
    );

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      subject: "Updated",
    }));

    expect(result.data.updated).toBe(true);
    expect(result.data.eventId).toBe("evt-1");
  });

  it("updates multiple fields at once", async () => {
    await tool.execute("id", {
      eventId: "evt-1",
      subject: "New Title",
      location: "Room B",
      importance: "high",
    });

    const body = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(body.subject).toBe("New Title");
    expect(body.location).toEqual({ displayName: "Room B" });
    expect(body.importance).toBe("high");
  });

  it("uses custom timezone for start/end updates", async () => {
    await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-04T14:00:00",
      endDateTime: "2026-04-04T15:00:00",
      timeZone: "Europe/London",
    });

    const body = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(body.start).toEqual({ dateTime: "2026-04-04T14:00:00", timeZone: "Europe/London" });
    expect(body.end).toEqual({ dateTime: "2026-04-04T15:00:00", timeZone: "Europe/London" });
  });

  it("returns error for missing eventId", async () => {
    const result = parseResult(await tool.execute("id", { subject: "No ID" }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("eventId");
  });

  it("returns error for no mutable fields provided (empty update)", async () => {
    const result = parseResult(await tool.execute("id", { eventId: "evt-1" }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("At least one field");
  });

  it("handles not_found error (404)", async () => {
    client._fetchMock.mockRejectedValue(
      new GraphApiError("Resource not found", "not_found", 404),
    );

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-missing",
      subject: "Update",
    }));

    expect(result.error.category).toBe("not_found");
  });

  it("sanitizes non-GraphApiError messages", async () => {
    client._fetchMock.mockRejectedValue(
      new Error("token=secret123 at https://internal"),
    );

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      subject: "Update",
    }));

    expect(result.error.category).toBe("transient");
    expect(result.error.message).not.toContain("secret123");
  });

  // ── Round 4: Update conflict detection ──────────────────────────────────

  it("blocks update when both dates changed and conflicts found", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-conflict",
        subject: "Existing Meeting",
        showAs: "busy",
        start: { dateTime: "2026-04-05T14:00:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T15:00:00", timeZone: "America/Detroit" },
      }],
    });

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("conflict");
    // fetchMock should not have been called for PATCH — only the beforeEach default stays
    // The conflict check should prevent the PATCH call
    expect(client._fetchMock).not.toHaveBeenCalled();
  });

  it("forceUpdate=true updates despite conflicts, includes warnings", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-conflict",
        subject: "Busy Slot",
        showAs: "busy",
        start: { dateTime: "2026-04-05T14:00:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T14:30:00", timeZone: "America/Detroit" },
      }],
    });

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
      forceUpdate: true,
    }));

    expect(result.data.updated).toBe(true);
    expect(result.data.warnings).toBeDefined();
    expect(result.data.warnings.conflictsDetected).toBe(1);
    expect(client._fetchMock).toHaveBeenCalled();
  });

  it("conflict error message says forceUpdate (not forceCreate)", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-c",
        subject: "Block",
        showAs: "busy",
        start: { dateTime: "2026-04-05T14:00:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T15:00:00", timeZone: "America/Detroit" },
      }],
    });

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
    }));

    expect(result.error.message).toContain("forceUpdate");
    expect(result.error.message).not.toContain("forceCreate");
  });

  it("no conflict check when only subject changed", async () => {
    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      subject: "Renamed",
    }));

    expect(result.data.updated).toBe(true);
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("no conflict check when only startDateTime provided", async () => {
    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
    }));

    expect(result.data.updated).toBe(true);
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("no conflict check when only endDateTime provided", async () => {
    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      endDateTime: "2026-04-05T15:00:00",
    }));

    expect(result.data.updated).toBe(true);
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("excludeEventId prevents self-conflict", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-1", // same ID as the event being updated
        subject: "Self",
        showAs: "busy",
        start: { dateTime: "2026-04-05T14:00:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T15:00:00", timeZone: "America/Detroit" },
      }],
    });

    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
    }));

    expect(result.data.updated).toBe(true);
  });

  it("checkConflicts=false skips check when both dates provided", async () => {
    const result = parseResult(await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
      checkConflicts: false,
    }));

    expect(result.data.updated).toBe(true);
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("conflict check uses correct timezone for update", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", {
      eventId: "evt-1",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
      timeZone: "Asia/Tokyo",
    });

    const preferHeader = client._fetchJsonMock.mock.calls[0][2];
    expect(preferHeader.Prefer).toContain("Asia/Tokyo");
  });
});

// ── calendar_delete tests ──────────────────────────────────────────────────

describe("calendar_delete", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createCalendarDeleteTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    client._fetchMock.mockResolvedValue(new Response(null, { status: 204 }));
    tool = createCalendarDeleteTool({ graphClient: client });
  });

  it("deletes event successfully", async () => {
    const result = parseResult(await tool.execute("id", { eventId: "evt-1" }));

    const [path, init] = client._fetchMock.mock.calls[0];
    expect(path).toBe("/me/events/evt-1");
    expect(init.method).toBe("DELETE");
    expect(result.data.deleted).toBe(true);
    expect(result.data.eventId).toBe("evt-1");
  });

  it("encodes eventId in path", async () => {
    await tool.execute("id", { eventId: "AAMkAD+special/chars=" });

    const path = client._fetchMock.mock.calls[0][0];
    expect(path).toBe(`/me/events/${encodeURIComponent("AAMkAD+special/chars=")}`);
  });

  it("returns error for missing eventId", async () => {
    const result = parseResult(await tool.execute("id", {}));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("eventId");
  });

  it("handles not_found error (404)", async () => {
    client._fetchMock.mockRejectedValue(
      new GraphApiError("Resource not found", "not_found", 404),
    );

    const result = parseResult(await tool.execute("id", { eventId: "evt-gone" }));

    expect(result.error.category).toBe("not_found");
  });

  it("sanitizes non-GraphApiError messages", async () => {
    client._fetchMock.mockRejectedValue(
      new Error("token=leaked at https://internal"),
    );

    const result = parseResult(await tool.execute("id", { eventId: "evt-1" }));

    expect(result.error.category).toBe("transient");
    expect(result.error.message).not.toContain("leaked");
  });
});

// ── calendar_create tests ──────────────────────────────────────────────────

describe("calendar_create", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createCalendarCreateTool>;

  const CREATED_EVENT: GraphEvent = {
    id: "evt-new-1",
    subject: "New Meeting",
    start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" },
    end: { dateTime: "2026-04-05T10:00:00", timeZone: "America/Detroit" },
    organizer: { emailAddress: { name: "Rod", address: "rod@test.com" } },
  };

  beforeEach(() => {
    client = createMockGraphClient();
    // Default: no conflicts (conflict check returns empty)
    client._fetchJsonMock.mockResolvedValue({ value: [] });
    client._fetchMock.mockResolvedValue(
      new Response(JSON.stringify(CREATED_EVENT), {
        status: 201,
        headers: { "Content-Type": "application/json" },
      }),
    );
    tool = createCalendarCreateTool({ graphClient: client });
  });

  it("creates event with required fields only", async () => {
    const result = parseResult(await tool.execute("id", {
      subject: "New Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    const [path, init] = client._fetchMock.mock.calls[0];
    expect(path).toBe("/me/events");
    expect(init.method).toBe("POST");

    const body = JSON.parse(init.body);
    expect(body.subject).toBe("New Meeting");
    expect(body.start.dateTime).toBe("2026-04-05T09:00:00");
    expect(body.end.dateTime).toBe("2026-04-05T10:00:00");

    expect(result.data.created).toBe(true);
  });

  it("returns created event formatted with formatEventSummary", async () => {
    const result = parseResult(await tool.execute("id", {
      subject: "New Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.data.event.id).toBe("evt-new-1");
    expect(result.data.event.subject).toBe("New Meeting");
    expect(result.data.event.organizer).toBe("Rod <rod@test.com>");
  });

  it("creates event with all optional fields", async () => {
    await tool.execute("id", {
      subject: "Full Event",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
      body: "<p>Agenda here</p>",
      location: "Room C",
      attendees: ["alice@test.com", "bob@test.com"],
      isAllDay: false,
      importance: "high",
      showAs: "tentative",
      isOnlineMeeting: true,
    });

    const body = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(body.body).toEqual({ contentType: "HTML", content: "<p>Agenda here</p>" });
    expect(body.location).toEqual({ displayName: "Room C" });
    expect(body.attendees).toEqual([
      { emailAddress: { address: "alice@test.com" }, type: "required" },
      { emailAddress: { address: "bob@test.com" }, type: "required" },
    ]);
    expect(body.importance).toBe("high");
    expect(body.showAs).toBe("tentative");
    expect(body.isOnlineMeeting).toBe(true);
  });

  it("uses custom timezone for start/end", async () => {
    await tool.execute("id", {
      subject: "UTC Meeting",
      startDateTime: "2026-04-05T14:00:00",
      endDateTime: "2026-04-05T15:00:00",
      timeZone: "UTC",
    });

    const body = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(body.start.timeZone).toBe("UTC");
    expect(body.end.timeZone).toBe("UTC");
  });

  it("uses default timezone America/Detroit when timeZone omitted", async () => {
    await tool.execute("id", {
      subject: "Local Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    });

    const body = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(body.start.timeZone).toBe("America/Detroit");
    expect(body.end.timeZone).toBe("America/Detroit");
  });

  it("returns error for missing subject", async () => {
    const result = parseResult(await tool.execute("id", {
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("subject");
  });

  it("returns error for missing startDateTime", async () => {
    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("startDateTime");
  });

  it("returns error for missing endDateTime", async () => {
    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("endDateTime");
  });

  it("returns error when endDateTime <= startDateTime", async () => {
    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T10:00:00",
      endDateTime: "2026-04-05T09:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("must be before");
  });

  it("treats empty attendees array as omitted", async () => {
    await tool.execute("id", {
      subject: "Solo Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
      attendees: [],
    });

    const body = JSON.parse(client._fetchMock.mock.calls[0][1].body);
    expect(body.attendees).toBeUndefined();
  });

  it("preserves GraphApiError category", async () => {
    client._fetchMock.mockRejectedValue(
      new GraphApiError("Forbidden", "permission", 403),
    );

    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("permission");
  });

  it("sanitizes non-GraphApiError messages", async () => {
    client._fetchMock.mockRejectedValue(
      new Error("token=secret at https://internal"),
    );

    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("transient");
    expect(result.error.message).not.toContain("secret");
  });

  // ── Round 2: Conflict detection happy path ────────────────────────────────

  it("default behavior blocks creation when conflicts found", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-conflict",
        subject: "Existing Meeting",
        start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T10:00:00", timeZone: "America/Detroit" },
        showAs: "busy",
        isCancelled: false,
      }],
    });

    const result = parseResult(await tool.execute("id", {
      subject: "New Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("conflict");
    expect(result.error.message).toContain("Existing Meeting");
    expect(client._fetchMock).not.toHaveBeenCalled();
  });

  it("forceCreate=true creates despite conflicts, includes warnings", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-conflict",
        subject: "Busy Event",
        start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T09:30:00", timeZone: "America/Detroit" },
        showAs: "busy",
      }],
    });

    const result = parseResult(await tool.execute("id", {
      subject: "New Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
      forceCreate: true,
    }));

    expect(result.data.created).toBe(true);
    expect(result.data.warnings).toBeDefined();
    expect(result.data.warnings.conflictsDetected).toBe(1);
    expect(result.data.warnings.conflicts[0].subject).toBe("Busy Event");
    expect(client._fetchMock).toHaveBeenCalled();
  });

  it("checkConflicts=false skips check entirely", async () => {
    const result = parseResult(await tool.execute("id", {
      subject: "Quick Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
      checkConflicts: false,
    }));

    expect(result.data.created).toBe(true);
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("creates normally when no conflicts found", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    const result = parseResult(await tool.execute("id", {
      subject: "Free Slot Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.data.created).toBe(true);
    expect(result.data.warnings).toBeUndefined();
    expect(client._fetchJsonMock).toHaveBeenCalled();
    expect(client._fetchMock).toHaveBeenCalled();
  });

  it("conflict error message includes event time and forceCreate hint", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [{
        id: "evt-c",
        subject: "Block",
        start: { dateTime: "2026-04-05T09:15:00", timeZone: "America/Detroit" },
        end: { dateTime: "2026-04-05T09:45:00", timeZone: "America/Detroit" },
        showAs: "busy",
      }],
    });

    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.message).toContain("09:15:00");
    expect(result.error.message).toContain("forceCreate");
  });

  it("Graph API error during conflict check propagates", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Throttled", "throttle", 429),
    );

    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.category).toBe("throttle");
    expect(client._fetchMock).not.toHaveBeenCalled();
  });

  // ── Round 3: Conflict edge cases ──────────────────────────────────────────

  it("multiple conflicts listed in error message", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [
        { id: "e1", subject: "Meeting A", showAs: "busy", start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" }, end: { dateTime: "2026-04-05T09:30:00", timeZone: "America/Detroit" } },
        { id: "e2", subject: "Meeting B", showAs: "busy", start: { dateTime: "2026-04-05T09:30:00", timeZone: "America/Detroit" }, end: { dateTime: "2026-04-05T10:00:00", timeZone: "America/Detroit" } },
        { id: "e3", subject: "Meeting C", showAs: "oof", start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" }, end: { dateTime: "2026-04-05T10:00:00", timeZone: "America/Detroit" } },
      ],
    });

    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    }));

    expect(result.error.message).toContain("3 conflicting");
  });

  it("mixed showAs — only conflicting statuses in warnings", async () => {
    client._fetchJsonMock.mockResolvedValue({
      value: [
        { id: "e1", subject: "Busy", showAs: "busy", start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" }, end: { dateTime: "2026-04-05T09:30:00", timeZone: "America/Detroit" } },
        { id: "e2", subject: "Free", showAs: "free", start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" }, end: { dateTime: "2026-04-05T09:30:00", timeZone: "America/Detroit" } },
        { id: "e3", subject: "OOF", showAs: "oof", start: { dateTime: "2026-04-05T09:00:00", timeZone: "America/Detroit" }, end: { dateTime: "2026-04-05T09:30:00", timeZone: "America/Detroit" } },
      ],
    });

    const result = parseResult(await tool.execute("id", {
      subject: "Test",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
      forceCreate: true,
    }));

    expect(result.data.warnings.conflictsDetected).toBe(2);
    expect(result.data.warnings.conflicts.map((c: any) => c.subject)).toEqual(["Busy", "OOF"]);
  });

  it("conflict check uses specified timezone in Prefer header", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", {
      subject: "London Meeting",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
      timeZone: "Europe/London",
    });

    const preferHeader = client._fetchJsonMock.mock.calls[0][2];
    expect(preferHeader.Prefer).toContain("Europe/London");
  });

  it("conflict check uses default timezone when omitted", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    await tool.execute("id", {
      subject: "Default TZ",
      startDateTime: "2026-04-05T09:00:00",
      endDateTime: "2026-04-05T10:00:00",
    });

    const preferHeader = client._fetchJsonMock.mock.calls[0][2];
    expect(preferHeader.Prefer).toContain("America/Detroit");
  });

  it("boundary adjacency: event ending at proposed start is not a conflict", async () => {
    // Graph calendarView returns events that overlap the window.
    // An event ending exactly when the new one starts should not appear in results.
    client._fetchJsonMock.mockResolvedValue({ value: [] });

    const result = parseResult(await tool.execute("id", {
      subject: "After Adjacent",
      startDateTime: "2026-04-05T10:00:00",
      endDateTime: "2026-04-05T11:00:00",
    }));

    expect(result.data.created).toBe(true);
    expect(result.data.warnings).toBeUndefined();
  });
});

// ── email_attachment_read tests ────────────────────────────────────────────

describe("email_attachment_read", () => {
  let client: ReturnType<typeof createMockGraphClient>;
  let tool: ReturnType<typeof createEmailAttachmentReadTool>;

  beforeEach(() => {
    client = createMockGraphClient();
    tool = createEmailAttachmentReadTool({ graphClient: client });
  });

  // ── Round 1: Happy path ─────────────────────────────────────────────────

  it("downloads file attachment by messageId and attachmentId", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-1",
      name: "report.pdf",
      contentType: "application/pdf",
      size: 12345,
      contentBytes: "SGVsbG8=",
    });

    const result = await tool.execute("id", {
      messageId: "msg-1",
      attachmentId: "att-1",
    });
    const parsed = parseResult(result);

    expect(parsed.data.id).toBe("att-1");
    expect(parsed.data.name).toBe("report.pdf");
    expect(parsed.data.contentType).toBe("application/pdf");
    expect(parsed.data.size).toBe(12345);
    expect(parsed.data.contentBytes).toBe("SGVsbG8=");
    expect(parsed.data.attachmentType).toBe("file");
    expect(parsed.error).toBeUndefined();

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages/msg-1/attachments/att-1",
    );
  });

  it("encodes messageId and attachmentId in URL path", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att/+id=",
      name: "file.txt",
      contentType: "text/plain",
      size: 100,
      contentBytes: "dGVzdA==",
    });

    await tool.execute("id", {
      messageId: "AAMkAD+special/chars=",
      attachmentId: "att/+id=",
    });

    const calledPath = client._fetchJsonMock.mock.calls[0][0] as string;
    expect(calledPath).toContain(encodeURIComponent("AAMkAD+special/chars="));
    expect(calledPath).toContain(encodeURIComponent("att/+id="));
  });

  it("returns content and details shape", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-1",
      name: "file.txt",
      contentType: "text/plain",
      size: 10,
      contentBytes: "dGVzdA==",
    });

    const result = await tool.execute("id", {
      messageId: "msg-1",
      attachmentId: "att-1",
    });

    expect(result.content[0].type).toBe("text");
    expect(result.details).toBeDefined();
    expect(result.details!.schemaVersion).toBe(1);
  });

  // ── Round 2: Edge cases ─────────────────────────────────────────────────

  it("returns user_input error for empty messageId", async () => {
    const result = parseResult(
      await tool.execute("id", { messageId: "", attachmentId: "att-1" }),
    );
    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("messageId");
  });

  it("returns user_input error for missing attachmentId", async () => {
    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1" }),
    );
    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("attachmentId");
  });

  it("returns user_input error for whitespace-only attachmentId", async () => {
    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "  " }),
    );
    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("attachmentId");
  });

  it("returns not_found error when attachment 404s", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Attachment not found.", "not_found", 404),
    );

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-gone" }),
    );
    expect(result.error.category).toBe("not_found");
  });

  it("returns not_found error when message 404s", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new GraphApiError("Message not found.", "not_found", 404),
    );

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-gone", attachmentId: "att-1" }),
    );
    expect(result.error.category).toBe("not_found");
  });

  it("returns metadata-only with tooLarge flag for oversized attachments", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-big",
      name: "huge.zip",
      contentType: "application/zip",
      size: 15_000_000,
      contentBytes: "AAAA",
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-big" }),
    );
    expect(result.data.contentBytes).toBeNull();
    expect(result.data.tooLarge).toBe(true);
    expect(result.data.warning).toContain("10MB");
  });

  it("allows attachment exactly at 10MB size limit", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-10mb",
      name: "exact.bin",
      contentType: "application/octet-stream",
      size: 10 * 1024 * 1024,
      contentBytes: "AAAA",
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-10mb" }),
    );
    expect(result.data.contentBytes).toBe("AAAA");
    expect(result.data.tooLarge).toBeUndefined();
  });

  it("handles itemAttachment type (info-only)", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.itemAttachment",
      id: "att-item",
      name: "Forwarded.eml",
      contentType: "message/rfc822",
      size: 5000,
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-item" }),
    );
    expect(result.data.attachmentType).toBe("item");
    expect(result.data.contentBytes).toBeNull();
    expect(result.data.note).toBeDefined();
  });

  it("handles referenceAttachment type (info-only)", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.referenceAttachment",
      id: "att-ref",
      name: "OneDriveDoc.docx",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      size: 0,
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-ref" }),
    );
    expect(result.data.attachmentType).toBe("reference");
    expect(result.data.contentBytes).toBeNull();
    expect(result.data.note).toBeDefined();
  });

  it("handles unknown @odata.type with missing contentBytes", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.someFutureType",
      id: "att-future",
      name: "mystery.dat",
      contentType: "application/octet-stream",
      size: 100,
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-future" }),
    );
    expect(result.data.attachmentType).toBe("unknown");
    expect(result.data.contentBytes).toBeNull();
    expect(result.data.note).toBeDefined();
  });

  it("handles missing contentBytes gracefully on fileAttachment", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-empty",
      name: "empty.txt",
      contentType: "text/plain",
      size: 0,
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-empty" }),
    );
    expect(result.data.contentBytes).toBeNull();
    expect(result.data.attachmentType).toBe("file");
    expect(result.error).toBeUndefined();
  });

  it("size guard takes precedence over contentBytes for large fileAttachment", async () => {
    client._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-huge",
      name: "big.bin",
      contentType: "application/octet-stream",
      size: 15_000_000,
      contentBytes: "AAAA".repeat(1000),
    });

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-huge" }),
    );
    expect(result.data.contentBytes).toBeNull();
    expect(result.data.tooLarge).toBe(true);
  });

  it("sanitizes non-GraphApiError messages", async () => {
    client._fetchJsonMock.mockRejectedValue(
      new Error("GET https://graph.microsoft.com/v1.0/me/messages/abc?token=secret"),
    );

    const result = parseResult(
      await tool.execute("id", { messageId: "msg-1", attachmentId: "att-1" }),
    );
    expect(result.error.category).toBe("transient");
    expect(result.error.message).not.toContain("secret");
    expect(result.error.message).not.toContain("graph.microsoft.com");
  });

  it("routes via resolveClient when account provided", async () => {
    const altClient = createMockGraphClient();
    altClient._fetchJsonMock.mockResolvedValue({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "att-1",
      name: "file.txt",
      contentType: "text/plain",
      size: 10,
      contentBytes: "dGVzdA==",
    });

    const resolveClient = vi.fn().mockReturnValue(altClient);
    const tool2 = createEmailAttachmentReadTool({ graphClient: client, resolveClient });

    await tool2.execute("id", {
      messageId: "msg-1",
      attachmentId: "att-1",
      account: "rod",
    });

    expect(resolveClient).toHaveBeenCalledWith("email_attachment_read", "rod");
    expect(altClient._fetchJsonMock).toHaveBeenCalled();
    expect(client._fetchJsonMock).not.toHaveBeenCalled();
  });

  it("returns policy denial from resolveClient", async () => {
    const resolveClient = vi.fn().mockImplementation(() => {
      throw new GraphApiError(
        "Tool email_attachment_read is not permitted for account 'openclaw'. Allowed accounts: ['rod'].",
        "user_input",
        403,
      );
    });
    const tool2 = createEmailAttachmentReadTool({ graphClient: client, resolveClient });

    const result = parseResult(
      await tool2.execute("id", { messageId: "msg-1", attachmentId: "att-1", account: "openclaw" }),
    );
    expect(result.error.category).toBe("user_input");
    expect(result.error.message).toContain("not permitted");
  });
});
