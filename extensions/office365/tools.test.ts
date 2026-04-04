import { describe, expect, it, vi, beforeEach } from "vitest";
import type { GraphClient } from "./src/graph-client.js";
import { createEmailListTool } from "./src/tools/email-list.js";
import { createEmailReadTool } from "./src/tools/email-read.js";
import { createEmailSendTool } from "./src/tools/email-send.js";
import { createEmailReplyTool } from "./src/tools/email-reply.js";
import { createEmailSearchTool } from "./src/tools/email-search.js";
import { GraphApiError } from "./src/types.js";

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

  it("dateFrom creates $filter and combines with $search", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "report", dateFrom: "2026-01-01" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages",
      expect.objectContaining({
        $filter: "receivedDateTime ge 2026-01-01T00:00:00Z",
        $search: expect.stringContaining("report"),
      }),
      expect.objectContaining({ ConsistencyLevel: "eventual" }),
    );
  });

  it("dateTo uses exclusive next-day boundary for date-only input", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { query: "report", dateTo: "2026-03-31" });

    expect(client._fetchJsonMock).toHaveBeenCalledWith(
      "/me/messages",
      expect.objectContaining({
        $filter: "receivedDateTime lt 2026-04-01T00:00:00Z",
      }),
      expect.any(Object),
    );
  });

  it("dateFrom + dateTo combines into one $filter", async () => {
    client._fetchJsonMock.mockResolvedValue({ value: [], "@odata.count": 0 });

    await tool.execute("id", { from: "alice@test.com", dateFrom: "2026-01-01", dateTo: "2026-03-31" });

    const query = client._fetchJsonMock.mock.calls[0][1];
    expect(query.$filter).toContain("receivedDateTime ge 2026-01-01T00:00:00Z");
    expect(query.$filter).toContain("receivedDateTime lt 2026-04-01T00:00:00Z");
    expect(query.$filter).toContain(" and ");
    expect(query.$search).toContain("from:alice@test.com");
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
});
