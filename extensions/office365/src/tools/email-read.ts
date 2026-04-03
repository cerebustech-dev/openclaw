import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { AttachmentMeta, GraphMessage } from "../types.js";
import { GraphApiError, toolSuccess, toolError } from "../types.js";

// ── Constants ───────────────────────────────────────────────────────────────

const MAX_BODY_BYTES = 100 * 1024; // 100KB
const MAX_ATTACHMENTS = 25;

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailReadSchema = Type.Object({
  messageId: Type.String({
    description: "The ID of the email message to read (from email_list results).",
  }),
  bodyFormat: Type.Optional(
    Type.Union([Type.Literal("text"), Type.Literal("html")], {
      description: 'Body content format. Default: "text". Use "html" for rich formatting.',
    }),
  ),
  markAsRead: Type.Optional(
    Type.Boolean({
      description: "Whether to mark the message as read after reading. Default: false.",
    }),
  ),
});

// ── Helpers ─────────────────────────────────────────────────────────────────

const SELECT_FIELDS =
  "id,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,sentDateTime,isRead,hasAttachments,body,importance,flag,conversationId";

function formatAddress(addr: GraphMessage["from"]): string {
  if (!addr?.emailAddress) return "(unknown)";
  const { name, address } = addr.emailAddress;
  return name ? `${name} <${address}>` : address;
}

function extractAddresses(
  recipients: GraphMessage["toRecipients"],
): string[] {
  return (recipients ?? [])
    .map((r) => r.emailAddress?.address)
    .filter((a): a is string => !!a);
}

function truncateBody(content: string): { text: string; truncated: boolean } {
  const bytes = new TextEncoder().encode(content);
  if (bytes.length <= MAX_BODY_BYTES) {
    return { text: content, truncated: false };
  }
  // Truncate by decoding the allowed byte range
  const truncated = new TextDecoder("utf-8", { fatal: false }).decode(
    bytes.slice(0, MAX_BODY_BYTES),
  );
  return {
    text: `${truncated}\n\n[... truncated at 100KB]`,
    truncated: true,
  };
}

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailReadTool(deps: { graphClient: GraphClient }) {
  return {
    name: "email_read",
    label: "Read Email",
    description:
      "Read the full content of a specific email message by ID. Returns subject, full body, all recipients, attachment metadata, and conversation thread ID.",
    parameters: EmailReadSchema,
    async execute(
      _toolCallId: string,
      args: unknown,
    ) {
      const p = args as Record<string, unknown>;
      const messageId = typeof p.messageId === "string" ? p.messageId.trim() : "";
      const bodyFormat = p.bodyFormat === "html" ? "html" : "text";
      const markAsRead = p.markAsRead === true;

      if (!messageId) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(toolError("user_input", "messageId is required."), null, 2),
          }],
        };
      }

      try {
        const encodedId = encodeURIComponent(messageId);

        // Fetch the message
        const extraHeaders: Record<string, string> = {};
        if (bodyFormat) {
          extraHeaders["Prefer"] = `outlook.body-content-type="${bodyFormat}"`;
        }

        const msg = await deps.graphClient.fetchJson<GraphMessage>(
          `/me/messages/${encodedId}`,
          { $select: SELECT_FIELDS },
          extraHeaders,
        );

        // Fetch attachment metadata if present (excludes inline, max 25)
        let attachments: AttachmentMeta[] | undefined;
        if (msg.hasAttachments) {
          const attachData = await deps.graphClient.fetchJson<{
            value: AttachmentMeta[];
          }>(
            `/me/messages/${encodedId}/attachments`,
            {
              $select: "id,name,contentType,size",
              $filter: "isInline eq false",
              $top: String(MAX_ATTACHMENTS),
            },
          );
          attachments = attachData.value;
        }

        // Mark as read if requested
        if (markAsRead && !msg.isRead) {
          await deps.graphClient.fetch(`/me/messages/${encodedId}`, {
            method: "PATCH",
            body: JSON.stringify({ isRead: true }),
          });
        }

        // Truncate body if too large
        const bodyContent = msg.body?.content ?? "";
        const { text: body, truncated } = truncateBody(bodyContent);

        const bodyType = msg.body?.contentType ?? bodyFormat;
        const bodyUnsafeHtml = truncated && bodyType === "html" ? true : undefined;

        const result = toolSuccess({
          id: msg.id,
          subject: msg.subject ?? "(no subject)",
          from: formatAddress(msg.from),
          to: extractAddresses(msg.toRecipients),
          cc: extractAddresses(msg.ccRecipients),
          bcc: extractAddresses(msg.bccRecipients),
          replyTo: extractAddresses(msg.replyTo),
          receivedDateTime: msg.receivedDateTime,
          sentDateTime: msg.sentDateTime,
          isRead: markAsRead ? true : (msg.isRead ?? false),
          hasAttachments: msg.hasAttachments ?? false,
          body,
          bodyType,
          bodyTruncated: truncated,
          ...(bodyUnsafeHtml ? { bodyUnsafeHtml } : {}),
          importance: msg.importance ?? "normal",
          flagStatus: msg.flag?.flagStatus ?? "notFlagged",
          conversationId: msg.conversationId,
          ...(attachments && attachments.length > 0 ? { attachments } : {}),
        });

        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
          details: result,
        };
      } catch (err) {
        const category = err instanceof GraphApiError ? err.category : "transient";
        const safeMsg = err instanceof GraphApiError
          ? err.message
          : "An unexpected error occurred. Check gateway logs for details.";
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(toolError(category, safeMsg), null, 2),
          }],
        };
      }
    },
  };
}
