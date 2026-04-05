import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import { GraphApiError, toolSuccess, toolErrorResult } from "../types.js";
import { validateAndMapAttachments } from "./_email-shared.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailReplySchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  messageId: Type.String({
    description: "The ID of the email message to reply to (from email_list or email_read results).",
  }),
  body: Type.String({
    description: "Reply body as HTML content.",
  }),
  replyAll: Type.Optional(
    Type.Boolean({
      description: "Whether to reply to all recipients. Default: false (reply to sender only).",
    }),
  ),
  attachments: Type.Optional(
    Type.Array(Type.Object({
      name: Type.String({ description: "Filename including extension." }),
      contentType: Type.String({ description: "MIME type (e.g. application/pdf)." }),
      contentBytes: Type.String({ description: "Base64-encoded file content." }),
    }), { description: "File attachments. Max 10 attachments, max 3MB each." }),
  ),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailReplyTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "email_reply",
    label: "Reply to Email",
    description:
      "Reply to an existing email message. Supports reply (sender only) and reply-all. Uses HTML body.",
    parameters: EmailReplySchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;

      // ── Validate required fields ────────────────────────────────────────
      const messageId = typeof p.messageId === "string" ? p.messageId.trim() : "";
      if (!messageId) {
        return toolErrorResult("user_input", "A 'messageId' is required.");
      }

      const body = typeof p.body === "string" ? p.body : "";
      if (!body) {
        return toolErrorResult("user_input", "A 'body' is required.");
      }

      const replyAll = p.replyAll === true;

      // ── Validate attachments ────────────────────────────────────────────
      const attachResult = validateAndMapAttachments(p.attachments);
      if (!attachResult.ok) {
        return toolErrorResult("user_input", attachResult.error);
      }

      // ── Send reply ──────────────────────────────────────────────────────
      try {
        const client = deps.resolveClient?.("email_reply", account) ?? deps.graphClient;
        const encodedId = encodeURIComponent(messageId);
        const action = replyAll ? "replyAll" : "reply";

        const message: Record<string, unknown> = {
          body: { contentType: "HTML", content: body },
        };
        if (attachResult.attachments) {
          message.attachments = attachResult.attachments;
        }

        await client.fetch(`/me/messages/${encodedId}/${action}`, {
          method: "POST",
          body: JSON.stringify({ message }),
        });

        const result = toolSuccess({
          replied: true,
          messageId,
          replyAll,
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
        return toolErrorResult(category, safeMsg);
      }
    },
  };
}
