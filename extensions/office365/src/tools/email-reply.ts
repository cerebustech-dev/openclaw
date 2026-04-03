import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import { GraphApiError, toolSuccess, toolError } from "../types.js";

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
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(
              toolError("user_input", "A 'messageId' is required."),
              null, 2,
            ),
          }],
        };
      }

      const body = typeof p.body === "string" ? p.body : "";
      if (!body) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(
              toolError("user_input", "A 'body' is required."),
              null, 2,
            ),
          }],
        };
      }

      const replyAll = p.replyAll === true;

      // ── Send reply ──────────────────────────────────────────────────────
      try {
        const client = deps.resolveClient?.("email_reply", account) ?? deps.graphClient;
        const encodedId = encodeURIComponent(messageId);
        const action = replyAll ? "replyAll" : "reply";

        await client.fetch(`/me/messages/${encodedId}/${action}`, {
          method: "POST",
          body: JSON.stringify({
            message: {
              body: { contentType: "HTML", content: body },
            },
          }),
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
