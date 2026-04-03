import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import { GraphApiError, toolSuccess, toolError } from "../types.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailSendSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  to: Type.Union([
    Type.Array(Type.String(), { minItems: 1 }),
    Type.String(),
  ], {
    description:
      "Recipient email address(es). A single string or array of strings.",
  }),
  subject: Type.String({
    description: "Email subject line.",
  }),
  body: Type.String({
    description: "Email body as HTML content.",
  }),
  cc: Type.Optional(
    Type.Array(Type.String(), {
      description: "CC recipient email addresses.",
    }),
  ),
  bcc: Type.Optional(
    Type.Array(Type.String(), {
      description: "BCC recipient email addresses.",
    }),
  ),
});

// ── Helpers ─────────────────────────────────────────────────────────────────

function toRecipients(
  input: unknown,
): Array<{ emailAddress: { address: string } }> | null {
  if (typeof input === "string") {
    const trimmed = input.trim();
    return trimmed ? [{ emailAddress: { address: trimmed } }] : null;
  }
  if (Array.isArray(input)) {
    const addrs = input
      .filter((v): v is string => typeof v === "string")
      .map((a) => a.trim())
      .filter(Boolean);
    return addrs.length > 0
      ? addrs.map((address) => ({ emailAddress: { address } }))
      : null;
  }
  return null;
}

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailSendTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "email_send",
    label: "Send Email",
    description:
      "Compose and send a new email via Microsoft 365. Supports HTML body, multiple recipients, CC, and BCC.",
    parameters: EmailSendSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;

      // ── Validate required fields ────────────────────────────────────────
      const recipients = toRecipients(p.to);
      if (!recipients) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(
              toolError("user_input", "At least one 'to' recipient is required."),
              null, 2,
            ),
          }],
        };
      }

      const subject = typeof p.subject === "string" ? p.subject.trim() : "";
      if (!subject) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(
              toolError("user_input", "A 'subject' is required."),
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

      // ── Build message ───────────────────────────────────────────────────
      const message: Record<string, unknown> = {
        subject,
        body: { contentType: "HTML", content: body },
        toRecipients: recipients,
      };

      const cc = toRecipients(p.cc);
      if (cc) message.ccRecipients = cc;

      const bcc = toRecipients(p.bcc);
      if (bcc) message.bccRecipients = bcc;

      // ── Send ────────────────────────────────────────────────────────────
      try {
        const client = deps.resolveClient?.("email_send", account) ?? deps.graphClient;
        await client.fetch("/me/sendMail", {
          method: "POST",
          body: JSON.stringify({ message }),
        });

        const result = toolSuccess({
          sent: true,
          to: recipients.map((r) => r.emailAddress.address),
          subject,
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
