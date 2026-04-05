import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphAttachment } from "../types.js";
import { GraphApiError, toolSuccess, toolError } from "../types.js";

// ── Constants ───────────────────────────────────────────────────────────────

const MAX_DOWNLOAD_SIZE = 10 * 1024 * 1024; // 10MB

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailAttachmentReadSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  messageId: Type.String({
    description: "The ID of the email message (from email_list or email_read results).",
  }),
  attachmentId: Type.String({
    description: "The ID of the attachment to download (from email_read results).",
  }),
});

// ── Helpers ─────────────────────────────────────────────────────────────────

function classifyAttachmentType(
  odataType: string | undefined,
): "file" | "item" | "reference" | "unknown" {
  if (!odataType) return "unknown";
  if (odataType === "#microsoft.graph.fileAttachment") return "file";
  if (odataType === "#microsoft.graph.itemAttachment") return "item";
  if (odataType === "#microsoft.graph.referenceAttachment") return "reference";
  return "unknown";
}

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailAttachmentReadTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "email_attachment_read",
    label: "Download Email Attachment",
    description:
      "Download the content of a specific email attachment by message ID and attachment ID. Returns base64-encoded content for file attachments, metadata-only for other types.",
    parameters: EmailAttachmentReadSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const messageId = typeof p.messageId === "string" ? p.messageId.trim() : "";
      const attachmentId = typeof p.attachmentId === "string" ? p.attachmentId.trim() : "";

      if (!messageId) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(toolError("user_input", "messageId is required."), null, 2),
          }],
        };
      }

      if (!attachmentId) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(toolError("user_input", "attachmentId is required."), null, 2),
          }],
        };
      }

      try {
        const client = deps.resolveClient?.("email_attachment_read", account) ?? deps.graphClient;
        const encodedMsgId = encodeURIComponent(messageId);
        const encodedAttId = encodeURIComponent(attachmentId);

        const attachment = await client.fetchJson<GraphAttachment>(
          `/me/messages/${encodedMsgId}/attachments/${encodedAttId}`,
        );

        const attachmentType = classifyAttachmentType(attachment["@odata.type"]);

        // Non-file attachment types: return metadata only
        if (attachmentType !== "file") {
          const typeLabel = attachmentType === "item"
            ? "an Outlook item (e.g. forwarded message)"
            : attachmentType === "reference"
              ? "a cloud file reference (e.g. OneDrive link)"
              : "an unsupported attachment type";

          const result = toolSuccess({
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            size: attachment.size,
            contentBytes: null,
            attachmentType,
            note: `This attachment is ${typeLabel}. Content download is only supported for file attachments.`,
          });

          return {
            content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
            details: result,
          };
        }

        // Size guard: refuse to return content for very large attachments
        if (attachment.size > MAX_DOWNLOAD_SIZE) {
          const sizeMB = (attachment.size / (1024 * 1024)).toFixed(1);
          const result = toolSuccess({
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            size: attachment.size,
            contentBytes: null,
            attachmentType,
            tooLarge: true,
            warning: `Attachment is ${sizeMB}MB, which exceeds the 10MB download limit. Use the attachment metadata to inform the user.`,
          });

          return {
            content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
            details: result,
          };
        }

        const result = toolSuccess({
          id: attachment.id,
          name: attachment.name,
          contentType: attachment.contentType,
          size: attachment.size,
          contentBytes: attachment.contentBytes ?? null,
          attachmentType,
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
