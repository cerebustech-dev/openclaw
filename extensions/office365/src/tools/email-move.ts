import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphMessage } from "../types.js";
import { GraphApiError, toolSuccess, toolErrorResult } from "../types.js";
import { formatMessageSummary, resolveFolder } from "./_email-shared.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailMoveSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  messageId: Type.String({
    description: "The ID of the email message to move (from email_list results).",
  }),
  destinationFolder: Type.String({
    description:
      "Target folder. Options: Inbox, SentItems, Drafts, DeletedItems, Archive, JunkEmail, or a folder ID.",
  }),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailMoveTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "email_move",
    label: "Move Email",
    description:
      "Move an email message to a different folder in Microsoft 365. Returns the moved message with its new ID (the ID may change after move).",
    parameters: EmailMoveSchema,
    async execute(
      _toolCallId: string,
      args: unknown,
    ) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const messageId = typeof p.messageId === "string" ? p.messageId.trim() : "";
      const destinationFolderInput = typeof p.destinationFolder === "string" ? p.destinationFolder.trim() : "";

      if (!messageId) {
        return toolErrorResult("user_input", "A 'messageId' is required.");
      }

      if (!destinationFolderInput) {
        return toolErrorResult("user_input", "A 'destinationFolder' is required.");
      }

      const folder = resolveFolder(destinationFolderInput);
      if (typeof folder === "object") {
        return toolErrorResult("user_input", folder.error);
      }

      try {
        const client = deps.resolveClient?.("email_move", account) ?? deps.graphClient;
        const encodedId = encodeURIComponent(messageId);

        const response = await client.fetch(
          `/me/messages/${encodedId}/move`,
          {
            method: "POST",
            body: JSON.stringify({ destinationId: folder }),
          },
        );

        const moved = (await response.json()) as GraphMessage;

        const result = toolSuccess({
          moved: true,
          previousMessageId: messageId,
          newMessageId: moved.id,
          message: formatMessageSummary(moved),
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
