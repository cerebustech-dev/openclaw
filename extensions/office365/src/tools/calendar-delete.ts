import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import { GraphApiError, toolSuccess, toolErrorResult } from "../types.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const CalendarDeleteSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  eventId: Type.String({
    description: "The ID of the event to delete.",
  }),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createCalendarDeleteTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "calendar_delete",
    label: "Delete Calendar Event",
    description:
      "Delete a calendar event from Microsoft 365 by event ID. Returns an error if the event does not exist.",
    parameters: CalendarDeleteSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const eventId = typeof p.eventId === "string" ? p.eventId.trim() : "";

      if (!eventId) {
        return toolErrorResult("user_input", "An 'eventId' is required.");
      }

      try {
        const client = deps.resolveClient?.("calendar_delete", account) ?? deps.graphClient;
        await client.fetch(
          `/me/events/${encodeURIComponent(eventId)}`,
          { method: "DELETE" },
        );

        const result = toolSuccess({ deleted: true, eventId });

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
