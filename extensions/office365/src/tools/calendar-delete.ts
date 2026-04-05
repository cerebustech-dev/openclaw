import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import { toolErrorResult, toolSuccessResult, catchAsToolError } from "../types.js";

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

        return toolSuccessResult({ deleted: true, eventId });
      } catch (err) {
        return catchAsToolError(err);
      }
    },
  };
}
