import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphEvent } from "../types.js";
import { GraphApiError, toolSuccess, toolErrorResult } from "../types.js";
import { DEFAULT_TIMEZONE, formatEventSummary, checkConflicts, formatConflictMessage } from "./_calendar-shared.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const CalendarUpdateSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  eventId: Type.String({
    description: "The ID of the event to update.",
  }),
  subject: Type.Optional(
    Type.String({ description: "New event subject/title." }),
  ),
  startDateTime: Type.Optional(
    Type.String({ description: "New start time (ISO 8601)." }),
  ),
  endDateTime: Type.Optional(
    Type.String({ description: "New end time (ISO 8601)." }),
  ),
  timeZone: Type.Optional(
    Type.String({
      description: 'IANA time zone for start/end. Default: "America/Detroit".',
    }),
  ),
  body: Type.Optional(
    Type.String({ description: "New event body as HTML content." }),
  ),
  location: Type.Optional(
    Type.String({ description: "New location display name." }),
  ),
  attendees: Type.Optional(
    Type.Array(Type.String(), {
      description: "New attendee email addresses (replaces existing).",
    }),
  ),
  isAllDay: Type.Optional(
    Type.Boolean({ description: "Whether this is an all-day event." }),
  ),
  importance: Type.Optional(
    Type.String({
      description: 'Event importance: "low", "normal", or "high".',
    }),
  ),
  showAs: Type.Optional(
    Type.String({
      description:
        'Free/busy status: "free", "tentative", "busy", "oof", "workingElsewhere", "unknown".',
    }),
  ),
  checkConflicts: Type.Optional(
    Type.Boolean({
      description: "Check for conflicting events when changing time. Default: true. Set false to skip. Only checked when both startDateTime and endDateTime are provided.",
    }),
  ),
  forceUpdate: Type.Optional(
    Type.Boolean({
      description: "Update even if conflicts are detected. Default: false.",
    }),
  ),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createCalendarUpdateTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "calendar_update",
    label: "Update Calendar Event",
    description:
      "Update an existing calendar event in Microsoft 365. Only provided fields are changed (partial update). Supports subject, time, location, attendees, importance, and more.",
    parameters: CalendarUpdateSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const eventId = typeof p.eventId === "string" ? p.eventId.trim() : "";

      if (!eventId) {
        return toolErrorResult("user_input", "An 'eventId' is required.");
      }

      // Build partial update body — only include provided fields
      const patch: Record<string, unknown> = {};

      const subject = typeof p.subject === "string" ? p.subject.trim() : undefined;
      if (subject !== undefined) patch.subject = subject;

      const tz = typeof p.timeZone === "string" ? p.timeZone.trim() : DEFAULT_TIMEZONE;

      const startDateTime = typeof p.startDateTime === "string" ? p.startDateTime.trim() : undefined;
      if (startDateTime) patch.start = { dateTime: startDateTime, timeZone: tz };

      const endDateTime = typeof p.endDateTime === "string" ? p.endDateTime.trim() : undefined;
      if (endDateTime) patch.end = { dateTime: endDateTime, timeZone: tz };

      const body = typeof p.body === "string" ? p.body : undefined;
      if (body !== undefined) patch.body = { contentType: "HTML", content: body };

      const location = typeof p.location === "string" ? p.location.trim() : undefined;
      if (location !== undefined) patch.location = { displayName: location };

      if (Array.isArray(p.attendees) && p.attendees.length > 0) {
        patch.attendees = p.attendees
          .filter((a): a is string => typeof a === "string")
          .map((a) => a.trim())
          .filter(Boolean)
          .map((address) => ({ emailAddress: { address }, type: "required" }));
      }

      if (typeof p.isAllDay === "boolean") patch.isAllDay = p.isAllDay;
      if (typeof p.importance === "string") patch.importance = p.importance.trim();
      if (typeof p.showAs === "string") patch.showAs = p.showAs.trim();

      if (Object.keys(patch).length === 0) {
        return toolErrorResult("user_input", "At least one field to update must be provided.");
      }

      // ── Conflict check (only when both dates change) ────────────────────
      type ConflictEntry = { id: string; subject: string; start: { dateTime: string; timeZone: string } | null; end: { dateTime: string; timeZone: string } | null; showAs: string };
      let conflictWarnings: ConflictEntry[] | undefined;

      if (startDateTime && endDateTime && p.checkConflicts !== false) {
        try {
          const client = deps.resolveClient?.("calendar_update", account) ?? deps.graphClient;
          const conflictResult = await checkConflicts(client, startDateTime, endDateTime, tz, eventId);
          if (conflictResult.hasConflicts && p.forceUpdate !== true) {
            return toolErrorResult("user_input", formatConflictMessage(
                    conflictResult.conflicts, startDateTime, endDateTime,
                    "forceUpdate", conflictResult.scanIncomplete,
                  ));
          }
          if (conflictResult.hasConflicts) {
            conflictWarnings = conflictResult.conflicts;
          }
        } catch (err) {
          const category = err instanceof GraphApiError ? err.category : "transient";
          const safeMsg = err instanceof GraphApiError
            ? err.message
            : "An unexpected error occurred. Check gateway logs for details.";
          return toolErrorResult(category, safeMsg);
        }
      }

      try {
        const client = deps.resolveClient?.("calendar_update", account) ?? deps.graphClient;
        const response = await client.fetch(
          `/me/events/${encodeURIComponent(eventId)}`,
          { method: "PATCH", body: JSON.stringify(patch) },
        );

        // PATCH typically returns 200 with updated event, but handle 204 edge case
        let event: ReturnType<typeof formatEventSummary> | undefined;
        try {
          const text = await response.text();
          if (text) {
            const parsed = JSON.parse(text) as GraphEvent;
            event = formatEventSummary(parsed);
          }
        } catch {
          // No parseable body — fall through
        }

        const warnings = conflictWarnings
          ? { warnings: { conflictsDetected: conflictWarnings.length, conflicts: conflictWarnings } }
          : {};
        const result = event
          ? toolSuccess({ updated: true, event, ...warnings })
          : toolSuccess({ updated: true, eventId, ...warnings });

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
