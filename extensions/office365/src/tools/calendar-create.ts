import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphEvent } from "../types.js";
import { GraphApiError, toolSuccess, toolError } from "../types.js";
import { DEFAULT_TIMEZONE, formatEventSummary, validateDateRange, checkConflicts, formatConflictMessage } from "./_calendar-shared.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const CalendarCreateSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  subject: Type.String({
    description: "Event subject/title.",
  }),
  startDateTime: Type.String({
    description: "Event start time (ISO 8601).",
  }),
  endDateTime: Type.String({
    description: "Event end time (ISO 8601).",
  }),
  timeZone: Type.Optional(
    Type.String({
      description: 'IANA time zone for start/end. Default: "America/Detroit".',
    }),
  ),
  body: Type.Optional(
    Type.String({ description: "Event body as HTML content." }),
  ),
  location: Type.Optional(
    Type.String({ description: "Location display name." }),
  ),
  attendees: Type.Optional(
    Type.Array(Type.String(), {
      description: "Attendee email addresses.",
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
  isOnlineMeeting: Type.Optional(
    Type.Boolean({ description: "Whether to create an online meeting link." }),
  ),
  checkConflicts: Type.Optional(
    Type.Boolean({
      description: "Check for conflicting events before creating. Default: true. Set false to skip.",
    }),
  ),
  forceCreate: Type.Optional(
    Type.Boolean({
      description: "Create even if conflicts are detected. Default: false.",
    }),
  ),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createCalendarCreateTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "calendar_create",
    label: "Create Calendar Event",
    description:
      "Create a new calendar event in Microsoft 365. Supports subject, time, location, attendees, online meeting, importance, and free/busy status.",
    parameters: CalendarCreateSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;

      // Validate required fields
      const subject = typeof p.subject === "string" ? p.subject.trim() : "";
      if (!subject) {
        return {
          content: [{ type: "text" as const, text: JSON.stringify(toolError("user_input", "A 'subject' is required."), null, 2) }],
        };
      }

      const startDateTime = typeof p.startDateTime === "string" ? p.startDateTime.trim() : "";
      if (!startDateTime) {
        return {
          content: [{ type: "text" as const, text: JSON.stringify(toolError("user_input", "A 'startDateTime' is required."), null, 2) }],
        };
      }

      const endDateTime = typeof p.endDateTime === "string" ? p.endDateTime.trim() : "";
      if (!endDateTime) {
        return {
          content: [{ type: "text" as const, text: JSON.stringify(toolError("user_input", "An 'endDateTime' is required."), null, 2) }],
        };
      }

      // Validate date range
      const dateError = validateDateRange(startDateTime, endDateTime);
      if (dateError) {
        return {
          content: [{ type: "text" as const, text: JSON.stringify(toolError("user_input", dateError.error), null, 2) }],
        };
      }

      const tz = typeof p.timeZone === "string" ? p.timeZone.trim() : DEFAULT_TIMEZONE;

      // ── Conflict check ──────────────────────────────────────────────────
      const shouldCheck = p.checkConflicts !== false;
      const force = p.forceCreate === true;
      type ConflictEntry = { id: string; subject: string; start: { dateTime: string; timeZone: string } | null; end: { dateTime: string; timeZone: string } | null; showAs: string };
      let conflictWarnings: ConflictEntry[] | undefined;

      if (shouldCheck) {
        try {
          const client = deps.resolveClient?.("calendar_create", account) ?? deps.graphClient;
          const conflictResult = await checkConflicts(client, startDateTime, endDateTime, tz);
          if (conflictResult.hasConflicts && !force) {
            return {
              content: [{
                type: "text" as const,
                text: JSON.stringify(
                  toolError("user_input", formatConflictMessage(
                    conflictResult.conflicts, startDateTime, endDateTime,
                    "forceCreate", conflictResult.scanIncomplete,
                  )),
                  null, 2,
                ),
              }],
            };
          }
          if (conflictResult.hasConflicts) {
            conflictWarnings = conflictResult.conflicts;
          }
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
      }

      // Build event body
      const event: Record<string, unknown> = {
        subject,
        start: { dateTime: startDateTime, timeZone: tz },
        end: { dateTime: endDateTime, timeZone: tz },
      };

      const body = typeof p.body === "string" ? p.body : undefined;
      if (body !== undefined) event.body = { contentType: "HTML", content: body };

      const location = typeof p.location === "string" ? p.location.trim() : undefined;
      if (location) event.location = { displayName: location };

      if (Array.isArray(p.attendees) && p.attendees.length > 0) {
        const addrs = p.attendees
          .filter((a): a is string => typeof a === "string")
          .map((a) => a.trim())
          .filter(Boolean);
        if (addrs.length > 0) {
          event.attendees = addrs.map((address) => ({
            emailAddress: { address },
            type: "required",
          }));
        }
      }

      if (typeof p.isAllDay === "boolean") event.isAllDay = p.isAllDay;
      if (typeof p.importance === "string") event.importance = p.importance.trim();
      if (typeof p.showAs === "string") event.showAs = p.showAs.trim();
      if (typeof p.isOnlineMeeting === "boolean") event.isOnlineMeeting = p.isOnlineMeeting;

      try {
        const client = deps.resolveClient?.("calendar_create", account) ?? deps.graphClient;
        const response = await client.fetch("/me/events", {
          method: "POST",
          body: JSON.stringify(event),
        });

        const created = (await response.json()) as GraphEvent;
        const result = toolSuccess({
          created: true,
          event: formatEventSummary(created),
          ...(conflictWarnings ? { warnings: { conflictsDetected: conflictWarnings.length, conflicts: conflictWarnings } } : {}),
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
          content: [{ type: "text" as const, text: JSON.stringify(toolError(category, safeMsg), null, 2) }],
        };
      }
    },
  };
}
