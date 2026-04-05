import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphEventListResponse } from "../types.js";
import { toolErrorResult, toolSuccessResult, catchAsToolError } from "../types.js";
import {
  DEFAULT_TIMEZONE,
  EVENT_SELECT_FIELDS,
  formatEventSummary,
  validateDateRange,
} from "./_calendar-shared.js";

// ── Schema ──────────────────────────────────────────────────────────────────

export const CalendarListSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  startDateTime: Type.Optional(
    Type.String({
      description:
        "Start of date range (ISO 8601). When used with endDateTime, queries calendarView which expands recurring events.",
    }),
  ),
  endDateTime: Type.Optional(
    Type.String({
      description:
        "End of date range (ISO 8601). Required when startDateTime is provided.",
    }),
  ),
  top: Type.Optional(
    Type.Number({
      description: "Number of events to return (1-50, default 10).",
      minimum: 1,
      maximum: 50,
    }),
  ),
  skip: Type.Optional(
    Type.Number({
      description: "Number of events to skip for pagination.",
      minimum: 0,
    }),
  ),
  filter: Type.Optional(
    Type.String({
      description:
        'OData $filter expression. Example: "importance eq \'high\'".',
    }),
  ),
  orderBy: Type.Optional(
    Type.String({
      description:
        'Sort order. Default: "start/dateTime asc".',
    }),
  ),
  timeZone: Type.Optional(
    Type.String({
      description:
        'IANA time zone for results. Default: "America/Detroit".',
    }),
  ),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createCalendarListTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "calendar_list",
    label: "List Calendar Events",
    description:
      "List calendar events from a Microsoft 365 calendar. Supports date range queries (expands recurring events), filtering, and pagination. Returns subject, time, location, attendees, and online meeting info.",
    parameters: CalendarListSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const startDateTime = typeof p.startDateTime === "string" ? p.startDateTime.trim() : undefined;
      const endDateTime = typeof p.endDateTime === "string" ? p.endDateTime.trim() : undefined;
      const top = Math.min(Math.max(Number(p.top) || 10, 1), 50);
      const filter = typeof p.filter === "string" ? p.filter.trim() : undefined;
      const skip = typeof p.skip === "number" ? Math.max(0, Math.floor(p.skip)) : undefined;
      const orderBy = typeof p.orderBy === "string" ? p.orderBy.trim() : undefined;
      const timeZone = typeof p.timeZone === "string" ? p.timeZone.trim() : undefined;

      // Validate date range
      const dateError = validateDateRange(startDateTime, endDateTime);
      if (dateError) {
        return toolErrorResult("user_input", dateError.error);
      }

      try {
        const client = deps.resolveClient?.("calendar_list", account) ?? deps.graphClient;

        const useCalendarView = !!(startDateTime && endDateTime);
        const basePath = useCalendarView ? "/me/calendarView" : "/me/events";

        const query: Record<string, string> = {
          $top: String(top),
          $select: EVENT_SELECT_FIELDS,
          $count: "true",
        };

        if (useCalendarView) {
          query.startDateTime = startDateTime!;
          query.endDateTime = endDateTime!;
        }

        if (skip) query.$skip = String(skip);
        if (filter) query.$filter = filter;
        query.$orderby = orderBy || "start/dateTime asc";

        const tz = timeZone || DEFAULT_TIMEZONE;
        const extraHeaders: Record<string, string> = {
          Prefer: `outlook.timezone="${tz}"`,
        };

        const data = await client.fetchJson<GraphEventListResponse>(
          basePath,
          query,
          extraHeaders,
        );

        return toolSuccessResult({
          events: (data.value ?? []).map(formatEventSummary),
          totalCount: data["@odata.count"] ?? null,
          hasMore: !!data["@odata.nextLink"],
          nextSkip: data["@odata.nextLink"] ? (skip ?? 0) + top : null,
        });
      } catch (err) {
        return catchAsToolError(err);
      }
    },
  };
}
