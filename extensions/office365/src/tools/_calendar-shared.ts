import type { GraphClient } from "../graph-client.js";
import type { GraphEvent, GraphEventListResponse } from "../types.js";

// ── Shared constants ────────────────────────────────────────────────────────

export const DEFAULT_TIMEZONE = "America/Detroit";

export const EVENT_SELECT_FIELDS =
  "id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled,isOnlineMeeting,onlineMeetingUrl,importance,showAs,webLink";

// ── Address formatting ─────────────────────────────────────────────────────

export function formatAddress(addr: GraphEvent["organizer"]): string {
  if (!addr?.emailAddress) return "(unknown)";
  const { name, address } = addr.emailAddress;
  return name ? `${name} <${address}>` : address;
}

// ── Event formatting ───────────────────────────────────────────────────────

export function formatEventSummary(event: GraphEvent) {
  return {
    id: event.id,
    subject: event.subject ?? "(no subject)",
    start: event.start ?? null,
    end: event.end ?? null,
    location: event.location?.displayName ?? "",
    organizer: formatAddress(event.organizer),
    attendees: (event.attendees ?? []).map((a) => ({
      email: a.emailAddress.address,
      name: a.emailAddress.name ?? "",
      type: a.type,
      response: a.status?.response ?? "none",
    })),
    isAllDay: event.isAllDay ?? false,
    isCancelled: event.isCancelled ?? false,
    isOnlineMeeting: event.isOnlineMeeting ?? false,
    onlineMeetingUrl: event.onlineMeetingUrl ?? null,
    importance: event.importance ?? "normal",
    showAs: event.showAs ?? "busy",
    bodyPreview: event.bodyPreview ?? "",
    webLink: event.webLink ?? null,
  };
}

// ── Date range validation ──────────────────────────────────────────────────

// ── Conflict detection ─────────────────────────────────────────────────────

const CONFLICT_SHOW_AS = new Set(["busy", "oof", "workingElsewhere"]);

export type ConflictCheckResult = {
  hasConflicts: boolean;
  scanIncomplete: boolean;
  conflicts: Array<{
    id: string;
    subject: string;
    start: { dateTime: string; timeZone: string } | null;
    end: { dateTime: string; timeZone: string } | null;
    showAs: string;
  }>;
};

export async function checkConflicts(
  client: GraphClient,
  startDateTime: string,
  endDateTime: string,
  timeZone: string,
  excludeEventId?: string,
): Promise<ConflictCheckResult> {
  const data = await client.fetchJson<GraphEventListResponse>(
    "/me/calendarView",
    {
      startDateTime,
      endDateTime,
      $top: "50",
      $select: EVENT_SELECT_FIELDS,
      $orderby: "start/dateTime asc",
    },
    { Prefer: `outlook.timezone="${timeZone}"` },
  );

  const conflicts = (data.value ?? [])
    .filter(
      (e) =>
        e.showAs != null &&
        CONFLICT_SHOW_AS.has(e.showAs) &&
        e.isCancelled !== true &&
        e.id !== excludeEventId,
    )
    .map((e) => ({
      id: e.id,
      subject: e.subject ?? "(no subject)",
      start: e.start ?? null,
      end: e.end ?? null,
      showAs: e.showAs!,
    }));

  return {
    hasConflicts: conflicts.length > 0,
    scanIncomplete: !!data["@odata.nextLink"],
    conflicts,
  };
}

export function formatConflictMessage(
  conflicts: ConflictCheckResult["conflicts"],
  startDateTime: string,
  endDateTime: string,
  overrideKey: string,
  scanIncomplete: boolean,
): string {
  const lines = conflicts.map((c) => {
    const startStr = c.start?.dateTime ?? "?";
    const endStr = c.end?.dateTime ?? "?";
    return `  - "${c.subject}" (${startStr} - ${endStr}, showAs: ${c.showAs})`;
  });

  let msg = `Calendar conflict detected: ${conflicts.length} conflicting event(s) in the requested time range (${startDateTime} to ${endDateTime}).\n\nConflicts:\n${lines.join("\n")}\n\nTo proceed anyway, set ${overrideKey}: true.`;

  if (scanIncomplete) {
    msg += "\n\nNote: More than 50 events exist in this time range. Additional conflicts may exist beyond those listed.";
  }

  return msg;
}

// ── Date range validation ──────────────────────────────────────────────────

export function validateDateRange(
  startDateTime?: string,
  endDateTime?: string,
): { error: string } | null {
  if (!startDateTime && !endDateTime) return null;

  if (startDateTime && !endDateTime) {
    return { error: "Both startDateTime and endDateTime are required for calendar view. Got only startDateTime." };
  }
  if (!startDateTime && endDateTime) {
    return { error: "Both startDateTime and endDateTime are required for calendar view. Got only endDateTime." };
  }

  const startMs = Date.parse(startDateTime!);
  const endMs = Date.parse(endDateTime!);

  if (isNaN(startMs)) {
    return { error: `Invalid startDateTime '${startDateTime}'. Use ISO 8601 format (e.g., 2026-04-04T09:00:00).` };
  }
  if (isNaN(endMs)) {
    return { error: `Invalid endDateTime '${endDateTime}'. Use ISO 8601 format (e.g., 2026-04-04T17:00:00).` };
  }
  if (startMs >= endMs) {
    return { error: "startDateTime must be before endDateTime." };
  }

  return null;
}
