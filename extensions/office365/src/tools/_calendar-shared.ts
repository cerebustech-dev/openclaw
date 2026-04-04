import type { GraphEvent } from "../types.js";

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
