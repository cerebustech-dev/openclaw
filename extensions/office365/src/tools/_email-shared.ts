import type { GraphMessage } from "../types.js";

// ── Shared constants ────────────────────────────────────────────────────────

export const SELECT_FIELDS =
  "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,importance,flag";

// ── Message formatting ──────────────────────────────────────────────────────

export function formatAddress(addr: GraphMessage["from"]): string {
  if (!addr?.emailAddress) return "(unknown)";
  const { name, address } = addr.emailAddress;
  return name ? `${name} <${address}>` : address;
}

export function formatMessageSummary(msg: GraphMessage) {
  return {
    id: msg.id,
    subject: msg.subject ?? "(no subject)",
    from: formatAddress(msg.from),
    to: (msg.toRecipients ?? [])
      .map((r) => r.emailAddress?.address)
      .filter(Boolean),
    receivedDateTime: msg.receivedDateTime,
    isRead: msg.isRead ?? false,
    hasAttachments: msg.hasAttachments ?? false,
    bodyPreview: msg.bodyPreview ?? "",
    importance: msg.importance ?? "normal",
    flagStatus: msg.flag?.flagStatus ?? "notFlagged",
  };
}
