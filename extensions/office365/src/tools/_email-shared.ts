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

// ── Attachment validation (send/reply) ─────────────────────────────────────

export const MAX_SEND_ATTACHMENT_SIZE = 3 * 1024 * 1024; // 3MB decoded
export const MAX_SEND_ATTACHMENTS = 10;

const BASE64_RE = /^[A-Za-z0-9+/]*={0,2}$/;

export type AttachmentPayload = {
  "@odata.type": "#microsoft.graph.fileAttachment";
  name: string;
  contentType: string;
  contentBytes: string;
};

export function validateAndMapAttachments(
  raw: unknown,
): { ok: true; attachments: AttachmentPayload[] | null } | { ok: false; error: string } {
  if (raw == null || (Array.isArray(raw) && raw.length === 0)) {
    return { ok: true, attachments: null };
  }

  if (!Array.isArray(raw)) {
    return { ok: false, error: "'attachments' must be an array." };
  }

  if (raw.length > MAX_SEND_ATTACHMENTS) {
    return {
      ok: false,
      error: `Too many attachments (${raw.length}). Maximum is ${MAX_SEND_ATTACHMENTS}.`,
    };
  }

  const mapped: AttachmentPayload[] = [];

  for (let i = 0; i < raw.length; i++) {
    const a = raw[i];
    if (!a || typeof a !== "object") {
      return { ok: false, error: `Attachment ${i + 1}: must be an object.` };
    }

    const obj = a as Record<string, unknown>;
    const name = typeof obj.name === "string" ? obj.name.trim() : "";
    const contentType = typeof obj.contentType === "string" ? obj.contentType.trim() : "";
    const rawBytes = typeof obj.contentBytes === "string" ? obj.contentBytes.trim() : "";

    if (!name) {
      return { ok: false, error: `Attachment ${i + 1}: 'name' is required.` };
    }
    if (!contentType) {
      return { ok: false, error: `Attachment ${i + 1}: 'contentType' is required.` };
    }
    if (!rawBytes) {
      return { ok: false, error: `Attachment ${i + 1}: 'contentBytes' is required.` };
    }

    // Strip line breaks (common in multi-line base64)
    const stripped = rawBytes.replace(/[\r\n]/g, "");

    // Validate base64 format
    if (!BASE64_RE.test(stripped)) {
      return { ok: false, error: `Attachment ${i + 1} ('${name}'): contentBytes is not valid base64.` };
    }

    // Check decoded size
    const paddingCount = stripped.endsWith("==") ? 2 : stripped.endsWith("=") ? 1 : 0;
    const decodedBytes = Math.floor((stripped.length * 3) / 4) - paddingCount;
    if (decodedBytes > MAX_SEND_ATTACHMENT_SIZE) {
      const sizeMB = (decodedBytes / (1024 * 1024)).toFixed(1);
      return {
        ok: false,
        error: `Attachment ${i + 1} ('${name}'): ${sizeMB}MB exceeds the 3MB per-attachment limit.`,
      };
    }

    mapped.push({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name,
      contentType,
      contentBytes: stripped,
    });
  }

  return { ok: true, attachments: mapped };
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
