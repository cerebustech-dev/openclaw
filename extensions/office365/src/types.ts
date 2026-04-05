// ── Graph API response types ────────────────────────────────────────────────

export type GraphEmailAddress = {
  emailAddress: { name?: string; address: string };
};

export type GraphMessage = {
  id: string;
  subject?: string;
  from?: GraphEmailAddress;
  toRecipients?: GraphEmailAddress[];
  ccRecipients?: GraphEmailAddress[];
  bccRecipients?: GraphEmailAddress[];
  replyTo?: GraphEmailAddress[];
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  bodyPreview?: string;
  body?: { contentType: string; content: string };
  importance?: string;
  flag?: { flagStatus: string };
  conversationId?: string;
};

export type GraphListResponse = {
  value?: GraphMessage[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
};

export type AttachmentMeta = {
  id: string;
  name: string;
  contentType: string;
  size: number;
};

export type GraphAttachment = AttachmentMeta & {
  "@odata.type"?: string;
  contentBytes?: string;
  contentId?: string;
  isInline?: boolean;
};

// ── Calendar event types ────────────────────────────────────────────────────

export type GraphEvent = {
  id: string;
  subject?: string;
  body?: { contentType: string; content: string };
  bodyPreview?: string;
  start?: { dateTime: string; timeZone: string };
  end?: { dateTime: string; timeZone: string };
  location?: { displayName?: string };
  organizer?: { emailAddress: { name?: string; address: string } };
  attendees?: Array<{
    emailAddress: { name?: string; address: string };
    type: string;
    status?: { response: string; time: string };
  }>;
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  importance?: string;
  showAs?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  webLink?: string;
};

export type GraphEventListResponse = {
  value?: GraphEvent[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
};

// ── Error taxonomy ──────────────────────────────────────────────────────────

export type GraphErrorCategory =
  | "auth"
  | "permission"
  | "throttle"
  | "transient"
  | "not_found"
  | "user_input";

export class GraphApiError extends Error {
  override readonly name = "GraphApiError";
  constructor(
    message: string,
    public readonly category: GraphErrorCategory,
    public readonly status: number,
    public readonly retryAfterMs?: number,
  ) {
    super(message);
  }
}

// ── Tool response envelope ──────────────────────────────────────────────────

export type ToolResponse<T> = {
  schemaVersion: 1;
  data?: T;
  error?: { category: GraphErrorCategory; message: string };
};

export function toolSuccess<T>(data: T): ToolResponse<T> {
  return { schemaVersion: 1, data };
}

export function toolError(
  category: GraphErrorCategory,
  message: string,
): ToolResponse<never> {
  return { schemaVersion: 1, error: { category, message } };
}

// ── Credential shape ────────────────────────────────────────────────────────

export type Office365Credential = {
  access: string;
  refresh: string;
  expires: number;
  email?: string;
};

// ── Plugin config ───────────────────────────────────────────────────────────

export type Office365AccountConfig = {
  name?: string;
  email?: string;
  scopes?: string[];
  tools?: string[];
  clientId?: string;
  tenantId?: string;
  clientSecret?: string;
  redirectUri?: string;
};

export type Office365Config = {
  clientId: string;
  tenantId: string;
  clientSecret: string;
  redirectUri: string;
  scopes: string[];
  defaultAccount?: string;
  accounts?: Record<string, Office365AccountConfig>;
};

export type ResolvedOffice365Account = {
  accountId: string;
  name: string;
  email?: string;
  config: Office365Config;
  tools: string[];
};
