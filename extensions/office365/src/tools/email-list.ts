import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphListResponse } from "../types.js";
import { GraphApiError, toolSuccess, toolError } from "../types.js";
import { SELECT_FIELDS, formatMessageSummary } from "./_email-shared.js";

// ── Folder mapping ──────────────────────────────────────────────────────────

const FOLDER_MAP: Record<string, string> = {
  inbox: "Inbox",
  sent: "SentItems",
  sentitems: "SentItems",
  drafts: "Drafts",
  deleted: "DeletedItems",
  deleteditems: "DeletedItems",
  archive: "Archive",
  junk: "JunkEmail",
  junkemail: "JunkEmail",
};

const KNOWN_FOLDERS = new Set(Object.values(FOLDER_MAP));

function resolveFolder(input: string): string | { error: string } {
  const lower = input.toLowerCase().replace(/[\s_-]/g, "");
  const mapped = FOLDER_MAP[lower];
  if (mapped) return mapped;
  if (KNOWN_FOLDERS.has(input)) return input;
  // Allow raw folder IDs (long Graph API identifiers: alphanumeric + base64/URL-safe chars)
  if (input.length > 20 && /^[A-Za-z0-9+=_-]+$/.test(input)) return input;
  return {
    error: `Unknown folder '${input}'. Use: Inbox, SentItems, Drafts, DeletedItems, Archive, JunkEmail, or a folder ID.`,
  };
}

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailListSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  folder: Type.Optional(
    Type.String({
      description:
        "Mail folder to list. Default: Inbox. Options: Inbox, SentItems, Drafts, DeletedItems, Archive, JunkEmail, or a folder ID.",
    }),
  ),
  top: Type.Optional(
    Type.Number({
      description: "Number of messages to return (1-50, default 10).",
      minimum: 1,
      maximum: 50,
    }),
  ),
  filter: Type.Optional(
    Type.String({
      description:
        'OData $filter expression. Examples: "isRead eq false", "from/emailAddress/address eq \'user@example.com\'".',
    }),
  ),
  search: Type.Optional(
    Type.String({
      description:
        "Free-text search query (searches subject, body, addresses). Cannot be combined with filter.",
    }),
  ),
  skip: Type.Optional(
    Type.Number({
      description: "Number of messages to skip for pagination.",
      minimum: 0,
    }),
  ),
  orderBy: Type.Optional(
    Type.String({
      description:
        'Sort order. Default: "receivedDateTime desc". Ignored when search is used.',
    }),
  ),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailListTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "email_list",
    label: "List Emails",
    description:
      "List email messages from a Microsoft 365 mailbox. Returns subject, sender, date, read status, and body preview. Supports filtering, search, and pagination.",
    parameters: EmailListSchema,
    async execute(
      _toolCallId: string,
      args: unknown,
    ) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const folderInput = (typeof p.folder === "string" ? p.folder.trim() : "") || "Inbox";
      const top = Math.min(Math.max(Number(p.top) || 10, 1), 50);
      const filter = typeof p.filter === "string" ? p.filter.trim() : undefined;
      const search = typeof p.search === "string" ? p.search.trim() : undefined;
      const skip = typeof p.skip === "number" ? Math.max(0, Math.floor(p.skip)) : undefined;
      const orderBy = typeof p.orderBy === "string" ? p.orderBy.trim() : undefined;

      // Validate folder
      const folder = resolveFolder(folderInput);
      if (typeof folder === "object") {
        return { content: [{ type: "text" as const, text: JSON.stringify(toolError("user_input", folder.error), null, 2) }] };
      }

      // $search and $filter cannot be combined
      if (search && filter) {
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify(
              toolError("user_input", "Cannot combine $search and $filter in the same request. Use one or the other."),
              null, 2,
            ),
          }],
        };
      }

      try {
        const client = deps.resolveClient?.("email_list", account) ?? deps.graphClient;

        const query: Record<string, string> = {
          $top: String(top),
          $select: SELECT_FIELDS,
          $count: "true",
        };

        if (skip) query.$skip = String(skip);

        if (search) {
          // $search requires ConsistencyLevel header and $count
          // $orderby is ignored by Graph API when $search is used
          query.$search = `"${search.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
        } else {
          query.$orderby = orderBy || "receivedDateTime desc";
          if (filter) query.$filter = filter;
        }

        const extraHeaders: Record<string, string> = {};
        if (search) {
          extraHeaders["ConsistencyLevel"] = "eventual";
        }

        const basePath =
          folder === "Inbox"
            ? "/me/messages"
            : `/me/mailFolders/${encodeURIComponent(folder)}/messages`;

        const data = await client.fetchJson<GraphListResponse>(
          basePath,
          query,
          extraHeaders,
        );

        const warnings: string[] = [];
        if (search && orderBy) {
          warnings.push(
            "Note: $orderby is ignored by Microsoft Graph when $search is used. Results are ranked by relevance.",
          );
        }

        const result = toolSuccess({
          messages: (data.value ?? []).map(formatMessageSummary),
          totalCount: data["@odata.count"] ?? null,
          hasMore: !!data["@odata.nextLink"],
          nextSkip: data["@odata.nextLink"] ? (skip ?? 0) + top : null,
          ...(warnings.length > 0 ? { warnings } : {}),
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
