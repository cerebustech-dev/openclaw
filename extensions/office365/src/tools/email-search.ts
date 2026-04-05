import { Type } from "@sinclair/typebox";
import type { GraphClient } from "../graph-client.js";
import type { GraphListResponse } from "../types.js";
import { GraphApiError, toolSuccess, toolErrorResult } from "../types.js";
import { SELECT_FIELDS, formatMessageSummary } from "./_email-shared.js";

// ── KQL builder ─────────────────────────────────────────────────────────────

function kqlQuoteIfNeeded(value: string): string {
  if (/[\s"(]/.test(value)) {
    return `"${value.replace(/"/g, '\\"')}"`;
  }
  return value;
}

function buildKql(params: {
  query?: string;
  from?: string;
  to?: string;
  subject?: string;
  hasAttachments?: boolean;
  dateFrom?: string;
  dateTo?: string;
}): string {
  const parts: string[] = [];

  if (params.from) parts.push(`from:${kqlQuoteIfNeeded(params.from)}`);
  if (params.to) parts.push(`to:${kqlQuoteIfNeeded(params.to)}`);
  if (params.subject) parts.push(`subject:${kqlQuoteIfNeeded(params.subject)}`);
  if (params.hasAttachments !== undefined) parts.push(`hasAttachments:${params.hasAttachments}`);
  if (params.dateFrom) parts.push(`received>=${params.dateFrom}`);
  if (params.dateTo) parts.push(`received<=${params.dateTo}`);
  if (params.query) parts.push(params.query);

  return parts.join(" ");
}

// ── Date parsing ────────────────────────────────────────────────────────────

const DATE_ONLY_RE = /^\d{4}-\d{2}-\d{2}$/;

function parseSearchDate(
  input: string,
  paramName: string,
): { iso: string; isDateOnly: boolean } | { error: string } {
  const trimmed = input.trim();
  if (DATE_ONLY_RE.test(trimmed)) {
    const d = new Date(`${trimmed}T00:00:00Z`);
    if (isNaN(d.getTime())) return { error: `Invalid ${paramName}: '${trimmed}'.` };
    return { iso: `${trimmed}T00:00:00Z`, isDateOnly: true };
  }
  const d = new Date(trimmed);
  if (isNaN(d.getTime())) return { error: `Invalid ${paramName}: '${trimmed}'. Use YYYY-MM-DD or ISO 8601.` };
  return { iso: d.toISOString(), isDateOnly: false };
}

function addOneDay(dateStr: string): string {
  const d = new Date(dateStr);
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

// ── Schema ──────────────────────────────────────────────────────────────────

export const EmailSearchSchema = Type.Object({
  account: Type.Optional(
    Type.String({
      description: "Account to use. Defaults based on tool type.",
    }),
  ),
  query: Type.Optional(
    Type.String({
      description:
        "Free-text search query. Searches subject, body, and addresses. Results are ranked by relevance, not date — combine with dateFrom/dateTo for recent results.",
    }),
  ),
  from: Type.Optional(
    Type.String({
      description: "Filter by sender email address.",
    }),
  ),
  to: Type.Optional(
    Type.String({
      description: "Filter by recipient email address.",
    }),
  ),
  subject: Type.Optional(
    Type.String({
      description: "Filter by subject text.",
    }),
  ),
  hasAttachments: Type.Optional(
    Type.Boolean({
      description: "Filter for messages with (true) or without (false) attachments.",
    }),
  ),
  dateFrom: Type.Optional(
    Type.String({
      description:
        "Start of date range (inclusive). YYYY-MM-DD or ISO 8601 datetime.",
    }),
  ),
  dateTo: Type.Optional(
    Type.String({
      description:
        "End of date range (inclusive for date-only, exact for datetime). YYYY-MM-DD or ISO 8601.",
    }),
  ),
  top: Type.Optional(
    Type.Number({
      description: "Number of results to return (1-50, default 10).",
      minimum: 1,
      maximum: 50,
    }),
  ),
});

// ── Tool factory ────────────────────────────────────────────────────────────

export function createEmailSearchTool(deps: {
  graphClient: GraphClient;
  resolveClient?: (toolName: string, accountId?: string) => GraphClient;
}) {
  return {
    name: "email_search",
    label: "Search Emails",
    description:
      "Search email messages across all folders in a Microsoft 365 mailbox. Supports free-text queries, structured field search (from, to, subject), date ranges, and attachment filtering. Results are relevance-ranked when using text search.",
    parameters: EmailSearchSchema,
    async execute(_toolCallId: string, args: unknown) {
      const p = args as Record<string, unknown>;
      const account = typeof p.account === "string" ? p.account.trim() : undefined;
      const query = typeof p.query === "string" ? p.query.trim() : undefined;
      const from = typeof p.from === "string" ? p.from.trim() : undefined;
      const to = typeof p.to === "string" ? p.to.trim() : undefined;
      const subject = typeof p.subject === "string" ? p.subject.trim() : undefined;
      const hasAttachments = typeof p.hasAttachments === "boolean" ? p.hasAttachments : undefined;
      const dateFromRaw = typeof p.dateFrom === "string" ? p.dateFrom.trim() : undefined;
      const dateToRaw = typeof p.dateTo === "string" ? p.dateTo.trim() : undefined;
      const top = Math.min(Math.max(Number(p.top) || 10, 1), 50);

      // Must provide at least one search criterion
      if (!query && !from && !to && !subject && hasAttachments === undefined && !dateFromRaw && !dateToRaw) {
        return toolErrorResult("user_input", "Provide at least one search criterion (query, from, to, subject, hasAttachments, dateFrom, or dateTo).");
      }

      // Parse and validate dates
      let dateFromIso: string | undefined;
      let dateToIso: string | undefined;

      if (dateFromRaw) {
        const parsed = parseSearchDate(dateFromRaw, "dateFrom");
        if ("error" in parsed) {
          return toolErrorResult("user_input", parsed.error);
        }
        dateFromIso = parsed.iso;
      }

      if (dateToRaw) {
        const parsed = parseSearchDate(dateToRaw, "dateTo");
        if ("error" in parsed) {
          return toolErrorResult("user_input", parsed.error);
        }
        // For date-only dateTo, advance to next day for exclusive boundary
        dateToIso = parsed.isDateOnly ? addOneDay(parsed.iso) : parsed.iso;
      }

      if (dateFromIso && dateToIso && dateFromIso >= dateToIso) {
        return toolErrorResult("user_input", "dateFrom must be before dateTo.");
      }

      try {
        const client = deps.resolveClient?.("email_search", account) ?? deps.graphClient;

        const oDataQuery: Record<string, string> = {
          $top: String(top),
          $select: SELECT_FIELDS,
          $count: "true",
        };

        // Build KQL from structured fields + free-text
        // Graph API blocks combining $search with $filter, so when we have
        // text/field criteria AND dates, dates go into KQL as received>=/<=
        // (day granularity). $filter is only used for date-only searches.
        const hasTextCriteria = !!(query || from || to || subject || hasAttachments !== undefined);
        const hasDates = !!(dateFromIso || dateToIso);

        // Extract YYYY-MM-DD for KQL date clauses
        const dateFromDate = dateFromIso?.slice(0, 10);
        const dateToDate = dateToRaw && dateToIso
          ? (DATE_ONLY_RE.test(dateToRaw) ? dateToRaw : dateToIso.slice(0, 10))
          : undefined;

        const kql = buildKql({
          query, from, to, subject, hasAttachments,
          // Only include dates in KQL when combining with text criteria
          dateFrom: hasTextCriteria && dateFromDate ? dateFromDate : undefined,
          dateTo: hasTextCriteria && dateToDate ? dateToDate : undefined,
        });
        const hasKql = kql.length > 0;

        if (hasKql) {
          // Wrap for OData $search: escape backslashes first, then quotes
          oDataQuery.$search = `"${kql.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
        }

        // $filter for date ranges ONLY when no text criteria (pure date search)
        if (hasDates && !hasTextCriteria) {
          const filterClauses: string[] = [];
          if (dateFromIso) filterClauses.push(`receivedDateTime ge ${dateFromIso}`);
          if (dateToIso) filterClauses.push(`receivedDateTime lt ${dateToIso}`);
          oDataQuery.$filter = filterClauses.join(" and ");
        }

        // $orderby only works without $search
        if (!hasKql) {
          oDataQuery.$orderby = "receivedDateTime desc";
        }

        // ConsistencyLevel required for $search
        const extraHeaders: Record<string, string> = {};
        if (hasKql) {
          extraHeaders.ConsistencyLevel = "eventual";
        }

        const data = await client.fetchJson<GraphListResponse>(
          "/me/messages",
          oDataQuery,
          Object.keys(extraHeaders).length > 0 ? extraHeaders : undefined,
        );

        const result = toolSuccess({
          messages: (data.value ?? []).map(formatMessageSummary),
          totalCount: data["@odata.count"] ?? null,
          hasMore: !!data["@odata.nextLink"],
          nextSkip: data["@odata.nextLink"] ? top : null,
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
        return toolErrorResult(category, safeMsg);
      }
    },
  };
}
