import { randomBytes } from "node:crypto";
import { writeFileSync, readFileSync, renameSync, mkdirSync, chmodSync, unlinkSync } from "node:fs";
import { dirname, join } from "node:path";
import type { PluginLogger } from "openclaw/plugin-sdk";
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import { ACCOUNT_ID_RE } from "./accounts.js";
import { refreshMicrosoftTokens } from "./oauth.js";
import {
  GraphApiError,
  type GraphErrorCategory,
  type Office365Config,
  type Office365Credential,
} from "./types.js";

// ── Constants ───────────────────────────────────────────────────────────────

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const FETCH_TIMEOUT_MS = 30_000;
const MAX_RETRY_AFTER_MS = 30_000;
const DEFAULT_RETRY_AFTER_MS = 5_000;
const TRANSIENT_RETRY_DELAY_MS = 1_000;

// ── GraphClient type ────────────────────────────────────────────────────────

export type GraphClient = {
  fetch(path: string, init?: RequestInit): Promise<Response>;
  fetchJson<T>(
    path: string,
    query?: Record<string, string>,
    extraHeaders?: Record<string, string>,
  ): Promise<T>;
  setCredential(cred: Office365Credential): void;
};

// ── Credential persistence ──────────────────────────────────────────────────

function credentialPath(stateDir: string, accountId = "default"): string {
  if (accountId === "default") {
    return join(stateDir, "office365-credentials.json");
  }
  return join(stateDir, `office365-credentials-${accountId}.json`);
}

function writeCredentialFile(path: string, cred: Office365Credential, logger?: PluginLogger): void {
  const dir = dirname(path);
  mkdirSync(dir, { recursive: true, mode: 0o700 });

  // Atomic write: write to temp file, then rename
  const tmpPath = `${path}.${randomBytes(4).toString("hex")}.tmp`;
  writeFileSync(tmpPath, JSON.stringify(cred, null, 2), { encoding: "utf8", mode: 0o600 });
  try {
    chmodSync(tmpPath, 0o600);
  } catch (err) {
    if (process.platform === "win32") {
      logger?.warn(
        "office365: chmod not supported on Windows, file permissions may be broader than intended",
      );
    } else {
      throw err;
    }
  }
  renameSync(tmpPath, path);
}

function readCredentialFile(path: string): Office365Credential | null {
  try {
    const raw = readFileSync(path, "utf8");
    const parsed = JSON.parse(raw) as Office365Credential;
    if (typeof parsed.access !== "string" || !parsed.access) {
      return null;
    }
    if (typeof parsed.refresh !== "string" || !parsed.refresh) {
      return null;
    }
    if (typeof parsed.expires !== "number" || parsed.expires <= 0) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

function deleteCredentialFile(path: string): void {
  try {
    unlinkSync(path);
  } catch {
    // ignore — file may not exist
  }
}

// ── Graph fetch with SSRF guard ─────────────────────────────────────────────

async function graphFetch(url: string, init: RequestInit): Promise<Response> {
  const { response, release } = await fetchWithSsrFGuard({
    url,
    init,
    timeoutMs: FETCH_TIMEOUT_MS,
  });
  try {
    // Null-body status codes (204, 304) must have null body per Fetch spec
    const isNullBody =
      response.status === 204 || response.status === 205 || response.status === 304;
    const body = isNullBody ? null : await response.arrayBuffer();
    return new Response(body, {
      status: response.status,
      statusText: response.statusText,
      headers: response.headers,
    });
  } finally {
    await release();
  }
}

// ── Error classification ────────────────────────────────────────────────────

function classifyHttpError(status: number): GraphErrorCategory {
  if (status === 401) {
    return "auth";
  }
  if (status === 403) {
    return "permission";
  }
  if (status === 404) {
    return "not_found";
  }
  if (status === 429) {
    return "throttle";
  }
  if (status >= 500) {
    return "transient";
  }
  return "user_input";
}

// ── Factory ─────────────────────────────────────────────────────────────────

export function createGraphClient(params: {
  config: Office365Config;
  stateDir: string;
  logger: PluginLogger;
  accountId?: string;
}): GraphClient {
  const { config, logger, accountId } = params;

  // Validate accountId before using it in file paths
  if (accountId && accountId !== "default" && !ACCOUNT_ID_RE.test(accountId)) {
    throw new Error(
      `Invalid account ID '${accountId}'. Must match ${ACCOUNT_ID_RE} (lowercase alphanumeric, hyphens, underscores).`,
    );
  }

  const credPath = credentialPath(params.stateDir, accountId);

  let cached: { access: string; expires: number } | null = null;
  let refreshInFlight: Promise<string> | null = null;

  async function doRefresh(): Promise<string> {
    // Read current credential from file
    const fileCred = readCredentialFile(credPath);
    if (!fileCred) {
      throw new GraphApiError(
        "Microsoft 365 not authenticated. Run `openclaw auth` to set up.",
        "auth",
        401,
      );
    }

    // If file credential is still valid, use it
    if (Date.now() < fileCred.expires) {
      cached = { access: fileCred.access, expires: fileCred.expires };
      return fileCred.access;
    }

    // Refresh the token
    logger.info("office365: refreshing expired access token");
    let refreshed: Office365Credential;
    try {
      refreshed = await refreshMicrosoftTokens(fileCred.refresh, config);
    } catch (err) {
      // If refresh token is revoked, delete corrupted credential
      const msg = err instanceof Error ? err.message : String(err);
      if (msg.includes("expired or revoked") || msg.includes("re-authenticate")) {
        deleteCredentialFile(credPath);
      }
      throw err;
    }
    refreshed.email = fileCred.email;

    writeCredentialFile(credPath, refreshed, logger);
    cached = { access: refreshed.access, expires: refreshed.expires };
    logger.info("office365: access token refreshed successfully");
    return refreshed.access;
  }

  async function resolveAccessToken(): Promise<string> {
    // Check in-memory cache first
    if (cached && Date.now() < cached.expires) {
      return cached.access;
    }

    // Single-flight: if a refresh is already running, await it
    if (refreshInFlight) {
      return refreshInFlight;
    }

    refreshInFlight = doRefresh().finally(() => {
      refreshInFlight = null;
    });
    return refreshInFlight;
  }

  async function doFetch(path: string, init?: RequestInit, isRetry = false): Promise<Response> {
    const token = await resolveAccessToken();
    const url = `${GRAPH_BASE}${path}`;
    const start = Date.now();

    const response = await graphFetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(init?.headers as Record<string, string>),
      },
    });

    const latency = Date.now() - start;
    logger.debug?.(
      `office365: ${init?.method ?? "GET"} ${path} → ${response.status} (${latency}ms)`,
    );

    if (response.ok) {
      return response;
    }

    // Retry logic (only on first attempt)
    if (!isRetry) {
      if (response.status === 401) {
        logger.warn("office365: 401 — clearing token cache and retrying");
        cached = null;
        return doFetch(path, init, true);
      }

      if (response.status === 429) {
        const retryAfterRaw = response.headers.get("Retry-After");
        const retryAfterMs = retryAfterRaw
          ? Math.min(Number.parseInt(retryAfterRaw, 10) * 1000, MAX_RETRY_AFTER_MS)
          : DEFAULT_RETRY_AFTER_MS;
        logger.warn(`office365: 429 — waiting ${retryAfterMs}ms before retry`);
        await new Promise((r) => setTimeout(r, retryAfterMs));
        return doFetch(path, init, true);
      }

      if (response.status >= 500) {
        logger.warn(`office365: ${response.status} — retrying after ${TRANSIENT_RETRY_DELAY_MS}ms`);
        await new Promise((r) => setTimeout(r, TRANSIENT_RETRY_DELAY_MS));
        return doFetch(path, init, true);
      }
    }

    // Non-retryable error
    const errorText = await response.text().catch(() => "");
    const category = classifyHttpError(response.status);
    let message = `Graph API ${path} failed (${response.status})`;

    if (category === "permission") {
      message +=
        ". Check that the Azure app has the required scopes (Mail.ReadWrite, Calendars.ReadWrite).";
    } else if (category === "not_found") {
      message += ". The requested resource was not found.";
    } else if (errorText) {
      message += `: ${errorText.slice(0, 500)}`;
    }

    throw new GraphApiError(message, category, response.status);
  }

  return {
    fetch: (path, init) => doFetch(path, init),

    async fetchJson<T>(
      path: string,
      query?: Record<string, string>,
      extraHeaders?: Record<string, string>,
    ): Promise<T> {
      const qs =
        query && Object.keys(query).length > 0 ? `?${new URLSearchParams(query).toString()}` : "";
      const response = await doFetch(
        `${path}${qs}`,
        extraHeaders ? { headers: extraHeaders } : undefined,
      );
      return (await response.json()) as T;
    },

    setCredential(cred: Office365Credential): void {
      writeCredentialFile(credPath, cred, logger);
      cached = { access: cred.access, expires: cred.expires };
      logger.info(`office365: credential stored for ${cred.email ?? "default"}`);
    },
  };
}
