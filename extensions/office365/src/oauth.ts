import { createHash, randomBytes } from "node:crypto";
import { createServer } from "node:http";
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import type { Office365Config, Office365Credential } from "./types.js";

// ── Constants ───────────────────────────────────────────────────────────────

const CALLBACK_PORT = 8080;
const CALLBACK_PATH = "/callback";
const REDIRECT_URI = `http://localhost:${CALLBACK_PORT}${CALLBACK_PATH}`;
const OAUTH_TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes
const EXPIRES_BUFFER_MS = 5 * 60 * 1000; // 5-minute safety buffer
const FETCH_TIMEOUT_MS = 15_000;

const WELL_KNOWN_TENANTS = new Set(["common", "organizations", "consumers"]);
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function validateTenantId(tenantId: string): void {
  if (WELL_KNOWN_TENANTS.has(tenantId)) {
    return;
  }
  if (GUID_RE.test(tenantId)) {
    return;
  }
  throw new Error(
    `Invalid tenantId "${tenantId}". Must be a GUID (8-4-4-4-12 hex) or one of: common, organizations, consumers.`,
  );
}

function authorizeUrl(tenantId: string): string {
  validateTenantId(tenantId);
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
}

function tokenUrl(tenantId: string): string {
  validateTenantId(tenantId);
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
}

const GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me";

// ── PKCE ────────────────────────────────────────────────────────────────────

export type PkcePair = { verifier: string; challenge: string };

export function generatePkce(): PkcePair {
  const verifier = randomBytes(32).toString("hex");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

// ── Authorize URL builder ───────────────────────────────────────────────────

export function buildAuthorizeUrl(
  config: Office365Config,
  challenge: string,
  state: string,
  loginHint?: string,
): string {
  const redirectUri = config.redirectUri || REDIRECT_URI;
  const qs = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    scope: config.scopes.join(" "),
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "consent",
  });
  if (loginHint) {
    qs.set("login_hint", loginHint);
  }
  return `${authorizeUrl(config.tenantId)}?${qs.toString()}`;
}

// ── Callback URL parsing (for VPS manual paste flow) ────────────────────────

export function parseCallbackInput(
  input: string,
  expectedState: string,
): { code: string; state: string } | { error: string } {
  const trimmed = input.trim();
  if (!trimmed) {
    return { error: "No input provided." };
  }

  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    // Handle query-string-only paste: ?code=...&state=... or code=...&state=...
    const qs = trimmed.startsWith("?") ? trimmed : `?${trimmed}`;
    try {
      url = new URL(`http://localhost/${qs}`);
    } catch {
      return {
        error: "Could not parse input. Paste the full redirect URL from your browser address bar.",
      };
    }
  }

  const code = url.searchParams.get("code")?.trim();
  const state = url.searchParams.get("state")?.trim();

  if (!code) {
    return {
      error:
        "Missing 'code' parameter. Make sure you copied the full URL from the browser address bar (it may have been truncated).",
    };
  }
  if (!state) {
    return {
      error:
        "Missing 'state' parameter. The URL may have been truncated. Copy the entire URL including all query parameters.",
    };
  }
  if (state !== expectedState) {
    return {
      error:
        "OAuth state mismatch — possible CSRF or wrong session. Did you use the correct Azure tenant? Run `openclaw auth` to retry.",
    };
  }
  return { code, state };
}

// ── Localhost callback server ───────────────────────────────────────────────

async function waitForLocalCallback(params: {
  redirectUri: string;
  expectedState: string;
  timeoutMs: number;
  onProgress?: (message: string) => void;
}): Promise<{ code: string; state: string }> {
  const redirectUrl = new URL(params.redirectUri);
  const hostname = redirectUrl.hostname || "localhost";
  const LOOPBACK_HOSTS = new Set(["localhost", "127.0.0.1", "::1", "[::1]"]);
  if (!LOOPBACK_HOSTS.has(hostname)) {
    throw new Error(
      `Refusing to bind callback server to non-loopback address "${hostname}". ` +
        "redirectUri must use localhost, 127.0.0.1, or ::1.",
    );
  }
  const port = redirectUrl.port ? Number.parseInt(redirectUrl.port, 10) : CALLBACK_PORT;
  const expectedPath = redirectUrl.pathname || CALLBACK_PATH;

  return new Promise<{ code: string; state: string }>((resolve, reject) => {
    let timeout: NodeJS.Timeout | null = null;

    const server = createServer((req, res) => {
      try {
        const requestUrl = new URL(req.url ?? "/", `http://${hostname}:${port}`);
        if (requestUrl.pathname !== expectedPath) {
          res.statusCode = 404;
          res.setHeader("Content-Type", "text/plain");
          res.end("Not found");
          return;
        }

        const error = requestUrl.searchParams.get("error");
        const code = requestUrl.searchParams.get("code")?.trim();
        const state = requestUrl.searchParams.get("state")?.trim();

        if (error) {
          // Only process error responses with a valid state to prevent CSRF
          // (e.g., <img src="http://localhost:8080/callback?error=access_denied">)
          if (!state || state !== params.expectedState) {
            res.statusCode = 400;
            res.setHeader("Content-Type", "text/plain");
            res.end("Invalid or missing state parameter");
            return;
          }
          res.statusCode = 400;
          res.setHeader("Content-Type", "text/plain");
          const desc = requestUrl.searchParams.get("error_description") ?? error;
          res.end(`Authentication failed: ${desc}`);
          finish(new Error(`OAuth error from Microsoft: ${desc}`));
          return;
        }

        if (!code || !state) {
          res.statusCode = 400;
          res.setHeader("Content-Type", "text/plain");
          res.end("Missing code or state");
          finish(new Error("Missing OAuth code or state in callback"));
          return;
        }

        if (state !== params.expectedState) {
          res.statusCode = 400;
          res.setHeader("Content-Type", "text/plain");
          res.end("Invalid state");
          finish(new Error("OAuth state mismatch"));
          return;
        }

        res.statusCode = 200;
        res.setHeader("Content-Type", "text/html; charset=utf-8");
        res.end(
          "<!doctype html><html><head><meta charset='utf-8'/></head>" +
            "<body><h2>Microsoft 365 OAuth complete</h2>" +
            "<p>You can close this window and return to OpenClaw.</p></body></html>",
        );
        finish(undefined, { code, state });
      } catch (err) {
        finish(err instanceof Error ? err : new Error("OAuth callback failed"));
      }
    });

    const finish = (err?: Error, result?: { code: string; state: string }) => {
      if (timeout) {
        clearTimeout(timeout);
        timeout = null;
      }
      try {
        server.close();
      } catch {
        // ignore close errors
      }
      if (err) {
        reject(err);
      } else if (result) {
        resolve(result);
      }
    };

    server.once("error", (err) => {
      finish(err instanceof Error ? err : new Error("OAuth callback server error"));
    });

    server.listen(port, hostname, () => {
      params.onProgress?.(`Waiting for OAuth callback on ${params.redirectUri}…`);
    });

    timeout = setTimeout(() => {
      finish(new Error("OAuth callback timeout. Run `openclaw auth` to try again."));
    }, params.timeoutMs);
  });
}

// ── Token exchange ──────────────────────────────────────────────────────────

async function fetchWithTimeout(url: string, init: RequestInit): Promise<Response> {
  const { response, release } = await fetchWithSsrFGuard({
    url,
    init,
    timeoutMs: FETCH_TIMEOUT_MS,
  });
  try {
    const body = await response.arrayBuffer();
    return new Response(body, {
      status: response.status,
      statusText: response.statusText,
      headers: response.headers,
    });
  } finally {
    await release();
  }
}

export async function exchangeCodeForTokens(
  code: string,
  verifier: string,
  config: Office365Config,
): Promise<Office365Credential> {
  const redirectUri = config.redirectUri || REDIRECT_URI;
  const body = new URLSearchParams({
    client_id: config.clientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
  });
  if (config.clientSecret) {
    body.set("client_secret", config.clientSecret);
  }

  const response = await fetchWithTimeout(tokenUrl(config.tenantId), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`Token exchange failed (${response.status}): ${errorText}`);
  }

  const data = (await response.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  if (!data.refresh_token) {
    throw new Error(
      "No refresh token received. Ensure 'offline_access' is in your scopes and re-consent.",
    );
  }

  const email = await getUserInfo(data.access_token);
  const expiresAt = Date.now() + data.expires_in * 1000 - EXPIRES_BUFFER_MS;

  return {
    access: data.access_token,
    refresh: data.refresh_token,
    expires: Math.max(expiresAt, Date.now() + 30_000),
    email,
  };
}

// ── Token refresh ───────────────────────────────────────────────────────────

export async function refreshMicrosoftTokens(
  refreshToken: string,
  config: Office365Config,
): Promise<Office365Credential> {
  const body = new URLSearchParams({
    client_id: config.clientId,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: config.scopes.join(" "),
  });
  if (config.clientSecret) {
    body.set("client_secret", config.clientSecret);
  }

  const response = await fetchWithTimeout(tokenUrl(config.tenantId), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    if (
      response.status === 400 &&
      (errorText.includes("invalid_grant") ||
        errorText.includes("AADSTS700082") ||
        errorText.includes("AADSTS50076"))
    ) {
      throw new Error(
        "Refresh token expired or revoked. Run `openclaw auth` to re-authenticate with Microsoft 365.",
      );
    }
    throw new Error(`Token refresh failed (${response.status}): ${errorText}`);
  }

  const data = (await response.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  const expiresAt = Date.now() + data.expires_in * 1000 - EXPIRES_BUFFER_MS;
  return {
    access: data.access_token,
    refresh: data.refresh_token ?? refreshToken, // RFC 6749 §6: new refresh_token is optional
    expires: Math.max(expiresAt, Date.now() + 30_000),
  };
}

// ── User info ───────────────────────────────────────────────────────────────

export async function getUserInfo(accessToken: string): Promise<string | undefined> {
  try {
    const response = await fetchWithTimeout(GRAPH_ME_URL, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (response.ok) {
      const data = (await response.json()) as {
        mail?: string;
        userPrincipalName?: string;
        displayName?: string;
      };
      return data.mail || data.userPrincipalName;
    }
  } catch {
    // non-critical — we can proceed without email
  }
  return undefined;
}

// ── Login orchestrator ──────────────────────────────────────────────────────

export type MicrosoftOAuthContext = {
  isRemote: boolean;
  openUrl: (url: string) => Promise<void>;
  log: (msg: string) => void;
  note: (message: string, title?: string) => Promise<void>;
  prompt: (message: string) => Promise<string>;
  progress: { update: (msg: string) => void; stop: (msg?: string) => void };
};

function shouldUseManualFlow(isRemote: boolean): boolean {
  if (isRemote) {
    return true;
  }
  // WSL2 cannot open Windows browser reliably in all setups
  if (process.platform === "linux" && !process.env.DISPLAY && !process.env.WAYLAND_DISPLAY) {
    return true;
  }
  return false;
}

export async function loginMicrosoftOAuth(
  ctx: MicrosoftOAuthContext,
  config: Office365Config,
): Promise<Office365Credential> {
  const needsManual = shouldUseManualFlow(ctx.isRemote);
  const redirectUri = config.redirectUri || REDIRECT_URI;

  await ctx.note(
    needsManual
      ? [
          "You are running in a remote/VPS environment.",
          "A URL will be shown for you to open in your LOCAL browser.",
          "After signing in, copy the full redirect URL from the browser address bar and paste it back here.",
          "",
          "Tip: The redirect will fail to load (that's expected) — just copy the URL.",
        ].join("\n")
      : [
          "Your browser will open for Microsoft authentication.",
          "Sign in with your Microsoft 365 account.",
          `The callback will be captured automatically on ${redirectUri}.`,
        ].join("\n"),
    "Microsoft 365 OAuth",
  );

  const { verifier, challenge } = generatePkce();
  const state = randomBytes(16).toString("hex");
  const authUrl = buildAuthorizeUrl(config, challenge, state);

  if (needsManual) {
    ctx.progress.update("OAuth URL ready");
    ctx.log(`\nOpen this URL in your LOCAL browser:\n\n${authUrl}\n`);
    ctx.progress.update("Waiting for you to paste the callback URL…");
    const callbackInput = await ctx.prompt("Paste the full redirect URL here: ");
    const parsed = parseCallbackInput(callbackInput, state);
    if ("error" in parsed) {
      throw new Error(parsed.error);
    }
    ctx.progress.update("Exchanging authorization code for tokens…");
    return exchangeCodeForTokens(parsed.code, verifier, config);
  }

  // Local flow: open browser + listen for callback
  ctx.progress.update("Complete sign-in in browser…");
  try {
    await ctx.openUrl(authUrl);
  } catch {
    ctx.log(`\nOpen this URL in your browser:\n\n${authUrl}\n`);
  }

  try {
    const { code } = await waitForLocalCallback({
      redirectUri,
      expectedState: state,
      timeoutMs: OAUTH_TIMEOUT_MS,
      onProgress: (msg) => ctx.progress.update(msg),
    });
    ctx.progress.update("Exchanging authorization code for tokens…");
    return exchangeCodeForTokens(code, verifier, config);
  } catch (err) {
    // If callback server fails (port in use, etc.), fall back to manual flow
    if (
      err instanceof Error &&
      (err.message.includes("EADDRINUSE") ||
        err.message.includes("port") ||
        err.message.includes("listen"))
    ) {
      ctx.progress.update("Local callback server failed. Switching to manual mode…");
      ctx.log(`\nOpen this URL in your LOCAL browser:\n\n${authUrl}\n`);
      const callbackInput = await ctx.prompt("Paste the full redirect URL here: ");
      const parsed = parseCallbackInput(callbackInput, state);
      if ("error" in parsed) {
        throw new Error(parsed.error, { cause: err });
      }
      ctx.progress.update("Exchanging authorization code for tokens…");
      return exchangeCodeForTokens(parsed.code, verifier, config);
    }
    throw err;
  }
}
