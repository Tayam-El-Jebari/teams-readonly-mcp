import { mkdir, readFile, rename, writeFile, chmod, stat } from 'node:fs/promises';
import { dirname } from 'node:path';
import * as z from 'zod';
import type { Config } from './config.js';

const ExpirySeconds = z.number().positive().refine((seconds) =>
  Number.isFinite(new Date(Date.now() + seconds * 1000).getTime()),
);
const TokenResponseSchema = z.object({
  access_token: z.string().min(1),
  refresh_token: z.string().min(1).optional(),
  expires_in: ExpirySeconds.optional(),
  scope: z.string().optional(),
});
type TokenResponse = z.infer<typeof TokenResponseSchema>;

const StoredTokenSchema = TokenResponseSchema.extend({
  expires_at: z.number().refine((value) => Number.isFinite(new Date(value).getTime())),
}).transform(({ expires_at, ...token }) => ({ ...token, expiresAtEpochMs: expires_at }));
type StoredToken = z.infer<typeof StoredTokenSchema>;

export interface PendingLogin {
  readonly deviceCode: string;
  readonly userCode: string;
  readonly verificationUri: string;
  readonly expiresAtEpochMs: number;
  readonly intervalSeconds: number;
}

/** Non-secret facts about the stored credential. Safe to return from a tool. */
export interface CredentialFacts {
  authenticated: boolean;
  account?: string;
  scopes: string[];
  /** ISO timestamp of access-token expiry, absent when unauthenticated. */
  expiresAt?: string;
  /** True when the access token is expired or within the renewal margin. */
  refreshDue: boolean;
  detail: string;
}

const RENEW_MARGIN_MS = 5 * 60 * 1000;

// Tool errors enter model context and logs, so this function exists to redact 
// recognizable token formats.
export function redactSensitiveData(text: string): string {
  return text
    .replace(/eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]+/g, '[redacted-jwt]')
    .replace(/("(?:access|refresh|id)_token"\s*:\s*")[^"]+"/g, '$1[redacted]"')
    .replace(/\b(0\.[A-Za-z0-9_.\-~]{20,})/g, '[redacted-token]')
    .replace(/(Bearer\s+)\S+/gi, '$1[redacted]');
}

function fail(message: string): never {
  throw new Error(redactSensitiveData(message));
}

export function scopesFromAccessToken(accessToken: string): string[] {
  const scope = claimsFromAccessToken(accessToken)?.scp;
  return typeof scope === 'string' ? scope.split(' ').filter(Boolean).sort() : [];
}

function accountFromAccessToken(accessToken: string): string | undefined {
  const claims = claimsFromAccessToken(accessToken);
  if (typeof claims?.upn === 'string') return claims.upn;
  if (typeof claims?.preferred_username === 'string') return claims.preferred_username;
  return undefined;
}

// Claims are unverified and must only be used for display.
function claimsFromAccessToken(accessToken: string): Record<string, unknown> | undefined {
  const payload = accessToken.split('.')[1];
  if (!payload) return undefined;
  try {
    const claims: unknown = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    const parsed = z.record(z.string(), z.unknown()).safeParse(claims);
    return parsed.success ? parsed.data : undefined;
  } catch {
    return undefined;
  }
}

// Rename prevents a crash mid-write from leaving a truncated credential.
async function persist(cfg: Config, token: TokenResponse): Promise<StoredToken> {
  const stored = {
    ...token,
    expires_at: Date.now() + (token.expires_in ?? 3600) * 1000,
  };
  const dir = dirname(cfg.tokenPath);
  await mkdir(dir, { recursive: true, mode: 0o700 });
  await chmod(dir, 0o700);
  const tmp = `${cfg.tokenPath}.tmp`;
  await writeFile(tmp, JSON.stringify(stored), { mode: 0o600 });
  await rename(tmp, cfg.tokenPath);
  return StoredTokenSchema.parse(stored);
}

async function readStored(cfg: Config): Promise<StoredToken | undefined> {
  try {
    const raw: unknown = JSON.parse(await readFile(cfg.tokenPath, 'utf8'));
    const parsed = StoredTokenSchema.safeParse(raw);
    return parsed.success ? parsed.data : undefined;
  } catch {
    // Missing or corrupt credentials require sign-in again.
    return undefined;
  }
}

export async function tokenFileMode(cfg: Config): Promise<string | undefined> {
  try {
    return ((await stat(cfg.tokenPath)).mode & 0o777).toString(8).padStart(3, '0');
  } catch {
    return undefined;
  }
}

function authority(cfg: Config, leaf: string): string {
  return `https://login.microsoftonline.com/${cfg.tenantId}/oauth2/v2.0/${leaf}`;
}

function parseTokenResponse(body: unknown): TokenResponse {
  const parsed = TokenResponseSchema.safeParse(body);
  if (!parsed.success) fail('token endpoint returned an invalid credential');
  return parsed.data;
}

async function postForm(url: string, body: Record<string, string>): Promise<unknown> {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(body).toString(),
  });
  let response: unknown;
  try {
    response = await res.json();
  } catch {
    fail(`${url} returned ${res.status} with a non-JSON body`);
  }
  if (!res.ok && !errorOf(response).error) {
    fail(`${url} returned HTTP ${res.status}`);
  }
  return response;
}

function errorOf(response: unknown): { error?: string; description?: string } {
  if (typeof response !== 'object' || response === null) return {};
  const r = response as { error?: unknown; error_description?: unknown };
  return {
    ...(typeof r.error === 'string' ? { error: r.error } : {}),
    ...(typeof r.error_description === 'string'
      ? { description: r.error_description.split('\n')[0] }
      : {}),
  };
}

export async function beginLogin(cfg: Config): Promise<PendingLogin> {
  const body = await postForm(authority(cfg, 'devicecode'), {
    client_id: cfg.clientId,
    scope: cfg.scope,
  });
  const { error, description } = errorOf(body);
  if (error) fail(`device code request refused (${error}): ${description ?? 'no detail'}`);

  const parsed = z.object({
    device_code: z.string().min(1),
    user_code: z.string().min(1),
    verification_uri: z.string().url(),
    expires_in: ExpirySeconds.default(900),
    interval: z.number().positive().max(2147483).default(5),
  }).safeParse(body);
  if (!parsed.success) {
    fail('device code endpoint returned an invalid response');
  }
  const deviceCode = parsed.data;
  return {
    deviceCode: deviceCode.device_code,
    userCode: deviceCode.user_code,
    verificationUri: deviceCode.verification_uri,
    expiresAtEpochMs: Date.now() + deviceCode.expires_in * 1000,
    intervalSeconds: deviceCode.interval,
  };
}

export async function completeLogin(
  cfg: Config,
  pending: PendingLogin,
  budgetMs: number,
): Promise<{ done: true; facts: CredentialFacts } | { done: false; detail: string }> {
  const deadlineEpochMs = Math.min(Date.now() + budgetMs, pending.expiresAtEpochMs);
  let waitMs = pending.intervalSeconds * 1000;

  while (Date.now() < deadlineEpochMs) {
    const body = await postForm(authority(cfg, 'token'), {
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: cfg.clientId,
      device_code: pending.deviceCode,
    });
    const { error, description } = errorOf(body);

    if (!error) {
      const stored = await persist(cfg, parseTokenResponse(body));
      return { done: true, facts: factsFrom(stored) };
    }
    switch (error) {
      case 'authorization_pending':
        break;
      case 'slow_down':
        waitMs += 5000;
        break;
      default:
        fail(`sign-in failed (${error}): ${description ?? 'no detail'}`);
    }
    const remainingMs = deadlineEpochMs - Date.now();
    if (remainingMs <= 0) break;
    await new Promise((resolve) => setTimeout(resolve, Math.min(waitMs, remainingMs)));
  }
  return {
    done: false,
    detail:
      `Sign-in not completed yet. Enter code ${pending.userCode} at ${pending.verificationUri}, ` +
      `then call teams_auth_complete again.`,
  };
}

// Save a rotated refresh token when supplied; otherwise retain the previous one.
async function refresh(cfg: Config, stored: StoredToken): Promise<StoredToken> {
  if (!stored.refresh_token) {
    fail('no refresh_token stored. Run teams_auth_login to sign in again.');
  }
  const body = await postForm(authority(cfg, 'token'), {
    grant_type: 'refresh_token',
    client_id: cfg.clientId,
    refresh_token: stored.refresh_token,
    scope: cfg.scope,
  });
  const { error, description } = errorOf(body);
  if (error) {
    fail(
      `refresh failed (${error}): ${description ?? 'no detail'}. ` +
        `The refresh token is likely expired or revoked (a password change or SSPR revokes it). ` +
        `Run teams_auth_login to sign in again.`,
    );
  }
  const token = parseTokenResponse(body);
  return persist(cfg, { ...token, refresh_token: token.refresh_token ?? stored.refresh_token });
}

// Reads must never start interactive sign-in, which would stall unattended runs.
export async function accessTokenForRead(cfg: Config): Promise<string> {
  const stored = await readStored(cfg);
  if (!stored) {
    fail('not authenticated. Run teams_auth_login first (this tool will not start a sign-in).');
  }
  const fresh = isRefreshDue(stored) ? await refresh(cfg, stored) : stored;
  return fresh.access_token;
}

function isRefreshDue(stored: StoredToken): boolean {
  return stored.expiresAtEpochMs - RENEW_MARGIN_MS <= Date.now();
}

function factsFrom(stored: StoredToken): CredentialFacts {
  const account = accountFromAccessToken(stored.access_token);
  const refreshDue = isRefreshDue(stored);
  return {
    authenticated: true,
    ...(account ? { account } : {}),
    scopes: scopesFromAccessToken(stored.access_token),
    expiresAt: new Date(stored.expiresAtEpochMs).toISOString(),
    refreshDue,
    detail: refreshDue
      ? 'Access token is expired or expiring; it will be refreshed on the next read.'
      : 'Credential is valid.',
  };
}

/** Non-secret status of the stored credential. Never touches the network. */
export async function credentialFacts(cfg: Config): Promise<CredentialFacts> {
  const stored = await readStored(cfg);
  if (!stored) {
    return {
      authenticated: false,
      scopes: [],
      refreshDue: false,
      detail: 'No credential stored. Run teams_auth_login to sign in.',
    };
  }
  return factsFrom(stored);
}
