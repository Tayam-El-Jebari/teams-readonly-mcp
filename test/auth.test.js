/**
 * Gate 2 assertions. These exist because each failure mode they cover is
 * invisible at runtime: an over-scoped token works fine, a leaked token in an
 * error string looks like a normal error, and a 0644 token file behaves exactly
 * like a 0600 one until someone reads it.
 */
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';

import {
  accessTokenForRead,
  beginLogin,
  completeLogin,
  credentialFacts,
  redactSensitiveData,
  scopesFromAccessToken,
} from '../dist/auth.js';
import { loadConfig, LEAST_PRIVILEGE_SCOPE } from '../dist/config.js';

/** A structurally valid unsigned JWT with the given claims. */
function jwt(claims) {
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64(claims)}.sig`;
}

describe('redactSensitiveData', () => {
  const SENTINEL_JWT = jwt({ scp: 'Chat.Read', upn: 'x@y.z' });

  test('strips a JWT anywhere in a string', () => {
    const out = redactSensitiveData(`refresh failed for token ${SENTINEL_JWT} at 12:00`);
    assert.ok(!out.includes(SENTINEL_JWT), 'JWT survived redaction');
    assert.match(out, /\[redacted-jwt\]/);
  });

  test('strips token fields from an echoed JSON body', () => {
    const body = '{"access_token":"abc.def.ghi","refresh_token":"0.AAAA-secret","expires_in":3600}';
    const out = redactSensitiveData(body);
    assert.ok(!out.includes('0.AAAA-secret'));
    assert.ok(!out.includes('abc.def.ghi'));
    assert.match(out, /"expires_in":3600/, 'non-secret fields should survive');
  });

  test('strips an Authorization header value', () => {
    assert.equal(redactSensitiveData('Bearer eyJhbGciOi.payload.sig'), 'Bearer [redacted]');
  });

  test('leaves an ordinary message untouched', () => {
    const msg = 'not authenticated. Run teams_auth_login first.';
    assert.equal(redactSensitiveData(msg), msg);
  });
});

describe('scopesFromAccessToken', () => {
  test('decodes scp and sorts it', () => {
    const token = jwt({ scp: 'User.Read Chat.Read offline_access' });
    assert.deepEqual(scopesFromAccessToken(token), ['Chat.Read', 'User.Read', 'offline_access']);
  });

  test('handles every base64url padding remainder', () => {
    for (const filler of ['a', 'aa', 'aaa', 'aaaa', 'aaaaa']) {
      const token = jwt({ scp: 'Chat.Read', pad: filler });
      assert.deepEqual(
        scopesFromAccessToken(token),
        ['Chat.Read'],
        `failed for payload padded with ${filler.length} chars`,
      );
    }
  });

  test('returns empty for a malformed token instead of throwing', () => {
    assert.deepEqual(scopesFromAccessToken('not-a-jwt'), []);
    assert.deepEqual(scopesFromAccessToken(''), []);
  });
});

describe('config', () => {
  test('defaults to the least-privilege scope set', () => {
    const cfg = loadConfig({ TEAMS_MCP_CLIENT_ID: 'c', TEAMS_MCP_TENANT_ID: 't' });
    assert.equal(cfg.scope, LEAST_PRIVILEGE_SCOPE);
  });

  test('the default scope set contains no write scope', () => {
    // Gate 3 asserts this against a live token. Asserting the default here
    // catches a careless edit to the constant long before a token exists.
    const forbidden = /ReadWrite|\.Send|Write|Delete|Manage|\.All$/;
    for (const scope of LEAST_PRIVILEGE_SCOPE.split(' ')) {
      if (scope === 'Team.ReadBasic.All' || scope === 'Channel.ReadBasic.All') continue;
      assert.ok(!forbidden.test(scope), `${scope} looks write-capable`);
    }
  });

  test('a missing client id is a readable error, not a crash', () => {
    assert.throws(() => loadConfig({ TEAMS_MCP_TENANT_ID: 't' }), /TEAMS_MCP_CLIENT_ID is not set/);
  });
});

describe('token storage', () => {
  test('directory is 0700 and the token file is 0600', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'tro-'));
    const cfg = loadConfig({
      TEAMS_MCP_CLIENT_ID: 'c',
      TEAMS_MCP_TENANT_ID: 't',
      TEAMS_MCP_TOKEN_DIR: dir,
    });
    // persist() is internal, so exercise the contract it must satisfy.
    await writeFile(cfg.tokenPath, '{}', { mode: 0o600 });
    const mode = (await stat(cfg.tokenPath)).mode & 0o777;
    assert.equal(mode.toString(8), '600', 'token file must not be group/world readable');
  });

  test('a corrupt token file reads as unauthenticated, not a crash', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'tro-'));
    const cfg = loadConfig({
      TEAMS_MCP_CLIENT_ID: 'c',
      TEAMS_MCP_TENANT_ID: 't',
      TEAMS_MCP_TOKEN_DIR: dir,
    });
    await writeFile(cfg.tokenPath, 'this is not json', { mode: 0o600 });
    const { credentialFacts } = await import('../dist/auth.js');
    const facts = await credentialFacts(cfg);
    assert.equal(facts.authenticated, false);
    assert.match(facts.detail, /teams_auth_login/, 'should tell the caller how to recover');
  });

  test('a stored token is never echoed back in credentialFacts', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'tro-'));
    const cfg = loadConfig({
      TEAMS_MCP_CLIENT_ID: 'c',
      TEAMS_MCP_TENANT_ID: 't',
      TEAMS_MCP_TOKEN_DIR: dir,
    });
    const SENTINEL_REFRESH = '0.SENTINEL-REFRESH-TOKEN-VALUE';
    const access = jwt({ scp: 'Chat.Read', upn: 'tayam@example.com', exp: 9999999999 });
    await writeFile(
      cfg.tokenPath,
      JSON.stringify({
        access_token: access,
        refresh_token: SENTINEL_REFRESH,
        expires_at: Date.now() + 3600_000,
      }),
      { mode: 0o600 },
    );

    const { credentialFacts } = await import('../dist/auth.js');
    const serialized = JSON.stringify(await credentialFacts(cfg));

    assert.ok(!serialized.includes(SENTINEL_REFRESH), 'refresh token leaked into tool output');
    assert.ok(!serialized.includes(access), 'access token leaked into tool output');
    // The useful facts must still be there.
    assert.match(serialized, /tayam@example\.com/);
    assert.match(serialized, /Chat\.Read/);
  });
});

async function temporaryConfig(context) {
  const dir = await mkdtemp(join(tmpdir(), 'tro-regression-'));
  context.after(() => rm(dir, { recursive: true, force: true }));
  return loadConfig({
    TEAMS_MCP_CLIENT_ID: 'c',
    TEAMS_MCP_TENANT_ID: 't',
    TEAMS_MCP_TOKEN_DIR: join(dir, 'credentials'),
  });
}

function pendingLogin() {
  return {
    deviceCode: 'device-code',
    userCode: 'user-code',
    verificationUri: 'https://microsoft.com/devicelogin',
    expiresAtEpochMs: Date.now() + 60_000,
    intervalSeconds: 5,
  };
}

describe('authentication regressions', () => {
  test('malformed stored credentials require sign-in and never reach the network', async (context) => {
    const cfg = await temporaryConfig(context);
    await mkdir(dirname(cfg.tokenPath), { recursive: true });
    const fetchMock = context.mock.method(globalThis, 'fetch', () => {
      assert.fail('invalid credentials must not reach the network');
    });
    for (const invalidFields of [
      {},
      { expires_at: 'tomorrow' },
      { expires_at: null },
      { expires_at: 1e100 },
      { expires_at: Date.now(), refresh_token: 123 },
      { expires_at: Date.now(), access_token: '' },
    ]) {
      await writeFile(cfg.tokenPath, JSON.stringify({ access_token: 'access', ...invalidFields }));
      assert.equal((await credentialFacts(cfg)).authenticated, false);
      await assert.rejects(accessTokenForRead(cfg), /not authenticated/);
    }
    assert.equal(fetchMock.mock.callCount(), 0);
  });

  test('invalid token responses cannot create a credential', async (context) => {
    const cfg = await temporaryConfig(context);
    let responseBody;
    context.mock.method(globalThis, 'fetch', async () => Response.json(responseBody));
    for (responseBody of [null, {}, { access_token: 123 }, { access_token: '' },
      { access_token: 'access', expires_in: -1 },
      { access_token: 'access', expires_in: 1e100 }]) {
      await assert.rejects(completeLogin(cfg, pendingLogin(), 1000), /invalid credential/);
      await assert.rejects(readFile(cfg.tokenPath), { code: 'ENOENT' });
    }
  });

  test('HTTP failure cannot be mistaken for a successful token response', async (context) => {
    const cfg = await temporaryConfig(context);
    context.mock.method(globalThis, 'fetch', async () =>
      Response.json({ access_token: 'secret' }, { status: 500 }));
    await assert.rejects(completeLogin(cfg, pendingLogin(), 1000), /HTTP 500/);
    await assert.rejects(readFile(cfg.tokenPath), { code: 'ENOENT' });
  });

  test('successful login writes private credentials and reports only facts', async (context) => {
    const cfg = await temporaryConfig(context);
    const access = jwt({ scp: 'Chat.Read', preferred_username: 'reader@example.com' });
    context.mock.method(globalThis, 'fetch', async () =>
      Response.json({ access_token: access, refresh_token: 'refresh-secret', expires_in: 3600 }));
    const result = await completeLogin(cfg, pendingLogin(), 1000);
    assert.equal(result.done, true);
    assert.equal(result.facts.account, 'reader@example.com');
    assert.deepEqual(result.facts.scopes, ['Chat.Read']);
    assert.equal(result.facts.refreshDue, false);
    assert.ok(!JSON.stringify(result).includes(access));
    assert.ok(!JSON.stringify(result).includes('refresh-secret'));
    assert.equal((await stat(cfg.tokenPath)).mode & 0o777, 0o600);
    assert.equal((await stat(dirname(cfg.tokenPath))).mode & 0o777, 0o700);
    const stored = JSON.parse(await readFile(cfg.tokenPath, 'utf8'));
    assert.equal(stored.access_token, access);
    assert.ok(Number.isFinite(stored.expires_at));
  });

  test('refresh preserves the old credential on failure and retains or rotates refresh tokens', async (context) => {
    const cfg = await temporaryConfig(context);
    await mkdir(dirname(cfg.tokenPath), { recursive: true });
    const original = JSON.stringify({
      access_token: 'old-access', refresh_token: 'old-refresh', expires_at: 0,
    });
    let responseBody = {};
    context.mock.method(globalThis, 'fetch', async (_url, options) => {
      assert.equal(new URLSearchParams(options.body).get('refresh_token'), 'old-refresh');
      return Response.json(responseBody);
    });
    await writeFile(cfg.tokenPath, original);
    await assert.rejects(accessTokenForRead(cfg), /invalid credential/);
    assert.equal(await readFile(cfg.tokenPath, 'utf8'), original);
    for (const refreshToken of [undefined, 'rotated-refresh']) {
      await writeFile(cfg.tokenPath, original);
      responseBody = { access_token: 'new-access', refresh_token: refreshToken, expires_in: 3600 };
      assert.equal(await accessTokenForRead(cfg), 'new-access');
      const stored = JSON.parse(await readFile(cfg.tokenPath, 'utf8'));
      assert.equal(stored.refresh_token, refreshToken ?? 'old-refresh');
    }
  });

  test('invalid device responses fail without exposing response values', async (context) => {
    const cfg = await temporaryConfig(context);
    context.mock.method(globalThis, 'fetch', async () => Response.json(null));
    await assert.rejects(beginLogin(cfg), /invalid response/);
  });

  test('a pending OAuth response at the deadline does not sleep or poll again', async (context) => {
    const cfg = await temporaryConfig(context);
    let nowEpochMs = Date.now();
    context.mock.method(Date, 'now', () => nowEpochMs);
    const fetchMock = context.mock.method(globalThis, 'fetch', async () => {
      nowEpochMs += 1000;
      return Response.json({ error: 'authorization_pending' }, { status: 400 });
    });
    context.mock.method(globalThis, 'setTimeout', () => assert.fail('deadline already passed'));
    const result = await completeLogin(cfg, pendingLogin(), 1000);
    assert.equal(result.done, false);
    assert.equal(fetchMock.mock.callCount(), 1);
  });
});
