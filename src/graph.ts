const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
const REQUEST_SPACING_MS = 1250; // Minimum spacing per endpoint
const GLOBAL_REQUEST_SPACING_MS = 65; // Minimum spacing across all endpoints
const MAX_RETRIES = 3;
const MAX_RETRY_WAIT_MS = 120_000;

interface GraphDependencies {
  fetch: typeof fetch;
  now: () => number;
  wait: (milliseconds: number) => Promise<void>;
}

export function graphUrl(path: string, query: Record<string, string> = {}): string {
  const url = new URL(`${GRAPH_ROOT}${path}`);
  url.search = new URLSearchParams(query).toString();
  return url.toString();
}

function validateGraphUrl(address: string): URL {
  const url = new URL(address);
  if (
    url.origin !== 'https://graph.microsoft.com' || url.username || url.password || url.hash ||
    !/^\/v1\.0\/(?:me\/chats|chats\/[^/]+\/messages)$/.test(url.pathname)
  ) {
    throw new Error('Refused an unexpected Graph URL. No credential was sent.');
  }
  return url;
}

export function validateNextLink(address: string, previousAddress: string): string {
  const next = validateGraphUrl(address);
  if (next.pathname !== validateGraphUrl(previousAddress).pathname) {
    throw new Error('Graph pagination changed the requested resource; read stopped.');
  }
  return next.toString();
}

export class GraphClient {
  // Request queue ensures sequential execution; both callbacks resolve to allow one failure to not block subsequent requests.
  private queue: Promise<void> = Promise.resolve();
  private readonly lastRequestEpochMs = new Map<string, number>();
  private lastGlobalRequestEpochMs = -Infinity;
  private readonly dependencies: GraphDependencies;

  constructor(
    private readonly accessToken: () => Promise<string>,
    dependencies: Partial<GraphDependencies> = {},
  ) {
    this.dependencies = {
      fetch,
      now: Date.now,
      wait: (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)),
      ...dependencies,
    };
  }

  async get(address: string): Promise<unknown> {
    const url = validateGraphUrl(address);
    const request = this.queue.then(() => this.request(url));
    this.queue = request.then(() => undefined, () => undefined);
    return request;
  }

  private async request(url: URL): Promise<unknown> {
    const { now, wait } = this.dependencies;
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      // Refresh token before each attempt to avoid reusing expired credentials across retry delays.
      const token = await this.accessToken();
      const nextRequestEpochMs = Math.max(
        (this.lastRequestEpochMs.get(url.pathname) ?? -Infinity) + REQUEST_SPACING_MS,
        this.lastGlobalRequestEpochMs + GLOBAL_REQUEST_SPACING_MS,
      );
      const delayMs = nextRequestEpochMs - now();
      if (delayMs > 0) await wait(delayMs);
      this.lastRequestEpochMs.set(url.pathname, now());
      this.lastGlobalRequestEpochMs = now();

      let response: Response;
      try {
        response = await this.dependencies.fetch(url.toString(), {
          method: 'GET',
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
          redirect: 'error',
          signal: AbortSignal.timeout(30_000),
        });
      } catch {
        throw new Error('Graph request failed or timed out. Check connectivity and retry.');
      }

      if (response.status === 429 || response.status === 503) {
        const retryAfter = response.headers.get('retry-after');
        await response.body?.cancel();
        if (attempt === MAX_RETRIES) {
          throw new Error('Graph remains busy after retries. Try again later.');
        }
        const retryDelayMs = this.parseRetryDelay(retryAfter, now());
        if (retryDelayMs > MAX_RETRY_WAIT_MS) {
          throw new Error('Graph requested a retry delay longer than two minutes. Try again later.');
        }
        await wait(retryDelayMs);
        continue;
      }

      if (!response.ok) {
        await response.body?.cancel();
        if (response.status === 401) {
          throw new Error('Graph rejected the credential. Run teams_auth_login to sign in again.');
        }
        if (response.status === 403) {
          throw new Error('Graph denied this read. Check Chat.Read consent and access to the conversation.');
        }
        if (response.status === 404) throw new Error('Conversation not found or no longer accessible.');
        throw new Error(`Graph read failed with HTTP ${response.status}. Try again later.`);
      }

      try {
        return await response.json();
      } catch {
        throw new Error('Graph returned an invalid JSON response. Try again later.');
      }
    }
    throw new Error('Graph read exhausted its retry budget.');
  }

  private parseRetryDelay(retryAfter: string | null, nowEpochMs: number): number {
    if (retryAfter === null) {
      return 1000 * 2 ** (Math.min(3, Math.floor(Math.random() * 4)));
    }
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) {
      return Math.max(0, seconds * 1000);
    }
    const dateDelayMs = Date.parse(retryAfter) - nowEpochMs;
    if (Number.isFinite(dateDelayMs)) {
      return Math.max(0, dateDelayMs);
    }
    return 1000 * 2;
  }
}