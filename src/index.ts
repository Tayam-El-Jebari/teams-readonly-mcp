#!/usr/bin/env node
/**
 * teams-readonly-mcp · read-only MCP server over my own Microsoft Teams chats.
 *
 * Gate 1 scaffolding: transport, one tool, no credential, no Graph calls.
 * The full tool surface and the gate sequence live in ../plan.md.
 *
 * Two invariants hold for every tool added to this file:
 *   1. No write tools. There is no Graph call here with a method other than GET,
 *      and no tool that mutates anything. See plan.md section 5.
 *   2. Token material never crosses the tool boundary. Tools report facts about
 *      the credential, never the credential.
 */
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import * as z from 'zod';

const NAME = 'teams-readonly-mcp';
const VERSION = '0.1.0';

/**
 * What the server is willing to say about the stored credential. Deliberately
 * describes the token rather than carrying it: no access_token, no
 * refresh_token, no Authorization header is representable in this shape.
 */
const AuthStatus = z.object({
  authenticated: z.boolean(),
  // Optional, not nullable. Both z.string().nullable() and
  // z.union([z.string(), z.null()]) serialize to type: ["string","null"], and
  // several MCP clients read `type` as a single string and either reject the
  // tool or silently drop the constraint. Absence carries the same meaning here
  // because `authenticated` is the flag, so an optional single-typed property is
  // both portable and a cleaner contract.
  account: z
    .string()
    .optional()
    .describe('UPN of the signed-in user. Absent when no credential is stored'),
  scopes: z.array(z.string()).describe('Scopes granted to the stored access token'),
  detail: z.string().describe('What to do next, in one sentence'),
});

function createServer(): McpServer {
  const server = new McpServer({ name: NAME, version: VERSION }, { capabilities: { tools: {} } });

  server.registerTool(
    'teams_auth_status',
    {
      title: 'Teams auth status',
      description:
        'Report whether a Microsoft Graph credential is stored and what it is allowed to read. ' +
        'Call this before any other tool: it is the precondition check, and it is what a ' +
        'scheduled briefing should use to decide whether to proceed or fail loudly. ' +
        'Returns facts about the token (account, granted scopes) and never the token itself. ' +
        'Does not trigger a login. If it reports unauthenticated, run teams_auth_login.',
      outputSchema: AuthStatus,
      // Advisory only: the spec tells clients to distrust annotations and the SDK
      // never gates on them. Set explicitly anyway, because the defaults are
      // pessimistic (destructiveHint and openWorldHint both default to true) and
      // readOnlyHint is what lets a host auto-approve without a prompt.
      // destructiveHint and idempotentHint are meaningful only when readOnlyHint
      // is false, so they are omitted.
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    () => {
      // Gate 1 has no credential store. Gate 2 replaces this with a real read.
      const status: z.infer<typeof AuthStatus> = {
        authenticated: false,
        scopes: [],
        detail:
          'No credential store yet (Gate 1 scaffolding). Gate 2 implements login and token storage.',
      };
      return {
        // Both shapes: structuredContent for callers that can use it, serialized
        // JSON in a text block for backwards compatibility, as the spec asks.
        content: [{ type: 'text' as const, text: JSON.stringify(status, null, 2) }],
        structuredContent: status,
      };
    },
  );

  return server;
}

// legacy defaults to 'serve', which pins a 2025-era instance from the same
// factory when a client opens with the old initialize handshake. Claude's
// negotiated revision over stdio is undocumented, so serving both eras is free
// insurance. Do not set legacy: 'reject'.
serveStdio(createServer, {
  onerror: (error: Error) => console.error(`[${NAME}] transport error: ${error.message}`),
});

// stdout carries JSON-RPC and nothing else: a single console.log here corrupts
// the stream. Diagnostics go to stderr, which Claude Desktop captures to
// ~/Library/Logs/Claude/mcp-server-<NAME>.log.
console.error(`[${NAME}] ${VERSION} listening on stdio`);
