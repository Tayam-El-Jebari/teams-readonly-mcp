import { homedir } from 'node:os';
import { join } from 'node:path';

export interface Config {
  readonly clientId: string;
  readonly tenantId: string;
  /** Space-separated OAuth scopes. */
  readonly scope: string;
  readonly tokenPath: string;
}

export const LEAST_PRIVILEGE_SCOPE =
  'offline_access Chat.Read User.Read Team.ReadBasic.All Channel.ReadBasic.All';

function required(name: string, value: string | undefined): string {
  if (!value) {
    throw new Error(
      `${name} is not set. Configure it in the MCP server's env block. `
    );
  }
  return value;
}

export function loadFirstNamesOnly(env: NodeJS.ProcessEnv = process.env): boolean {
  return env['TEAMS_MCP_FIRST_NAMES_ONLY']?.trim().toLowerCase() !== 'false';
}

export type NameMode = 'initials' | 'first-name' | 'full';

export function loadNameMode(env: NodeJS.ProcessEnv = process.env): NameMode {
  const mode = env['TEAMS_MCP_NAME_MODE']?.trim().toLowerCase();
  if (mode !== undefined) {
    if (mode === 'initials' || mode === 'first-name' || mode === 'full') return mode;
    throw new Error('TEAMS_MCP_NAME_MODE must be initials, first-name, or full.');
  }
  if (env['TEAMS_MCP_FIRST_NAMES_ONLY'] !== undefined) {
    return loadFirstNamesOnly(env) ? 'first-name' : 'full';
  }
  return 'initials';
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  const dir = env['TEAMS_MCP_TOKEN_DIR'] ?? join(homedir(), '.config', 'teams-readonly-mcp');
  return {
    clientId: required('TEAMS_MCP_CLIENT_ID', env['TEAMS_MCP_CLIENT_ID']),
    tenantId: required('TEAMS_MCP_TENANT_ID', env['TEAMS_MCP_TENANT_ID']),
    scope: env['TEAMS_MCP_SCOPE'] ?? LEAST_PRIVILEGE_SCOPE,
    tokenPath: join(dir, 'token.json'),
  };
}
