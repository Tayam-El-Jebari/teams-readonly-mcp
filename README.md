# teams-readonly-mcp

An MCP server that reads **your own** Microsoft Teams chats, and cannot write to them.

Built for a straightforward Teams sign-in flow, reduced exposure of people's
names, and a tool surface that cannot post as you. Once your own Entra app
registration is configured, device-code sign-in and automatic token refresh
let an assistant read conversations without handling credentials itself.

Names default to first name plus surname initials, such as `Karel v. d. s.`.
This reduces name exposure; it does not anonymize message content or identities.
Write tools are absent, rather than hidden behind settings you must remember
to disable in each client.

A daily briefing motivated the project, but the connector can support other
read-only Teams workflows. It handles Graph filtering, pagination, and bounded
output in code so each automation does not have to rediscover those rules.

## Why read-only, and what that actually means

"Read-only" here is three things, not a promise:

1. **No write tools exist.** The tool surface is the capability surface. There is
   nothing to misconfigure, because the code to send a message was never
   written.
2. **Least-privilege scopes by default.** The default configuration requests
  read scopes only. `TEAMS_MCP_SCOPE` is configurable, so verify the granted
  scopes with `teams_auth_status`; the server does not enforce a read-only
  credential if you configure broader permissions.
3. **One HTTP chokepoint that only issues `GET`** against Graph. The only `POST`
   in the codebase is the OAuth token endpoint.

Reading does not mark anything as read. Graph only moves your read state via a
separate `POST /chats/{id}/markChatReadForUser`, which needs `Chat.ReadWrite`
and is never called.

## Tools

| Tool | What it does |
|---|---|
| `teams_auth_status` | Whether a credential is stored, which scopes it holds, when it expires. Never returns token material. Touches no network. Use this as a precondition check |
| `teams_auth_login` | Starts an interactive device-code sign-in, returns the URL and code |
| `teams_auth_complete` | Waits for you to finish signing in, stores the credential |
| `teams_list_conversations` | Your chats: name, type, last activity, and a bounded member preview with count and clipping flag |
| `teams_read_conversation` | Messages from one chat, optionally since a timestamp. HTML stripped, mentions resolved, system messages filtered |
| `teams_sweep` | Reads many chats in one call, isolating failures per chat so a partial sweep can never look complete |
| `teams_list_channels` | Teams and channels you belong to |
| `teams_read_channel` | Channel messages. **Requires admin consent**, see below |

## Setup

### 1. Register the app (needed only once per organisation/department)

Use an app registration supplied by your organisation, or create one in the
Microsoft Entra portal according to its policies:

1. **App registrations** → **New registration** → single tenant.
2. **Authentication** → **Advanced settings** → **Allow public client flows** =
   **Yes**. No client secret.
3. **API permissions** → **Microsoft Graph** → **Delegated**, add:
   `offline_access`, `Chat.Read`, `User.Read`, `Team.ReadBasic.All`,
   `Channel.ReadBasic.All`.
4. **For channel messages**, also add `ChannelMessage.Read.All` and ask an
  administrator to grant consent. It is optional for chat-only use.
5. **Overview** → copy the **Application (client) ID** and **Directory (tenant) ID**.

### 2. Install

```bash
git clone https://github.com/Tayam-El-Jebari/teams-readonly-mcp
cd teams-readonly-mcp
pnpm install
pnpm build
```

Requires Node 20.19 or newer. [pnpm](https://pnpm.io/installation) is recommended.

### 3. Connect your AI app

Choose your app below and follow its instructions. You only need one.
If your organisation manages your AI app, ask IT to add this connector for you.

<details>
<summary>Claude Desktop: copy and paste a configuration</summary>

1. Open **Claude Desktop → Settings → Developer → Edit Config**.
2. Open `claude_desktop_config.json` in a text editor. If it is empty or contains
  only `{}`, paste the configuration below. If it already contains connections,
  add only the `teams-readonly` entry inside its existing `mcpServers` object.
  Do not replace your other connections; ask IT for help merging it if needed.
3. Replace the four placeholder values using the instructions below the example.
4. Save the file, fully quit Claude Desktop, and reopen it.

```json
{
  "mcpServers": {
    "teams-readonly": {
      "command": "/absolute/path/to/node",
      "args": ["/absolute/path/to/teams-readonly-mcp/dist/index.js"],
      "env": {
        "TEAMS_MCP_CLIENT_ID": "<your app registration's client ID>",
        "TEAMS_MCP_TENANT_ID": "<your tenant ID>"
      }
    }
  }
}
```

Replace these four values, keeping the surrounding double quotes:

- `/absolute/path/to/node`: open Terminal on macOS and run `which node`, or
  Command Prompt on Windows and run `where node`. Paste the full path it prints
  (use the first result if there are several).
- `/absolute/path/to/teams-readonly-mcp/dist/index.js`: in the terminal you used
  for step 2, inside the downloaded project folder, run
  `node -p "require('node:path').resolve('dist/index.js')"`. Paste the full path
  it prints.
- `<your app registration's client ID>`: paste the **Application (client) ID**
  from step 1, replacing the angle brackets too.
- `<your tenant ID>`: paste the **Directory (tenant) ID** from step 1.

On Windows, replace each backslash in both paths with `/` when pasting into the
configuration, for example `C:/Program Files/nodejs/node.exe`.

**Want channel messages too?** After an administrator grants the permission in
step 1, replace the `env` section above with this one, filling in both IDs:

```json
"env": {
  "TEAMS_MCP_CLIENT_ID": "<your app registration's client ID>",
  "TEAMS_MCP_TENANT_ID": "<your tenant ID>",
  "TEAMS_MCP_SCOPE": "offline_access Chat.Read User.Read Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All"
}
```

This is a replacement section, not a complete configuration file. Save and
restart Claude Desktop after changing it, then sign in as described in step 4.

</details>

<details>
<summary>Claude Code command</summary>

```bash
claude mcp add teams-readonly \
  --env TEAMS_MCP_CLIENT_ID="<client-id>" \
  --env TEAMS_MCP_TENANT_ID="<tenant-id>" \
  -- node /absolute/path/to/teams-readonly-mcp/dist/index.js
```

Replace the IDs and project path before running the command. For channel messages,
after admin consent, add the following option before `--`:

```text
--env TEAMS_MCP_SCOPE="offline_access Chat.Read User.Read Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All"
```

</details>

### 4. Sign in

Ask your assistant to run `teams_auth_login`, open the URL it gives you, enter
the code, and sign in. Then run `teams_auth_complete` and check
`teams_auth_status` for the expected read permissions.

Tokens refresh automatically when possible. Sign in again if access is revoked,
your organisation requires it, or you change the requested permissions.

## Configuration

| Variable | Required | Default |
|---|---|---|
| `TEAMS_MCP_CLIENT_ID` | yes | none |
| `TEAMS_MCP_TENANT_ID` | yes | none |
| `TEAMS_MCP_SCOPE` | no | `offline_access Chat.Read User.Read Team.ReadBasic.All Channel.ReadBasic.All` |
| `TEAMS_MCP_TOKEN_DIR` | no | `~/.config/teams-readonly-mcp` |
| `TEAMS_MCP_NAME_MODE` | no | `initials` (`first-name` and `full` also supported) |

`TEAMS_MCP_NAME_MODE` controls member names, generated chat names, user senders,
and resolved mentions, not arbitrary names in message text or chat topics.
The legacy `TEAMS_MCP_FIRST_NAMES_ONLY` setting is still respected when the new
variable is absent: `true` selects `first-name`, and `false` selects `full`.

Concise listings include at most five member names. `membersTruncated` reports
whether the server clipped that preview; `memberCount` counts members returned
by Graph, not a verified membership total. Detailed listings include all returned
member names, using the same name mode.

The token is written `0600` in a `0700` directory, via a temp file and an atomic
rename so a crash cannot leave a truncated credential behind. Keep it out of
your client's config file, which is plaintext and routinely pasted into bug
reports.

## Things worth knowing before you rely on it

- **A partial listing says nothing about omitted chats' activity.** Chat lists
  follow up to 20 Graph pages and also have an output-size limit. `activeSince`
  and `types` filter results locally; neither guarantees completeness or ordering.
  Name lookup searches independently of the listing output budget, up to the
  same page limit, and refuses to claim a unique match if that search is partial.
  Use known IDs for critical chats and report unreachable targets separately
  from successfully read chats with no matching messages.
- **Graph's terms of use cap change-polling at once per day.** A daily briefing
  is fine. A five-minute cron is a terms violation.
- **Chat message reads are throttled to 1 request per second per chat**, so a
  sweep over many chats is paced, not instant.
- **There is no delta endpoint for delegated chat reads.** Incremental reads use
  a per-chat filter against a timestamp you supply.
- **`$filter` is silently ignored** unless paired with `$orderby` on the same
  property. This server always pairs them, and a test enforces it, because the
  failure returns `200 OK` with the filter dropped: you would get the newest 50
  messages regardless of your timestamp, with nothing to indicate anything went
  wrong.

## Development

```bash
pnpm build     # tsc
pnpm test      # node --test
pnpm inspect   # tools/list via the MCP Inspector
```

Add `--strict` when inspecting. It catches schema portability problems that
individual clients will otherwise hit at runtime, such as a nullable field
serialising to `type: ["string","null"]`, which some clients reject outright.

## Credits

Teams returns message bodies as HTML with `<at id="0">Name</at>` mention tags
that must be correlated against a parallel `mentions[]` array. The approach to
that is owed to [floriscornel/teams-mcp](https://github.com/floriscornel/teams-mcp)
(MIT), which solved it first.

## License

MIT
