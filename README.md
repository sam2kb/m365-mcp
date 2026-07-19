# m365-mcp

[![npm](https://img.shields.io/npm/v/%40sam2kb%2Fm365-mcp?logo=npm&label=npm)](https://www.npmjs.com/package/@sam2kb/m365-mcp)
[![ClawHub](https://img.shields.io/badge/ClawHub-m365--mcp-f97316)](https://clawhub.ai/sam2kb/skills/m365-mcp)

**Production-grade Microsoft 365 MCP server** — Email, Calendar, Contacts, OneDrive, Teams, Tasks, and Users via delegated OAuth.

Built from the best of both `office365-connector` (delegated OAuth, multi-account) and `mcp-microsoft365` (MCP protocol, full scope), with all the gaps fixed: pagination, rate limiting, retry logic, proper timezone handling, and TypeScript throughout.

## Why this exists

This comparison reflects the code shipped in
[`mcp-microsoft365` v1.0.0](https://clawhub.ai/makhatib/skills/mcp-microsoft365)
and
[`office365-connector` v2.0.0](https://clawhub.ai/tirandagan/skills/office365-connector)
as reviewed on July 18, 2026.

| Feature | `mcp-microsoft365` | `office365-connector` | **`m365-mcp`** |
|---|---|---|---|
| Auth flow | Client credentials (app-only) | Device code + client secret | **Public-client device code** |
| Access model | Application permissions (tenant-wide) | Delegated (signed-in user) | **Delegated (signed-in user)** |
| Client secret | Required | Required and stored locally | **None** ✅ |
| MCP server | ✅ | ❌ (CLI scripts) | ✅ |
| Multi-account | ❌ | ✅ | ✅ |
| Pagination | ❌ | ❌ | ✅ |
| Graph 429 retries | ❌ | ❌ (not implemented in shipped code) | ✅ |
| Read-only mode | ❌ | ❌ | ✅ |
| Email | ✅ | ✅ | ✅ |
| Calendar | ✅ | ✅ | ✅ |
| Contacts | ❌ | ❌ (documented, not shipped) | ✅ |
| OneDrive | ✅ | ❌ | ✅ |
| Teams | ✅ | ❌ | ✅ |
| Tasks | ✅ | ❌ | ✅ |
| Users | ✅ | ❌ | ✅ |
| Implementation | TypeScript | JavaScript CLI | TypeScript |
| Interface coverage | 19 MCP tools | CLI scripts | **42 MCP tools** |

`m365-mcp` does not store a client secret or password. Tenant and client IDs are
non-secret identifiers; sensitive OAuth access and refresh tokens are stored
locally as permission-restricted plaintext files. See [Security](#security) for
the storage and revocation details.

## Scope

**Delegated permissions** — the app acts AS YOU, not as the tenant. It can only access YOUR data:

- `Mail.Read` / `Mail.ReadWrite` / `Mail.Send`
- `Calendars.Read` / `Calendars.ReadWrite`
- `Contacts.Read` / `Contacts.ReadWrite`
- `Files.Read.All` (your OneDrive)
- `Tasks.ReadWrite` (To Do)
- `Chat.Read` / `Chat.ReadWrite` (Teams)
- `User.Read` + `offline_access`

These are delegated scopes, not tenant-wide application permissions. Set
`M365_MCP_READ_ONLY=true` before authentication and when running the server to
request only the read variants and disable mutating MCP tools.

Organization-wide mail application permissions such as `Mail.Read.All` are not used.

## Setup

### 1. Create Azure Entra ID App Registration

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. **New registration**
   - Name: `m365-mcp`
   - Supported account types: **Single tenant** (or multi-tenant if you manage your own tenant)
   - Redirect URI: leave this blank; device-code flow does not use one
3. Click **Register**
4. Go to **Authentication** → **Advanced settings**
5. Set **"Allow public client flows"** to **Yes**, then click **Save**

> This is required for device-code OAuth. Do not create or configure a client
> secret; `m365-mcp` is a public client.

### 2. Add API Permissions

Under **API Permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:

```
offline_access
User.Read
Mail.Read
Mail.ReadWrite
Mail.Send
Calendars.Read
Calendars.ReadWrite
Contacts.Read
Contacts.ReadWrite
Files.Read.All
Tasks.ReadWrite
Chat.Read
Chat.ReadWrite
```

Click **Grant admin consent** (or each user will consent individually during login).

> Organization-wide user listing and lookup additionally requires delegated `User.Read.All` with admin consent. This broader scope is not requested by default.

### 3. Install

Install the published command-line tools:

```bash
npm install --global @sam2kb/m365-mcp
```

Or build from source:

```bash
# Clone
git clone https://github.com/sam2kb/m365-mcp.git
cd m365-mcp

# Install dependencies
npm install

# Build
npm run build
```

### 4. Add Your Account & Authenticate

```bash
m365-mcp-auth add work <tenant-id> <client-id> you@company.com "Work account"
m365-mcp-auth login --account=work
```

> **`AADSTS7000218` or a missing `client_secret` / `client_assertion` error:**
> Microsoft is treating the registration as a confidential client. Confirm that
> the configured Application (client) ID belongs to the registration you edited,
> then return to **Authentication** and verify **Allow public client flows** is
> saved as **Yes**. Do not work around this error by adding a client secret.

When running from a source checkout, use `node dist/auth-cli.js` instead of
`m365-mcp-auth`.

Follow the on-screen URL + code to sign in. Access and refresh tokens are stored
as plaintext JSON in `~/.m365-mcp/auth/` by default, protected with directory
mode `0700` and file mode `0600` where supported. Set `M365_MCP_AUTH_DIR` to
use another protected location.

### 5. Configure an MCP Client

The server uses standard MCP over stdio and works with any MCP client (Claude Desktop, Cursor, Continue, etc.).

#### OpenClaw

For the global npm installation, add this to your mcporter config at `~/.openclaw/mcporter.json`:

```json
{
  "mcpServers": {
    "m365": {
      "command": "m365-mcp",
      "args": [],
      "env": {
        "M365_ACCOUNT": "work",
        "M365_TIMEZONE": "America/Chicago"
      }
    }
  }
}
```

Or via mcporter CLI:

```bash
mcporter config add m365 --stdio "m365-mcp" \
  --env M365_ACCOUNT=work \
  --env M365_TIMEZONE=America/Chicago
```

For a source checkout, use `"command": "node"` with
`"args": ["/absolute/path/to/m365-mcp/dist/index.js"]`.

Then restart OpenClaw for the server to load.

#### Other MCP Clients

Use the same JSON config in your client's MCP server configuration — `m365-mcp` is a standard stdio MCP server with no client-specific requirements.

## Multi-Account

```bash
# Add more accounts
node dist/auth-cli.js add personal <tenant2> <client2> you@outlook.com "Personal"
node dist/auth-cli.js add client <tenant3> <client3> you@client.com "Consulting"

# Authenticate each
node dist/auth-cli.js login --account=personal
node dist/auth-cli.js login --account=client

# Set default
node dist/auth-cli.js default work

# List
node dist/auth-cli.js list
```

Each account needs its own App Registration in its respective tenant. Tokens are isolated per account.

## Available Tools (42)

### 📧 Mail (9 tools)
| Tool | Description |
|---|---|
| `m365_mail_list` | List emails (folder, filter, search, paginated) |
| `m365_mail_read` | Read full email by ID |
| `m365_mail_send` | Send email (to/cc/bcc, HTML or plain) |
| `m365_mail_reply` | Reply / reply-all to an email |
| `m365_mail_search` | Search emails across folders |
| `m365_mail_move` | Move email to another folder |
| `m365_mail_delete` | Delete email |
| `m365_mail_mark_read` | Mark as read/unread |
| `m365_mail_folders` | List all mail folders with counts |

### 📅 Calendar (7 tools)
| Tool | Description |
|---|---|
| `m365_calendar_list` | Events in a date range (paginated) |
| `m365_calendar_today` | Today's events, nicely formatted |
| `m365_calendar_week` | Week view grouped by day |
| `m365_calendar_create` | Create event (optional Teams meeting) |
| `m365_calendar_update` | Update event |
| `m365_calendar_delete` | Cancel event with message |
| `m365_calendar_availability` | Free/busy lookup |

### 👤 Contacts (10 tools)
| Tool | Description |
|---|---|
| `m365_contacts_list` | List by folder and filter by category |
| `m365_contacts_search` | Search by name, email, or company |
| `m365_contacts_read` | Full contact details |
| `m365_contacts_create` | Create with multiple emails, phones, categories, addresses, and work or personal details |
| `m365_contacts_update` | Update or clear contact fields and category assignments |
| `m365_contacts_delete` | Move a contact to Deleted Items |
| `m365_contacts_folders_list` | List top-level or child contact folders |
| `m365_contacts_folder_create` | Create a child contact folder |
| `m365_contacts_folder_update` | Rename or move a contact folder |
| `m365_contacts_folder_delete` | Delete a non-default contact folder |

Outlook categories are tags stored on each contact; contact folders are containers. Category filtering is case-insensitive and can be combined with folder selection. Folder tools manage top-level folders and one direct child level, keeping the interface predictable.

### 📁 OneDrive (5 tools)
| Tool | Description |
|---|---|
| `m365_files_list` | List files/folders (paginated) |
| `m365_files_search` | Search files |
| `m365_files_read` | Read text file content |
| `m365_files_info` | File/folder metadata |
| `m365_files_create_folder` | Create folder |

### 💬 Teams (3 tools)
| Tool | Description |
|---|---|
| `m365_teams_chats` | List your chats |
| `m365_teams_messages` | Get chat messages |
| `m365_teams_send` | Send chat message |

### ✅ Tasks (5 tools)
| Tool | Description |
|---|---|
| `m365_tasks_lists` | List To Do lists |
| `m365_tasks_list` | List tasks in a list |
| `m365_tasks_create` | Create task |
| `m365_tasks_update` | Update task |
| `m365_tasks_delete` | Delete task |

### 👥 Users (3 tools)
| Tool | Description |
|---|---|
| `m365_users_list` | List org users |
| `m365_users_profile` | Get user profile |
| `m365_users_manager` | Get user's manager |

## Development

```bash
npm install
npm run dev    # tsx watch mode
npm run build  # compile TypeScript
```

## Security

- **Delegated OAuth** — the app acts as the authenticated user; it does not use
  application credentials or tenant-wide mail permissions.
- **Device code flow** — you never type your password into anything but Microsoft's login page.
- **Sensitive reads** — mail, files, calendar entries, contacts, Teams messages,
  tasks, and user profiles can enter the MCP client's model context.
- **Real side effects** — send, reply, move, create, update, and delete tools
  change Microsoft 365 data. Configure your MCP client to require explicit user
  approval before it invokes them.
- **MCP safety annotations** — every tool declares read-only, destructive,
  idempotent, and open-world hints for clients that enforce tool policies.
- **Enforced read-only mode** — set `M365_MCP_READ_ONLY=true` both when
  authenticating and running the server. The auth flow requests read-only Graph
  scopes, mutating tools are omitted from discovery, and direct calls are blocked.
- **Local token storage** — access and refresh tokens are plaintext JSON protected
  by `0700` directories and `0600` files where supported. Protect
  `~/.m365-mcp/auth/` or your configured `M365_MCP_AUTH_DIR`.
- **Auto-refresh** — tokens refreshed before expiry, expired refresh tokens trigger re-auth.
- **No telemetry, no analytics, no third-party calls** besides `login.microsoftonline.com` and `graph.microsoft.com`.

To delete a local token, remove its account:

```bash
m365-mcp-auth remove work
```

If a token or device may be compromised, also revoke the application's consent
from the Microsoft account or Entra ID portal. Re-authenticate after changing
`M365_MCP_READ_ONLY` so the stored token reflects the intended scope set.

## License

MIT
