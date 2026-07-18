# m365-mcp

**Production-grade Microsoft 365 MCP server** — Email, Calendar, Contacts, OneDrive, Teams, Tasks, and Users via delegated OAuth.

Built from the best of both `office365-connector` (delegated OAuth, multi-account) and `mcp-microsoft365` (MCP protocol, full scope), with all the gaps fixed: pagination, rate limiting, retry logic, proper timezone handling, and TypeScript throughout.

## Why this exists

| Feature | mcp-microsoft365 | office365-connector | **m365-mcp** |
|---|---|---|---|
| Auth model | Client credentials (tenant-wide) | Delegated (device code) | **Delegated (device code)** ✅ |
| MCP server | ✅ | ❌ (CLI scripts) | ✅ |
| Multi-account | ❌ | ✅ | ✅ |
| Pagination | ❌ | ❌ | ✅ |
| Rate limit handling | ❌ | Partial | ✅ |
| Email | ✅ | ✅ | ✅ |
| Calendar | ✅ | ✅ | ✅ |
| Contacts | ❌ | ✅ | ✅ |
| OneDrive | ✅ | ❌ | ✅ |
| Teams | ✅ | ❌ | ✅ |
| Tasks | ✅ | ❌ | ✅ |
| Users | ✅ | ❌ | ✅ |
| TypeScript | ✅ | ❌ (JS) | ✅ |
| **Total tools** | 19 | ~12 (CLI) | **38** |

## Scope

**Delegated permissions** — the app acts AS YOU, not as the tenant. It can only access YOUR data:

- `Mail.Read` / `Mail.ReadWrite` / `Mail.Send`
- `Calendars.Read` / `Calendars.ReadWrite`
- `Contacts.Read` / `Contacts.ReadWrite`
- `Files.Read.All` (your OneDrive)
- `Tasks.ReadWrite` (To Do)
- `Chat.Read` / `Chat.ReadWrite` (Teams)
- `User.Read` + `offline_access`

No tenant-wide `Mail.Read.All` or `Files.Read.All` needed.

## Setup

### 1. Create Azure Entra ID App Registration

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. **New registration**
   - Name: `m365-mcp`
   - Supported account types: **Single tenant** (or multi-tenant if you manage your own tenant)
   - Redirect URI: `http://localhost` (not used for device code, but required)
3. Click **Register**
4. Go to **Authentication** → **Advanced settings** → set **"Allow public client flows"** to **Yes** → **Save**
   > This is required for device-code OAuth to work without a client secret.

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
# Add account
node dist/auth-cli.js add work <tenant-id> <client-id> you@company.com "Work account"

# Authenticate (device code flow)
node dist/auth-cli.js login --account=work
```

Follow the on-screen URL + code to sign in. Tokens are stored securely in `~/.m365-mcp/auth/` by default. Set `M365_MCP_AUTH_DIR` to use another location.

### 5. Configure an MCP Client

The server uses standard MCP over stdio. For example, add it to an OpenClaw mcporter config:

```json
{
  "mcpServers": {
    "m365": {
      "command": "node",
      "args": ["/path/to/m365-mcp/dist/index.js"],
      "env": {
        "M365_ACCOUNT": "work",
        "M365_TIMEZONE": "America/Chicago"
      }
    }
  }
}
```

Or via CLI:

```bash
mcporter config add m365 --stdio "node /path/to/m365-mcp/dist/index.js" \
  --env M365_ACCOUNT=work \
  --env M365_TIMEZONE=America/Chicago
```

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

## Available Tools (38)

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

### 👤 Contacts (6 tools)
| Tool | Description |
|---|---|
| `m365_contacts_list` | List contacts |
| `m365_contacts_search` | Search by name/email/company |
| `m365_contacts_read` | Full contact details |
| `m365_contacts_create` | Create contact |
| `m365_contacts_update` | Update contact |
| `m365_contacts_delete` | Delete contact |

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

- **Delegated OAuth** — app acts as the authenticated user. No tenant-wide access.
- **Device code flow** — you never type your password into anything but Microsoft's login page.
- **Tokens stored with 0600 permissions** in `~/.m365-mcp/auth/` by default, configurable with `M365_MCP_AUTH_DIR`.
- **Auto-refresh** — tokens refreshed before expiry, expired refresh tokens trigger re-auth.
- **No telemetry, no analytics, no third-party calls** besides `login.microsoftonline.com` and `graph.microsoft.com`.

## License

MIT
