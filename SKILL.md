---
name: m365-mcp
description: Production-grade Microsoft 365 MCP server with delegated OAuth, multi-account support, pagination, rate limiting, and 38 tools covering email, calendar, contacts, OneDrive, Teams, tasks, and users.
metadata:
  openclaw:
    homepage: https://github.com/sam2kb/m365-mcp#readme
    requires:
      bins:
        - m365-mcp
        - m365-mcp-auth
    envVars:
      - name: M365_ACCOUNT
        required: false
        description: Optional account name selected from the local account store.
      - name: M365_TIMEZONE
        required: false
        description: Optional IANA timezone for calendar operations; defaults to UTC.
      - name: M365_MCP_AUTH_DIR
        required: false
        description: Optional absolute path for the local OAuth account and token store.
      - name: M365_MCP_READ_ONLY
        required: false
        description: Set true to request read-only scopes and disable all mutating tools.
    install:
      - kind: node
        package: "@sam2kb/m365-mcp"
        bins:
          - m365-mcp
          - m365-mcp-auth
---

# m365-mcp

Production-grade Microsoft 365 MCP server combining the best of office365-connector (delegated OAuth, multi-account) and mcp-microsoft365 (MCP protocol, full scope), with production-ready features: pagination, rate limiting, retry logic, and full TypeScript.

## Requirements

- Node.js 18+
- Azure Entra ID App Registration with delegated Microsoft Graph permissions
- Device code OAuth (no client-credentials, no tenant-wide access)

## Capabilities

### Email (9 tools)

- List, read, send, reply, search, move, delete, mark read/unread, list folders

### Calendar (7 tools)

- List events, today view, week view, create (with Teams meeting), update, delete, free/busy

### Contacts (6 tools)

- List, search, read, create, update, delete

### OneDrive (5 tools)

- List files, search, read content, metadata, create folders

### Teams (3 tools)

- List chats, read messages, send messages

### Tasks (5 tools)

- List lists, list tasks, create, update, delete

### Users (3 tools)

- List org users, profile lookup, manager lookup

## Security and consent

- The server contacts only Microsoft's OAuth and Graph services:
  `login.microsoftonline.com` and `graph.microsoft.com`.
- Device-code OAuth grants delegated access as the signed-in user. Read tools can
  expose private mail, files, calendars, contacts, Teams chats, tasks, and user data.
- Send, reply, move, create, update, and delete tools change real Microsoft 365
  data. Configure the MCP client to require explicit user approval for those tools.
- Tools publish standard MCP read-only, destructive, idempotent, and open-world
  annotations so compatible clients can apply confirmation policies.
- Set `M365_MCP_READ_ONLY=true` for an enforced least-privilege mode. It requests
  read-only OAuth scopes, omits mutating tools from discovery, and rejects direct
  calls to them.
- Refresh and access tokens are stored as plaintext JSON under
  `~/.m365-mcp/auth/` (or `M365_MCP_AUTH_DIR`), protected with directory mode
  `0700` and file mode `0600` where supported. Protect that directory and revoke
  the app's Microsoft account consent if a token or device is compromised.

## Setup

ClawHub installs the prebuilt npm package automatically. For a manual install:

```bash
npm install --global @sam2kb/m365-mcp

# Add account
m365-mcp-auth add work <tenantId> <clientId> you@company.com

# Authenticate
m365-mcp-auth login --account=work
```

Then configure `m365-mcp` as a stdio MCP server in your client.

See the [project README](https://github.com/sam2kb/m365-mcp#readme) for the full Azure setup guide.
