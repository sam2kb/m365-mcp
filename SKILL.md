---
name: m365-mcp
description: Production-grade Microsoft 365 MCP server with delegated OAuth, multi-account support, pagination, rate limiting, and 35+ tools covering email, calendar, contacts, OneDrive, Teams, tasks, and users.
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

## Setup

```bash
cd m365-mcp
npm install
npm run build

# Add account
node dist/auth-cli.js add work <tenantId> <clientId> you@company.com

# Authenticate
node dist/auth-cli.js login --account=work
```

Then configure as an MCP server in your mcporter config.

See [README.md](README.md) for full Azure setup guide.

## License

MIT
