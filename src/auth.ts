/**
 * OAuth 2.0 Device Code Flow + Multi-Account Token Management
 *
 * Uses delegated permissions — the app acts AS YOU, not as the tenant.
 * Tokens stored in ~/.m365-mcp/auth/ by default with 0600 permissions.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import type { AccountConfig, AccountsStore, TokenData } from "./types.js";

export function resolveAuthDir(
  env: NodeJS.ProcessEnv = process.env,
  homeDir = os.homedir()
): string {
  const configuredDir = env.M365_MCP_AUTH_DIR?.trim();
  return configuredDir ? path.resolve(configuredDir) : path.join(homeDir, ".m365-mcp", "auth");
}

const AUTH_DIR = resolveAuthDir();
const ACCOUNTS_PATH = path.join(AUTH_DIR, "accounts.json");
const TOKENS_DIR = path.join(AUTH_DIR, "tokens");

// Scopes for delegated access — reads/writes YOUR data only, not the whole tenant
const SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.Read",
  "Calendars.ReadWrite",
  "Contacts.Read",
  "Contacts.ReadWrite",
  "Files.Read.All",
  "Tasks.ReadWrite",
  "Chat.Read",
  "Chat.ReadWrite",
].join(" ");

// ─── helpers ────────────────────────────────────────────────────────

export function ensureDirs(): void {
  fs.mkdirSync(AUTH_DIR, { recursive: true, mode: 0o700 });
  fs.mkdirSync(TOKENS_DIR, { recursive: true, mode: 0o700 });
}

export function getTokenPath(accountName: string): string {
  return path.join(TOKENS_DIR, `${accountName}.json`);
}

// ─── account store ──────────────────────────────────────────────────

export function loadAccounts(): AccountsStore {
  ensureDirs();
  if (!fs.existsSync(ACCOUNTS_PATH)) {
    return { default: null, accounts: {} };
  }
  try {
    return JSON.parse(fs.readFileSync(ACCOUNTS_PATH, "utf8"));
  } catch {
    return { default: null, accounts: {} };
  }
}

function saveAccounts(store: AccountsStore): void {
  ensureDirs();
  fs.writeFileSync(ACCOUNTS_PATH, JSON.stringify(store, null, 2), { mode: 0o600 });
}

export function addAccount(
  name: string,
  tenantId: string,
  clientId: string,
  email?: string,
  description?: string
): void {
  const store = loadAccounts();
  store.accounts[name] = {
    tenantId,
    clientId,
    email,
    description,
    addedAt: store.accounts[name]?.addedAt ?? new Date().toISOString(),
  };
  if (!store.default) store.default = name;
  saveAccounts(store);
}

export function removeAccount(name: string): void {
  const store = loadAccounts();
  delete store.accounts[name];
  const tokenPath = getTokenPath(name);
  if (fs.existsSync(tokenPath)) fs.unlinkSync(tokenPath);
  if (store.default === name) {
    store.default = Object.keys(store.accounts)[0] ?? null;
  }
  saveAccounts(store);
}

export function setDefaultAccount(name: string): void {
  const store = loadAccounts();
  if (!store.accounts[name]) throw new Error(`Account "${name}" not found`);
  store.default = name;
  saveAccounts(store);
}

export function resolveAccount(accountName?: string): AccountConfig {
  const store = loadAccounts();
  const name = accountName ?? store.default;
  if (!name) throw new Error(
    "No account configured. Run: m365-mcp auth add <name> <tenantId> <clientId>"
  );
  const acc = store.accounts[name];
  if (!acc) throw new Error(`Account "${name}" not found`);
  return { name, ...acc, tokenPath: getTokenPath(name) };
}

export function listAccounts(): { default: string | null; accounts: (AccountConfig & { isDefault: boolean })[] } {
  const store = loadAccounts();
  return {
    default: store.default,
    accounts: Object.entries(store.accounts).map(([name, acc]) => ({
      ...acc,
      name,
      isDefault: name === store.default,
      tokenPath: getTokenPath(name),
    })),
  };
}

// ─── token store ────────────────────────────────────────────────────

export function loadTokens(config: AccountConfig): TokenData | null {
  if (!fs.existsSync(config.tokenPath)) return null;
  try {
    const data = JSON.parse(fs.readFileSync(config.tokenPath, "utf8"));
    if (!data.access_token || !data.refresh_token) return null;
    return data;
  } catch {
    return null;
  }
}

function saveTokens(config: AccountConfig, tokens: TokenData): TokenData {
  ensureDirs();
  // Preemptively rotate the refresh token — store, don't re-use the old one
  const data: TokenData = {
    access_token: tokens.access_token,
    refresh_token: tokens.refresh_token,
    expires_at: tokens.expires_at,
    scope: tokens.scope ?? SCOPES,
  };
  fs.writeFileSync(config.tokenPath, JSON.stringify(data, null, 2), { mode: 0o600 });
  return data;
}

// ─── HTTPS helper ───────────────────────────────────────────────────

async function httpsPost(url: string, body: URLSearchParams): Promise<any> {
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const data = await resp.json();
  if (!resp.ok) {
    throw new Error(
      `Auth error: ${data.error_description ?? data.error ?? resp.statusText}`
    );
  }
  return data;
}

// ─── device code flow ───────────────────────────────────────────────

export async function requestDeviceCode(config: AccountConfig): Promise<{
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
}> {
  return httpsPost(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/devicecode`,
    new URLSearchParams({ client_id: config.clientId, scope: SCOPES })
  );
}

async function pollForToken(
  deviceCode: string,
  config: AccountConfig
): Promise<TokenData> {
  const data = await httpsPost(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: config.clientId,
      device_code: deviceCode,
    })
  );
  return {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_at: Date.now() + (data.expires_in - 60) * 1000,
    scope: data.scope,
  };
}

// ─── token refresh ──────────────────────────────────────────────────

async function refreshAccessToken(config: AccountConfig, refreshToken: string): Promise<TokenData> {
  const data = await httpsPost(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    new URLSearchParams({
      grant_type: "refresh_token",
      client_id: config.clientId,
      refresh_token: refreshToken,
      scope: SCOPES,
    })
  );
  return {
    access_token: data.access_token,
    // OAuth servers may omit refresh_token when they do not rotate it.
    refresh_token: data.refresh_token ?? refreshToken,
    expires_at: Date.now() + (data.expires_in - 60) * 1000,
    scope: data.scope,
  };
}

// ─── public API ─────────────────────────────────────────────────────

/** Serialises refresh attempts per account without mixing tokens across accounts. */
const refreshLocks = new Map<string, Promise<string>>();

/**
 * Get a valid access token (auto-refresh if expired or within 5 min of expiry).
 * Uses a mutex to avoid concurrent refresh races.
 */
export async function getAccessToken(accountName?: string): Promise<string> {
  const config = resolveAccount(accountName);
  let tokens = loadTokens(config);
  if (!tokens) {
    throw new Error(
      `Not authenticated for account "${config.name}". Run: m365-mcp auth login [--account=${config.name}]`
    );
  }

  // Still valid with 5-minute buffer
  if (tokens.expires_at > Date.now() + 5 * 60 * 1000) {
    return tokens.access_token;
  }

  // Serialise refresh — if another caller is already refreshing, wait for it
  let refreshLock = refreshLocks.get(config.tokenPath);
  if (!refreshLock) {
    refreshLock = doRefresh(config, tokens.refresh_token).finally(() => {
      refreshLocks.delete(config.tokenPath);
    });
    refreshLocks.set(config.tokenPath, refreshLock);
  }
  return refreshLock;
}

async function doRefresh(config: AccountConfig, refreshToken: string): Promise<string> {
  // Double-check tokens haven't been refreshed by a concurrent caller
  const current = loadTokens(config);
  if (current && current.expires_at > Date.now() + 5 * 60 * 1000) {
    return current.access_token;
  }

  let attempt = 0;
  let delay = 1000;
  while (attempt < 3) {
    try {
      const fresh = await refreshAccessToken(config, current?.refresh_token ?? refreshToken);
      saveTokens(config, fresh);
      return fresh.access_token;
    } catch (err: any) {
      attempt++;
      if (err.message.includes("expired") || err.message.includes("invalid_grant")) {
        throw new Error(
          `Refresh token expired for "${config.name}". Re-authenticate: m365-mcp auth login`
        );
      }
      if (attempt >= 3) throw err;
      await new Promise((r) => setTimeout(r, delay));
      delay *= 2;
    }
  }
  throw new Error(`Failed to refresh token after 3 attempts`);
}

/**
 * Full device-code authentication flow.
 * Calls onPrompt(verificationUrl, userCode) so the caller can display instructions.
 */
export async function authenticate(
  accountName: string | undefined,
  onPrompt: (verificationUrl: string, userCode: string) => void
): Promise<TokenData> {
  const config = resolveAccount(accountName);

  // Check existing
  const existing = loadTokens(config);
  if (existing && existing.expires_at > Date.now() + 5 * 60 * 1000) {
    return existing;
  }

  // Request device code
  const dc = await requestDeviceCode(config);
  onPrompt(dc.verification_uri, dc.user_code);

  // Poll
  const deadline = Date.now() + dc.expires_in * 1000;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, dc.interval * 1000));
    try {
      const tokens = await pollForToken(dc.device_code, config);
      return saveTokens(config, tokens);
    } catch (err: any) {
      if (err.message.includes("authorization_pending")) {
        continue;
      }
      if (
        err.message.includes("authorization_declined") ||
        err.message.includes("access_denied")
      ) {
        throw new Error("User declined authorization.");
      }
      if (
        err.message.includes("expired_token") ||
        err.message.includes("AADSTS70019")
      ) {
        throw new Error("Device code expired. Please try again.");
      }
      throw err;
    }
  }
  throw new Error("Authentication timed out.");
}

export async function getTokenStatus(
  accountName?: string
): Promise<{ account: string; valid: boolean; expires: string | null }> {
  const config = resolveAccount(accountName);
  const tokens = loadTokens(config);
  if (!tokens) return { account: config.name, valid: false, expires: null };
  const valid = tokens.expires_at > Date.now();
  return {
    account: config.name,
    valid,
    expires: new Date(tokens.expires_at).toISOString(),
  };
}
