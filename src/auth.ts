/**
 * OAuth 2.0 Device Code Flow + Multi-Account Token Management
 *
 * Uses delegated permissions — the app acts AS YOU, not as the tenant.
 * Tokens stored in ~/.m365-mcp/auth/ by default with 0600 permissions.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import { randomUUID } from "node:crypto";
import type { AccountConfig, AccountsStore, TokenData } from "./types.js";
import { envFlag } from "./security.js";

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

// Delegated scopes: Graph combines consent with the signed-in user's access.
export function resolveScopes(env: NodeJS.ProcessEnv = process.env): string {
  const scopes = [
    "offline_access",
    "User.Read",
    "Mail.Read",
    "Calendars.Read",
    "Contacts.Read",
    "People.Read",
    "Files.Read.All",
    "Tasks.Read",
    "Chat.Read",
  ];

  if (!envFlag(env.M365_MCP_READ_ONLY)) {
    scopes.push(
      "Mail.ReadWrite",
      "Mail.Send",
      "Calendars.ReadWrite",
      "Contacts.ReadWrite",
      "Tasks.ReadWrite",
      "Chat.ReadWrite",
      "Files.ReadWrite"
    );
  }

  if (envFlag(env.M365_MCP_DIRECTORY_READ)) scopes.push("User.Read.All");
  return scopes.join(" ");
}

const SCOPES = resolveScopes();

// ─── helpers ────────────────────────────────────────────────────────

export function ensureDirs(): void {
  fs.mkdirSync(AUTH_DIR, { recursive: true, mode: 0o700 });
  fs.mkdirSync(TOKENS_DIR, { recursive: true, mode: 0o700 });
}

export function getTokenPath(accountName: string): string {
  validateAccountName(accountName);
  return path.join(TOKENS_DIR, `${accountName}.json`);
}

function validateAccountName(name: string): void {
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,63}$/.test(name) || /\.$/.test(name)) {
    throw new Error("Account name must start with a letter or digit and contain at most 64 letters, digits, dots, underscores or hyphens");
  }
}

function writePrivateJson(file: string, data: unknown): void {
  const temporary = `${file}.${randomUUID()}.tmp`;
  try {
    fs.writeFileSync(temporary, JSON.stringify(data, null, 2), { mode: 0o600, flag: "wx" });
    fs.renameSync(temporary, file);
  } finally {
    if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
  }
}

function hasRequestedScopes(scope: string): boolean {
  const granted = new Set(scope.split(/\s+/));
  const required = SCOPES.split(" ").filter((value) => value !== "offline_access");
  if (!required.every((value) => granted.has(value))) return false;
  return !envFlag(process.env.M365_MCP_READ_ONLY) ||
    ![...granted].some((value) => /(?:\.ReadWrite(?:\.All)?|\.Send)$/.test(value));
}

// ─── account store ──────────────────────────────────────────────────

export function loadAccounts(): AccountsStore {
  ensureDirs();
  if (!fs.existsSync(ACCOUNTS_PATH)) {
    return { default: null, accounts: {} };
  }
  try {
    const store = JSON.parse(fs.readFileSync(ACCOUNTS_PATH, "utf8"));
    if (!store || !store.accounts || typeof store.accounts !== "object" || Array.isArray(store.accounts) ||
      (store.default !== null && typeof store.default !== "string")) throw new Error("Invalid account registry");
    return store;
  } catch (error) {
    throw new Error("Cannot read account registry; restore accounts.json before changing accounts", { cause: error });
  }
}

function saveAccounts(store: AccountsStore): void {
  ensureDirs();
  writePrivateJson(ACCOUNTS_PATH, store);
}

export function addAccount(
  name: string,
  tenantId: string,
  clientId: string,
  email?: string,
  description?: string
): void {
  validateAccountName(name);
  if (!tenantId.trim() || !clientId.trim()) throw new Error("tenantId and clientId are required");
  const store = loadAccounts();
  const existing = Object.hasOwn(store.accounts, name) ? store.accounts[name] : undefined;
  if (existing && (existing.tenantId !== tenantId || existing.clientId !== clientId)) {
    const tokenPath = getTokenPath(name);
    if (fs.existsSync(tokenPath)) fs.unlinkSync(tokenPath);
  }
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
  validateAccountName(name);
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
  validateAccountName(name);
  if (!Object.hasOwn(store.accounts, name)) throw new Error(`Account "${name}" not found`);
  store.default = name;
  saveAccounts(store);
}

export function resolveAccount(accountName?: string): AccountConfig {
  const store = loadAccounts();
  const name = accountName ?? store.default;
  if (!name) throw new Error(
    "No account configured. Run: m365-mcp-auth add <name> <tenantId> <clientId>"
  );
  validateAccountName(name);
  if (!Object.hasOwn(store.accounts, name)) throw new Error(`Account "${name}" not found`);
  const acc = store.accounts[name];
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
    if (typeof data?.access_token !== "string" || !data.access_token ||
        typeof data.refresh_token !== "string" || !data.refresh_token ||
        !Number.isFinite(data.expires_at) || data.expires_at <= 0 || typeof data.scope !== "string") return null;
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
  writePrivateJson(config.tokenPath, data);
  return data;
}

// ─── HTTPS helper ───────────────────────────────────────────────────

class OAuthError extends Error {
  constructor(
    readonly code: string | undefined,
    description: string
  ) {
    super(`Auth error${code ? ` [${code}]` : ""}: ${description}`);
    this.name = "OAuthError";
  }
}

async function httpsPost(url: string, body: URLSearchParams): Promise<any> {
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
    signal: AbortSignal.timeout(30_000),
  });
  const data = await resp.json();
  if (!resp.ok) {
    throw new OAuthError(
      typeof data.error === "string" ? data.error : undefined,
      data.error_description ?? resp.statusText
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
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/devicecode`,
    new URLSearchParams({ client_id: config.clientId, scope: SCOPES })
  );
}

async function pollForToken(
  deviceCode: string,
  config: AccountConfig
): Promise<TokenData> {
  const data = await httpsPost(
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`,
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
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`,
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
      `Not authenticated for account "${config.name}". Run: m365-mcp-auth login [--account=${config.name}]`
    );
  }

  // Still valid with 5-minute buffer
  if (tokens.expires_at > Date.now() + 5 * 60 * 1000) {
    if (!hasRequestedScopes(tokens.scope)) throw new Error(`Permissions changed for "${config.name}". Re-authenticate: m365-mcp-auth login --account=${config.name}`);
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
  if (current && current.expires_at > Date.now() + 5 * 60 * 1000 && hasRequestedScopes(current.scope)) {
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
          `Refresh token expired for "${config.name}". Re-authenticate: m365-mcp-auth login`
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
  if (existing && existing.expires_at > Date.now() + 5 * 60 * 1000 && hasRequestedScopes(existing.scope)) {
    return existing;
  }

  // Request device code
  const dc = await requestDeviceCode(config);
  onPrompt(dc.verification_uri, dc.user_code);

  // Poll
  const deadline = Date.now() + dc.expires_in * 1000;
  let interval = dc.interval ?? 5;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, interval * 1000));
    if (Date.now() >= deadline) break;
    try {
      const tokens = await pollForToken(dc.device_code, config);
      return saveTokens(config, tokens);
    } catch (err: any) {
      const errorCode = err instanceof OAuthError ? err.code : undefined;
      if (errorCode === "slow_down") { interval += 5; continue; }
      if (
        errorCode === "authorization_pending" ||
        err.message.includes("AADSTS70016")
      ) {
        continue;
      }
      if (
        errorCode === "authorization_declined" ||
        errorCode === "access_denied"
      ) {
        throw new Error("User declined authorization.");
      }
      if (
        errorCode === "expired_token" ||
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
