/**
 * Tests for auth.ts — account management, token storage, refresh logic
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import * as path from "node:path";
import * as fs from "node:fs";

// ─── mock filesystem ────────────────────────────────────────────────

const fsState = new Map<string, string>();

vi.mock("node:fs", () => ({
  existsSync: vi.fn((p: string) => fsState.has(p.toString())),
  mkdirSync: vi.fn(),
  readFileSync: vi.fn((p: string) => {
    const content = fsState.get(p.toString());
    if (content === undefined) throw new Error("ENOENT");
    return content;
  }),
  writeFileSync: vi.fn((p: string, data: string) => {
    fsState.set(p.toString(), data);
  }),
  renameSync: vi.fn((source: string, destination: string) => {
    const content = fsState.get(source);
    if (content === undefined) throw new Error("ENOENT");
    fsState.set(destination, content);
    fsState.delete(source);
  }),
  unlinkSync: vi.fn((p: string) => {
    fsState.delete(p.toString());
  }),
}));

vi.mock("node:os", () => ({
  homedir: vi.fn(() => "/home/testuser"),
}));

import {
  resolveAuthDir,
  ensureDirs,
  loadAccounts,
  addAccount,
  removeAccount,
  setDefaultAccount,
  resolveAccount,
  listAccounts,
  loadTokens,
  getTokenPath,
  getAccessToken,
  getTokenStatus,
  authenticate,
  resolveScopes,
} from "../auth.js";
import type { AccountConfig } from "../types.js";

// ─── helpers ────────────────────────────────────────────────────────

function seedTokens(accountName: string, tokens: { access_token: string; refresh_token: string; expires_at: number; scope: string }) {
  const tokenPath = getTokenPath(accountName);
  fsState.set(tokenPath, JSON.stringify(tokens));
}

function clearFs() {
  fsState.clear();
}

// ─── tests ──────────────────────────────────────────────────────────

describe("auth", () => {
  beforeEach(() => {
    clearFs();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // ── account store ───────────────────────────────────────────────

  describe("account management", () => {
    it("loadAccounts returns empty store when no file exists", () => {
      const store = loadAccounts();
      expect(store).toEqual({ default: null, accounts: {} });
    });

    it("addAccount creates store and sets as default if first", () => {
      addAccount("work", "tenant-1", "client-1", "me@corp.com", "Work");
      const store = loadAccounts();
      expect(store.default).toBe("work");
      expect(store.accounts["work"]).toMatchObject({
        tenantId: "tenant-1",
        clientId: "client-1",
        email: "me@corp.com",
        description: "Work",
      });
    });

    it("addAccount preserves existing default", () => {
      addAccount("work", "t1", "c1");
      addAccount("personal", "t2", "c2");
      const store = loadAccounts();
      expect(store.default).toBe("work");
      expect(Object.keys(store.accounts)).toHaveLength(2);
    });

    it("removeAccount removes account and token file", () => {
      addAccount("work", "t1", "c1");
      seedTokens("work", { access_token: "at", refresh_token: "rt", expires_at: Date.now() + 999999, scope: "openid" });

      removeAccount("work");
      const store = loadAccounts();
      expect(store.accounts["work"]).toBeUndefined();
      expect(store.default).toBeNull();

      // Token file should be gone
      const tokenPath = getTokenPath("work");
      expect(fsState.has(tokenPath)).toBe(false);
    });

    it("removeAccount updates default to next account", () => {
      addAccount("a", "t1", "c1");
      addAccount("b", "t2", "c2");
      removeAccount("a");
      expect(loadAccounts().default).toBe("b");
    });

    it("setDefaultAccount changes default", () => {
      addAccount("a", "t1", "c1");
      addAccount("b", "t2", "c2");
      setDefaultAccount("b");
      expect(loadAccounts().default).toBe("b");
    });

    it("setDefaultAccount throws for unknown account", () => {
      expect(() => setDefaultAccount("nope")).toThrow("not found");
    });

    it("resolveAccount returns default when no name given", () => {
      addAccount("work", "t1", "c1");
      const cfg = resolveAccount();
      expect(cfg.name).toBe("work");
      expect(cfg.tenantId).toBe("t1");
      expect(cfg.tokenPath).toContain("work.json");
    });

    it("resolveAccount returns named account", () => {
      addAccount("a", "t1", "c1");
      addAccount("b", "t2", "c2");
      const cfg = resolveAccount("b");
      expect(cfg.name).toBe("b");
    });

    it("resolveAccount throws when no accounts exist", () => {
      expect(() => resolveAccount()).toThrow("No account configured");
    });

    it("listAccounts returns full list with default marker", () => {
      addAccount("a", "t1", "c1");
      addAccount("b", "t2", "c2");
      const { default: def, accounts } = listAccounts();
      expect(def).toBe("a");
      expect(accounts).toHaveLength(2);
      expect(accounts[0].isDefault).toBe(true);
      expect(accounts[1].isDefault).toBe(false);
    });
  });

  // ── token store ──────────────────────────────────────────────────

  describe("token management", () => {
    it("loadTokens returns null when no file", () => {
      const cfg: AccountConfig = {
        name: "test",
        tenantId: "t1",
        clientId: "c1",
        addedAt: new Date().toISOString(),
        tokenPath: "/home/testuser/.m365-mcp/auth/tokens/test.json",
      };
      expect(loadTokens(cfg)).toBeNull();
    });

    it("loadTokens returns null when JSON is invalid", () => {
      const tokenPath = "/home/testuser/.m365-mcp/auth/tokens/bad.json";
      fsState.set(tokenPath, "not json");
      const cfg: AccountConfig = {
        name: "bad",
        tenantId: "t1",
        clientId: "c1",
        addedAt: new Date().toISOString(),
        tokenPath,
      };
      expect(loadTokens(cfg)).toBeNull();
    });

    it("loadTokens returns null when missing required fields", () => {
      const tokenPath = "/home/testuser/.m365-mcp/auth/tokens/partial.json";
      fsState.set(tokenPath, JSON.stringify({ access_token: "at" }));
      const cfg: AccountConfig = {
        name: "partial",
        tenantId: "t1",
        clientId: "c1",
        addedAt: new Date().toISOString(),
        tokenPath,
      };
      expect(loadTokens(cfg)).toBeNull();
    });

    it("loadTokens returns parsed tokens", () => {
      const tokenPath = "/home/testuser/.m365-mcp/auth/tokens/valid.json";
      const data = { access_token: "at", refresh_token: "rt", expires_at: Date.now() + 99999, scope: "openid" };
      fsState.set(tokenPath, JSON.stringify(data));
      const cfg: AccountConfig = {
        name: "valid",
        tenantId: "t1",
        clientId: "c1",
        addedAt: new Date().toISOString(),
        tokenPath,
      };
      const tokens = loadTokens(cfg);
      expect(tokens).toEqual(data);
    });
  });

  // ── getAccessToken ───────────────────────────────────────────────

  describe("getAccessToken", () => {
    it("throws when not authenticated", async () => {
      addAccount("work", "t1", "c1");
      await expect(getAccessToken("work")).rejects.toThrow("Not authenticated");
    });

    it("returns token when still valid with buffer", async () => {
      addAccount("work", "t1", "c1");
      const future = Date.now() + 30 * 60 * 1000; // 30 min from now
      seedTokens("work", { access_token: "valid-token", refresh_token: "rt", expires_at: future, scope: resolveScopes() });

      const token = await getAccessToken("work");
      expect(token).toBe("valid-token");
    });

    it("refreshes different accounts independently and preserves unrotated refresh tokens", async () => {
      addAccount("work", "t1", "c1");
      addAccount("personal", "t2", "c2");
      const expired = Date.now() - 60_000;
      seedTokens("work", { access_token: "old-work", refresh_token: "work-refresh", expires_at: expired, scope: "openid" });
      seedTokens("personal", { access_token: "old-personal", refresh_token: "personal-refresh", expires_at: expired, scope: "openid" });

      vi.spyOn(globalThis, "fetch").mockImplementation(async (_url, init) => {
        const body = new URLSearchParams(String(init?.body));
        const refreshToken = body.get("refresh_token")!;
        return {
          ok: true,
          json: async () => ({
            access_token: `${refreshToken}-access`,
            expires_in: 3600,
            scope: "openid",
          }),
        } as Response;
      });

      await expect(Promise.all([
        getAccessToken("work"),
        getAccessToken("personal"),
      ])).resolves.toEqual(["work-refresh-access", "personal-refresh-access"]);

      expect(loadTokens(resolveAccount("work"))?.refresh_token).toBe("work-refresh");
      expect(loadTokens(resolveAccount("personal"))?.refresh_token).toBe("personal-refresh");
      expect(fetch).toHaveBeenCalledTimes(2);
    });
  });

  describe("authenticate", () => {
    it("continues polling when Microsoft returns authorization_pending", async () => {
      addAccount("work", "tenant-1", "client-1");
      const onPrompt = vi.fn();

      vi.spyOn(globalThis, "fetch")
        .mockResolvedValueOnce({
          ok: true,
          json: async () => ({
            device_code: "device-code",
            user_code: "ABCD-EFGH",
            verification_uri: "https://microsoft.com/devicelogin",
            expires_in: 60,
            interval: 0,
          }),
        } as Response)
        .mockResolvedValueOnce({
          ok: false,
          statusText: "Bad Request",
          json: async () => ({
            error: "authorization_pending",
            error_description: "Waiting for the user to complete authentication.",
          }),
        } as Response)
        .mockResolvedValueOnce({
          ok: true,
          json: async () => ({
            access_token: "access-token",
            refresh_token: "refresh-token",
            expires_in: 3600,
            scope: "openid offline_access",
          }),
        } as Response);

      await expect(authenticate("work", onPrompt)).resolves.toMatchObject({
        access_token: "access-token",
        refresh_token: "refresh-token",
      });
      expect(onPrompt).toHaveBeenCalledWith(
        "https://microsoft.com/devicelogin",
        "ABCD-EFGH"
      );
      expect(fetch).toHaveBeenCalledTimes(3);
    });
  });

  // ── getTokenStatus ───────────────────────────────────────────────

  describe("getTokenStatus", () => {
    it("reports invalid when no tokens", async () => {
      addAccount("work", "t1", "c1");
      const status = await getTokenStatus("work");
      expect(status.valid).toBe(false);
      expect(status.expires).toBeNull();
    });

    it("reports valid when token is fresh", async () => {
      addAccount("work", "t1", "c1");
      const future = Date.now() + 60_000;
      seedTokens("work", { access_token: "at", refresh_token: "rt", expires_at: future, scope: "openid" });

      const status = await getTokenStatus("work");
      expect(status.valid).toBe(true);
      expect(status.expires).toBeTruthy();
    });

    it("reports valid = false when token is expired", async () => {
      addAccount("work", "t1", "c1");
      const past = Date.now() - 60_000;
      seedTokens("work", { access_token: "at", refresh_token: "rt", expires_at: past, scope: "openid" });

      const status = await getTokenStatus("work");
      expect(status.valid).toBe(false);
    });
  });

  describe("resolveAuthDir", () => {
    it("uses a client-neutral default under the home directory", () => {
      expect(resolveAuthDir({}, "/home/testuser")).toBe(path.join("/home/testuser", ".m365-mcp", "auth"));
    });

    it("honors M365_MCP_AUTH_DIR", () => {
      expect(resolveAuthDir({ M365_MCP_AUTH_DIR: "/srv/m365-auth" }, "/unused")).toBe(path.resolve("/srv/m365-auth"));
    });
  });

  describe("resolveScopes", () => {
    it("requests least-privilege read scopes in read-only mode", () => {
      const scopes = resolveScopes({ M365_MCP_READ_ONLY: "true" }).split(" ");
      expect(scopes).toContain("Mail.Read");
      expect(scopes).toContain("Tasks.Read");
      expect(scopes).toContain("People.Read");
      expect(scopes).not.toContain("User.Read.All");
      expect(scopes).not.toContain("Mail.Send");
      expect(scopes).not.toContain("Mail.ReadWrite");
      expect(scopes).not.toContain("Tasks.ReadWrite");
    });

    it("preserves write scopes by default", () => {
      const scopes = resolveScopes({}).split(" ");
      expect(scopes).toContain("Mail.Send");
      expect(scopes).toContain("Calendars.ReadWrite");
      expect(scopes).toContain("Chat.ReadWrite");
      expect(scopes).not.toContain("User.Read.All");
    });
  });

  // ── getTokenPath ─────────────────────────────────────────────────

  describe("getTokenPath", () => {
    it("returns path under tokens dir", () => {
      const p = getTokenPath("myaccount");
      expect(p).toBe(path.join("/home/testuser", ".m365-mcp", "auth", "tokens", "myaccount.json"));
    });
  });
});

describe("auth lifecycle regressions", () => {
  beforeEach(() => { fsState.clear(); });
  afterEach(() => { vi.restoreAllMocks(); vi.useRealTimers(); });

  it.each(["../other", "a/b", "a\\b", "..", "", "bad:name", "name."])("rejects unsafe account label %s", (name) => {
    expect(() => addAccount(name, "tenant", "client")).toThrow("Account name");
  });

  it("invalidates credentials when the tenant/client registration changes", () => {
    addAccount("work", "tenant-a", "client-a");
    seedTokens("work", { access_token: "old", refresh_token: "old-refresh", expires_at: Date.now() + 3600000, scope: resolveScopes() });
    addAccount("work", "tenant-b", "client-b");
    expect(loadTokens(resolveAccount("work"))).toBeNull();
  });

  it("preserves credentials for an unchanged registration", () => {
    addAccount("work", "tenant", "client");
    seedTokens("work", { access_token: "token", refresh_token: "refresh", expires_at: Date.now() + 3600000, scope: resolveScopes() });
    addAccount("work", "tenant", "client", undefined, "Updated label");
    expect(loadTokens(resolveAccount("work"))?.access_token).toBe("token");
  });

  it("does not overwrite a corrupted account registry", () => {
    const registry = path.join(resolveAuthDir(), "accounts.json");
    fsState.set(registry, "{broken");
    expect(() => addAccount("work", "tenant", "client")).toThrow("Cannot read account registry");
    expect(fsState.get(registry)).toBe("{broken");
  });

  it("keeps the previous registry intact if atomic replacement fails", () => {
    addAccount("work", "tenant", "client");
    const registry = path.join(resolveAuthDir(), "accounts.json");
    const before = fsState.get(registry);
    vi.mocked(fs.renameSync).mockImplementationOnce(() => { throw new Error("disk failure"); });
    expect(() => addAccount("new", "tenant", "client")).toThrow("disk failure");
    expect(fsState.get(registry)).toBe(before);
    expect([...fsState.keys()].some((name) => name.endsWith(".tmp"))).toBe(false);
  });

  it("requires fresh authentication for cached credentials with missing scopes", async () => {
    addAccount("work", "tenant", "client");
    seedTokens("work", { access_token: "old", refresh_token: "refresh", expires_at: Date.now() + 3600000, scope: "User.Read" });
    await expect(getAccessToken("work")).rejects.toThrow("Permissions changed");
  });

  it("requests file writes only in write mode and directory access only by opt-in", () => {
    expect(resolveScopes({})).toContain("Files.ReadWrite");
    expect(resolveScopes({ M365_MCP_READ_ONLY: "true" })).not.toContain("Files.ReadWrite");
    expect(resolveScopes({})).not.toContain("User.Read.All");
    expect(resolveScopes({ M365_MCP_DIRECTORY_READ: "true" })).toContain("User.Read.All");
  });

  it("re-prompts when a valid cached token lacks requested permissions", async () => {
    vi.useFakeTimers();
    addAccount("work", "tenant", "client");
    seedTokens("work", { access_token: "old", refresh_token: "refresh", expires_at: Date.now() + 3600000, scope: "User.Read" });
    vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response(JSON.stringify({ device_code: "device", user_code: "CODE", verification_uri: "https://microsoft.com/devicelogin", expires_in: 60, interval: 1 })))
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "new", refresh_token: "new-refresh", expires_in: 3600, scope: resolveScopes() })));
    const prompt = vi.fn();
    const request = authenticate("work", prompt);
    await vi.runAllTimersAsync();
    await expect(request).resolves.toMatchObject({ access_token: "new" });
    expect(prompt).toHaveBeenCalledOnce();
  });

  it("increases the polling interval after slow_down", async () => {
    vi.useFakeTimers();
    addAccount("work", "tenant", "client");
    const fetchMock = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response(JSON.stringify({ device_code: "device", user_code: "CODE", verification_uri: "https://microsoft.com/devicelogin", expires_in: 60, interval: 1 })))
      .mockResolvedValueOnce(new Response(JSON.stringify({ error: "slow_down" }), { status: 400 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "new", refresh_token: "refresh", expires_in: 3600, scope: resolveScopes() })));
    const request = authenticate("work", vi.fn());
    await vi.advanceTimersByTimeAsync(1000);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    await vi.advanceTimersByTimeAsync(5999);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    await vi.advanceTimersByTimeAsync(1);
    await expect(request).resolves.toMatchObject({ access_token: "new" });
  });
});
