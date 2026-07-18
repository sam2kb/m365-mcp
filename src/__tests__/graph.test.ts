/**
 * Tests for GraphClient — retry logic, pagination, download, error handling
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

// We mock the auth module so getAccessToken just returns a fake token
vi.mock("../auth.js", () => ({
  getAccessToken: vi.fn().mockResolvedValue("fake-token"),
}));

import { GraphClient } from "../graph.js";

// ─── helpers ────────────────────────────────────────────────────────

function mockFetch(responses: Array<{ status: number; body?: unknown; headers?: Record<string, string>; jsonError?: boolean }>) {
  let call = 0;
  vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
    const r = responses[call] ?? responses[responses.length - 1];
    call++;
    const headers = new Headers(r.headers ?? { "Content-Type": "application/json" });
    return {
      status: r.status,
      ok: r.status >= 200 && r.status < 300,
      headers,
      json: async () => {
        if (r.jsonError) throw new SyntaxError("Unexpected end of JSON input");
        return r.body;
      },
      arrayBuffer: async () => {
        const enc = new TextEncoder();
        return enc.encode(String(r.body)).buffer as ArrayBuffer;
      },
      statusText: r.status === 429 ? "Too Many Requests" : r.status >= 500 ? "Server Error" : "OK",
    } as Response;
  });
  return call;
}

// ─── tests ──────────────────────────────────────────────────────────

describe("GraphClient", () => {
  let client: GraphClient;

  beforeEach(() => {
    client = new GraphClient("America/Chicago", "test-account");
    vi.restoreAllMocks();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.useRealTimers();
  });

  // ── request ─────────────────────────────────────────────────────

  describe("request retry logic", () => {
    it("returns parsed JSON on 200", async () => {
      mockFetch([{ status: 200, body: { value: "ok" } }]);
      const result = await (client as any).request("/test");
      expect(result).toEqual({ value: "ok" });
    });

    it("retries on 429 with Retry-After header", async () => {
      mockFetch([
        { status: 429, body: {}, headers: { "Retry-After": "0" } },
        { status: 200, body: { value: "retried" } },
      ]);
      const result = await (client as any).request("/test");
      expect(result).toEqual({ value: "retried" });
      expect(fetch).toHaveBeenCalledTimes(2);
    });

    it("retries on 429 with exponential backoff when no Retry-After", async () => {
      mockFetch([
        { status: 429, body: {} },
        { status: 429, body: {} },
        { status: 200, body: { value: "finally" } },
      ]);
      const result = await (client as any).request("/test");
      expect(result).toEqual({ value: "finally" });
      expect(fetch).toHaveBeenCalledTimes(3);
    });

    it("throws after MAX_RETRIES 429s", async () => {
      vi.useFakeTimers();
      const responses = Array.from({ length: 5 }, () => ({ status: 429, body: {} }));
      mockFetch(responses);
      const promise = expect((client as any).request("/test")).rejects.toThrow("Rate limited");
      await vi.runAllTimersAsync();
      await promise;
    });

    it("retries on 5xx server errors", async () => {
      mockFetch([
        { status: 503, body: {} },
        { status: 200, body: { value: "ok" } },
      ]);
      const result = await (client as any).request("/test");
      expect(result).toEqual({ value: "ok" });
      expect(fetch).toHaveBeenCalledTimes(2);
    });

    it("throws on 5xx after max retries", async () => {
      vi.useFakeTimers();
      const responses = Array.from({ length: 5 }, () => ({ status: 500, body: {} }));
      mockFetch(responses);
      const promise = expect((client as any).request("/test")).rejects.toThrow("server error");
      await vi.runAllTimersAsync();
      await promise;
    });

    it("throws on 4xx errors (no retry)", async () => {
      mockFetch([{ status: 400, body: { error: { message: "Bad request" } } }]);
      await expect((client as any).request("/test")).rejects.toThrow("Bad request");
    });

    it("handles 204 No Content gracefully", async () => {
      mockFetch([{ status: 204, body: null }]);
      const result = await (client as any).request("/test");
      expect(result).toBeUndefined();
    });

    it("accepts successful Graph actions with an empty response body", async () => {
      mockFetch([{ status: 202, jsonError: true }]);
      const result = await client.post("/me/sendMail", { message: {} });
      expect(result).toBeUndefined();
    });
  });

  // ── getAll (pagination) ──────────────────────────────────────────

  describe("getAll pagination", () => {
    it("collects all pages via @odata.nextLink", async () => {
      let fetchCount = 0;
      vi.spyOn(globalThis, "fetch").mockImplementation(async (url) => {
        fetchCount++;
        if (fetchCount === 1) {
          return {
            status: 200,
            ok: true,
            headers: new Headers({ "Content-Type": "application/json" }),
            json: async () => ({
              value: [{ id: "1" }, { id: "2" }],
              "@odata.nextLink": "https://graph.microsoft.com/v1.0/next",
            }),
          } as Response;
        }
        return {
          status: 200,
          ok: true,
          headers: new Headers({ "Content-Type": "application/json" }),
          json: async () => ({ value: [{ id: "3" }] }),
        } as Response;
      });

      const results = await client.getAll("/test", 5);
      expect(results).toHaveLength(3);
      expect(results.map((r: any) => r.id)).toEqual(["1", "2", "3"]);
    });

    it("respects maxPages limit", async () => {
      let fetchCount = 0;
      vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
        fetchCount++;
        return {
          status: 200,
          ok: true,
          headers: new Headers({ "Content-Type": "application/json" }),
          json: async () => ({
            value: [{ id: String(fetchCount) }],
            "@odata.nextLink": `https://graph.microsoft.com/v1.0/page${fetchCount + 1}`,
          }),
        } as Response;
      });

      const results = await client.getAll("/test", 3);
      expect(results).toHaveLength(3);
      expect(fetchCount).toBe(3);
    });

    it("stops pagination at the requested item limit", async () => {
      let fetchCount = 0;
      vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
        fetchCount++;
        return {
          status: 200,
          ok: true,
          headers: new Headers({ "Content-Type": "application/json" }),
          json: async () => ({
            value: [{ id: "1" }, { id: "2" }, { id: "3" }],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/next",
          }),
        } as Response;
      });

      const results = await client.getAll("/test", 10, 2);
      expect(results.map((result: any) => result.id)).toEqual(["1", "2"]);
      expect(fetchCount).toBe(1);
    });
  });

  // ── paginate (generator) ─────────────────────────────────────────

  describe("paginate generator", () => {
    it("yields pages lazily", async () => {
      let fetchCount = 0;
      vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
        fetchCount++;
        const body: any = { value: [{ id: String(fetchCount) }] };
        if (fetchCount < 3) {
          body["@odata.nextLink"] = `https://graph.microsoft.com/v1.0/page${fetchCount + 1}`;
        }
        return {
          status: 200,
          ok: true,
          headers: new Headers({ "Content-Type": "application/json" }),
          json: async () => body,
        } as Response;
      });

      const pages: any[][] = [];
      for await (const page of client.paginate("/test", 5)) {
        pages.push(page);
      }
      expect(pages).toHaveLength(3);
      expect(pages.flat().map((i: any) => i.id)).toEqual(["1", "2", "3"]);
    });
  });

  // ── HTTP methods ─────────────────────────────────────────────────

  describe("HTTP verb helpers", () => {
    it("get delegates to request", async () => {
      mockFetch([{ status: 200, body: { data: "get" } }]);
      const result = await client.get("/test");
      expect(result).toEqual({ data: "get" });
    });

    it("post sends JSON body", async () => {
      mockFetch([{ status: 200, body: { id: "new" } }]);
      const result = await client.post("/test", { name: "x" });
      expect(result).toEqual({ id: "new" });
    });

    it("patch sends PATCH", async () => {
      mockFetch([{ status: 200, body: { updated: true } }]);
      const result = await client.patch("/test", { name: "y" });
      expect(result).toEqual({ updated: true });
    });

    it("delete sends DELETE", async () => {
      mockFetch([{ status: 204, body: null }]);
      await expect(client.delete("/test")).resolves.toBeUndefined();
    });
  });

  // ── download ─────────────────────────────────────────────────────

  describe("download", () => {
    it("downloads and decodes text content", async () => {
      mockFetch([{ status: 200, body: "hello world" }]);
      const content = await client.download("https://example.com/file.txt");
      expect(content).toBe("hello world");
    });

    it("truncates to maxBytes", async () => {
      const longText = "x".repeat(60_000);
      mockFetch([{ status: 200, body: longText }]);
      const content = await client.download("https://example.com/file.txt", 10_000);
      expect(content.length).toBeLessThanOrEqual(10_000);
    });

    it("retries on 429", async () => {
      mockFetch([
        { status: 429, body: "", headers: { "Retry-After": "0" } },
        { status: 200, body: "ok" },
      ]);
      const content = await client.download("https://example.com/file.txt");
      expect(content).toBe("ok");
      expect(fetch).toHaveBeenCalledTimes(2);
    });

    it("throws on non-retryable error", async () => {
      mockFetch([{ status: 404, body: "" }]);
      await expect(client.download("https://example.com/missing.txt")).rejects.toThrow("Download failed");
    });
  });

  // ── dateTimeStr ──────────────────────────────────────────────────

  describe("dateTimeStr", () => {
    it("wraps ISO string with configured timezone", () => {
      const result = client.dateTimeStr("2026-07-18T14:00:00");
      expect(result).toEqual({
        dateTime: "2026-07-18T14:00:00",
        timeZone: "America/Chicago",
      });
    });
  });
});
