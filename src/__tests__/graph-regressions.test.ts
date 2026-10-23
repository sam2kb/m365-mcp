import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
vi.mock("../auth.js", () => ({ getAccessToken: vi.fn().mockResolvedValue("selected-account-token") }));
import { GraphClient } from "../graph.js";
import { getAccessToken } from "../auth.js";

beforeEach(() => { vi.mocked(getAccessToken).mockResolvedValue("selected-account-token"); });
afterEach(() => { vi.restoreAllMocks(); vi.useRealTimers(); });

describe("request safety", () => {
  it("does not replay a write after the server may have committed it", async () => {
    let writes = 0;
    vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
      writes++;
      throw new TypeError("connection lost after committing the operation");
    });
    await expect(new GraphClient().post("/me/sendMail", { message: {} })).rejects.toThrow("outcome is unknown");
    expect(writes).toBe(1);
  });
  it.each([500, 502, 503, 504])("does not replay a write after HTTP %s", async (status) => {
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response('{"error":{"code":"Unavailable","message":"Failure"}}', { status }));
    await expect(new GraphClient().post("/me/events", { subject: "Meeting" })).rejects.toThrow("check the target");
    expect(fetchMock).toHaveBeenCalledOnce();
  });
  it.each(["get", "download"] as const)("honors 503 Retry-After for %s", async (method) => {
    vi.useFakeTimers();
    const fetchMock = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response("busy", { status: 503, headers: { "Retry-After": "60" } }))
      .mockResolvedValueOnce(new Response(method === "get" ? '{"ok":true}' : "ok"));
    const client = new GraphClient();
    const request = method === "get" ? client.get("/me") : client.download("https://example.com/file");
    await vi.advanceTimersByTimeAsync(59_999);
    expect(fetchMock).toHaveBeenCalledOnce();
    await vi.advanceTimersByTimeAsync(1);
    await request;
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
  it("honors an HTTP-date Retry-After and permits safe read-only POST retry", async () => {
    vi.useFakeTimers(); vi.setSystemTime(new Date("2026-01-01T00:00:00Z"));
    const fetchMock = vi.spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(new Response("busy", { status: 503, headers: { "Retry-After": "Thu, 01 Jan 2026 00:00:10 GMT" } }))
      .mockResolvedValueOnce(new Response('{"value":[]}'));
    const request = new GraphClient().post("/me/calendar/getSchedule", {}, { retrySafe: true });
    await vi.advanceTimersByTimeAsync(9999);
    expect(fetchMock).toHaveBeenCalledOnce();
    await vi.advanceTimersByTimeAsync(1);
    await expect(request).resolves.toEqual({ value: [] });
  });
  it("does not swallow malformed JSON on a successful resource response", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response("broken", { status: 200 }));
    await expect(new GraphClient().get("/me")).rejects.toThrow("unable to parse");
  });
  it("protects Authorization case-insensitively while allowing a timezone override", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response("{}"));
    await new GraphClient("America/New_York").get("/me", { authorization: "forged", Prefer: 'outlook.timezone="UTC"' });
    const headers = new Headers(fetchMock.mock.calls[0][1]?.headers);
    expect(headers.get("Authorization")).toBe("Bearer selected-account-token");
    expect(headers.get("Prefer")).toBe('outlook.timezone="UTC"');
    expect(fetchMock.mock.calls[0][1]?.signal).toBeInstanceOf(AbortSignal);
  });
  it("rejects an off-origin pagination link before fetching it", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(JSON.stringify({ value: [], "@odata.nextLink": "https://attacker.invalid/" })));
    await expect(new GraphClient().getAll("/me/messages")).rejects.toThrow("graph.microsoft.com");
    expect(fetchMock).toHaveBeenCalledOnce();
  });
});

describe("bounded downloads", () => {
  it.each([-1, 1.5, Infinity, NaN, 1_000_001, "10"])("rejects invalid maxBytes %s before fetching", async (limit) => {
    const fetchMock = vi.spyOn(globalThis, "fetch");
    await expect(new GraphClient().download("https://example.com/file", limit as number)).rejects.toThrow("maxBytes");
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("returns zero bytes without a network request", async () => {
    const fetchMock = vi.spyOn(globalThis, "fetch");
    expect(await new GraphClient().download("https://example.com/file", 0)).toBe("");
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("cancels an oversized response at the limit without reading another chunk", async () => {
    const pull = vi.fn((controller: ReadableStreamDefaultController<Uint8Array>) => controller.enqueue(new TextEncoder().encode("abcdef")));
    const cancel = vi.fn();
    const body = new ReadableStream({ pull, cancel }, { highWaterMark: 0 });
    vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(body));
    expect(await new GraphClient().download("https://example.com/file", 3)).toBe("abc");
    expect(pull).toHaveBeenCalledOnce();
    expect(cancel).toHaveBeenCalledOnce();
  });
  it("cancels at an exact chunk boundary", async () => {
    const cancel = vi.fn();
    let pulls = 0;
    const body = new ReadableStream({ pull(controller) { pulls++; controller.enqueue(new Uint8Array([65, 66])); }, cancel }, { highWaterMark: 0 });
    vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(body));
    expect(await new GraphClient().download("https://example.com/file", 2)).toBe("AB");
    expect(pulls).toBe(1); expect(cancel).toHaveBeenCalledOnce();
  });
  it("decodes UTF-8 split across chunks and preserves truncated-prefix behavior", async () => {
    const stream = () => new ReadableStream<Uint8Array>({ start(controller) {
      controller.enqueue(new Uint8Array([0xe2])); controller.enqueue(new Uint8Array([0x82, 0xac])); controller.close();
    } });
    vi.spyOn(globalThis, "fetch").mockImplementation(async () => new Response(stream()));
    expect(await new GraphClient().download("https://example.com/file", 3)).toBe("€");
    expect(await new GraphClient().download("https://example.com/file", 2)).toBe("�");
  });
  it("propagates a stream failure", async () => {
    const body = new ReadableStream({ start(controller) { controller.error(new Error("broken stream")); } });
    vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(body));
    await expect(new GraphClient().download("https://example.com/file")).rejects.toThrow("broken stream");
  });
  it("accepts an empty response body", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(null));
    expect(await new GraphClient().download("https://example.com/file")).toBe("");
  });
});

it("normalizes explicit datetime offsets instead of relabeling them", () => {
  expect(new GraphClient("America/Chicago").dateTimeStr("2026-07-18T14:00:00+02:00")).toEqual({ dateTime: "2026-07-18T12:00:00.000", timeZone: "UTC" });
});
