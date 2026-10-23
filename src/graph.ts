/**
 * Typed Microsoft Graph API client with:
 * - Automatic pagination (@odata.nextLink)
 * - Rate-limit handling (429 → exponential backoff + retry)
 * - Transient error retry (5xx)
 * - Configurable timezone
 */

import type { GraphError, GraphListResponse } from "./types.js";
import { getAccessToken } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES = 4;
const RETRY_BACKOFF_MS = 1000;
const REQUEST_TIMEOUT_MS = 30_000;
const MAX_DOWNLOAD_BYTES = 1_000_000;
const TRANSIENT_STATUSES = new Set([500, 502, 503, 504]);

function retryDelayMs(retryAfter: string | null, fallbackMs: number): number {
  if (!retryAfter?.trim()) return fallbackMs;

  const seconds = Number(retryAfter);
  if (Number.isFinite(seconds) && seconds >= 0) return seconds * 1000;

  if (/^-?\d+(\.\d+)?$/.test(retryAfter.trim())) return fallbackMs;
  const retryAt = Date.parse(retryAfter);
  return Number.isNaN(retryAt) ? fallbackMs : Math.max(0, retryAt - Date.now());
}

export class GraphClient {
  timezone: string;

  constructor(timezone = "UTC", private accountName?: string) {
    this.timezone = timezone;
  }

  // ─── request ────────────────────────────────────────────────────

  private async request<T>(
    path: string,
    method = "GET",
    body?: unknown,
    raw = false,
    additionalHeaders: Readonly<Record<string, string>> = {},
    retrySafe = method === "GET"
  ): Promise<T> {
    const url = path.startsWith("https://") ? path : `${GRAPH_BASE}${path}`;
    if (new URL(url).origin !== new URL(GRAPH_BASE).origin) {
      throw new Error("Graph API request URL must use graph.microsoft.com");
    }

    let attempt = 0;
    let delay = RETRY_BACKOFF_MS;

    while (true) {
      const headers = new Headers({
        "Content-Type": "application/json",
        Accept: "application/json",
        Prefer: `outlook.timezone="${this.timezone}"`,
      });
      for (const [key, value] of Object.entries(additionalHeaders)) headers.set(key, value);
      // Headers is case-insensitive, so callers cannot override the selected token.
      headers.set("Authorization", `Bearer ${await getAccessToken(this.accountName)}`);
      const opts: RequestInit = {
        method,
        headers,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      };
      if (body !== undefined && method !== "GET") {
        opts.body = JSON.stringify(body);
      }

      let resp: Response;
      try {
        resp = await fetch(url, opts);
      } catch (error) {
        if (!retrySafe) {
          throw new Error("Graph write outcome is unknown after a connection failure. Check the target before retrying.", { cause: error });
        }
        if (attempt >= MAX_RETRIES) throw error;
        await new Promise((r) => setTimeout(r, delay));
        attempt++;
        delay *= 2;
        continue;
      }

      // 204 No Content
      if (resp.status === 204) return undefined as T;

      // Rate-limited — retry with Retry-After header or exponential backoff
      if (resp.status === 429) {
        if (attempt >= MAX_RETRIES) {
          throw new Error(`Rate limited after ${MAX_RETRIES} retries`);
        }
        const retryAfter = resp.headers.get("Retry-After");
        const waitMs = retryDelayMs(retryAfter, delay);
        await resp.body?.cancel();
        await new Promise((r) => setTimeout(r, waitMs));
        attempt++;
        delay *= 2;
        continue;
      }

      // Transient server error — retry
      if (TRANSIENT_STATUSES.has(resp.status) && retrySafe) {
        if (attempt >= MAX_RETRIES) {
          throw new Error(`Graph API server error ${resp.status} after ${MAX_RETRIES} retries`);
        }
        const waitMs = retryDelayMs(resp.headers.get("Retry-After"), delay);
        await resp.body?.cancel();
        await new Promise((r) => setTimeout(r, waitMs));
        attempt++;
        delay *= 2;
        continue;
      }

      // Parse response
      if (raw && resp.ok) {
        return resp as unknown as T;
      }

      let parsed: any;
      try {
        parsed = await resp.json();
      } catch {
        // Graph action endpoints such as sendMail commonly return 202 with
        // an empty body. A successful empty response is not an error.
        if (resp.status === 202) return undefined as T;
        throw new Error(`Graph API ${resp.status}: unable to parse response`);
      }

      if (!resp.ok) {
        const err = parsed?.error as GraphError["error"] | undefined;
        throw new Error(
          `Graph API ${resp.status}${err?.code ? ` [${err.code}]` : ""}: ${err?.message ?? resp.statusText}${!retrySafe && resp.status >= 500 ? ". Write outcome may be unknown; check the target before retrying." : ""}`
        );
      }

      return parsed as T;
    }
  }

  // ─── public methods ──────────────────────────────────────────────

  /** Single-page GET */
  async get<T>(
    path: string,
    additionalHeaders: Readonly<Record<string, string>> = {}
  ): Promise<T> {
    return this.request<T>(path, "GET", undefined, false, additionalHeaders);
  }

  /** GET that follows @odata.nextLink pagination, collecting all results */
  async getAll<T>(
    path: string,
    maxPages = 10,
    maxItems = Infinity,
    additionalHeaders: Readonly<Record<string, string>> = {}
  ): Promise<T[]> {
    if (!Number.isInteger(maxPages) || maxPages < 1) throw new Error("maxPages must be a positive integer");
    if (maxItems !== Infinity && (!Number.isInteger(maxItems) || maxItems < 0)) throw new Error("maxItems must be a nonnegative integer");
    const results: T[] = [];
    let url: string | undefined = path;
    let pages = 0;

    while (url && pages < maxPages && results.length < maxItems) {
      // If url is a nextLink, it's already absolute
      const data: GraphListResponse<T> = await this.request<GraphListResponse<T>>(
        url,
        "GET",
        undefined,
        false,
        additionalHeaders
      );
      if (data.value) results.push(...data.value.slice(0, maxItems - results.length));
      url = data["@odata.nextLink"];
      pages++;
    }

    return results;
  }

  /** Paginated GET — yields pages lazily */
  async *paginate<T>(
    path: string,
    maxPages = 10,
    additionalHeaders: Readonly<Record<string, string>> = {}
  ): AsyncGenerator<T[]> {
    let url: string | undefined = path;
    let pages = 0;

    while (url && pages < maxPages) {
      const data: GraphListResponse<T> = await this.request<GraphListResponse<T>>(
        url,
        "GET",
        undefined,
        false,
        additionalHeaders
      );
      if (data.value) yield data.value;
      url = data["@odata.nextLink"];
      pages++;
    }
  }

  /** POST */
  async post<T>(path: string, body: unknown, options: { retrySafe?: boolean } = {}): Promise<T> {
    return this.request<T>(path, "POST", body, false, {}, options.retrySafe ?? false);
  }

  /** PATCH */
  async patch<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>(path, "PATCH", body);
  }

  /** DELETE */
  async delete(path: string): Promise<void> {
    await this.request(path, "DELETE");
  }

  /** Download raw file content with retry on transient failures */
  async download(url: string, maxBytes = 50_000): Promise<string> {
    if (!Number.isInteger(maxBytes) || maxBytes < 0 || maxBytes > MAX_DOWNLOAD_BYTES) {
      throw new Error(`maxBytes must be an integer between 0 and ${MAX_DOWNLOAD_BYTES}`);
    }
    const downloadUrl = new URL(url);
    if (downloadUrl.protocol !== "https:") {
      throw new Error("Download URL must use HTTPS");
    }
    if (maxBytes === 0) return "";

    // Graph may return a pre-authenticated SharePoint/OneDrive download URL.
    // Never forward the Microsoft Graph bearer token to another origin.
    const headers: HeadersInit = {};
    if (downloadUrl.origin === new URL(GRAPH_BASE).origin) {
      const token = await getAccessToken(this.accountName);
      headers.Authorization = `Bearer ${token}`;
    }

    let attempt = 0;
    let delay = RETRY_BACKOFF_MS;

    while (true) {
      let resp: Response;
      try {
        resp = await fetch(url, { headers, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
      } catch (error) {
        if (attempt >= MAX_RETRIES) throw error;
        await new Promise((r) => setTimeout(r, delay));
        attempt++;
        delay *= 2;
        continue;
      }

      if (resp.status === 429 && attempt < MAX_RETRIES) {
        const retryAfter = resp.headers.get("Retry-After");
        await resp.body?.cancel();
        await new Promise((r) => setTimeout(r, retryDelayMs(retryAfter, delay)));
        attempt++;
        delay *= 2;
        continue;
      }

      if (TRANSIENT_STATUSES.has(resp.status) && attempt < MAX_RETRIES) {
        const waitMs = retryDelayMs(resp.headers.get("Retry-After"), delay);
        await resp.body?.cancel();
        await new Promise((r) => setTimeout(r, waitMs));
        attempt++;
        delay *= 2;
        continue;
      }

      if (!resp.ok) {
        throw new Error(`Download failed: ${resp.status} ${resp.statusText}`);
      }

      if (!resp.body) return "";
      const reader = resp.body.getReader();
      const bytes = new Uint8Array(maxBytes);
      let length = 0;
      try {
        while (length < maxBytes) {
          const { value, done } = await reader.read();
          if (done) break;
          const take = Math.min(value.byteLength, maxBytes - length);
          bytes.set(value.subarray(0, take), length);
          length += take;
        }
      } finally {
        // Cancellation closes the connection even if the server ignores Range.
        try { await reader.cancel(); } catch { /* Preserve a read failure. */ }
        reader.releaseLock();
      }
      return new TextDecoder().decode(bytes.subarray(0, length));
    }
  }

  /** Format a dateTimeTimeZone object for Graph API */
  dateTimeStr(iso: string): { dateTime: string; timeZone: string } {
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,7})?)?(?:Z|[+-]\d{2}:\d{2})?$/i.test(iso) || !Number.isFinite(Date.parse(iso))) {
      throw new Error("Expected an ISO 8601 date-time");
    }
    // An explicit offset is an instant; do not reinterpret it in the configured zone.
    if (/(?:Z|[+-]\d{2}:\d{2})$/i.test(iso)) {
      return { dateTime: new Date(iso).toISOString().replace(/Z$/, ""), timeZone: "UTC" };
    }
    return { dateTime: iso, timeZone: this.timezone };
  }
}
