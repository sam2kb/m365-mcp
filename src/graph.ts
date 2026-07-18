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

function retryDelayMs(retryAfter: string | null, fallbackMs: number): number {
  if (!retryAfter) return fallbackMs;

  const seconds = Number(retryAfter);
  if (Number.isFinite(seconds) && seconds >= 0) return seconds * 1000;

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
    raw = false
  ): Promise<T> {
    const token = await getAccessToken(this.accountName);
    const url = path.startsWith("https://") ? path : `${GRAPH_BASE}${path}`;
    if (new URL(url).origin !== new URL(GRAPH_BASE).origin) {
      throw new Error("Graph API request URL must use graph.microsoft.com");
    }

    let attempt = 0;
    let delay = RETRY_BACKOFF_MS;

    while (true) {
      const opts: RequestInit = {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          Accept: "application/json",
          Prefer: `outlook.timezone="${this.timezone}"`,
        },
      };
      if (body && method !== "GET") {
        opts.body = JSON.stringify(body);
      }

      let resp: Response;
      try {
        resp = await fetch(url, opts);
      } catch (error) {
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
        await new Promise((r) => setTimeout(r, waitMs));
        attempt++;
        delay *= 2;
        continue;
      }

      // Transient server error — retry
      if (resp.status >= 500 && resp.status < 600) {
        if (attempt >= MAX_RETRIES) {
          throw new Error(`Graph API server error ${resp.status} after ${MAX_RETRIES} retries`);
        }
        await new Promise((r) => setTimeout(r, delay));
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
        if (resp.ok) return undefined as T;
        throw new Error(`Graph API ${resp.status}: unable to parse response`);
      }

      if (!resp.ok) {
        const err = parsed?.error as GraphError["error"] | undefined;
        throw new Error(
          `Graph API ${resp.status}: ${err?.message ?? resp.statusText}`
        );
      }

      return parsed as T;
    }
  }

  // ─── public methods ──────────────────────────────────────────────

  /** Single-page GET */
  async get<T>(path: string): Promise<T> {
    return this.request<T>(path);
  }

  /** GET that follows @odata.nextLink pagination, collecting all results */
  async getAll<T>(path: string, maxPages = 10, maxItems = Infinity): Promise<T[]> {
    const results: T[] = [];
    let url: string | undefined = path;
    let pages = 0;

    while (url && pages < maxPages && results.length < maxItems) {
      // If url is a nextLink, it's already absolute
      const data: GraphListResponse<T> = await this.request<GraphListResponse<T>>(url);
      if (data.value) results.push(...data.value.slice(0, maxItems - results.length));
      url = data["@odata.nextLink"];
      pages++;
    }

    return results;
  }

  /** Paginated GET — yields pages lazily */
  async *paginate<T>(path: string, maxPages = 10): AsyncGenerator<T[]> {
    let url: string | undefined = path;
    let pages = 0;

    while (url && pages < maxPages) {
      const data: GraphListResponse<T> = await this.request<GraphListResponse<T>>(url);
      if (data.value) yield data.value;
      url = data["@odata.nextLink"];
      pages++;
    }
  }

  /** POST */
  async post<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>(path, "POST", body);
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
    const token = await getAccessToken(this.accountName);

    let attempt = 0;
    let delay = RETRY_BACKOFF_MS;

    while (true) {
      const resp = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (resp.status === 429 && attempt < MAX_RETRIES) {
        const retryAfter = resp.headers.get("Retry-After");
        await new Promise((r) => setTimeout(r, retryDelayMs(retryAfter, delay)));
        attempt++;
        delay *= 2;
        continue;
      }

      if (resp.status >= 500 && resp.status < 600 && attempt < MAX_RETRIES) {
        await new Promise((r) => setTimeout(r, delay));
        attempt++;
        delay *= 2;
        continue;
      }

      if (!resp.ok) {
        throw new Error(`Download failed: ${resp.status} ${resp.statusText}`);
      }

      const buf = await resp.arrayBuffer();
      // Truncate to text-safe size
      const slice = buf.byteLength > maxBytes ? buf.slice(0, maxBytes) : buf;
      return new TextDecoder().decode(slice);
    }
  }

  /** Format a dateTimeTimeZone object for Graph API */
  dateTimeStr(iso: string): { dateTime: string; timeZone: string } {
    return { dateTime: iso, timeZone: this.timezone };
  }
}
