/**
 * Tests for OneDrive files tools
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerFilesTools, filesToolSchemas } from "../tools/files.js";

function mockClient(overrides: Partial<GraphClient> = {}): GraphClient {
  return {
    timezone: "UTC",
    getAll: vi.fn().mockResolvedValue([]),
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    paginate: vi.fn(),
    download: vi.fn(),
    dateTimeStr: vi.fn((iso: string) => ({ dateTime: iso, timeZone: "UTC" })),
    ...overrides,
  } as unknown as GraphClient;
}

describe("files tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerFilesTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerFilesTools(client);
  });

  describe("files_list", () => {
    it("lists root by default", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.files_list({});
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("/drive/root/children");
    });

    it("lists subfolder when path provided", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.files_list({ path: "Shared Files/2026 #1" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain(":/Shared%20Files/2026%20%231:");
    });

    it("formats items with type indicators", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "f1", name: "report.docx", size: 12345, folder: undefined, file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, webUrl: "https://...", lastModifiedDateTime: "2026-07-18T10:00:00Z" },
        { id: "f2", name: "Photos", size: 0, folder: { childCount: 15 }, file: undefined, webUrl: "https://...", lastModifiedDateTime: "2026-07-17T10:00:00Z" },
      ]);
      const result = await tools.files_list({});
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "f1", name: "report.docx", isFolder: false });
      expect(parsed[1]).toMatchObject({ id: "f2", name: "Photos", isFolder: true, childCount: 15 });
    });
  });

  describe("files_search", () => {
    it("searches with query", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.files_search({ query: "O'Brien" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("search(q='O%27%27Brien')");
    });
  });

  describe("files_read", () => {
    it("downloads content via @microsoft.graph.downloadUrl", async () => {
      (client.get as any).mockResolvedValue({ "@microsoft.graph.downloadUrl": "https://download.url" });
      (client.download as any).mockResolvedValue("file content here");
      const result = await tools.files_read({ itemId: "f/1+2" });
      expect(client.get).toHaveBeenCalledWith("/me/drive/items/f%2F1%2B2");
      expect(result).toBe("file content here");
    });

    it("returns metadata when no download URL", async () => {
      (client.get as any).mockResolvedValue({ id: "f1", name: "folder", folder: {} });
      const result = await tools.files_read({ itemId: "f1" });
      const parsed = JSON.parse(result);
      expect(parsed.name).toBe("folder");
    });
  });

  describe("files_create_folder", () => {
    it("creates folder in root by default", async () => {
      (client.post as any).mockResolvedValue({ id: "new-folder", name: "NewFolder", webUrl: "https://..." });
      await tools.files_create_folder({ name: "NewFolder" });
      expect(client.post).toHaveBeenCalledWith("/me/drive/root/children", { name: "NewFolder", folder: {} });
    });

    it("creates in parent path", async () => {
      (client.post as any).mockResolvedValue({ id: "f2", name: "Sub", webUrl: "https://..." });
      await tools.files_create_folder({ name: "Sub", parentPath: "Shared Files/2026 #1" });
      expect(client.post).toHaveBeenCalledWith("/me/drive/root:/Shared%20Files/2026%20%231:/children", { name: "Sub", folder: {} });
    });
  });

  describe("tool schemas", () => {
    it("has 5 files tool schemas", () => {
      expect(filesToolSchemas).toHaveLength(5);
    });
  });
});
