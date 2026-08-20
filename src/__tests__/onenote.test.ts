import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphClient } from "../graph.js";
import { onenoteToolSchemas, registerOneNoteTools } from "../tools/onenote.js";

function mockClient(): GraphClient {
  return {
    timezone: "UTC",
    getAll: vi.fn().mockResolvedValue([]),
    get: vi.fn(),
    getText: vi.fn(),
    postContent: vi.fn(),
  } as unknown as GraphClient;
}

describe("OneNote tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerOneNoteTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerOneNoteTools(client);
  });

  it("lists notebooks through paginated Graph access", async () => {
    (client.getAll as any).mockResolvedValue([{ id: "n1", displayName: "Work", isDefault: true }]);
    expect(JSON.parse(await tools.onenote_notebooks_list({ top: 5 }))[0]).toMatchObject({ id: "n1", name: "Work" });
    expect(client.getAll).toHaveBeenCalledWith(expect.stringContaining("/me/onenote/notebooks?$top=5"), 10, 5);
  });

  it("lists all sections or scopes them to an encoded notebook", async () => {
    await tools.onenote_sections_list({});
    expect(client.getAll).toHaveBeenLastCalledWith(expect.stringContaining("/me/onenote/sections?"), 10, 50);
    await tools.onenote_sections_list({ notebookId: "notebook/1" });
    expect(client.getAll).toHaveBeenLastCalledWith(expect.stringContaining("/notebooks/notebook%2F1/sections?"), 10, 50);
  });

  it("lists pages in a required section", async () => {
    await tools.onenote_pages_list({ sectionId: "s/1", top: 3 });
    expect(client.getAll).toHaveBeenCalledWith(expect.stringContaining("/sections/s%2F1/pages?$top=3"), 10, 3);
  });

  it("reads metadata and bounded HTML concurrently", async () => {
    (client.get as any).mockResolvedValue({ id: "p1", title: "Plan" });
    (client.getText as any).mockResolvedValue({ content: "<html>plan</html>", truncated: true });
    const result = JSON.parse(await tools.onenote_page_read({ pageId: "p/1", maxBytes: 999_999 }));
    expect(client.getText).toHaveBeenCalledWith("/me/onenote/pages/p%2F1/content", 500_000, "text/html");
    expect(result).toMatchObject({ id: "p1", html: "<html>plan</html>", truncated: true, contentType: "text/html" });
  });

  it("searches pages using Graph's OneNote search query", async () => {
    await tools.onenote_pages_search({ query: "quarterly plan", top: 7 });
    expect(client.getAll).toHaveBeenCalledWith(expect.stringContaining("pages?search=quarterly%20plan&$top=7"), 10, 7);
  });

  it("creates XHTML and escapes the title", async () => {
    (client.postContent as any).mockResolvedValue({ id: "new", title: "R&D" });
    const result = JSON.parse(await tools.onenote_page_create({ sectionId: "s1", title: "R&D <plan>", contentHtml: "<p>Hello</p>" }));
    expect(client.postContent).toHaveBeenCalledWith(
      "/me/onenote/sections/s1/pages",
      expect.stringContaining("<title>R&amp;D &lt;plan&gt;</title>"),
      "application/xhtml+xml",
    );
    expect(result.success).toBe(true);
  });

  it("defines five read tools and one create tool", () => {
    expect(onenoteToolSchemas).toHaveLength(6);
  });
});
