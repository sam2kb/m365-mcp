/** Microsoft OneNote tools — notebooks, sections, pages, search, read, create. */

import type { GraphClient } from "../graph.js";
import type { OneNoteNotebook, OneNotePage, OneNoteSection } from "../types.js";

const user = "/me/onenote";
const pageSelect = "id,title,createdDateTime,lastModifiedDateTime,level,order,contentUrl,parentSection,links";

function webUrl(item: { links?: { oneNoteWebUrl?: { href: string } } }): string | undefined {
  return item.links?.oneNoteWebUrl?.href;
}

function pageSummary(page: OneNotePage) {
  return {
    id: page.id,
    title: page.title,
    created: page.createdDateTime,
    modified: page.lastModifiedDateTime,
    level: page.level,
    order: page.order,
    section: page.parentSection,
    webUrl: webUrl(page),
  };
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[char]!);
}

export function registerOneNoteTools(client: GraphClient) {
  return {
    async onenote_notebooks_list(args: { top?: number }): Promise<string> {
      const top = args.top ?? 50;
      const notebooks = await client.getAll<OneNoteNotebook>(
        `${user}/notebooks?$top=${top}&$select=id,displayName,createdDateTime,lastModifiedDateTime,isDefault,isShared,links`, 10, top,
      );
      return JSON.stringify(notebooks.map((notebook) => ({
        id: notebook.id, name: notebook.displayName, isDefault: notebook.isDefault,
        isShared: notebook.isShared, created: notebook.createdDateTime,
        modified: notebook.lastModifiedDateTime, webUrl: webUrl(notebook),
      })), null, 2);
    },

    async onenote_sections_list(args: { notebookId?: string; top?: number }): Promise<string> {
      const top = args.top ?? 50;
      const parent = args.notebookId
        ? `${user}/notebooks/${encodeURIComponent(args.notebookId)}`
        : user;
      const sections = await client.getAll<OneNoteSection>(
        `${parent}/sections?$top=${top}&$select=id,displayName,createdDateTime,lastModifiedDateTime,isDefault,pagesUrl,parentNotebook,links`, 10, top,
      );
      return JSON.stringify(sections.map((section) => ({
        id: section.id, name: section.displayName, notebook: section.parentNotebook,
        isDefault: section.isDefault, created: section.createdDateTime,
        modified: section.lastModifiedDateTime, webUrl: webUrl(section),
      })), null, 2);
    },

    async onenote_pages_list(args: { sectionId: string; top?: number }): Promise<string> {
      const top = args.top ?? 50;
      const pages = await client.getAll<OneNotePage>(
        `${user}/sections/${encodeURIComponent(args.sectionId)}/pages?$top=${top}&$orderby=lastModifiedDateTime desc&$select=${pageSelect}`, 10, top,
      );
      return JSON.stringify(pages.map(pageSummary), null, 2);
    },

    async onenote_page_read(args: { pageId: string; maxBytes?: number }): Promise<string> {
      const pageId = encodeURIComponent(args.pageId);
      const maxBytes = Math.min(Math.max(args.maxBytes ?? 100_000, 1), 500_000);
      const [page, body] = await Promise.all([
        client.get<OneNotePage>(`${user}/pages/${pageId}?$select=${pageSelect}`),
        client.getText(`${user}/pages/${pageId}/content`, maxBytes, "text/html"),
      ]);
      return JSON.stringify({ ...pageSummary(page), contentType: "text/html", html: body.content, truncated: body.truncated }, null, 2);
    },

    async onenote_pages_search(args: { query: string; top?: number }): Promise<string> {
      const top = args.top ?? 20;
      const pages = await client.getAll<OneNotePage>(
        `${user}/pages?search=${encodeURIComponent(args.query)}&$top=${top}&$select=${pageSelect}`, 10, top,
      );
      return JSON.stringify(pages.map(pageSummary), null, 2);
    },

    async onenote_page_create(args: { sectionId: string; title: string; contentHtml?: string }): Promise<string> {
      const html = `<!DOCTYPE html><html><head><title>${escapeHtml(args.title)}</title></head><body>${args.contentHtml ?? ""}</body></html>`;
      const page = await client.postContent<OneNotePage>(
        `${user}/sections/${encodeURIComponent(args.sectionId)}/pages`, html, "application/xhtml+xml",
      );
      return JSON.stringify({ ...pageSummary(page), success: true }, null, 2);
    },
  };
}

export const onenoteToolSchemas = [
  { name: "m365_onenote_notebooks_list", description: "List OneNote notebooks", inputSchema: { type: "object", properties: { top: { type: "number", description: "Maximum notebooks (default 50)" } } } },
  { name: "m365_onenote_sections_list", description: "List all OneNote sections or sections in one notebook", inputSchema: { type: "object", properties: { notebookId: { type: "string", description: "Optional notebook ID" }, top: { type: "number", description: "Maximum sections (default 50)" } } } },
  { name: "m365_onenote_pages_list", description: "List pages in a OneNote section", inputSchema: { type: "object", properties: { sectionId: { type: "string", description: "Section ID" }, top: { type: "number", description: "Maximum pages (default 50)" } }, required: ["sectionId"] } },
  { name: "m365_onenote_page_read", description: "Read OneNote page metadata and HTML content", inputSchema: { type: "object", properties: { pageId: { type: "string", description: "Page ID" }, maxBytes: { type: "number", description: "Maximum UTF-8 content bytes (default 100000, max 500000)" } }, required: ["pageId"] } },
  { name: "m365_onenote_pages_search", description: "Search OneNote pages by title or content", inputSchema: { type: "object", properties: { query: { type: "string", description: "Search terms" }, top: { type: "number", description: "Maximum pages (default 20)" } }, required: ["query"] } },
  { name: "m365_onenote_page_create", description: "Create a OneNote page in a section from HTML", inputSchema: { type: "object", properties: { sectionId: { type: "string", description: "Destination section ID" }, title: { type: "string", description: "Page title" }, contentHtml: { type: "string", description: "Optional HTML for the page body" } }, required: ["sectionId", "title"] } },
];
