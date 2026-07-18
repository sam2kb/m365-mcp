/**
 * OneDrive files tools — list, search, read, info, create folder
 */

import type { DriveItem } from "../types.js";
import type { GraphClient } from "../graph.js";

function encodeDrivePath(value: string): string {
  return value.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

export function registerFilesTools(client: GraphClient) {
  const user = "/me";

  return {
    async files_list(args: { path?: string; top?: number }): Promise<string> {
      const top = args.top || 50;
      const folderPath = args.path ? `:/${encodeDrivePath(args.path)}:` : "";
      const path = `${user}/drive/root${folderPath}/children?$top=${top}&$select=id,name,size,folder,file,webUrl,lastModifiedDateTime`;
      const items = await client.getAll<DriveItem>(path, 10, top);
      return JSON.stringify(
        items.map((f) => ({
          id: f.id,
          name: f.name,
          size: f.size,
          isFolder: !!f.folder,
          childCount: f.folder?.childCount,
          mimeType: f.file?.mimeType,
          modified: f.lastModifiedDateTime,
          webUrl: f.webUrl,
        })),
        null,
        2
      );
    },

    async files_search(args: { query: string; top?: number }): Promise<string> {
      const top = args.top || 20;
      const query = encodeURIComponent(args.query.replace(/'/g, "''")).replace(/'/g, "%27");
      const path = `${user}/drive/root/search(q='${query}')?$top=${top}&$select=id,name,size,file,webUrl,parentReference`;
      const items = await client.getAll<DriveItem>(path, 10, top);
      return JSON.stringify(
        items.map((f) => ({
          id: f.id,
          name: f.name,
          size: f.size,
          path: f.parentReference?.path,
          webUrl: f.webUrl,
        })),
        null,
        2
      );
    },

    async files_read(args: { itemId: string; maxBytes?: number }): Promise<string> {
      // Get item metadata first to find download URL
      const itemId = encodeURIComponent(args.itemId);
      const meta = await client.get<any>(
        `${user}/drive/items/${itemId}`
      );

      const downloadUrl = meta["@microsoft.graph.downloadUrl"];
      if (!downloadUrl) {
        // For folders or items without direct download, return metadata
        return JSON.stringify(meta, null, 2);
      }

      const content = await client.download(downloadUrl, args.maxBytes ?? 50_000);
      return content;
    },

    async files_info(args: { itemId: string }): Promise<string> {
      const itemId = encodeURIComponent(args.itemId);
      const data = await client.get<DriveItem>(
        `${user}/drive/items/${itemId}?$select=id,name,size,folder,file,webUrl,lastModifiedDateTime,createdDateTime,parentReference`
      );
      return JSON.stringify(data, null, 2);
    },

    async files_create_folder(args: { name: string; parentPath?: string }): Promise<string> {
      const parentPath = args.parentPath ? `:/${encodeDrivePath(args.parentPath)}:` : "";
      const path = `${user}/drive/root${parentPath}/children`;
      const data = await client.post<any>(path, {
        name: args.name,
        folder: {},
      });
      return JSON.stringify({ id: data.id, name: data.name, webUrl: data.webUrl, success: true }, null, 2);
    },
  };
}

export const filesToolSchemas = [
  {
    name: "m365_files_list",
    description: "List files and folders in your OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: "Folder path (e.g. Documents/Work)" },
        top: { type: "number", description: "Max results", default: 50 },
      },
    },
  },
  {
    name: "m365_files_search",
    description: "Search for files in your OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query" },
        top: { type: "number", description: "Max results", default: 20 },
      },
      required: ["query"],
    },
  },
  {
    name: "m365_files_read",
    description: "Read a file's content (text files only, max 50KB by default)",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "File item ID from list/search" },
        maxBytes: { type: "number", description: "Max bytes to read (default 50000)" },
      },
      required: ["itemId"],
    },
  },
  {
    name: "m365_files_info",
    description: "Get file or folder metadata",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "Item ID" },
      },
      required: ["itemId"],
    },
  },
  {
    name: "m365_files_create_folder",
    description: "Create a new folder in OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Folder name" },
        parentPath: { type: "string", description: "Parent folder path (default: root)" },
      },
      required: ["name"],
    },
  },
];
