/**
 * Mail tools — list, read, send, reply, search, move, delete, folders
 */

import type { EmailMessage } from "../types.js";
import type { GraphClient } from "../graph.js";

export function registerMailTools(client: GraphClient) {
  const user = "/me";

  return {
    // ── list ─────────────────────────────────────────────────────
    async mail_list(args: {
      folder?: string;
      top?: number;
      filter?: string;
      search?: string;
    }): Promise<string> {
      const folder = args.folder || "inbox";
      const top = args.top || 20;
      let path = `${user}/mailFolders/${encodeURIComponent(folder)}/messages?$top=${top}&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments,importance,webLink`;

      if (args.filter) path += `&$filter=${encodeURIComponent(args.filter)}`;
      if (args.search) path += `&$search="${encodeURIComponent(args.search)}"`;

      const msgs = await client.getAll<EmailMessage>(path, 10, top);
      return JSON.stringify(
        msgs.map((m) => ({
          id: m.id,
          subject: m.subject,
          from: m.from?.emailAddress?.address,
          received: m.receivedDateTime,
          isRead: m.isRead,
          hasAttachments: m.hasAttachments,
          importance: m.importance,
          preview: m.bodyPreview?.slice(0, 200),
        })),
        null,
        2
      );
    },

    // ── read ─────────────────────────────────────────────────────
    async mail_read(args: { messageId: string }): Promise<string> {
      const messageId = encodeURIComponent(args.messageId);
      const m = await client.get<EmailMessage>(
        `${user}/messages/${messageId}`
      );
      return JSON.stringify(
        {
          id: m.id,
          subject: m.subject,
          from: m.from?.emailAddress,
          to: m.toRecipients?.map((r) => r.emailAddress),
          cc: m.ccRecipients?.map((r) => r.emailAddress),
          received: m.receivedDateTime,
          isRead: m.isRead,
          importance: m.importance,
          hasAttachments: m.hasAttachments,
          body: m.body?.content?.slice(0, 10_000),
          webLink: m.webLink,
        },
        null,
        2
      );
    },

    // ── send ─────────────────────────────────────────────────────
    async mail_send(args: {
      to: string;
      subject: string;
      body?: string;
      cc?: string;
      bcc?: string;
      html?: boolean;
    }): Promise<string> {
      const toList = args.to.split(",").map((s) => ({
        emailAddress: { address: s.trim() },
      }));
      const msg: any = {
        subject: args.subject,
        body: {
          contentType: args.html ? "HTML" : "Text",
          content: args.body || "",
        },
        toRecipients: toList,
      };
      if (args.cc) {
        msg.ccRecipients = args.cc.split(",").map((s) => ({
          emailAddress: { address: s.trim() },
        }));
      }
      if (args.bcc) {
        msg.bccRecipients = args.bcc.split(",").map((s) => ({
          emailAddress: { address: s.trim() },
        }));
      }
      await client.post(`${user}/sendMail`, { message: msg });
      return JSON.stringify({ success: true, to: args.to, subject: args.subject });
    },

    // ── reply ────────────────────────────────────────────────────
    async mail_reply(args: {
      messageId: string;
      body: string;
      replyAll?: boolean;
    }): Promise<string> {
      const messageId = encodeURIComponent(args.messageId);
      await client.post(
        args.replyAll
          ? `${user}/messages/${messageId}/replyAll`
          : `${user}/messages/${messageId}/reply`,
        { comment: args.body }
      );
      return JSON.stringify({ success: true });
    },

    // ── search ───────────────────────────────────────────────────
    async mail_search(args: { query: string; top?: number }): Promise<string> {
      const top = args.top || 20;
      const path = `${user}/messages?$search="${encodeURIComponent(args.query)}"&$top=${top}&$orderby=receivedDateTime desc`;
      const msgs = await client.getAll<EmailMessage>(path, 10, top);
      return JSON.stringify(
        msgs.map((m) => ({
          id: m.id,
          subject: m.subject,
          from: m.from?.emailAddress?.address,
          received: m.receivedDateTime,
          preview: m.bodyPreview,
        })),
        null,
        2
      );
    },

    // ── move ─────────────────────────────────────────────────────
    async mail_move(args: { messageId: string; folderId: string }): Promise<string> {
      const messageId = encodeURIComponent(args.messageId);
      await client.post(
        `${user}/messages/${messageId}/move`,
        { destinationId: args.folderId }
      );
      return JSON.stringify({ success: true });
    },

    // ── delete ───────────────────────────────────────────────────
    async mail_delete(args: { messageId: string }): Promise<string> {
      const messageId = encodeURIComponent(args.messageId);
      await client.delete(`${user}/messages/${messageId}`);
      return JSON.stringify({ success: true });
    },

    // ── mark read / unread ───────────────────────────────────────
    async mail_mark_read(args: { messageId: string; isRead: boolean }): Promise<string> {
      const messageId = encodeURIComponent(args.messageId);
      await client.patch(`${user}/messages/${messageId}`, {
        isRead: args.isRead,
      });
      return JSON.stringify({ success: true, isRead: args.isRead });
    },

    // ── list folders ─────────────────────────────────────────────
    async mail_folders(): Promise<string> {
      const folders = await client.getAll<any>(
        `${user}/mailFolders?$select=id,displayName,childFolderCount,totalItemCount,unreadItemCount`
      );
      return JSON.stringify(
        folders.map((f) => ({
          id: f.id,
          name: f.displayName,
          total: f.totalItemCount,
          unread: f.unreadItemCount,
          childFolders: f.childFolderCount,
        })),
        null,
        2
      );
    },
  };
}

export const mailToolSchemas = [
  {
    name: "m365_mail_list",
    description: "List emails from your mailbox with optional folder, filter, and search",
    inputSchema: {
      type: "object",
      properties: {
        folder: { type: "string", description: "Folder name: inbox, sentitems, drafts, etc." },
        top: { type: "number", description: "Number of emails (max 100 per page, paginated)" },
        filter: { type: "string", description: "OData filter, e.g. isRead eq false" },
        search: { type: "string", description: "Search query within folder" },
      },
    },
  },
  {
    name: "m365_mail_read",
    description: "Read a specific email by ID, including full body",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "Email message ID" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "m365_mail_send",
    description: "Send an email from your account",
    inputSchema: {
      type: "object",
      properties: {
        to: { type: "string", description: "Recipient email(s), comma-separated" },
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body (plain text or HTML)" },
        cc: { type: "string", description: "CC recipients, comma-separated" },
        bcc: { type: "string", description: "BCC recipients, comma-separated" },
        html: { type: "boolean", description: "Set true if body is HTML", default: false },
      },
      required: ["to", "subject"],
    },
  },
  {
    name: "m365_mail_reply",
    description: "Reply to an email",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "ID of email to reply to" },
        body: { type: "string", description: "Reply body text" },
        replyAll: { type: "boolean", description: "Reply to all recipients" },
      },
      required: ["messageId", "body"],
    },
  },
  {
    name: "m365_mail_search",
    description: "Search emails across all folders",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query (supports KQL)" },
        top: { type: "number", description: "Max results", default: 20 },
      },
      required: ["query"],
    },
  },
  {
    name: "m365_mail_move",
    description: "Move an email to a different folder",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "Email message ID" },
        folderId: { type: "string", description: "Destination folder ID (use m365_mail_folders to list)" },
      },
      required: ["messageId", "folderId"],
    },
  },
  {
    name: "m365_mail_delete",
    description: "Delete an email (moves to Deleted Items)",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "Email message ID" },
      },
      required: ["messageId"],
    },
  },
  {
    name: "m365_mail_mark_read",
    description: "Mark an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string", description: "Email message ID" },
        isRead: { type: "boolean", description: "True = mark read, False = mark unread" },
      },
      required: ["messageId", "isRead"],
    },
  },
  {
    name: "m365_mail_folders",
    description: "List all mail folders with counts",
    inputSchema: { type: "object", properties: {} },
  },
];
