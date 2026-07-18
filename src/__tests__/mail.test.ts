/**
 * Tests for mail tools
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerMailTools, mailToolSchemas } from "../tools/mail.js";

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

describe("mail tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerMailTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerMailTools(client);
  });

  describe("mail_folders", () => {
    it("returns formatted folder list", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "inbox", displayName: "Inbox", totalItemCount: 100, unreadItemCount: 5, childFolderCount: 2 },
        { id: "sent", displayName: "Sent Items", totalItemCount: 50, unreadItemCount: 0, childFolderCount: 0 },
      ]);
      const result = await tools.mail_folders();
      const parsed = JSON.parse(result);
      expect(parsed).toHaveLength(2);
      expect(parsed[0]).toEqual({ id: "inbox", name: "Inbox", total: 100, unread: 5, childFolders: 2 });
    });
  });

  describe("mail_list", () => {
    it("builds correct path for inbox", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.mail_list({ folder: "inbox", top: 10 });
      expect(client.getAll).toHaveBeenCalledWith(
        expect.stringContaining("/mailFolders/inbox/messages"),
        10,
        10
      );
    });

    it("includes filter and search params", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.mail_list({ filter: "isRead eq false", search: "report" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("$filter=isRead%20eq%20false");
      expect(path).toContain('$search="report"');
    });

    it("formats messages with key fields", async () => {
      (client.getAll as any).mockResolvedValue([
        {
          id: "msg1",
          subject: "Hello",
          from: { emailAddress: { address: "a@b.com" } },
          receivedDateTime: "2026-07-18T10:00:00Z",
          isRead: false,
          hasAttachments: true,
          importance: "high",
          bodyPreview: "Hi there...",
        },
      ]);
      const result = await tools.mail_list({});
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "msg1", subject: "Hello", from: "a@b.com" });
    });
  });

  describe("mail_read", () => {
    it("returns full message details", async () => {
      (client.get as any).mockResolvedValue({
        id: "msg1",
        subject: "Test",
        from: { emailAddress: { name: "Alice", address: "alice@b.com" } },
        toRecipients: [{ emailAddress: { address: "me@b.com" } }],
        ccRecipients: [],
        receivedDateTime: "2026-07-18T10:00:00Z",
        isRead: true,
        importance: "normal",
        hasAttachments: false,
        body: { content: "Full body" },
        webLink: "https://outlook.office.com/...",
      });
      const result = await tools.mail_read({ messageId: "msg/1+2" });
      expect(client.get).toHaveBeenCalledWith("/me/messages/msg%2F1%2B2");
      const parsed = JSON.parse(result);
      expect(parsed.subject).toBe("Test");
      expect(parsed.body).toBe("Full body");
    });
  });

  describe("mail_send", () => {
    it("sends email with to, cc, bcc", async () => {
      await tools.mail_send({
        to: "a@b.com, c@d.com",
        subject: "Hello",
        body: "World",
        cc: "cc@b.com",
        bcc: "bcc@b.com",
      });
      expect(client.post).toHaveBeenCalledWith("/me/sendMail", {
        message: expect.objectContaining({
          subject: "Hello",
          toRecipients: [
            { emailAddress: { address: "a@b.com" } },
            { emailAddress: { address: "c@d.com" } },
          ],
          ccRecipients: [{ emailAddress: { address: "cc@b.com" } }],
          bccRecipients: [{ emailAddress: { address: "bcc@b.com" } }],
        }),
      });
    });

    it("sends HTML email when html flag is set", async () => {
      await tools.mail_send({ to: "a@b.com", subject: "H", body: "<b>hi</b>", html: true });
      expect(client.post).toHaveBeenCalledWith("/me/sendMail", {
        message: expect.objectContaining({
          body: { contentType: "HTML", content: "<b>hi</b>" },
        }),
      });
    });
  });

  describe("mail_move", () => {
    it("moves message to target folder", async () => {
      await tools.mail_move({ messageId: "msg1", folderId: "folder2" });
      expect(client.post).toHaveBeenCalledWith("/me/messages/msg1/move", { destinationId: "folder2" });
    });
  });

  describe("mail_delete", () => {
    it("deletes message", async () => {
      await tools.mail_delete({ messageId: "msg1" });
      expect(client.delete).toHaveBeenCalledWith("/me/messages/msg1");
    });
  });

  describe("mail_mark_read", () => {
    it("marks as read", async () => {
      await tools.mail_mark_read({ messageId: "msg1", isRead: true });
      expect(client.patch).toHaveBeenCalledWith("/me/messages/msg1", { isRead: true });
    });

    it("marks as unread", async () => {
      await tools.mail_mark_read({ messageId: "msg1", isRead: false });
      expect(client.patch).toHaveBeenCalledWith("/me/messages/msg1", { isRead: false });
    });
  });

  describe("tool schemas", () => {
    it("has 9 mail tool schemas", () => {
      expect(mailToolSchemas).toHaveLength(9);
    });

    it("mail_read requires messageId", () => {
      const schema = mailToolSchemas.find((s) => s.name === "m365_mail_read")!;
      expect(schema.inputSchema.required).toContain("messageId");
    });
  });
});
