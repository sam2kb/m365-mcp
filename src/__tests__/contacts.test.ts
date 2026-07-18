/**
 * Tests for contacts tools
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerContactsTools, contactsToolSchemas } from "../tools/contacts.js";

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

describe("contacts tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerContactsTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerContactsTools(client);
  });

  describe("contacts_list", () => {
    it("lists contacts with summary fields", async () => {
      (client.getAll as any).mockResolvedValue([
        { id: "c1", displayName: "Alice", emailAddresses: [{ address: "alice@b.com" }], companyName: "Acme", jobTitle: "CEO", businessPhones: ["555-0100"], mobilePhone: "555-0200" },
      ]);
      const result = await tools.contacts_list({ top: 10 });
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({ id: "c1", name: "Alice", email: "alice@b.com", company: "Acme" });
    });
  });

  describe("contacts_search", () => {
    it("searches with query", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.contacts_search({ query: "Alice" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain('$search="Alice"');
    });
  });

  describe("contacts_create", () => {
    it("creates contact with all fields", async () => {
      (client.post as any).mockResolvedValue({ id: "new", displayName: "Bob" });
      const result = await tools.contacts_create({
        givenName: "Bob",
        surname: "Smith",
        email: "bob@b.com",
        company: "Corp",
        jobTitle: "Dev",
        mobilePhone: "555-1234",
      });
      expect(client.post).toHaveBeenCalledWith("/me/contacts", expect.objectContaining({
        givenName: "Bob",
        surname: "Smith",
        emailAddresses: [{ address: "bob@b.com", name: "Bob Smith" }],
        companyName: "Corp",
      }));
      const parsed = JSON.parse(result);
      expect(parsed.success).toBe(true);
    });
  });

  describe("contacts_update", () => {
    it("patches provided fields", async () => {
      await tools.contacts_update({ contactId: "c1", jobTitle: "CTO" });
      expect(client.patch).toHaveBeenCalledWith("/me/contacts/c1", { jobTitle: "CTO" });
    });
  });

  describe("tool schemas", () => {
    it("has 6 contact tool schemas", () => {
      expect(contactsToolSchemas).toHaveLength(6);
    });
  });
});
