/**
 * Tests for contacts tools.
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerContactsTools, contactsToolSchemas } from "../tools/contacts.js";

async function* noPages(): AsyncGenerator<unknown[]> {
  return;
}

function mockClient(overrides: Partial<GraphClient> = {}): GraphClient {
  return {
    timezone: "UTC",
    getAll: vi.fn().mockResolvedValue([]),
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    paginate: vi.fn(() => noPages()),
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

  it("lists contacts with categories and folder metadata", async () => {
    (client.getAll as any).mockResolvedValue([{
      id: "c1",
      displayName: "Alice",
      emailAddresses: [{ address: "alice@b.com" }],
      companyName: "Acme",
      categories: ["Customers"],
      parentFolderId: "f1",
    }]);

    const parsed = JSON.parse(await tools.contacts_list({ top: 10 }));
    expect(parsed[0]).toMatchObject({
      id: "c1",
      name: "Alice",
      categories: ["Customers"],
      parentFolderId: "f1",
    });
  });

  it("filters categories across pages without relying on Graph filter support", async () => {
    async function* pages() {
      yield [{ id: "c1", categories: ["Other"] }];
      yield [{ id: "c2", categories: ["Customers"] }];
    }
    (client.paginate as any).mockReturnValue(pages());

    const parsed = JSON.parse(await tools.contacts_list({ category: "customers", top: 5 }));
    expect(parsed.map((contact: any) => contact.id)).toEqual(["c2"]);
    expect(client.getAll).not.toHaveBeenCalled();
  });

  it("searches inside a child folder", async () => {
    await tools.contacts_search({
      query: "Alice",
      folderId: "child",
      parentFolderId: "parent",
    });
    const path = (client.getAll as any).mock.calls[0][0] as string;
    expect(path).toContain("/me/contactFolders/parent/childFolders/child/contacts");
    expect(path).toContain("$search=%22Alice%22");
  });

  it("creates a rich categorized contact in a folder", async () => {
    (client.post as any).mockResolvedValue({ id: "new", displayName: "Bob", categories: ["Vendors"] });
    const result = await tools.contacts_create({
      folderId: "f1",
      givenName: "Bob",
      surname: "Smith",
      emailAddresses: [{ name: "Bob Smith", address: "bob@b.com" }],
      businessPhones: ["555-1234", "555-5678"],
      companyName: "Corp",
      categories: ["Vendors"],
      businessAddress: { city: "Chicago", state: "IL" },
    });

    expect(client.post).toHaveBeenCalledWith(
      "/me/contactFolders/f1/contacts",
      expect.objectContaining({
        emailAddresses: [{ name: "Bob Smith", address: "bob@b.com" }],
        businessPhones: ["555-1234", "555-5678"],
        categories: ["Vendors"],
        businessAddress: { city: "Chicago", state: "IL" },
      })
    );
    expect(JSON.parse(result).success).toBe(true);
  });

  it("keeps convenience fields backward compatible", async () => {
    (client.post as any).mockResolvedValue({ id: "new", displayName: "Bob" });
    await tools.contacts_create({
      givenName: "Bob",
      surname: "Smith",
      email: "bob@b.com",
      company: "Corp",
      businessPhone: "555-1234",
      notes: "Met at conference",
    });

    expect(client.post).toHaveBeenCalledWith("/me/contacts", expect.objectContaining({
      emailAddresses: [{ address: "bob@b.com", name: "Bob Smith" }],
      companyName: "Corp",
      businessPhones: ["555-1234"],
      personalNotes: "Met at conference",
    }));
  });

  it("updates and clears fields instead of dropping empty values", async () => {
    await tools.contacts_update({
      contactId: "c1",
      mobilePhone: "",
      emailAddresses: [],
      categories: [],
    });
    expect(client.patch).toHaveBeenCalledWith("/me/contacts/c1", {
      mobilePhone: "",
      categories: [],
      emailAddresses: [],
    });
  });

  it("rejects an update with no changed fields", async () => {
    await expect(tools.contacts_update({ contactId: "c1" })).rejects.toThrow(
      "at least one contact field"
    );
  });

  it("lists, creates, updates, and deletes contact folders", async () => {
    (client.getAll as any).mockResolvedValue([{ id: "child", displayName: "Vendors" }]);
    await tools.contacts_folders_list({ parentFolderId: "parent" });
    expect(client.getAll).toHaveBeenCalledWith(
      "/me/contactFolders/parent/childFolders?$top=100&$orderby=displayName",
      10,
      100
    );

    (client.post as any).mockResolvedValue({ id: "new-folder", displayName: "Customers" });
    await tools.contacts_folder_create({ parentFolderId: "parent", displayName: "Customers" });
    expect(client.post).toHaveBeenCalledWith(
      "/me/contactFolders/parent/childFolders",
      { displayName: "Customers" }
    );

    (client.patch as any).mockResolvedValue({ id: "child", displayName: "Key Vendors" });
    await tools.contacts_folder_update({
      folderId: "child",
      parentFolderId: "parent",
      displayName: "Key Vendors",
    });
    expect(client.patch).toHaveBeenCalledWith(
      "/me/contactFolders/parent/childFolders/child",
      { displayName: "Key Vendors" }
    );

    await tools.contacts_folder_delete({ folderId: "child", parentFolderId: "parent" });
    expect(client.delete).toHaveBeenCalledWith(
      "/me/contactFolders/parent/childFolders/child"
    );
  });

  it("exposes 10 contact tool schemas", () => {
    expect(contactsToolSchemas).toHaveLength(10);
    expect(new Set(contactsToolSchemas.map((schema) => schema.name)).size).toBe(10);
  });
});
