/**
 * Contacts tools — list, search, read, create, update, delete
 */

import type { Contact } from "../types.js";
import type { GraphClient } from "../graph.js";

export function registerContactsTools(client: GraphClient) {
  const user = "/me";

  return {
    async contacts_list(args: { top?: number; folderId?: string }): Promise<string> {
      const top = args.top || 50;
      const folderPath = args.folderId
        ? `/contactFolders/${encodeURIComponent(args.folderId)}`
        : "";
      const path = `${user}${folderPath}/contacts?$top=${top}&$orderby=displayName`;
      const contacts = await client.getAll<Contact>(path, 10, top);
      return JSON.stringify(
        contacts.map((c) => ({
          id: c.id,
          name: c.displayName,
          email: c.emailAddresses?.[0]?.address,
          company: c.companyName,
          jobTitle: c.jobTitle,
          phones: c.businessPhones,
          mobile: c.mobilePhone,
        })),
        null,
        2
      );
    },

    async contacts_search(args: { query: string; top?: number }): Promise<string> {
      const top = args.top || 20;
      const path = `${user}/contacts?$search="${encodeURIComponent(args.query)}"&$top=${top}`;
      const contacts = await client.getAll<Contact>(path, 10, top);
      return JSON.stringify(
        contacts.map((c) => ({
          id: c.id,
          name: c.displayName,
          email: c.emailAddresses?.map((e) => e.address).join(", "),
          company: c.companyName,
          jobTitle: c.jobTitle,
        })),
        null,
        2
      );
    },

    async contacts_read(args: { contactId: string }): Promise<string> {
      const contactId = encodeURIComponent(args.contactId);
      const c = await client.get<Contact>(`${user}/contacts/${contactId}`);
      return JSON.stringify(c, null, 2);
    },

    async contacts_create(args: {
      givenName?: string;
      surname?: string;
      email?: string;
      company?: string;
      jobTitle?: string;
      mobilePhone?: string;
      businessPhone?: string;
      notes?: string;
    }): Promise<string> {
      const contact: any = {};
      if (args.givenName) contact.givenName = args.givenName;
      if (args.surname) contact.surname = args.surname;
      if (args.email) {
        contact.emailAddresses = [{ address: args.email, name: `${args.givenName || ""} ${args.surname || ""}`.trim() }];
      }
      if (args.company) contact.companyName = args.company;
      if (args.jobTitle) contact.jobTitle = args.jobTitle;
      if (args.mobilePhone) contact.mobilePhone = args.mobilePhone;
      if (args.businessPhone) contact.businessPhones = [args.businessPhone];
      if (args.notes) contact.personalNotes = args.notes;

      const data = await client.post<any>(`${user}/contacts`, contact);
      return JSON.stringify({ id: data.id, displayName: data.displayName, success: true }, null, 2);
    },

    async contacts_update(args: {
      contactId: string;
      givenName?: string;
      surname?: string;
      email?: string;
      company?: string;
      jobTitle?: string;
      mobilePhone?: string;
    }): Promise<string> {
      const patch: any = {};
      if (args.givenName) patch.givenName = args.givenName;
      if (args.surname) patch.surname = args.surname;
      if (args.email) patch.emailAddresses = [{ address: args.email }];
      if (args.company) patch.companyName = args.company;
      if (args.jobTitle) patch.jobTitle = args.jobTitle;
      if (args.mobilePhone) patch.mobilePhone = args.mobilePhone;

      const contactId = encodeURIComponent(args.contactId);
      await client.patch(`${user}/contacts/${contactId}`, patch);
      return JSON.stringify({ success: true, contactId: args.contactId });
    },

    async contacts_delete(args: { contactId: string }): Promise<string> {
      const contactId = encodeURIComponent(args.contactId);
      await client.delete(`${user}/contacts/${contactId}`);
      return JSON.stringify({ success: true });
    },
  };
}

export const contactsToolSchemas = [
  {
    name: "m365_contacts_list",
    description: "List your contacts",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", description: "Max results", default: 50 },
        folderId: { type: "string", description: "Contact folder ID (optional)" },
      },
    },
  },
  {
    name: "m365_contacts_search",
    description: "Search your contacts by name, email, or company",
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
    name: "m365_contacts_read",
    description: "Read a contact's full details",
    inputSchema: {
      type: "object",
      properties: {
        contactId: { type: "string", description: "Contact ID" },
      },
      required: ["contactId"],
    },
  },
  {
    name: "m365_contacts_create",
    description: "Create a new contact",
    inputSchema: {
      type: "object",
      properties: {
        givenName: { type: "string", description: "First name" },
        surname: { type: "string", description: "Last name" },
        email: { type: "string", description: "Email address" },
        company: { type: "string", description: "Company name" },
        jobTitle: { type: "string", description: "Job title" },
        mobilePhone: { type: "string", description: "Mobile phone" },
        businessPhone: { type: "string", description: "Business phone" },
        notes: { type: "string", description: "Personal notes" },
      },
    },
  },
  {
    name: "m365_contacts_update",
    description: "Update an existing contact",
    inputSchema: {
      type: "object",
      properties: {
        contactId: { type: "string", description: "Contact ID" },
        givenName: { type: "string" },
        surname: { type: "string" },
        email: { type: "string" },
        company: { type: "string" },
        jobTitle: { type: "string" },
        mobilePhone: { type: "string" },
      },
      required: ["contactId"],
    },
  },
  {
    name: "m365_contacts_delete",
    description: "Delete a contact",
    inputSchema: {
      type: "object",
      properties: {
        contactId: { type: "string", description: "Contact ID" },
      },
      required: ["contactId"],
    },
  },
];
