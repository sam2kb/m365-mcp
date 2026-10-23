/**
 * Contacts tools — rich contact CRUD, categories, and contact folders.
 */

import type { Contact } from "../types.js";
import type { GraphClient } from "../graph.js";
import { normalizeTop } from "../validation.js";

interface ContactLocation {
  folderId?: string;
  parentFolderId?: string;
}

interface PhysicalAddressInput {
  street?: string;
  city?: string;
  state?: string;
  countryOrRegion?: string;
  postalCode?: string;
  postOfficeBox?: string;
}

interface ContactInput extends ContactLocation {
  givenName?: string;
  middleName?: string;
  surname?: string;
  title?: string;
  generation?: string;
  displayName?: string;
  nickName?: string;
  email?: string;
  emailAddresses?: Array<{ name?: string; address: string }>;
  mobilePhone?: string;
  businessPhone?: string;
  businessPhones?: string[];
  homePhones?: string[];
  company?: string;
  companyName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  profession?: string;
  businessHomePage?: string;
  notes?: string;
  personalNotes?: string;
  birthday?: string;
  spouseName?: string;
  assistantName?: string;
  manager?: string;
  children?: string[];
  categories?: string[];
  imAddresses?: string[];
  fileAs?: string;
  homeAddress?: PhysicalAddressInput;
  businessAddress?: PhysicalAddressInput;
  otherAddress?: PhysicalAddressInput;
}

interface ContactUpdateInput extends ContactInput {
  contactId: string;
}

interface ContactFolder {
  id: string;
  displayName?: string;
  parentFolderId?: string;
  childFolderCount?: number;
}

interface GraphPerson {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  userPrincipalName?: string;
  companyName?: string;
  department?: string;
  jobTitle?: string;
  officeLocation?: string;
  phones?: Array<{
    type?: string;
    number?: string;
  }>;
  scoredEmailAddresses?: Array<{
    address?: string;
    relevanceScore?: number;
  }>;
  personType?: {
    class?: string;
    subclass?: string;
  };
}

type PeopleSource = "mailbox" | "directory" | "both";
type PeoplePersonType = "any" | "organizationUser" | "personalContact" | "group";

interface PeopleListInput {
  top?: number;
  search?: string;
  source?: PeopleSource;
  personType?: PeoplePersonType;
  includeProfile?: boolean;
}

const PEOPLE_BASIC_FIELDS = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "userPrincipalName",
  "scoredEmailAddresses",
  "personType",
] as const;

const PEOPLE_PROFILE_FIELDS = [
  "companyName",
  "department",
  "jobTitle",
  "officeLocation",
  "phones",
] as const;

const PEOPLE_SOURCE_HEADERS: Record<PeopleSource, string> = {
  mailbox: "Mailbox",
  directory: "Directory",
  both: "Mailbox,Directory",
};

const PEOPLE_TYPE_FILTERS: Record<Exclude<PeoplePersonType, "any">, string> = {
  organizationUser: "personType/class eq 'Person' and personType/subclass eq 'OrganizationUser'",
  personalContact: "personType/class eq 'Person' and personType/subclass eq 'PersonalContact'",
  group: "personType/class eq 'Group'",
};

const CONTACT_STRING_FIELDS = [
  "givenName", "middleName", "surname", "title", "generation", "displayName",
  "nickName", "mobilePhone", "jobTitle", "department", "officeLocation",
  "profession", "businessHomePage", "birthday", "spouseName", "assistantName",
  "manager", "fileAs",
] as const;

const CONTACT_ARRAY_FIELDS = [
  "homePhones", "children", "categories", "imAddresses",
] as const;

function encodeId(id: string): string {
  return encodeURIComponent(id);
}

function normalizePeopleSource(value: PeopleSource | undefined): PeopleSource {
  const source = value ?? "mailbox";
  if (!Object.hasOwn(PEOPLE_SOURCE_HEADERS, source)) {
    throw new Error('source must be "mailbox", "directory", or "both"');
  }
  return source;
}

function normalizePeoplePersonType(value: PeoplePersonType | undefined): PeoplePersonType {
  const personType = value ?? "any";
  if (personType !== "any" && !Object.hasOwn(PEOPLE_TYPE_FILTERS, personType)) {
    throw new Error(
      'personType must be "any", "organizationUser", "personalContact", or "group"'
    );
  }
  return personType;
}

function contactCollectionPath(location: ContactLocation): string {
  if (location.parentFolderId && !location.folderId) {
    throw new Error("folderId is required when parentFolderId is provided");
  }
  if (location.parentFolderId && location.folderId) {
    return `/me/contactFolders/${encodeId(location.parentFolderId)}/childFolders/${encodeId(location.folderId)}/contacts`;
  }
  if (location.folderId) {
    return `/me/contactFolders/${encodeId(location.folderId)}/contacts`;
  }
  return "/me/contacts";
}

function contactResourcePath(location: ContactLocation, contactId: string): string {
  return `${contactCollectionPath(location)}/${encodeId(contactId)}`;
}

function folderResourcePath(folderId: string, parentFolderId?: string): string {
  if (parentFolderId) {
    return `/me/contactFolders/${encodeId(parentFolderId)}/childFolders/${encodeId(folderId)}`;
  }
  return `/me/contactFolders/${encodeId(folderId)}`;
}

function buildContactPayload(args: ContactInput): Record<string, unknown> {
  const payload: Record<string, unknown> = {};

  for (const field of CONTACT_STRING_FIELDS) {
    if (args[field] !== undefined) payload[field] = args[field];
  }
  for (const field of CONTACT_ARRAY_FIELDS) {
    if (args[field] !== undefined) payload[field] = args[field];
  }

  if (args.emailAddresses !== undefined) {
    payload.emailAddresses = args.emailAddresses;
  } else if (args.email !== undefined) {
    const name = [args.givenName, args.surname].filter(Boolean).join(" ");
    payload.emailAddresses = [{ address: args.email, ...(name ? { name } : {}) }];
  }

  if (args.businessPhones !== undefined) {
    payload.businessPhones = args.businessPhones;
  } else if (args.businessPhone !== undefined) {
    payload.businessPhones = args.businessPhone ? [args.businessPhone] : [];
  }

  if (args.companyName !== undefined) {
    payload.companyName = args.companyName;
  } else if (args.company !== undefined) {
    payload.companyName = args.company;
  }

  if (args.personalNotes !== undefined) {
    payload.personalNotes = args.personalNotes;
  } else if (args.notes !== undefined) {
    payload.personalNotes = args.notes;
  }

  for (const field of ["homeAddress", "businessAddress", "otherAddress"] as const) {
    if (args[field] !== undefined) payload[field] = args[field];
  }

  return payload;
}

function summarizeContact(contact: Contact) {
  return {
    id: contact.id,
    name: contact.displayName,
    emails: contact.emailAddresses ?? [],
    company: contact.companyName,
    jobTitle: contact.jobTitle,
    businessPhones: contact.businessPhones ?? [],
    mobilePhone: contact.mobilePhone,
    categories: contact.categories ?? [],
    parentFolderId: contact.parentFolderId,
  };
}

type CategoryMatch = "any" | "all";

interface CategoryFilter {
  category?: string;
  categories?: string[];
  categoryMatch?: CategoryMatch;
}

function normalizeCategoryFilter(filter: CategoryFilter): {
  categories: string[];
  match: CategoryMatch;
} {
  if (filter.category !== undefined && filter.categories !== undefined) {
    throw new Error("Use category or categories, not both");
  }
  const categories = filter.categories ?? (filter.category !== undefined ? [filter.category] : []);
  if (categories.some((category) => category.length === 0)) {
    throw new Error("Category names cannot be empty");
  }
  const match = filter.categoryMatch ?? "any";
  if (match !== "any" && match !== "all") {
    throw new Error('categoryMatch must be "any" or "all"');
  }
  return {
    categories: [...new Set(categories.map((category) => category.toLocaleLowerCase()))],
    match,
  };
}

async function getContacts(
  client: GraphClient,
  path: string,
  top: number,
  filter: CategoryFilter
): Promise<Contact[]> {
  const { categories, match } = normalizeCategoryFilter(filter);
  if (categories.length === 0) return client.getAll<Contact>(path, 10, top);

  const matches: Contact[] = [];
  for await (const page of client.paginate<Contact>(path, 10)) {
    for (const contact of page) {
      const assigned = new Set(
        (contact.categories ?? []).map((category) => category.toLocaleLowerCase())
      );
      const matched = match === "all"
        ? categories.every((category) => assigned.has(category))
        : categories.some((category) => assigned.has(category));
      if (matched) {
        matches.push(contact);
        if (matches.length === top) return matches;
      }
    }
  }
  return matches;
}

export function registerContactsTools(client: GraphClient) {
  return {
    async contacts_list(args: ContactLocation & CategoryFilter & { top?: number }): Promise<string> {
      const top = normalizeTop(args.top, 50);
      const path = `${contactCollectionPath(args)}?$top=${Math.min(top, 100)}&$orderby=displayName`;
      const contacts = await getContacts(client, path, top, args);
      return JSON.stringify(contacts.map(summarizeContact), null, 2);
    },

    async contacts_search(args: ContactLocation & CategoryFilter & { query: string; top?: number }): Promise<string> {
      const top = normalizeTop(args.top, 20);
      const path = `${contactCollectionPath(args)}?$search=%22${encodeURIComponent(args.query)}%22&$top=${Math.min(top, 100)}`;
      const contacts = await getContacts(client, path, top, args);
      return JSON.stringify(contacts.map(summarizeContact), null, 2);
    },

    async contacts_read(args: ContactLocation & { contactId: string }): Promise<string> {
      const contact = await client.get<Contact>(contactResourcePath(args, args.contactId));
      return JSON.stringify(contact, null, 2);
    },

    async contacts_create(args: ContactInput): Promise<string> {
      const contact = buildContactPayload(args);
      if (Object.keys(contact).length === 0) throw new Error("Provide at least one contact field");
      const data = await client.post<Contact>(contactCollectionPath(args), contact);
      return JSON.stringify({ id: data.id, displayName: data.displayName, categories: data.categories ?? [], success: true }, null, 2);
    },

    async contacts_update(args: ContactUpdateInput): Promise<string> {
      const patch = buildContactPayload(args);
      if (Object.keys(patch).length === 0) throw new Error("Provide at least one contact field to update");
      await client.patch(contactResourcePath(args, args.contactId), patch);
      return JSON.stringify({ success: true, contactId: args.contactId });
    },

    async contacts_delete(args: ContactLocation & { contactId: string }): Promise<string> {
      await client.delete(contactResourcePath(args, args.contactId));
      return JSON.stringify({ success: true, contactId: args.contactId });
    },

    async contacts_categories_list(args: ContactLocation = {}): Promise<string> {
      const path = `${contactCollectionPath(args)}?$top=100&$select=id,categories`;
      const contacts = await client.getAll<Contact>(path, 100, 10_000);
      const categories = new Map<string, { name: string; count: number }>();
      for (const contact of contacts) {
        const seen = new Set<string>();
        for (const name of contact.categories ?? []) {
          const key = name.toLocaleLowerCase();
          if (seen.has(key)) continue;
          seen.add(key);
          const existing = categories.get(key);
          if (existing) existing.count++;
          else categories.set(key, { name, count: 1 });
        }
      }
      return JSON.stringify(
        [...categories.values()].sort((a, b) => a.name.localeCompare(b.name)),
        null,
        2
      );
    },

    async contacts_folders_list(args: { parentFolderId?: string; top?: number } = {}): Promise<string> {
      const top = normalizeTop(args.top, 100);
      const base = args.parentFolderId
        ? `/me/contactFolders/${encodeId(args.parentFolderId)}/childFolders`
        : "/me/contactFolders";
      const folders = await client.getAll<ContactFolder>(`${base}?$top=${Math.min(top, 100)}&$orderby=displayName`, 10, top);
      return JSON.stringify(folders, null, 2);
    },

    async contacts_folder_create(args: { parentFolderId: string; displayName: string }): Promise<string> {
      const path = `/me/contactFolders/${encodeId(args.parentFolderId)}/childFolders`;
      const folder = await client.post<ContactFolder>(path, { displayName: args.displayName });
      return JSON.stringify({ ...folder, success: true }, null, 2);
    },

    async contacts_folder_update(args: { folderId: string; parentFolderId?: string; displayName?: string; moveToParentFolderId?: string }): Promise<string> {
      const patch: Record<string, string> = {};
      if (args.displayName !== undefined) patch.displayName = args.displayName;
      if (args.moveToParentFolderId !== undefined) patch.parentFolderId = args.moveToParentFolderId;
      if (Object.keys(patch).length === 0) throw new Error("Provide displayName or moveToParentFolderId");
      const folder = await client.patch<ContactFolder>(folderResourcePath(args.folderId, args.parentFolderId), patch);
      return JSON.stringify({ ...folder, success: true }, null, 2);
    },

    async contacts_folder_delete(args: { folderId: string; parentFolderId?: string }): Promise<string> {
      await client.delete(folderResourcePath(args.folderId, args.parentFolderId));
      return JSON.stringify({ success: true, folderId: args.folderId });
    },

    async contacts_people_list(args: PeopleListInput = {}): Promise<string> {
      const top = normalizeTop(args.top, 100);
      const source = normalizePeopleSource(args.source);
      const personType = normalizePeoplePersonType(args.personType);
      const fields = args.includeProfile
        ? [...PEOPLE_BASIC_FIELDS, ...PEOPLE_PROFILE_FIELDS]
        : [...PEOPLE_BASIC_FIELDS];
      const query = [
        `$top=${Math.min(top, 100)}`,
        `$select=${fields.join(",")}`,
      ];
      if (args.search) query.push(`$search=%22${encodeURIComponent(args.search)}%22`);
      if (personType !== "any") {
        query.push(`$filter=${encodeURIComponent(PEOPLE_TYPE_FILTERS[personType])}`);
      }
      const path = `/me/people?${query.join("&")}`;
      const people = await client.getAll<GraphPerson>(path, 10, top, {
        "X-PeopleQuery-QuerySources": PEOPLE_SOURCE_HEADERS[source],
      });
      return JSON.stringify(
        people.map((person) => {
          const primaryEmail = person.scoredEmailAddresses?.[0];
          return {
            id: person.id,
            displayName: person.displayName,
            givenName: person.givenName,
            surname: person.surname,
            email: primaryEmail?.address ?? person.userPrincipalName,
            relevanceScore: primaryEmail?.relevanceScore,
            personType: person.personType,
            ...(args.includeProfile ? {
              companyName: person.companyName,
              department: person.department,
              jobTitle: person.jobTitle,
              officeLocation: person.officeLocation,
              phones: person.phones ?? [],
            } : {}),
          };
        }),
        null,
        2
      );
    },
  };
}

const locationProperties = {
  folderId: { type: "string", description: "Optional contact folder ID; omit for the default Contacts folder" },
  parentFolderId: { type: "string", description: "Parent folder ID when folderId identifies a direct child folder" },
};

const categoryFilterProperties = {
  category: { type: "string", minLength: 1, description: "One exact category name (case-insensitive)" },
  categories: {
    type: "array",
    minItems: 1,
    uniqueItems: true,
    description: "Exact category names (case-insensitive); use categoryMatch for any/all behavior",
    items: { type: "string", minLength: 1 },
  },
  categoryMatch: {
    type: "string",
    enum: ["any", "all"],
    default: "any",
    description: "Match any supplied category or require all supplied categories",
  },
};

const addressSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    street: { type: "string" }, city: { type: "string" }, state: { type: "string" },
    countryOrRegion: { type: "string" }, postalCode: { type: "string" }, postOfficeBox: { type: "string" },
  },
};

const stringArraySchema = (description: string) => ({ type: "array", description, items: { type: "string" } });

const contactFieldProperties = {
  givenName: { type: "string", description: "First name" },
  middleName: { type: "string" }, surname: { type: "string", description: "Last name" },
  title: { type: "string", description: "Courtesy title, such as Dr. or Ms." },
  generation: { type: "string", description: "Name suffix, such as Jr." },
  displayName: { type: "string" }, nickName: { type: "string" },
  email: { type: "string", description: "Single-email convenience field" },
  emailAddresses: {
    type: "array", description: "Complete email list; an empty array clears all email addresses",
    items: { type: "object", additionalProperties: false, properties: { name: { type: "string" }, address: { type: "string" } }, required: ["address"] },
  },
  mobilePhone: { type: "string" },
  businessPhone: { type: "string", description: "Single-business-phone convenience field" },
  businessPhones: stringArraySchema("Complete business phone list; empty clears it"),
  homePhones: stringArraySchema("Complete home phone list; empty clears it"),
  company: { type: "string", description: "Company-name convenience field" },
  companyName: { type: "string" }, jobTitle: { type: "string" }, department: { type: "string" },
  officeLocation: { type: "string" }, profession: { type: "string" }, businessHomePage: { type: "string" },
  notes: { type: "string", description: "Personal-notes convenience field" }, personalNotes: { type: "string" },
  birthday: { type: "string", description: "ISO 8601 date-time in UTC" }, spouseName: { type: "string" },
  assistantName: { type: "string" }, manager: { type: "string" },
  children: stringArraySchema("Children names; empty clears the list"),
  categories: stringArraySchema("Outlook category names; empty removes all categories"),
  imAddresses: stringArraySchema("Instant-messaging addresses; empty clears the list"),
  fileAs: { type: "string" }, homeAddress: addressSchema, businessAddress: addressSchema, otherAddress: addressSchema,
};

export const contactsToolSchemas = [
  {
    name: "m365_contacts_list",
    description: "List contacts, optionally within a folder and filtered by one or more exact category names",
    inputSchema: { type: "object", properties: { top: { type: "number", minimum: 1, maximum: 500, default: 50 }, ...categoryFilterProperties, ...locationProperties } },
  },
  {
    name: "m365_contacts_search",
    description: "Search contacts by name, email, or company, optionally within a folder and filtered by exact categories",
    inputSchema: { type: "object", properties: { query: { type: "string" }, top: { type: "number", minimum: 1, maximum: 500, default: 20 }, ...categoryFilterProperties, ...locationProperties }, required: ["query"] },
  },
  {
    name: "m365_contacts_read", description: "Read a contact's full details",
    inputSchema: { type: "object", properties: { contactId: { type: "string" }, ...locationProperties }, required: ["contactId"] },
  },
  {
    name: "m365_contacts_create",
    description: "Create a contact with names, multiple emails and phones, categories, addresses, and work or personal details",
    inputSchema: { type: "object", properties: { ...contactFieldProperties, ...locationProperties } },
  },
  {
    name: "m365_contacts_update",
    description: "Update or clear contact fields; use empty strings or arrays to clear supported values",
    inputSchema: { type: "object", properties: { contactId: { type: "string" }, ...contactFieldProperties, ...locationProperties }, required: ["contactId"] },
  },
  {
    name: "m365_contacts_delete", description: "Move a contact to Deleted Items",
    inputSchema: { type: "object", properties: { contactId: { type: "string" }, ...locationProperties }, required: ["contactId"] },
  },
  {
    name: "m365_contacts_categories_list",
    description: "List unique category names assigned to contacts in the default or selected folder, with usage counts",
    inputSchema: { type: "object", properties: { ...locationProperties } },
  },
  {
    name: "m365_contacts_folders_list", description: "List top-level contact folders or the direct children of one folder",
    inputSchema: { type: "object", properties: { parentFolderId: { type: "string" }, top: { type: "number", minimum: 1, maximum: 500, default: 100 } } },
  },
  {
    name: "m365_contacts_folder_create", description: "Create a child contact folder under an existing contact folder",
    inputSchema: { type: "object", properties: { parentFolderId: { type: "string" }, displayName: { type: "string" } }, required: ["parentFolderId", "displayName"] },
  },
  {
    name: "m365_contacts_folder_update", description: "Rename or move a top-level or direct-child contact folder",
    inputSchema: { type: "object", properties: { folderId: { type: "string" }, parentFolderId: { type: "string", description: "Current parent when this is a child folder" }, displayName: { type: "string" }, moveToParentFolderId: { type: "string", description: "New parent folder ID" } }, required: ["folderId"] },
  },
  {
    name: "m365_contacts_folder_delete", description: "Delete a non-default top-level or direct-child contact folder",
    inputSchema: { type: "object", properties: { folderId: { type: "string" }, parentFolderId: { type: "string", description: "Current parent when this is a child folder" } }, required: ["folderId"] },
  },
  {
    name: "m365_contacts_people_list",
    description: "Find relevance-ranked people for recipient selection using mailbox signals, the organization directory, or both; does not expose private Outlook Contact Lists",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", minimum: 1, maximum: 500, default: 100 },
        search: { type: "string", description: "Fuzzy search by name or email alias" },
        source: {
          type: "string",
          enum: ["mailbox", "directory", "both"],
          default: "mailbox",
          description: "Search mailbox-relevant people, the organization directory, or both",
        },
        personType: {
          type: "string",
          enum: ["any", "organizationUser", "personalContact", "group"],
          default: "any",
          description: "Limit results to one documented People API identity type",
        },
        includeProfile: {
          type: "boolean",
          default: false,
          description: "Include company, department, job title, office, and phone details",
        },
      },
    },
  },
];
