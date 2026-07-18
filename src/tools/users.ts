/**
 * Users tools — list org users, get profile, get manager
 */

import type { OrgUser } from "../types.js";
import type { GraphClient } from "../graph.js";

export function registerUsersTools(client: GraphClient) {
  return {
    async users_list(args: { top?: number; filter?: string; search?: string }): Promise<string> {
      let path = `/users?$top=${args.top || 50}&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation`;
      if (args.filter) path += `&$filter=${encodeURIComponent(args.filter)}`;
      if (args.search) path += `&$search="${encodeURIComponent(args.search)}"`;
      const users = await client.getAll<OrgUser>(path);
      return JSON.stringify(users, null, 2);
    },

    async users_profile(args: { user?: string }): Promise<string> {
      const userId = args.user || "me";
      const data = await client.get<OrgUser>(
        `${
          userId === "me" ? "/me" : `/users/${userId}`
        }?$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones`
      );
      return JSON.stringify(data, null, 2);
    },

    async users_manager(args: { user?: string }): Promise<string> {
      const userId = args.user || "me";
      const data = await client.get<any>(
        `${
          userId === "me" ? "/me" : `/users/${userId}`
        }/manager?$select=id,displayName,mail,userPrincipalName,jobTitle`
      );
      return JSON.stringify(data, null, 2);
    },
  };
}

export const usersToolSchemas = [
  {
    name: "m365_users_list",
    description: "List users in your organization (requires User.Read.All — may need admin consent)",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", description: "Max results", default: 50 },
        filter: { type: "string", description: "OData filter" },
        search: { type: "string", description: "Search by displayName or mail" },
      },
    },
  },
  {
    name: "m365_users_profile",
    description: "Get a user's profile (or your own if no user specified)",
    inputSchema: {
      type: "object",
      properties: {
        user: { type: "string", description: "User email or ID (defaults to 'me')" },
      },
    },
  },
  {
    name: "m365_users_manager",
    description: "Get a user's manager (or your own)",
    inputSchema: {
      type: "object",
      properties: {
        user: { type: "string", description: "User email or ID (defaults to 'me')" },
      },
    },
  },
];
