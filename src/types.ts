/**
 * Shared types for m365-mcp
 */

export interface AccountConfig {
  name: string;
  tenantId: string;
  clientId: string;
  email?: string;
  description?: string;
  addedAt: string;
  tokenPath: string;
}

export interface TokenData {
  access_token: string;
  refresh_token: string;
  expires_at: number;
  scope: string;
}

export interface AccountsStore {
  default: string | null;
  accounts: Record<string, {
    tenantId: string;
    clientId: string;
    email?: string;
    description?: string;
    addedAt: string;
  }>;
}

export interface GraphListResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

export interface GraphError {
  error: {
    code: string;
    message: string;
    innerError?: {
      "request-id": string;
      date: string;
    };
  };
}

export interface EmailAddress {
  name?: string;
  address: string;
}

export interface EmailMessage {
  id: string;
  subject: string;
  from?: { emailAddress: EmailAddress };
  toRecipients?: { emailAddress: EmailAddress }[];
  ccRecipients?: { emailAddress: EmailAddress }[];
  receivedDateTime: string;
  sentDateTime: string;
  isRead: boolean;
  isDraft: boolean;
  bodyPreview: string;
  body: { contentType: string; content: string };
  hasAttachments: boolean;
  importance: string;
  webLink: string;
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName: string };
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  organizer?: { emailAddress: EmailAddress };
  attendees?: { emailAddress: EmailAddress; type: string; status: { response: string } }[];
  body?: { contentType: string; content: string };
  isAllDay: boolean;
  recurrence?: any;
  webLink: string;
}

export interface Contact {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  emailAddresses?: { name?: string; address: string }[];
  businessPhones?: string[];
  mobilePhone?: string;
  companyName?: string;
  jobTitle?: string;
  homeAddress?: any;
  businessAddress?: any;
  personalNotes?: string;
}

export interface DriveItem {
  id: string;
  name: string;
  size: number;
  folder?: { childCount: number };
  file?: { mimeType: string };
  webUrl: string;
  createdDateTime?: string;
  lastModifiedDateTime: string;
  parentReference?: { path: string };
  "@microsoft.graph.downloadUrl"?: string;
}

export interface TeamsChat {
  id: string;
  topic?: string;
  chatType: string;
  members?: { displayName: string; userId: string }[];
  lastUpdatedDateTime: string;
}

export interface TeamsMessage {
  id: string;
  from?: { user?: { displayName: string; id: string } };
  body: { contentType: string; content: string };
  createdDateTime: string;
}

export interface TodoList {
  id: string;
  displayName: string;
  isOwner: boolean;
  wellknownListName: string;
}

export interface TodoTask {
  id: string;
  title: string;
  status: string;
  importance: string;
  dueDateTime?: { dateTime: string; timeZone: string };
  body?: { contentType: string; content: string };
  createdDateTime: string;
}

export interface OrgUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
}
