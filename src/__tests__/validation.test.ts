import { describe, expect, it } from "vitest";
import { createArgumentValidator } from "../validation.js";
import { mailToolSchemas } from "../tools/mail.js";
import { calendarToolSchemas } from "../tools/calendar.js";
import { contactsToolSchemas } from "../tools/contacts.js";
import { filesToolSchemas } from "../tools/files.js";
import { teamsToolSchemas } from "../tools/teams.js";
import { tasksToolSchemas } from "../tools/tasks.js";
import { usersToolSchemas } from "../tools/users.js";

const validate = createArgumentValidator([...mailToolSchemas, ...calendarToolSchemas, ...contactsToolSchemas, ...filesToolSchemas, ...teamsToolSchemas, ...tasksToolSchemas, ...usersToolSchemas]);
describe("advertised tool argument validation", () => {
  it.each([
    ["m365_mail_send", {}], ["m365_mail_send", { to: [], subject: "x" }],
    ["m365_mail_list", { top: -1 }], ["m365_teams_messages", { chatId: "x", top: 1.5 }],
    ["m365_files_read", { itemId: "x", maxBytes: 1e12 }],
    ["m365_contacts_create", { emailAddresses: [{ name: "Missing address" }] }],
    ["m365_calendar_decline", { eventId: "x", sendResponse: "false" }],
  ])("rejects invalid %s input", (tool, args) => expect(() => validate(tool as string, args)).toThrow("Invalid arguments"));
  it("accepts valid mail recipients and explicit false/zero values", () => {
    expect(() => validate("m365_mail_send", { to: "a@example.com", subject: "", html: false })).not.toThrow();
    expect(() => validate("m365_files_read", { itemId: "x", maxBytes: 0 })).not.toThrow();
  });
});
