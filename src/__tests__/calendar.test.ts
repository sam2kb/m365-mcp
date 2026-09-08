/**
 * Tests for calendar tools
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerCalendarTools, calendarToolSchemas } from "../tools/calendar.js";

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

describe("calendar tools", () => {
  let client: GraphClient;
  let tools: ReturnType<typeof registerCalendarTools>;

  beforeEach(() => {
    client = mockClient();
    tools = registerCalendarTools(client);
  });

  describe("calendar_list", () => {
    it("uses calendarView when date range provided", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.calendar_list({ startDateTime: "2026-07-01T00:00:00Z", endDateTime: "2026-07-31T23:59:59Z" });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("calendarView");
      expect(path).toContain("startDateTime=");
      expect(path).toContain("endDateTime=");
    });

    it("uses events endpoint when no date range", async () => {
      (client.getAll as any).mockResolvedValue([]);
      await tools.calendar_list({});
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("/events");
    });

    it("formats events with key fields", async () => {
      (client.getAll as any).mockResolvedValue([
        {
          id: "ev1",
          subject: "Meeting",
          start: { dateTime: "2026-07-18T10:00:00.0000000", timeZone: "UTC" },
          end: { dateTime: "2026-07-18T11:00:00.0000000", timeZone: "UTC" },
          isAllDay: false,
          location: { displayName: "Room 1" },
          isOnlineMeeting: true,
          onlineMeetingUrl: "https://teams.microsoft.com/...",
          organizer: { emailAddress: { name: "Alice", address: "alice@b.com" } },
          attendees: [{}, {}],
          isOrganizer: false,
          responseStatus: { response: "tentativelyAccepted", time: "2026-07-01T12:00:00Z" },
          webLink: "https://outlook.office.com/...",
        },
      ]);
      const result = await tools.calendar_list({});
      const parsed = JSON.parse(result);
      expect(parsed[0]).toMatchObject({
        id: "ev1",
        subject: "Meeting",
        isAllDay: false,
        location: "Room 1",
        isOnline: true,
        organizer: "Alice",
        attendees: 2,
        isOrganizer: false,
        responseStatus: "tentativelyAccepted",
      });
      const path = (client.getAll as any).mock.calls[0][0] as string;
      expect(path).toContain("isOrganizer,responseStatus");
    });
  });

  describe("calendar_create", () => {
    it("creates event with attendees and online meeting", async () => {
      (client.post as any).mockResolvedValue({
        id: "new-ev",
        subject: "Sync",
        webLink: "https://...",
        onlineMeetingUrl: null,
      });
      const result = await tools.calendar_create({
        subject: "Sync",
        start: "2026-07-18T14:00:00",
        end: "2026-07-18T15:00:00",
        body: "Discuss Q3",
        location: "Board Room",
        attendees: "a@b.com, c@d.com",
        isOnline: true,
      });
      expect(client.post).toHaveBeenCalledWith("/me/events", expect.objectContaining({
        subject: "Sync",
        isOnlineMeeting: true,
        attendees: [
          { emailAddress: { address: "a@b.com" }, type: "required" },
          { emailAddress: { address: "c@d.com" }, type: "required" },
        ],
      }));
      const parsed = JSON.parse(result);
      expect(parsed.id).toBe("new-ev");
    });
  });

  describe("calendar_update", () => {
    it("clears notes only when an empty body is explicitly supplied", async () => {
      await tools.calendar_update({ eventId: "event/+?", calendarId: "calendar/+?", body: "", bodyContentType: "text" });
      expect(client.patch).toHaveBeenCalledWith("/me/calendars/calendar%2F%2B%3F/events/event%2F%2B%3F", {
        body: { contentType: "Text", content: "" },
      });
    });

    it("keeps the legacy HTML default when replacing notes", async () => {
      await tools.calendar_update({ eventId: "ev1", body: "<p>Updated notes</p>" });
      expect(client.patch).toHaveBeenCalledWith("/me/events/ev1", {
        body: { contentType: "HTML", content: "<p>Updated notes</p>" },
      });
    });

    it("patches only provided fields", async () => {
      await tools.calendar_update({ eventId: "ev1", subject: "Updated", location: "Room 2" });
      expect(client.patch).toHaveBeenCalledWith("/me/events/ev1", {
        subject: "Updated",
        location: { displayName: "Room 2" },
      });
    });
  });

  describe("calendar_get", () => {
    it("returns complete short notes with their original format", async () => {
      vi.mocked(client.get).mockResolvedValue({ id: "ev1", body: { contentType: "html", content: "<p>Notes</p>" } });
      const result = JSON.parse(await tools.calendar_get({ eventId: "ev1" }));
      expect(client.get).toHaveBeenCalledWith("/me/events/ev1?$select=id,subject,start,end,body,webLink");
      expect(result).toMatchObject({ body: { contentType: "html", content: "<p>Notes</p>" }, bodyTruncated: false, bodyLength: 12 });
    });

    it.each([undefined, 0, 5, 50_000])("bounds note output with maxBodyChars=%s", async (maxBodyChars) => {
      vi.mocked(client.get).mockResolvedValue({ body: { contentType: "text", content: "x".repeat(50_001) } });
      const result = JSON.parse(await tools.calendar_get({ eventId: "event/+?", calendarId: "calendar/+?", maxBodyChars }));
      expect(client.get).toHaveBeenCalledWith("/me/calendars/calendar%2F%2B%3F/events/event%2F%2B%3F?$select=id,subject,start,end,body,webLink");
      expect(result.body.content).toHaveLength(maxBodyChars ?? 10_000);
      expect(result.bodyTruncated).toBe(true);
      expect(result.bodyLength).toBe(50_001);
    });

    it.each([-1, 50_001, 1.5, NaN, Infinity])("rejects an invalid output limit %s before Graph access", async (maxBodyChars) => {
      await expect(tools.calendar_get({ eventId: "ev1", maxBodyChars })).rejects.toThrow("maxBodyChars");
      expect(client.get).not.toHaveBeenCalled();
    });

    it("represents an absent body without claiming truncation", async () => {
      vi.mocked(client.get).mockResolvedValue({ id: "ev1" });
      expect(JSON.parse(await tools.calendar_get({ eventId: "ev1" }))).toMatchObject({
        body: { contentType: "text", content: "" }, bodyTruncated: false, bodyLength: 0,
      });
    });

    it("propagates a missing event instead of falling back to another calendar", async () => {
      vi.mocked(client.get).mockRejectedValue(new Error("Graph API 404: Not found"));
      await expect(tools.calendar_get({ eventId: "ev1", calendarId: "selected" })).rejects.toThrow("404");
      expect(client.get).toHaveBeenCalledTimes(1);
    });
  });

  describe("specific calendar notes", () => {
    it.each([{}, { startDateTime: "2026-07-01T00:00:00Z", endDateTime: "2026-07-31T23:59:59Z" }])(
      "scopes list requests without adding full bodies to the list",
      async (range) => {
        await tools.calendar_list({ calendarId: "calendar/+?", ...range });
        const path = vi.mocked(client.getAll).mock.calls[0][0];
        expect(path).toMatch(/^\/me\/calendars\/calendar%2F%2B%3F\/(events|calendarView)\?/);
        expect(path).not.toContain("body");
      },
    );

    it("creates literal text notes in the selected calendar", async () => {
      vi.mocked(client.post).mockResolvedValue({ id: "new" });
      await tools.calendar_create({ calendarId: "selected", subject: "Test", start: "2026-07-18T12:00:00", end: "2026-07-18T12:30:00", body: "a < b\nNext line", bodyContentType: "text" });
      expect(client.post).toHaveBeenCalledWith("/me/calendars/selected/events", expect.objectContaining({
        body: { contentType: "Text", content: "a < b\nNext line" },
      }));
    });

    it("rejects an empty explicit calendar rather than silently selecting the default", async () => {
      await expect(tools.calendar_list({ calendarId: " " })).rejects.toThrow("calendarId");
      expect(client.getAll).not.toHaveBeenCalled();
    });
  });

  describe("calendar_delete", () => {
    it("keeps the legacy organizer-cancel behavior", async () => {
      await tools.calendar_delete({ eventId: "ev1", comment: "No longer needed" });
      expect(client.post).toHaveBeenCalledWith("/me/events/ev1/cancel", { comment: "No longer needed" });
    });
  });

  describe("calendar_cancel", () => {
    it("cancels an organizer-owned event with a comment", async () => {
      await tools.calendar_cancel({ eventId: "organizer/event", comment: "Cancelled" });
      expect(client.post).toHaveBeenCalledWith("/me/events/organizer%2Fevent/cancel", { comment: "Cancelled" });
    });
  });

  describe("calendar_decline", () => {
    it("declines an attendee invitation and sends a response by default", async () => {
      await tools.calendar_decline({ eventId: "invite1", comment: "Unable to attend" });
      expect(client.post).toHaveBeenCalledWith("/me/events/invite1/decline", {
        comment: "Unable to attend",
        sendResponse: true,
      });
    });

    it("can decline without sending a response", async () => {
      await tools.calendar_decline({ eventId: "invite1", sendResponse: false });
      expect(client.post).toHaveBeenCalledWith("/me/events/invite1/decline", {
        comment: "",
        sendResponse: false,
      });
    });
  });

  describe("calendar_availability", () => {
    it("queries free/busy for multiple users", async () => {
      (client.post as any).mockResolvedValue({ value: [] });
      await tools.calendar_availability({
        users: "a@b.com, c@d.com",
        startDateTime: "2026-07-18T00:00:00",
        endDateTime: "2026-07-18T23:59:59",
      });
      expect(client.post).toHaveBeenCalledWith("/me/calendar/getSchedule", expect.objectContaining({
        schedules: ["a@b.com", "c@d.com"],
      }));
    });
  });

  describe("tool schemas", () => {
    it("has 10 calendar tool schemas without an ambiguous remove action", () => {
      expect(calendarToolSchemas).toHaveLength(10);
      expect(calendarToolSchemas.some((schema) => schema.name === "m365_calendar_remove")).toBe(false);
    });
  });
});
