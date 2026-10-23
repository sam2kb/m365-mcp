/**
 * Tests for calendar tools
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import type { GraphClient } from "../graph.js";
import { registerCalendarTools, calendarToolSchemas, tzDayBounds } from "../tools/calendar.js";

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
        onlineMeeting: { joinUrl: "https://teams.microsoft.com/join" },
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
        onlineMeetingProvider: "teamsForBusiness",
        attendees: [
          { emailAddress: { address: "a@b.com" }, type: "required" },
          { emailAddress: { address: "c@d.com" }, type: "required" },
        ],
      }));
      const parsed = JSON.parse(result);
      expect(parsed.id).toBe("new-ev");
      expect(parsed.onlineMeetingUrl).toBe("https://teams.microsoft.com/join");
    });
  });

  describe("calendar_update", () => {
    it("patches only provided fields", async () => {
      await tools.calendar_update({ eventId: "ev1", subject: "Updated", location: "Room 2" });
      expect(client.patch).toHaveBeenCalledWith("/me/events/ev1", {
        subject: "Updated",
        location: { displayName: "Room 2" },
      });
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
      }), { retrySafe: true });
    });
  });

  describe("tool schemas", () => {
    it("has 9 calendar tool schemas without an ambiguous remove action", () => {
      expect(calendarToolSchemas).toHaveLength(9);
      expect(calendarToolSchemas.some((schema) => schema.name === "m365_calendar_remove")).toBe(false);
    });
  });
});

describe("calendar timezone regressions", () => {
  afterEach(() => vi.useRealTimers());
  it.each([
    ["2026-03-08T12:00:00Z", "2026-03-08T05:00:00.000Z", "2026-03-09T03:59:59.999Z"],
    ["2026-11-01T12:00:00Z", "2026-11-01T04:00:00.000Z", "2026-11-02T04:59:59.999Z"],
  ])("uses separate midnight offsets across DST on %s", (now, start, end) => {
    vi.useFakeTimers(); vi.setSystemTime(new Date(now));
    expect(tzDayBounds("America/New_York")).toEqual({ startDateTime: start, endDateTime: end });
  });
  it("uses the configured date when the host is already on the next UTC day", async () => {
    vi.useFakeTimers(); vi.setSystemTime(new Date("2026-07-19T01:00:00Z"));
    const client = mockClient({ timezone: "America/Los_Angeles" });
    vi.mocked(client.getAll).mockResolvedValue([{
      subject: "Local evening", start: { dateTime: "2026-07-19T01:00:00", timeZone: "UTC" },
      end: { dateTime: "2026-07-19T02:00:00", timeZone: "UTC" },
    }]);
    const result = await registerCalendarTools(client).calendar_today();
    expect(result).toContain("July 18");
    expect(result).toContain("6:00 PM");
    expect(client.getAll).toHaveBeenCalledWith(expect.any(String), 10, 50, { Prefer: 'outlook.timezone="UTC"' });
  });
  it("rejects half a date range", async () => {
    await expect(registerCalendarTools(mockClient()).calendar_list({ startDateTime: "2026-07-01T00:00:00Z" })).rejects.toThrow("both");
  });
  it("clears event body/location when empty strings are supplied", async () => {
    const client = mockClient();
    await registerCalendarTools(client).calendar_update({ eventId: "event", body: "", location: "" });
    expect(client.patch).toHaveBeenCalledWith("/me/events/event", { body: { contentType: "HTML", content: "" }, location: { displayName: "" } });
  });
});
