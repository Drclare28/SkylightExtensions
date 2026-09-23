/* SkylightICS
 * Minimal, dependency-free iCalendar (RFC 5545) builder.
 *
 * Design decisions:
 *  - Standalone (non-recurring) events use UTC timestamps ("Z") so no VTIMEZONE
 *    block is required.
 *  - Recurring events carry a TZID plus a VTIMEZONE block generated from the
 *    Intl API (see timezones.js) so DST transitions are honoured.
 *  - All-day events use DATE values (VALUE=DATE), end dates exclusive.
 *
 * Skylight.buildIcs(options) -> string
 * options: {
 *   title: string,
 *   location?: string,
 *   description?: string,
 *   start: Date, end: Date,
 *   allDay: boolean,
 *   timeZone?: string,          // IANA id, required when recurrence is set
 *   recurrence?: {
 *     freq: 'daily'|'weekly'|'monthly'|'yearly',
 *     interval?: number,
 *     byDay?: string[],         // ['MO','WE'] for weekly (or monthly nth-day)
 *     byMonthDay?: number,      // for monthly/yearly
 *     byMonth?: number,         // for yearly
 *     until?: Date | null       // inclusive last occurrence (converted to UTC)
 *   },
 *   uid?: string,
 *   now?: Date
 * }
 */
(function () {
  "use strict";

  if (typeof window === "undefined") return;

  var encoder = new TextEncoder();
  var decoder = new TextDecoder();

  function pad(value, width) {
    var s = String(Math.abs(value));
    while (s.length < (width || 2)) s = "0" + s;
    return s;
  }

  function isLeap(y) { return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; }

  function daysInMonth(y, m) {
    return [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
  }

  function escapeText(value) {
    return String(value == null ? "" : value)
      .replace(/\\/g, "\\\\")
      .replace(/;/g, "\\;")
      .replace(/,/g, "\\,")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .replace(/\n/g, "\\n");
  }

  // Fold a content line to <=75 octets (continuation lines <=74 with the
  // leading space). Never splits a UTF-8 sequence.
  function foldLine(content) {
    var bytes = encoder.encode(content);
    var chunks = [];
    var start = 0;
    while (start < bytes.length) {
      var max = chunks.length === 0 ? 75 : 74;
      var cut = Math.min(max, bytes.length - start);
      while (cut > 1 && start + cut < bytes.length && (bytes[start + cut] & 0xC0) === 0x80) cut--;
      if (cut === 0) cut = 1;
      chunks.push(decoder.decode(bytes.subarray(start, start + cut)));
      start += cut;
    }
    return chunks.join("\r\n ");
  }

  function icsDateTimeUTC(d) {
    return pad(d.getUTCFullYear(), 4) + pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()) +
           "T" + pad(d.getUTCHours()) + pad(d.getUTCMinutes()) + pad(d.getUTCSeconds()) + "Z";
  }

  function icsDate(d) {
    return pad(d.getFullYear(), 4) + pad(d.getMonth() + 1) + pad(d.getDate());
  }

  function icsDateTimeLocal(parts) {
    return pad(parts.y, 4) + pad(parts.mo) + pad(parts.d) + "T" + pad(parts.h) + pad(parts.mi) + pad(parts.s || 0);
  }

  function wallPartsInZone(timeZone, d) {
    var wp = window.SkylightTimezones && window.SkylightTimezones.wallParts
      ? window.SkylightTimezones.wallParts(timeZone, d.getTime())
      : null;
    if (!wp) {
      // Fallback: treat the local Date as the wall clock (device timezone).
      return { y: d.getFullYear(), mo: d.getMonth() + 1, d: d.getDate(), h: d.getHours(), mi: d.getMinutes(), s: d.getSeconds() };
    }
    return wp;
  }

  var WEEKDAY_TO_LETTER = { SUN: "SU", MON: "MO", TUE: "TU", WED: "WE", THU: "TH", FRI: "FR", SAT: "SA" };
  var LETTER_TO_OFFSET = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };

  function weekdayAdd(start, days) {
    var d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + days);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function pruneByDay(date, byDay) {
    var target = LETTER_TO_OFFSET[byDay] == null ? null : LETTER_TO_OFFSET[byDay];
    var cur = date.getDay();
    var diff = ((target - cur) % 7 + 7) % 7;
    return weekdayAdd(date, diff);
  }

  function buildRule(r) {
    var parts = ["FREQ=" + (r.freq || "WEEKLY").toUpperCase()];
    if (r.interval && r.interval > 1) parts.push("INTERVAL=" + r.interval);

    if (r.freq === "weekly") {
      if (r.byDay && r.byDay.length) parts.push("BYDAY=" + r.byDay.join(","));
    } else if (r.freq === "monthly") {
      if (r.byDay && r.byDay.length && typeof r.byDay[0] === "string" && /^[-+]?[0-9]?[A-Z]{2}$/.test(r.byDay[0])) {
        parts.push("BYDAY=" + r.byDay.join(","));
      } else if (r.byMonthDay) {
        parts.push("BYMONTHDAY=" + r.byMonthDay);
      }
    } else if (r.freq === "yearly") {
      if (r.byMonth) parts.push("BYMONTH=" + r.byMonth);
      if (r.byDay && r.byDay.length && /^[-+]?[0-9]?[A-Z]{2}$/.test(r.byDay[0])) {
        parts.push("BYDAY=" + r.byDay.join(","));
      } else if (r.byMonthDay) {
        parts.push("BYMONTHDAY=" + r.byMonthDay);
      }
    }

    if (r.until instanceof Date && !isNaN(r.until.getTime())) {
      // UNTIL for recurring (timezone-aware) events is emitted in UTC, which
      // the major calendars accept.
      parts.push("UNTIL=" + icsDateTimeUTC(clampUntil(r.until)));
    }
    return parts.join(";");
  }

  function clampUntil(until) {
    // Ensure UNTIL covers the final entire occurrence day: end of the UTC day.
    var d = new Date(Date.UTC(until.getUTCFullYear(), until.getUTCMonth(), until.getUTCDate(), 23, 59, 59));
    return d;
  }

  function guid() {
    if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
    var s = "12345678-1234-4";
    var now = Date.now();
    var hex = "";
    for (var k = 0; k < 16; k++) hex += ((now + Math.floor(Math.random() * 16)) & 0xF).toString(16);
    return s + "-" + hex;
  }

  function buildVtimezone(timeZone, startYear, untilYear) {
    var builder = window.SkylightTimezones && window.SkylightTimezones.buildVTimezone;
    if (!builder) return "";
    return builder(timeZone, startYear, untilYear);
  }

  function buildIcs(o) {
    var title = escapeText(o.title);
    var location = escapeText(o.location);
    var description = escapeText(o.description);

    var lines = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//Skylight Extensions//Send to Skylight//EN",
      "CALSCALE:GREGORIAN"
    ];

    var hasTZ = !!(o.timeZone && o.recurrence);
    if (hasTZ) {
      var startYear = o.start.getFullYear();
      var untilYear = o.recurrence && o.recurrence.until ? o.recurrence.until.getUTCFullYear() : startYear + 1;
      var vtz = buildVtimezone(o.timeZone, startYear, untilYear);
      if (vtz) {
        var vtzLines = vtz.split(/\r?\n/);
        for (var i = 0; i < vtzLines.length; i++) {
          if (vtzLines[i]) lines.push(vtzLines[i]);
        }
      }
    }

    lines.push("BEGIN:VEVENT");
    lines.push("UID:" + (o.uid || guid() + "@skylight-extensions.local"));
    if (o.now instanceof Date && !isNaN(o.now.getTime())) {
      lines.push("DTSTAMP:" + icsDateTimeUTC(o.now));
    } else {
      lines.push("DTSTAMP:" + icsDateTimeUTC(new Date()));
    }

    if (o.allDay) {
      lines.push("DTSTART;VALUE=DATE:" + icsDate(o.start));
      lines.push("DTEND;VALUE=DATE:" + icsDate(o.end));
    } else if (hasTZ) {
      var sp = wallPartsInZone(o.timeZone, o.start);
      var ep = wallPartsInZone(o.timeZone, o.end);
      lines.push("DTSTART;TZID=" + o.timeZone + ":" + icsDateTimeLocal(sp));
      lines.push("DTEND;TZID=" + o.timeZone + ":" + icsDateTimeLocal(ep));
    } else {
      lines.push("DTSTART:" + icsDateTimeUTC(o.start));
      lines.push("DTEND:" + icsDateTimeUTC(o.end));
    }

    if (title) lines.push("SUMMARY:" + title);
    else lines.push("SUMMARY:" + escapeText("Appointment"));
    if (location) lines.push("LOCATION:" + location);
    if (description) lines.push("DESCRIPTION:" + description);

    if (o.recurrence) {
      var rule = buildRule(o.recurrence);
      if (rule) lines.push("RRULE:" + rule);
    }

    lines.push("END:VEVENT");
    lines.push("END:VCALENDAR");

    return lines.map(foldLine).join("\r\n") + "\r\n";
  }

  // Small recurrence helpers shared with the app UI.
  function buildWeeklyByDay(mask, fallbackWeekday) {
    var letters = [];
    var bits = [["SU", 1], ["MO", 2], ["TU", 4], ["WE", 8], ["TH", 16], ["FR", 32], ["SA", 64]];
    for (var i = 0; i < bits.length; i++) {
      if (mask & bits[i][1]) letters.push(bits[i][0]);
    }
    if (letters.length === 0 && fallbackWeekday != null) letters.push(fallbackWeekday);
    return letters.length ? letters : null;
  }

  window.Skylight = window.Skylight || {};
  window.Skylight.ICS = {
    buildIcs: buildIcs,
    buildWeeklyByDay: buildWeeklyByDay,
    escapeText: escapeText,
    icsDate: icsDate,
    icsDateTimeUTC: icsDateTimeUTC,
    WEEKDAY_TO_OFFSET: LETTER_TO_OFFSET
  };
})();