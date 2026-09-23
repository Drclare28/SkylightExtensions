/* SkylightTimezones
 * Builds an ICS VTIMEZONE component for an IANA time zone using only the
 * Intl.DateTimeFormat API. It samples offsets across a few years, detects the
 * DST transitions, and derives YEARLY rules (or falls back to per-year
 * transitions). Libraries such as JSJoda/luxon/moment are intentionally avoided.
 *
 * Exposed globals (all optional, used defensively by ics.js):
 *   SkylightTimezones.buildVTimezone(tz, startYear, untilYear) -> string
 *   SkylightTimezones.wallParts(tz, utcMs) -> {y,mo,d,h,mi,s} | null
 */
(function () {
  "use strict";

  if (typeof window === "undefined") return;

  var WEEKDAY_NAMES = { 0: "SU", 1: "MO", 2: "TU", 3: "WE", 4: "TH", 5: "FR", 6: "SA" };

  var formatterCache = {};

  function pad(value, width) {
    var s = String(Math.abs(value));
    while (s.length < (width || 2)) s = "0" + s;
    return s;
  }

  function fmtOffset(minutes) {
    var sign = minutes < 0 ? "-" : "+";
    var abs = Math.abs(minutes);
    return sign + pad(Math.floor(abs / 60)) + pad(abs % 60);
  }

  function formatterFor(timeZone) {
    if (formatterCache[timeZone]) return formatterCache[timeZone];
    try {
      var f = new Intl.DateTimeFormat("en-US", {
        timeZone: timeZone,
        year: "numeric", month: "2-digit", day: "2-digit",
        hour: "2-digit", minute: "2-digit", second: "2-digit",
        hourCycle: "h23"
      });
      formatterCache[timeZone] = f;
      return f;
    } catch (e) {
      formatterCache[timeZone] = null;
      return null;
    }
  }

  // Parse the wall-clock parts of a UTC instant rendered in the given zone.
  function wallClockAsUTC(timeZone, utcMs) {
    var formatter = formatterFor(timeZone);
    if (!formatter) return null;
    var parts = formatter.formatToParts(new Date(utcMs));
    var vals = {};
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (p.type !== "literal") vals[p.type] = parseInt(p.value, 10);
    }
    if (vals.year === undefined || vals.month === undefined ||
        vals.day === undefined || vals.hour === undefined) return null;
    var hour = vals.hour || 0;
    if (hour === 24) hour = 0; // some environments emit "24" at midnight
    return Date.UTC(vals.year, vals.month - 1, vals.day, hour,
                    vals.minute || 0, vals.second || 0);
  }

  function getOffsetMinutes(timeZone, utcMs) {
    var asUTC = wallClockAsUTC(timeZone, utcMs);
    if (asUTC === null) return null;
    return Math.round((asUTC - utcMs) / 60000);
  }

  // Wall-clock parts of a UTC instant shifted by a known offset.
  function localParts(utcMs, offsetMinutes) {
    var d = new Date(utcMs + offsetMinutes * 60000);
    return {
      y: d.getUTCFullYear(),
      mo: d.getUTCMonth() + 1,
      d: d.getUTCDate(),
      h: d.getUTCHours(),
      mi: d.getUTCMinutes(),
      s: d.getUTCSeconds()
    };
  }

  function wallDayOfWeek(parts) {
    return new Date(Date.UTC(parts.y, parts.mo - 1, parts.d)).getUTCDay();
  }

  function binarySearch(timeZone, loMs, hiMs, fromOffset) {
    while (hiMs - loMs > 60000) {
      var mid = loMs + Math.floor((hiMs - loMs) / 2);
      var midOff = getOffsetMinutes(timeZone, mid);
      if (midOff === null) break;
      if (midOff === fromOffset) loMs = mid; else hiMs = mid;
    }
    return { utcMs: hiMs, to: getOffsetMinutes(timeZone, hiMs) };
  }

  // Probe the zone at noon UTC every day throughout the calendar year and
  // recover every offset transition. Returns [{utcMs, from, to, isDaylight}].
  function findTransitions(timeZone, year) {
    var transitions = [];
    var prevOffset = null;
    var dayMs = 86400000;
    for (var i = 0; i < 366; i++) {
      var utcMs = Date.UTC(year, 0, 1, 12, 0, 0) + i * dayMs;
      var cur = getOffsetMinutes(timeZone, utcMs);
      if (cur === null) continue;
      if (prevOffset !== null && cur !== prevOffset) {
        var lo = utcMs - dayMs;
        var hit = binarySearch(timeZone, lo, utcMs, prevOffset);
        transitions.push({
          utcMs: hit.utcMs,
          from: prevOffset,
          to: hit.to,
          isDaylight: hit.to > prevOffset
        });
      }
      prevOffset = cur;
    }
    return transitions;
  }

  function describe(parts) {
    return {
      month: parts.mo,
      byDay: {
        weekday: WEEKDAY_NAMES[wallDayOfWeek(parts)],
        index: Math.ceil(parts.d / 7)
      },
      byMonthDay: { day: parts.d }
    };
  }

  function ruleSignature(desc) {
    return desc.month + ":" + desc.byDay.index + desc.byDay.weekday;
  }

  // Group transitions by lifecycle (standard vs daylight) and derive either a
  // YEARLY rule or an explicit list of per-year transitions.
  function deriveUnits(timeZone, yearRange) {
    var all = [];
    for (var y = yearRange[0]; y <= yearRange[1]; y++) {
      all = all.concat(findTransitions(timeZone, y));
    }
    if (all.length) all.sort(function (a, b) { return a.utcMs - b.utcMs; });

    // Collapse near-duplicate detections of the same transition (same from/to
    // offsets within a 2-minute window, e.g. the same DST change found on both
    // sides of a sampled year boundary). Keep the earliest detection.
    var deduped = [];
    for (var i = 0; i < all.length; i++) {
      var t = all[i];
      var prev = deduped[deduped.length - 1];
      if (prev && prev.from === t.from && prev.to === t.to &&
          (t.utcMs - prev.utcMs) <= 120000) {
        continue;
      }
      deduped.push(t);
    }
    all = deduped;

    if (all.length === 0) {
      // No transitions: fixed offset year-round.
      var anchor = Date.UTC(yearRange[1], 5, 1, 12, 0, 0);
      var off = getOffsetMinutes(timeZone, anchor);
      if (off === null) return null; // unresolvable zone
      return [{
        kind: "standard",
        offsetFrom: off,
        offsetTo: off,
        dstart: "19700101T000000",
        rrule: null,
        explicit: null
      }];
    }

    var units = [];
    var groups = { standard: [], daylight: [] };
    all.forEach(function (t) {
      var key = t.isDaylight ? "daylight" : "standard";
      groups[key].push(t);
    });

    function buildGroup(key) {
      var list = groups[key] || [];
      if (!list.length) return;
      list.sort(function (a, b) { return a.utcMs - b.utcMs; });
      var first = list[0];
      var firstParts = localParts(first.utcMs, first.from);
      var firstDesc = describe(firstParts);
      var dstart = localPartsForVTimezone(first.utcMs, first.from);

      // Check for a consistent formula across every sampled year.
      var byDayOk = true, byMonthDayOk = true;
      for (var i = 0; i < list.length; i++) {
        var p = localParts(list[i].utcMs, list[i].from);
        var d = describe(p);
        if (ruleSignature(d) !== ruleSignature(firstDesc)) byDayOk = false;
        if (d.month !== firstDesc.month || d.byMonthDay.day !== firstDesc.byMonthDay.day) byMonthDayOk = false;
      }

      var rrule = null;
      if (byDayOk && list.length >= 2) {
        rrule = "FREQ=YEARLY;BYMONTH=" + firstDesc.month +
                ";BYDAY=" + firstDesc.byDay.index + firstDesc.byDay.weekday;
      } else if (byMonthDayOk && list.length >= 2) {
        rrule = "FREQ=YEARLY;BYMONTH=" + firstDesc.month +
                ";BYMONTHDAY=" + firstDesc.byMonthDay.day;
      }

      var explicit = null;
      if (!rrule) {
        explicit = list.map(function (t) {
          return {
            dstart: localPartsForVTimezone(t.utcMs, t.from),
            offsetFrom: t.from,
            offsetTo: t.to
          };
        });
      }

      units.push({
        kind: key,
        offsetFrom: first.from,
        offsetTo: first.to,
        dstart: dstart,
        rrule: rrule,
        explicit: explicit
      });
    }

    buildGroup("daylight");
    buildGroup("standard");
    return units;
  }

  function localPartsForVTimezone(utcMs, offsetMinutes) {
    var p = localParts(utcMs, offsetMinutes);
    return pad(p.y, 4) + pad(p.mo) + pad(p.d) + "T" + pad(p.h) + pad(p.mi) + "00";
  }

  function componentForUnit(tz, unit) {
    var lines = ["BEGIN:" + (unit.kind === "daylight" ? "DAYLIGHT" : "STANDARD")];
    if (unit.explicit) {
      // Explicit per-year transitions; each row carries its own offsets so the
      // component DTSTART is emitted once by the list itself.
      for (var i = 0; i < unit.explicit.length; i++) {
        var e = unit.explicit[i];
        lines.push("DTSTART:" + e.dstart);
        lines.push("TZOFFSETFROM:" + fmtOffset(e.offsetFrom));
        lines.push("TZOFFSETTO:" + fmtOffset(e.offsetTo));
      }
    } else {
      lines.push("DTSTART:" + unit.dstart);
      if (unit.rrule) lines.push("RRULE:" + unit.rrule);
      lines.push("TZOFFSETFROM:" + fmtOffset(unit.offsetFrom));
      lines.push("TZOFFSETTO:" + fmtOffset(unit.offsetTo));
    }
    lines.push("END:" + (unit.kind === "daylight" ? "DAYLIGHT" : "STANDARD"));
    return lines.join("\r\n");
  }

  function buildVTimezone(timeZone, startYear, untilYear) {
    if (!timeZone) return "";
    if (formatterFor(timeZone) === null) return ""; // unknown/unresolvable zone

    if (!startYear) startYear = new Date().getUTCFullYear();
    var hiYear = (untilYear && untilYear > startYear + 1) ? Math.min(untilYear, startYear + 3) : startYear + 1;
    var loYear = startYear - 1 > 0 ? startYear - 1 : startYear;
    var yearRange = [loYear, hiYear];

    var units = deriveUnits(timeZone, yearRange);
    if (!units) return "";

    var lines = ["BEGIN:VTIMEZONE", "TZID:" + timeZone];
    for (var i = 0; i < units.length; i++) {
      lines.push(componentForUnit(timeZone, units[i]));
    }
    lines.push("END:VTIMEZONE");
    return lines.join("\r\n");
  }

  // Public: wall parts of utcMs rendered in `timeZone`.
  function wallParts(timeZone, utcMs) {
    var asUTC = wallClockAsUTC(timeZone, utcMs);
    if (asUTC === null) return null;
    var d = new Date(asUTC);
    return {
      y: d.getUTCFullYear(),
      mo: d.getUTCMonth() + 1,
      d: d.getUTCDate(),
      h: d.getUTCHours(),
      mi: d.getUTCMinutes(),
      s: d.getUTCSeconds()
    };
  }

  window.SkylightTimezones = {
    buildVTimezone: buildVTimezone,
    wallParts: wallParts,
    getOffsetMinutes: getOffsetMinutes
  };
})();