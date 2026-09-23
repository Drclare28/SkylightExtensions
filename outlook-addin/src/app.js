/* Send to Skylight - taskpane logic (Office.js).
 * Reads the open appointment, builds an .ics + friendly text, and opens a
 * compose form to the Skylight Magic Import email, with fallbacks.
 *
 * Requirements: Mailbox 1.7 (recurrence). Features newer than 1.7 are
 * feature-detected and degraded gracefully.
 */
(function () {
  "use strict";

  var SETTINGS_KEY = "skylightEmail";
  var ICS_FILENAME = "skylight-event.ics";

  var $ = function (id) { return document.getElementById(id); };

  var DAYS = [
    ["SU", 1, "Sunday"],
    ["MO", 2, "Monday"],
    ["TU", 4, "Tuesday"],
    ["WE", 8, "Wednesday"],
    ["TH", 16, "Thursday"],
    ["FR", 32, "Friday"],
    ["SA", 64, "Saturday"]
  ];

  var DAY_MAP = {
    su: "SU", sun: "SU", sunday: "SU",
    mo: "MO", mon: "MO", monday: "MO",
    tu: "TU", tue: "TU", tuesday: "TU",
    we: "WE", wed: "WE", wednesday: "WE",
    th: "TH", thu: "TH", thursday: "TH",
    fr: "FR", fri: "FR", friday: "FR",
    sa: "SA", sat: "SA", saturday: "SA"
  };

  var WEEK_MAP = {
    first: 1, "1": 1, 1: 1,
    second: 2, "2": 2, 2: 2,
    third: 3, "3": 3, 3: 3,
    fourth: 4, "4": 4, 4: 4,
    last: -1, "-1": -1, 5: -1
  };

  var MONTH_MAP = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
    apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7,
    aug: 8, august: 8, sep: 9, september: 9, oct: 10, october: 10,
    nov: 11, november: 11, dec: 12, december: 12
  };

  var WEEK_NUMBER_NAMES = ["", "first", "second", "third", "fourth", "last"];
  var MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];

  var state = {
    recurrence: null,   // mapped {freq, interval, byDay, byMonthDay, byMonth, until}
    timeZone: "",
    itemStart: null,
    itemEnd: null,
    bodyText: "",
    category: "",
    ics: "",
    fallbackMailtoUrl: ""
  };

  /* ---------- small helpers ---------- */

  function pad(v) { return String(v).padStart(2, "0"); }

  function isoLocal(d) {
    if (!d || isNaN(d.getTime())) return "";
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) +
           "T" + pad(d.getHours()) + ":" + pad(d.getMinutes());
  }

  function parseDateTime(v) {
    if (!v) return null;
    var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(:(\d{2}))?$/.exec(v);
    if (!m) return null;
    return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[7] || 0));
  }

  function isValidEmail(v) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || "").trim());
  }

  function showError(msg) {
    var box = $("errorBox");
    if (box) {
      box.textContent = msg;
      box.style.display = msg ? "block" : "none";
    }
  }

  function clearError() {
    showError("");
  }

  function showStatus(msg) {
    var box = $("statusBox");
    if (box) {
      box.textContent = msg;
      box.style.display = msg ? "block" : "none";
    }
  }

  function clearStatus() {
    showStatus("");
  }

  function base64Utf8(str) {
    var bytes = new TextEncoder().encode(str);
    var bin = "";
    for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  function resolveTimeZone() {
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.mailbox &&
          Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.timeZone) {
        return Office.context.mailbox.userProfile.timeZone;
      }
    } catch (e) { /* ignored */ }
    try { return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC"; }
    catch (e) { return "UTC"; }
  }

  function getLocation(item) {
    if (!item || !item.location) return "";
    if (typeof item.location === "object" && "displayName" in item.location) {
      return item.location.displayName || "";
    }
    return String(item.location || "");
  }

  function readCategories(item) {
    try {
      if (item && item.categories && typeof item.categories === "object" && item.categories.length) {
        return item.categories.map(function (c) { return c.displayName || c; })[0] || "";
      }
    } catch (e) { /* ignored */ }
    return "";
  }

  /* ---------- recurrence mapping ---------- */

  function toDayLetter(val) {
    if (!val) return null;
    return DAY_MAP[String(val).trim().toLowerCase()] || null;
  }

  function parseDays(raw) {
    if (!raw) return [];
    if (Array.isArray(raw)) {
      return raw.map(toDayLetter).filter(Boolean);
    }
    if (typeof raw === "string") {
      return raw.split(/[\s,]+/).map(toDayLetter).filter(Boolean);
    }
    if (typeof raw === "number") {
      var bits = [["SU", 1], ["MO", 2], ["TU", 4], ["WE", 8], ["TH", 16], ["FR", 32], ["SA", 64]];
      var out = [];
      for (var i = 0; i < bits.length; i++) {
        if (raw & bits[i][1]) out.push(bits[i][0]);
      }
      return out;
    }
    return [];
  }

  function mapRecurrence(rec) {
    if (!rec) return null;

    var type = (rec.recurrenceType || rec.patternType || "").toLowerCase();
    var props = rec.recurrenceProperties || rec.recurrence || {};
    var seriesTime = rec.seriesTime;

    var interval = props.interval && props.interval > 1 ? Number(props.interval) : undefined;
    var mapped = { freq: null };

    if (type === "daily") {
      mapped.freq = "daily";
      if (props.days) {
        var d = parseDays(props.days);
        if (d.length && d.length < 7) {
          mapped.freq = "weekly";
          mapped.byDay = d;
        }
      }
    } else if (type === "weekly") {
      mapped.freq = "weekly";
      mapped.byDay = parseDays(props.days || props.dayOfWeekMask);
      if (!mapped.byDay.length) mapped.byDay = ["MO"];
    } else if (type === "monthly") {
      mapped.freq = "monthly";
      if (props.dayOfMonth) {
        mapped.byMonthDay = Number(props.dayOfMonth);
      } else if (props.weekNumber && (props.dayOfWeek || props.days)) {
        var wk = WEEK_MAP[String(props.weekNumber).toLowerCase()] || 1;
        var dow = toDayLetter(props.dayOfWeek || (Array.isArray(props.days) ? props.days[0] : null));
        if (dow) {
          mapped.byDay = [(wk > 0 ? String(wk) : "-1") + dow];
        }
      }
    } else if (type === "yearly") {
      mapped.freq = "yearly";
      var mo = null;
      if (props.month) {
        if (typeof props.month === "string") {
          mo = MONTH_MAP[props.month.toLowerCase()] || parseInt(props.month, 10);
        } else if (typeof props.month === "number") {
          mo = props.month >= 1 && props.month <= 12 ? props.month : props.month + 1;
        }
      }
      if (mo) mapped.byMonth = mo;

      if (props.dayOfMonth) {
        mapped.byMonthDay = Number(props.dayOfMonth);
      } else if (props.weekNumber && (props.dayOfWeek || props.days)) {
        var wk2 = WEEK_MAP[String(props.weekNumber).toLowerCase()] || 1;
        var dow2 = toDayLetter(props.dayOfWeek || (Array.isArray(props.days) ? props.days[0] : null));
        if (dow2) {
          mapped.byDay = [(wk2 > 0 ? String(wk2) : "-1") + dow2];
        }
      }
    }

    if (!mapped.freq) return null;
    if (interval) mapped.interval = interval;

    // Series end date
    var until = null;
    if (seriesTime) {
      var end = typeof seriesTime.getEndDate === "function" ? seriesTime.getEndDate() : seriesTime.endDate;
      if (end) {
        var dEnd = new Date(end);
        if (!isNaN(dEnd.getTime()) && dEnd.getUTCFullYear() < 4000) {
          until = dEnd;
        }
      }
    }
    mapped.until = until;

    return mapped;
  }

  /* ---------- friendly text ---------- */

  function friendlyDateTime(d) {
    if (!d || isNaN(d.getTime())) return "";
    return d.toLocaleString(undefined, {
      weekday: "long", year: "numeric", month: "long", day: "numeric",
      hour: "numeric", minute: "2-digit"
    });
  }

  function friendlyDate(d) {
    if (!d || isNaN(d.getTime())) return "";
    return d.toLocaleDateString(undefined, {
      weekday: "long", year: "numeric", month: "long", day: "numeric"
    });
  }

  function describeRecurrence(mapped) {
    if (!mapped) return "";
    var f = mapped.freq;
    if (f === "daily") {
      return mapped.interval && mapped.interval > 1 ? ("Every " + mapped.interval + " days") : "Every day";
    }
    var everyPrefix = "Every" + (mapped.interval && mapped.interval > 1 ? (" " + mapped.interval) : "");
    if (f === "weekly") {
      var names = DAYS.filter(function (pair) {
        return mapped.byDay && mapped.byDay.indexOf(pair[0]) !== -1;
      }).map(function (pair) { return pair[2]; });
      return everyPrefix + " week" + (names.length ? " on " + names.join(", ") : "");
    }
    if (f === "monthly") {
      if (mapped.byDay && /^[-+0-9][0-9]?[A-Z]{2}$/.test(mapped.byDay[0])) {
        var wk = /^([-+]?\d)/.exec(mapped.byDay[0]);
        var dowCode = mapped.byDay[0].slice(-2);
        var dow = DAYS.filter(function (p) { return p[0] === dowCode; });
        var wkNum = wk && +wk[1] > 0 ? +wk[1] : 5;
        return everyPrefix + " month on the " + WEEK_NUMBER_NAMES[wkNum] + " " + (dow[0] ? dow[0][2] : "");
      }
      return everyPrefix + " month on day " + (mapped.byMonthDay != null ? mapped.byMonthDay : "");
    }
    if (f === "yearly") {
      var moName = mapped.byMonth ? MONTH_NAMES[(+mapped.byMonth) - 1] : "";
      if (mapped.byDay && /^[-+0-9][0-9]?[A-Z]{2}$/.test(mapped.byDay[0])) {
        var wk2 = /^([-+]?\d)/.exec(mapped.byDay[0]);
        var dowCode2 = mapped.byDay[0].slice(-2);
        var dow2 = DAYS.filter(function (p) { return p[0] === dowCode2; });
        var wkNum2 = wk2 && +wk2[1] > 0 ? +wk2[1] : 5;
        return everyPrefix + " year in " + moName + " on the " +
          WEEK_NUMBER_NAMES[wkNum2] + " " + (dow2[0] ? dow2[0][2] : "");
      }
      return everyPrefix + " year in " + moName + " on day " + (mapped.byMonthDay != null ? mapped.byMonthDay : "");
    }
    return "";
  }

  function buildTextBody() {
    var allDay = $("chkAllDay").checked;
    var start = parseDateTime($("txtStart").value) || state.itemStart || new Date();
    var end = parseDateTime($("txtEnd").value) || state.itemEnd || new Date(start.getTime() + 30 * 60000);
    var title = ($("txtTitle").value || "").trim();
    var location = ($("txtLocation").value || "").trim();
    var details = ($("txtDetails").value || "").trim();

    var lines = [title || "Appointment"];
    if (allDay) {
      lines.push("Start: " + friendlyDate(start) + " (all day)");
      if (end && (end.getTime() - start.getTime() > 86400000)) {
        var last = new Date(end.getFullYear(), end.getMonth(), end.getDate() - 1);
        lines.push("End: " + friendlyDate(last));
      }
    } else {
      lines.push("Start: " + friendlyDateTime(start));
      lines.push("End: " + friendlyDateTime(end));
    }
    if (location) lines.push("Location: " + location);
    var rep = describeRecurrence(state.recurrence);
    if (rep) lines.push("Repeats: " + rep);
    if (details) lines.push("Details: " + details);
    return lines.join("\n");
  }

  /* ---------- ICS construction ---------- */

  function buildIcs() {
    var allDay = $("chkAllDay").checked;
    var start = parseDateTime($("txtStart").value) || state.itemStart || new Date();
    var end = parseDateTime($("txtEnd").value) || state.itemEnd || new Date(start.getTime() + 30 * 60000);
    var opts = {
      title: $("txtTitle").value || "Appointment",
      location: $("txtLocation").value || "",
      description: $("txtDetails").value || "",
      start: start,
      end: end,
      allDay: allDay,
      timeZone: state.timeZone || resolveTimeZone(),
      recurrence: state.recurrence,
      now: new Date()
    };
    var content = window.Skylight && window.Skylight.ICS
      ? window.Skylight.ICS.buildIcs(opts)
      : "";
    state.ics = content;
    return content;
  }

  /* ---------- sending ---------- */

  function sendToSkylight() {
    clearError();
    clearStatus();
    var email = ($("txtSkylightEmail").value || "").trim();
    if (!isValidEmail(email)) {
      showError("Enter a valid Skylight import email.");
      return;
    }
    saveEmail();

    var title = ($("txtTitle").value || "").trim() || "Appointment";
    var category = ($("txtCategory").value || "").trim();

    var ics = buildIcs();
    var body = buildTextBody();
    var subject = category ? (category + " - " + title) : title;
    var formData = {
      toRecipients: [email],
      subject: subject,
      body: { type: "text", content: body }
    };

    // Download ICS so user always has the file ready to attach
    downloadIcs(false);

    var mb = typeof Office !== "undefined" && Office.context ? Office.context.mailbox : null;
    if (mb && typeof mb.displayNewMessageForm === "function") {
      try {
        mb.displayNewMessageForm(formData);
        showStatus("Compose window opened! skylight-event.ics has also been saved to your downloads.");
        return;
      } catch (e) {
        showFallback();
        return;
      }
    }
    if (mb && typeof mb.displayNewMessageFormAsync === "function") {
      mb.displayNewMessageFormAsync(formData, function (result) {
        if (result && result.status === Office.AsyncResultStatus.Failed) {
          showFallback();
        } else {
          showStatus("Compose window opened! skylight-event.ics has also been saved to your downloads.");
        }
      });
      return;
    }
    showFallback();
  }

  function showFallback() {
    var email = ($("txtSkylightEmail").value || "").trim();
    var title = ($("txtTitle").value || "").trim() || "Appointment";
    var category = ($("txtCategory").value || "").trim();
    var subject = category ? (category + " - " + title) : title;

    var mailto = "mailto:" + encodeURIComponent(email) +
      "?subject=" + encodeURIComponent(subject) +
      "&body=" + encodeURIComponent(buildTextBody());
    state.fallbackMailtoUrl = mailto;

    var fallbackArea = $("fallbackArea");
    if (fallbackArea) fallbackArea.style.display = "block";
    showStatus("Compose window unavailable. Use 'Open in Email App' or copy/download the .ics file below.");
  }

  /* ---------- fallbacks: copy / download ---------- */

  function downloadIcs(showConfirmation) {
    buildIcs();
    var blob = new Blob([state.ics], { type: "text/calendar;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = ICS_FILENAME;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    if (showConfirmation !== false) {
      showStatus("Downloaded " + ICS_FILENAME);
    }
  }

  function copyIcs() {
    buildIcs();
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(state.ics).then(function () {
        showStatus("ICS content copied to clipboard!");
      }, function () {
        legacyCopy();
      });
      return;
    }
    legacyCopy();
  }

  function legacyCopy() {
    var ta = document.createElement("textarea");
    ta.value = state.ics;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
      showStatus("ICS content copied to clipboard!");
    } catch (e) {
      showError("Copy failed - please use Download instead.");
    }
    ta.remove();
  }

  /* ---------- settings ---------- */

  function loadEmail() {
    var v = "";
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
        v = Office.context.roamingSettings.get(SETTINGS_KEY) || "";
      }
    } catch (e) { /* ignored */ }
    if (!v) {
      try { v = localStorage.getItem(SETTINGS_KEY) || ""; } catch (e2) { /* ignored */ }
    }
    $("txtSkylightEmail").value = v;
  }

  function saveEmail() {
    var v = ($("txtSkylightEmail").value || "").trim();
    if (!isValidEmail(v)) return;
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.set(SETTINGS_KEY, v);
        Office.context.roamingSettings.saveAsync(function () {});
      }
    } catch (e) { /* ignored */ }
    try { localStorage.setItem(SETTINGS_KEY, v); } catch (e2) { /* ignored */ }
  }

  /* ---------- recurrence UI ---------- */

  function showRecurrence(description) {
    $("recurrenceText").textContent = description || "No";
    $("recurrenceBox").style.display = description ? "block" : "none";
  }

  function wireReload() {
    ["txtTitle", "txtCategory", "txtStart", "txtEnd", "chkAllDay", "txtLocation", "txtDetails"].forEach(function (id) {
      var el = $(id);
      if (el) el.addEventListener("input", function () { validateAndEnable(); });
      if (el) el.addEventListener("change", function () { validateAndEnable(); });
    });
  }

  function validateAndEnable() {
    var okEmail = isValidEmail($("txtSkylightEmail").value);
    var hasTitle = !!($("txtTitle").value || "").trim();
    $("btnSend").disabled = !(okEmail && hasTitle);
    if (okEmail) clearError();
  }

  /* ---------- sample data for test / browser mode ---------- */

  function populateSampleData() {
    var now = new Date();
    var tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 17, 30);
    var tomorrowEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 18, 30);

    state.itemStart = tomorrow;
    state.itemEnd = tomorrowEnd;
    state.timeZone = resolveTimeZone();
    state.recurrence = {
      freq: "weekly",
      interval: 1,
      byDay: ["TU", "TH"],
      until: new Date(now.getFullYear(), now.getMonth() + 2, now.getDate())
    };

    $("txtTitle").value = "Soccer Practice";
    $("txtCategory").value = "Kids Sports";
    $("txtLocation").value = "Memorial Park, Field 3";
    $("txtDetails").value = "Bring water bottle and shin guards.";
    $("txtStart").value = isoLocal(state.itemStart);
    $("txtEnd").value = isoLocal(state.itemEnd);
    $("chkAllDay").checked = false;

    showRecurrence(describeRecurrence(state.recurrence));
    showStatus("Browser testing mode active. Edit fields or click 'Download .ics' / 'Copy .ics'.");
  }

  /* ---------- reading Outlook item ---------- */

  function readItem(item) {
    // 1. Title / Subject
    if (typeof item.subject === "object" && typeof item.subject.getAsync === "function") {
      item.subject.getAsync(function (res) {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) {
          $("txtTitle").value = res.value || "";
          validateAndEnable();
        }
      });
    } else {
      $("txtTitle").value = item.subject || "";
    }

    // 2. Location
    if (typeof item.location === "object" && typeof item.location.getAsync === "function") {
      item.location.getAsync(function (res) {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) {
          $("txtLocation").value = getLocation({ location: res.value });
        }
      });
    } else {
      $("txtLocation").value = getLocation(item);
    }

    // 3. Category
    if (typeof item.categories === "object" && typeof item.categories.getAsync === "function") {
      item.categories.getAsync(function (res) {
        if (res && res.status === Office.AsyncResultStatus.Succeeded) {
          $("txtCategory").value = readCategories({ categories: res.value });
        }
      });
    } else {
      $("txtCategory").value = readCategories(item);
    }

    // 4. Start / End
    function applyDates(start, end) {
      state.itemStart = start && !isNaN(start.getTime()) ? start : new Date();
      state.itemEnd = end && !isNaN(end.getTime()) ? end : new Date(state.itemStart.getTime() + 30 * 60000);
      if (state.itemEnd <= state.itemStart) {
        state.itemEnd = new Date(state.itemStart.getTime() + 30 * 60000);
      }
      $("txtStart").value = isoLocal(state.itemStart);
      $("txtEnd").value = isoLocal(state.itemEnd);
      validateAndEnable();
    }

    if (item.start && typeof item.start.getAsync === "function") {
      item.start.getAsync(function (resStart) {
        var s = resStart && resStart.status === Office.AsyncResultStatus.Succeeded ? new Date(resStart.value) : null;
        if (item.end && typeof item.end.getAsync === "function") {
          item.end.getAsync(function (resEnd) {
            var e = resEnd && resEnd.status === Office.AsyncResultStatus.Succeeded ? new Date(resEnd.value) : null;
            applyDates(s, e);
          });
        } else {
          applyDates(s, item.end ? new Date(item.end) : null);
        }
      });
    } else {
      applyDates(
        item.start ? new Date(item.start) : null,
        item.end ? new Date(item.end) : null
      );
    }

    // 5. All day
    if (item.allDayEvent !== undefined) {
      $("chkAllDay").checked = !!item.allDayEvent;
    } else if (state.itemStart && state.itemEnd) {
      var dur = state.itemEnd - state.itemStart;
      if (dur >= 86400000 && state.itemStart.getHours() === 0 && state.itemStart.getMinutes() === 0) {
        $("chkAllDay").checked = true;
      }
    }

    // 6. Body details
    if (item.body && typeof item.body.getAsync === "function") {
      item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result && result.status === Office.AsyncResultStatus.Succeeded) {
          state.bodyText = result.value || "";
          $("txtDetails").value = state.bodyText;
        }
      });
    }

    // 7. Time zone
    state.timeZone = resolveTimeZone();

    // 8. Recurrence (Office.js Mailbox 1.7+)
    if (item.recurrence && typeof item.recurrence.getAsync === "function") {
      item.recurrence.getAsync(function (result) {
        if (result && result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          state.recurrence = mapRecurrence(result.value);
          showRecurrence(describeRecurrence(state.recurrence));
        } else {
          showRecurrence("");
        }
      });
    } else if (item.recurrence) {
      state.recurrence = mapRecurrence(item.recurrence);
      showRecurrence(describeRecurrence(state.recurrence));
    } else {
      showRecurrence("");
    }
  }

  /* ---------- initialization ---------- */

  function init() {
    wireReload();
    loadEmail();
    $("txtSkylightEmail").addEventListener("change", function () { saveEmail(); validateAndEnable(); });
    $("txtSkylightEmail").addEventListener("input", validateAndEnable);
    $("btnSend").addEventListener("click", sendToSkylight);
    $("btnMailto").addEventListener("click", function () {
      if (state.fallbackMailtoUrl) {
        window.location.href = state.fallbackMailtoUrl;
      } else {
        showFallback();
      }
    });
    $("btnCopyIcs").addEventListener("click", copyIcs);
    $("btnDownloadIcs").addEventListener("click", function () { downloadIcs(true); });
    validateAndEnable();
  }

  function start() {
    init();
    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady(function (info) {
        var hasItem = false;
        try {
          if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
            readItem(Office.context.mailbox.item);
            hasItem = true;
          }
        } catch (e) {
          showError("Could not read appointment: " + e.message);
        }
        if (!hasItem) {
          // If Office.onReady resolved in a browser test or without an active appointment
          populateSampleData();
        }
      });
    } else {
      document.addEventListener("DOMContentLoaded", function () {
        populateSampleData();
      });
    }
  }

  start();
})();