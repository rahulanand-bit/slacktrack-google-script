/***********************
 * SlackTrack MVP + Phase 2 (Apps Script)
 * - Button-based attendance updates
 * - Morning/evening reminders
 * - Chat intent parsing via Gemini (queued processing)
 ***********************/

const CFG = {
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  NAME_COL: 1,          // A: Employee Name
  SLACK_ID_COL: 2,      // B: Slack User ID
  EMAIL_COL: 3,         // C: Email
  LAST_UPDATED_COL: 4,  // D: Last Updated at
  FIRST_DAY_COL: 6,     // F onward day columns
  CHAT_QUEUE_KEY: "CHAT_EVENT_QUEUE_JSON",
  CHAT_QUEUE_LIMIT: 200
};

/**
 * Slack entrypoint
 * Handles:
 * 1) Slack Events API (url_verification + message.im)
 * 2) Interactive button payload
 * 3) Optional slash command (/attendance)
 */
function doPost(e) {
  const start = new Date();
  try {
    const tz = getProp_("TIMEZONE", false) || Session.getScriptTimeZone();

    // 1) Try JSON body first (Slack Events API)
    const raw = e?.postData?.contents || "";
    if (raw && raw.trim().startsWith("{")) {
      const body = JSON.parse(raw);

      // URL verification challenge
      if (body.type === "url_verification" && body.challenge) {
        return ContentService.createTextOutput(body.challenge);
      }

      // Events API callbacks (message.im)
      if (body.type === "event_callback") {
        const expectedTeamId = getProp_("SLACK_TEAM_ID", false);
        if (expectedTeamId && body?.team_id && body.team_id !== expectedTeamId) {
          return ContentService.createTextOutput("ok");
        }

        // Deduplicate by event_id
        const eventId = body.event_id || `evt_${Date.now()}`;
        if (isDuplicate_(`evt:${eventId}`, 24 * 60 * 60)) {
          return ContentService.createTextOutput("ok");
        }

        const ev = body.event || {};
        if (
          ev.type === "message" &&
          ev.channel_type === "im" &&
          !ev.bot_id &&
          !ev.subtype
        ) {
          enqueueChatEvent_(ev);
        }

        // ACK quickly to avoid Slack timeout
        return ContentService.createTextOutput("ok");
      }
    }

    // 2) Interactive payload (buttons)
    const payload = getSlackPayload_(e);
    if (payload) {
      const expectedTeamId = getProp_("SLACK_TEAM_ID", false);
      if (expectedTeamId && payload?.team?.id !== expectedTeamId) {
        return jsonOut_({ ok: false, error: "Team mismatch" });
      }

      const userId = payload?.user?.id;
      const actionId = payload?.actions?.[0]?.action_id;
      const attendanceValue = mapActionToValue_(actionId);

      if (!userId || !attendanceValue) {
        return jsonOut_({
          response_type: "ephemeral",
          text: "Invalid action payload."
        });
      }

      const actionTs = payload?.action_ts || String(Date.now());
      const dedupeKey = `dedupe:${userId}:${actionId}:${actionTs}`;
      if (isDuplicate_(dedupeKey, 6 * 60 * 60)) {
        return jsonOut_({ ok: true, message: "Duplicate ignored" });
      }

      const result = updateTodayAttendance_(userId, attendanceValue, tz);

      if (!result.ok) {
        console.error("Update failed:", JSON.stringify(result));
        return jsonOut_({
          response_type: "ephemeral",
          text: `Update failed: ${result.error}`
        });
      }

      console.log(`Updated ${userId} -> ${attendanceValue} in ${new Date() - start}ms`);
      return jsonOut_({
        response_type: "ephemeral",
        replace_original: false,
        text: `Done. Marked *${attendanceValue}* for today.`
      });
    }

    // 3) Optional slash command
    const command = e?.parameter?.command || null;
    const userId = e?.parameter?.user_id || null;
    if (command === "/attendance" && userId) {
      sendAttendanceButtonsDM_(userId);
      return ContentService.createTextOutput("Check your DM from SlackTrack.");
    }

    console.log("No payload/command. postData:", raw);
    return ContentService.createTextOutput("OK");
  } catch (err) {
    console.error("doPost error:", err);
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getSlackPayload_(e) {
  const p1 = e?.parameter?.payload;
  if (p1) return JSON.parse(p1);

  const raw = e?.postData?.contents;
  if (!raw) return null;

  const match = raw.match(/(?:^|&)payload=([^&]+)/);
  if (match && match[1]) {
    const decoded = decodeURIComponent(match[1].replace(/\+/g, " "));
    return JSON.parse(decoded);
  }
  return null;
}

/* =========================
   Chat queue + processor
========================= */

function enqueueChatEvent_(eventObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const current = props.getProperty(CFG.CHAT_QUEUE_KEY);
    const queue = current ? JSON.parse(current) : [];

    queue.push({
      user: eventObj.user || "",
      channel: eventObj.channel || "",
      text: eventObj.text || "",
      ts: eventObj.ts || String(Date.now())
    });

    // keep bounded queue
    while (queue.length > CFG.CHAT_QUEUE_LIMIT) queue.shift();

    props.setProperty(CFG.CHAT_QUEUE_KEY, JSON.stringify(queue));
  } finally {
    lock.releaseLock();
  }
}

function dequeueChatEvents_(maxItems) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const current = props.getProperty(CFG.CHAT_QUEUE_KEY);
    const queue = current ? JSON.parse(current) : [];
    const n = Math.max(1, maxItems || 10);

    const batch = queue.splice(0, n);
    props.setProperty(CFG.CHAT_QUEUE_KEY, JSON.stringify(queue));
    return batch;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Add a time trigger (every minute or every 5 min):
 * function name: processQueuedChatEvents
 */
function processQueuedChatEvents() {
  const tz = getProp_("TIMEZONE", false) || Session.getScriptTimeZone();
  const batch = dequeueChatEvents_(10);
  if (!batch.length) return;

  batch.forEach(ev => {
    try {
      handleChatMessage_(ev.user, ev.channel, ev.text, tz);
    } catch (err) {
      console.error("Chat processing failed:", err);
      // send fallback message
      try {
        sendTextDM_(ev.channel, "I couldn't process that. Please try: `WFH today` or `leave 3 Mar to 5 Mar`.");
      } catch (_) {}
    }
  });
}

function handleChatMessage_(userId, channelId, text, tz) {
  const parsed = parseIntentWithGemini_(text, tz);

  if (parsed.needs_clarification) {
    sendTextDM_(channelId, parsed.clarification_question || "Can you clarify your request?");
    return;
  }

  if (parsed.intent === "status_check") {
    const status = getStatusForDate_(userId, new Date(), tz);
    sendTextDM_(channelId, `Your status today: ${status || "Not updated"}`);
    return;
  }

  if (parsed.intent !== "attendance_update" && parsed.intent !== "leave") {
    sendTextDM_(channelId, "Try messages like: `WFH today`, `WFO tomorrow`, `leave 3 Mar to 5 Mar`, `half day on 7 Mar`.");
    return;
  }

  const value = parsed.attendance_value;
  if (!value || !["WFO", "WFH", "-1", "-0.5"].includes(value)) {
    sendTextDM_(channelId, "Please specify one of: WFO, WFH, leave, half day.");
    return;
  }

  const dates = resolveIntentDates_(parsed, tz);
  if (!dates.length) {
    sendTextDM_(channelId, "I couldn't determine the date(s). Please mention dates like `3 Mar` or `3 Mar to 5 Mar`.");
    return;
  }

  let okCount = 0;
  const failed = [];
  dates.forEach(d => {
    const r = updateAttendanceForDate_(userId, value, d, tz);
    if (r.ok) okCount++;
    else failed.push(`${formatDateYmd_(d, tz)} (${r.error})`);
  });

  const dateLabel = dates.length === 1
    ? formatDateYmd_(dates[0], tz)
    : `${formatDateYmd_(dates[0], tz)} to ${formatDateYmd_(dates[dates.length - 1], tz)}`;

  if (!failed.length) {
    sendTextDM_(channelId, `Done. Marked *${value}* for ${dateLabel}.`);
  } else {
    sendTextDM_(channelId, `Partially updated: ${okCount}/${dates.length}. Failed: ${failed.join("; ")}`);
  }
}

/* =========================
   Gemini intent parsing
========================= */

function parseIntentWithGemini_(text, tz) {
  const apiKey = getProp_("GEMINI_API_KEY");
  const today = formatDateYmd_(new Date(), tz);

  const prompt = `
You are an attendance intent parser for Slack.

Return ONLY JSON. No markdown, no explanation.

Schema:
{
  "intent": "attendance_update|leave|status_check|unknown",
  "attendance_value": "WFO|WFH|-1|-0.5|null",
  "dates": ["YYYY-MM-DD"],
  "from_date": "YYYY-MM-DD|null",
  "to_date": "YYYY-MM-DD|null",
  "needs_clarification": true|false,
  "clarification_question": "string|null"
}

Rules:
- "leave" => attendance_value "-1" unless half day is explicit then "-0.5"
- "half day" => "-0.5"
- "wfh" => "WFH", "wfo" => "WFO"
- If user asks "status today" => intent "status_check"
- If date not provided for update/leave, assume today (${today})
- If ambiguous, set needs_clarification true with a short question.
- For ranges, set from_date and to_date.
- For explicit multiple dates, fill dates array.

User text:
"""${text}"""
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 400 }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const raw = resp.getContentText();
  const json = JSON.parse(raw);
  const out = json?.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
  const parsed = parseJsonLoose_(out);

  return {
    intent: parsed.intent || "unknown",
    attendance_value: parsed.attendance_value || null,
    dates: Array.isArray(parsed.dates) ? parsed.dates : [],
    from_date: parsed.from_date || null,
    to_date: parsed.to_date || null,
    needs_clarification: !!parsed.needs_clarification,
    clarification_question: parsed.clarification_question || null
  };
}

function parseJsonLoose_(text) {
  const cleaned = String(text || "")
    .replace(/```json/gi, "")
    .replace(/```/g, "")
    .trim();

  const start = cleaned.indexOf("{");
  const end = cleaned.lastIndexOf("}");
  if (start === -1 || end === -1 || end < start) return {};
  return JSON.parse(cleaned.slice(start, end + 1));
}

/* =========================
   Date resolution
========================= */

function resolveIntentDates_(parsed, tz) {
  if (Array.isArray(parsed.dates) && parsed.dates.length) {
    return parsed.dates.map(d => parseYmd_(d)).filter(Boolean);
  }

  if (parsed.from_date && parsed.to_date) {
    const from = parseYmd_(parsed.from_date);
    const to = parseYmd_(parsed.to_date);
    if (!from || !to || from > to) return [];

    const dates = [];
    const cursor = new Date(from);
    while (cursor <= to && dates.length < 60) {
      dates.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    return dates;
  }

  // default today for attendance update / leave
  if (parsed.intent === "attendance_update" || parsed.intent === "leave") {
    return [new Date()];
  }

  return [];
}

function parseYmd_(s) {
  const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const d = Number(m[3]);
  const dt = new Date(y, mo, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) return null;
  return dt;
}

function formatDateYmd_(d, tz) {
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

/* =========================
   Sheet creation + update
========================= */

function getOrCreateMonthSheet_(dateObj, tz) {
  const ss = SpreadsheetApp.openById(getProp_("SPREADSHEET_ID"));
  const zone = tz || getProp_("TIMEZONE", false) || Session.getScriptTimeZone();
  const targetName = Utilities.formatDate(dateObj, zone, "MMMM yyyy");

  let sh = ss.getSheetByName(targetName);
  if (sh) return sh;

  const template = getLatestMonthTemplateSheet_(ss, targetName);
  if (template) {
    sh = template.copyTo(ss);
    sh.setName(targetName);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(ss.getNumSheets());
    clearMonthData_(sh);
    formatMonthSheetHeader_(sh, dateObj, zone);
    return sh;
  }

  sh = ss.insertSheet(targetName);
  initializeNewMonthSheet_(sh, dateObj, zone);
  return sh;
}

function updateTodayAttendance_(slackUserId, attendanceValue, tz) {
  return updateAttendanceForDate_(slackUserId, attendanceValue, new Date(), tz);
}

function updateAttendanceForDate_(slackUserId, attendanceValue, dateObj, tz) {
  try {
    const sh = getOrCreateMonthSheet_(dateObj, tz);

    const row = findOrCreateUserRow_(sh, slackUserId, dateObj, tz);
    if (!row) return { ok: false, error: `Could not resolve row for Slack user ${slackUserId}` };

    const dayNum = Number(Utilities.formatDate(dateObj, tz, "d"));
    const col = findDayColumn_(sh, dayNum);
    if (!col) return { ok: false, error: `Day column not found for day ${dayNum}` };

    sh.getRange(row, col).setValue(attendanceValue);
    sh.getRange(row, CFG.LAST_UPDATED_COL).setValue(new Date());
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function getStatusForDate_(slackUserId, dateObj, tz) {
  const sh = getOrCreateMonthSheet_(dateObj, tz);
  const row = findRowBySlackId_(sh, slackUserId);
  if (!row) return "";
  const dayNum = Number(Utilities.formatDate(dateObj, tz, "d"));
  const col = findDayColumn_(sh, dayNum);
  if (!col) return "";
  return String(sh.getRange(row, col).getValue() || "").trim();
}

/* =========================
   Month sheet helpers
========================= */

function getLatestMonthTemplateSheet_(ss, excludeName) {
  const sheets = ss.getSheets();
  const monthMap = {
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
    july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  let best = null;
  let bestTime = -1;

  sheets.forEach(s => {
    const name = s.getName();
    if (name === excludeName) return;
    const m = name.match(/^([A-Za-z]+)\s+(\d{4})$/);
    if (!m) return;

    const monthName = m[1].toLowerCase();
    const year = Number(m[2]);
    if (!(monthName in monthMap)) return;

    const t = new Date(year, monthMap[monthName], 1).getTime();
    if (t > bestTime) {
      bestTime = t;
      best = s;
    }
  });

  return best;
}

function clearMonthData_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.DATA_START_ROW) return;

  sh.getRange(CFG.DATA_START_ROW, CFG.LAST_UPDATED_COL, lastRow - CFG.DATA_START_ROW + 1, 1).clearContent();

  const dayCols = getDayColumns_(sh);
  dayCols.forEach(col => {
    sh.getRange(CFG.DATA_START_ROW, col, lastRow - CFG.DATA_START_ROW + 1, 1).clearContent();
  });
}

function refreshCurrentMonthSheetFormat() {
  const tz = getProp_("TIMEZONE", false) || Session.getScriptTimeZone();
  const today = new Date();
  const sh = getOrCreateMonthSheet_(today, tz);
  rewriteDayHeaders_(sh, today, tz);
  formatMonthSheetHeader_(sh, today, tz);
}

function initializeNewMonthSheet_(sh, dateObj, tz) {
  sh.getRange(1, 1).setValue("Employee Name");
  sh.getRange(1, 2).setValue("SlackUserID");
  sh.getRange(1, 3).setValue("Email");
  sh.getRange(1, 4).setValue("Last Updated at");
  sh.getRange(1, 5).setValue("Client");

  rewriteDayHeaders_(sh, dateObj, tz);

  formatMonthSheetHeader_(sh, dateObj, tz);
  sh.setFrozenRows(1);
}

function rewriteDayHeaders_(sh, dateObj, tz) {
  const year = Number(Utilities.formatDate(dateObj, tz, "yyyy"));
  const month = Number(Utilities.formatDate(dateObj, tz, "M")) - 1;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const headerValues = [Array.from({ length: daysInMonth }, (_, index) => {
    return getDayHeaderLabel_(new Date(year, month, index + 1));
  })];

  sh.getRange(CFG.HEADER_ROW, CFG.FIRST_DAY_COL, 1, daysInMonth).setValues(headerValues);
}

function formatMonthSheetHeader_(sh, dateObj, tz) {
  const year = Number(Utilities.formatDate(dateObj, tz, "yyyy"));
  const month = Number(Utilities.formatDate(dateObj, tz, "M")) - 1;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const lastHeaderCol = CFG.FIRST_DAY_COL + daysInMonth - 1;
  const headerRange = sh.getRange(CFG.HEADER_ROW, 1, 1, lastHeaderCol);

  headerRange
    .setBackground("#d9eaf7")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sh.getRange(CFG.HEADER_ROW, 1, 1, CFG.FIRST_DAY_COL - 1).setWrap(false);
  sh.getRange(CFG.HEADER_ROW, CFG.FIRST_DAY_COL, 1, daysInMonth).setWrap(true);
  sh.setRowHeight(CFG.HEADER_ROW, 36);

  getWeekendDayColumns_(dateObj, tz).forEach(col => {
    sh.getRange(CFG.HEADER_ROW, col).setBackground("#f4cccc");
  });

  formatWeekendColumnsForUsedRows_(sh, dateObj, tz);
}

function getWeekendDayColumns_(dateObj, tz) {
  const year = Number(Utilities.formatDate(dateObj, tz, "yyyy"));
  const month = Number(Utilities.formatDate(dateObj, tz, "M")) - 1;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cols = [];

  for (let d = 1; d <= daysInMonth; d++) {
    const weekday = new Date(year, month, d).getDay();
    if (weekday === 0 || weekday === 6) cols.push(CFG.FIRST_DAY_COL + d - 1);
  }
  return cols;
}

function formatWeekendColumnsForUsedRows_(sh, dateObj, tz) {
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.DATA_START_ROW) return;

  const year = Number(Utilities.formatDate(dateObj, tz, "yyyy"));
  const month = Number(Utilities.formatDate(dateObj, tz, "M")) - 1;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const rowCount = lastRow - CFG.DATA_START_ROW + 1;

  sh.getRange(CFG.DATA_START_ROW, CFG.FIRST_DAY_COL, rowCount, daysInMonth).setBackground("#ffffff");
  getWeekendDayColumns_(dateObj, tz).forEach(col => {
    sh.getRange(CFG.DATA_START_ROW, col, rowCount, 1).setBackground("#f4cccc");
  });
}

function formatWeekendColumnsForRow_(sh, row, dateObj, tz) {
  getWeekendDayColumns_(dateObj, tz).forEach(col => {
    sh.getRange(row, col).setBackground("#f4cccc");
  });
}

function getDayHeaderLabel_(dateObj) {
  const dayNum = dateObj.getDate();
  const weekdayMap = ["S", "M", "T", "W", "Th", "F", "S"];
  return `${dayNum} ${weekdayMap[dateObj.getDay()]}`;
}

function extractDayNumber_(headerValue) {
  const s = String(headerValue || "").trim();
  if (!s) return NaN;

  const direct = Number(s);
  if (!isNaN(direct)) return direct;

  const match = s.match(/^(\d{1,2})\b/);
  return match ? Number(match[1]) : NaN;
}

/* =========================
   Lookup helpers
========================= */

function findOrCreateUserRow_(sh, slackUserId, dateObj, tz) {
  const existing = findRowBySlackId_(sh, slackUserId);
  if (existing) return existing;

  const nextRow = Math.max(sh.getLastRow() + 1, CFG.DATA_START_ROW);
  sh.getRange(nextRow, CFG.NAME_COL).setValue(getUserDisplayName_(slackUserId));
  sh.getRange(nextRow, CFG.SLACK_ID_COL).setValue(slackUserId);
  sh.getRange(nextRow, CFG.EMAIL_COL).setValue("");
  formatWeekendColumnsForRow_(sh, nextRow, dateObj, tz);
  return nextRow;
}

function getUserDisplayName_(slackUserId) {
  const raw = getProp_("USER_MAP_JSON", false);
  if (!raw) return slackUserId;

  try {
    const userMap = JSON.parse(raw);
    const mappedName = userMap && typeof userMap === "object" ? userMap[slackUserId] : "";
    return String(mappedName || "").trim() || slackUserId;
  } catch (err) {
    console.error(`Invalid USER_MAP_JSON: ${err}`);
    return slackUserId;
  }
}

function findRowBySlackId_(sh, slackUserId) {
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.DATA_START_ROW) return 0;

  const numRows = lastRow - CFG.DATA_START_ROW + 1;
  const values = sh.getRange(CFG.DATA_START_ROW, CFG.SLACK_ID_COL, numRows, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const v = String(values[i][0] || "").trim();
    if (v === slackUserId) return CFG.DATA_START_ROW + i;
  }
  return 0;
}

function findDayColumn_(sh, dayNum) {
  const dayCols = getDayColumns_(sh);
  for (let i = 0; i < dayCols.length; i++) {
    const col = dayCols[i];
    const v = sh.getRange(CFG.HEADER_ROW, col).getValue();
    if (extractDayNumber_(v) === dayNum) return col;
  }
  return 0;
}

function getDayColumns_(sh) {
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const cols = [];

  for (let c = 1; c <= headers.length; c++) {
    const dayNum = extractDayNumber_(headers[c - 1]);
    if (!isNaN(dayNum) && dayNum >= 1 && dayNum <= 31) cols.push(c);
  }
  return cols;
}

/* =========================
   Slack message helpers
========================= */

function sendAttendanceButtonsDM_(userId) {
  const channelId = openIm_(userId);
  if (!channelId) throw new Error("Could not open DM channel");

  const text = "Please update your attendance for today.";
  const blocks = [
    { type: "section", text: { type: "mrkdwn", text } },
    {
      type: "actions",
      elements: [
        { type: "button", text: { type: "plain_text", text: "WFO" }, style: "primary", action_id: "wfo" },
        { type: "button", text: { type: "plain_text", text: "WFH" }, action_id: "wfh" },
        { type: "button", text: { type: "plain_text", text: "Leave (-1)" }, style: "danger", action_id: "leave_full" },
        { type: "button", text: { type: "plain_text", text: "Half Day (-0.5)" }, action_id: "leave_half" }
      ]
    }
  ];

  slackApi_("chat.postMessage", { channel: channelId, text, blocks });
}

function sendTextDM_(channelId, text) {
  slackApi_("chat.postMessage", { channel: channelId, text });
}

function openIm_(userId) {
  const res = slackApi_("conversations.open", { users: userId });
  return res && res.channel ? res.channel.id : null;
}

function slackApi_(method, body) {
  const token = getProp_("SLACK_BOT_TOKEN");
  const resp = UrlFetchApp.fetch(`https://slack.com/api/${method}`, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const json = JSON.parse(resp.getContentText() || "{}");
  if (!json.ok) throw new Error(`Slack ${method} failed: ${json.error || "unknown"}`);
  return json;
}

/* =========================
   Utility
========================= */

function mapActionToValue_(actionId) {
  switch (actionId) {
    case "wfo": return "WFO";
    case "wfh": return "WFH";
    case "leave_full": return "-1";
    case "leave_half": return "-0.5";
    default: return null;
  }
}

function isDuplicate_(key, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const exists = cache.get(key);
  if (exists) return true;
  cache.put(key, "1", Math.min(ttlSeconds, 21600)); // max 6h in cache
  return false;
}


function getProp_(key, required = true) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (required && (!v || !String(v).trim())) {
    throw new Error(`Missing script property: ${key}`);
  }
  return v;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   Reminder scheduler helpers
========================= */

function getReminderUserIds_() {
  return (getProp_("SLACK_USER_IDS") || "")
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);
}

function hasTodayAttendance_(sh, slackUserId, tz) {
  const row = findRowBySlackId_(sh, slackUserId);
  if (!row) return false;

  const dayNum = Number(Utilities.formatDate(new Date(), tz, "d"));
  const dayCol = findDayColumn_(sh, dayNum);
  if (!dayCol) return false;

  const val = String(sh.getRange(row, dayCol).getValue() || "").trim();
  return val !== "";
}

function sendMorningReminders() {
  const userIds = getReminderUserIds_();
  userIds.forEach(userId => {
    try {
      sendAttendanceButtonsDM_(userId);
      console.log(`Morning reminder sent: ${userId}`);
    } catch (err) {
      console.error(`Morning reminder failed for ${userId}: ${err}`);
    }
  });
}

function sendEveningReminders() {
  const tz = getProp_("TIMEZONE", false) || Session.getScriptTimeZone();
  const sh = getOrCreateMonthSheet_(new Date(), tz);
  const userIds = getReminderUserIds_();

  userIds.forEach(userId => {
    try {
      const updated = hasTodayAttendance_(sh, userId, tz);
      if (!updated) {
        sendAttendanceButtonsDM_(userId);
        console.log(`Evening reminder sent: ${userId}`);
      } else {
        console.log(`Evening skipped (already updated): ${userId}`);
      }
    } catch (err) {
      console.error(`Evening reminder failed for ${userId}: ${err}`);
    }
  });
}





