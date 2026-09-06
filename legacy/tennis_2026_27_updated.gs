/***** CONFIG *****/

// Always use the spreadsheet's timezone for date handling
const SPREADSHEET_TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

const DOC_URL_TWO_PANE    = 'https://docs.google.com/spreadsheets/d/12oZlgXc5SZuxytPnUX6jeXjGan0UvHM0SqjZ0dgAfao/edit';
const LOCATION_TWO        = 'Longfellow Club';
const COURT_NUMBER        = '1';   // fixed court number
const CC_ADDRESS          = 'bahec700@gmail.com';
const EMAIL_FROM_NAME_TWO = '';    // e.g., 'Ivan' if you want a display name

// Different start times by sheet
const START_TIME_BY_SHEET = {
  Saturday: '9:30 AM',
  Tuesday:  '7:30 PM'
};

/**
 * If a header uses a nickname or just first+initial and you want to force a mapping,
 * put it here:  "Header Text" : "Exact Full Name from Directory"
 * Example given: Jim B -> Jim Bryant
 */
const CUSTOM_NAME_ALIASES = {
  'Jim B': 'Jim Bryant',
  // add more like:
  // 'Alex K': 'Alex Kuznetsov',
};

/***** SHEET SPECS *****/
const TWO_PANE_SPECS = {
  Saturday: {
    headerRow: 14,
    directory: { namesRange: 'A3:A11', emailsRange: 'H3:H11' },
    parts: [
      // 8 scheduled-player columns: B:I
      { firstRow: 15, lastRow: 33, dateCol: 'A', playersStartCol: 'B', playersEndCol: 'I' },
      // Right pane: date in L, players M:T
      { firstRow: 15, lastRow: 32, dateCol: 'L', playersStartCol: 'M', playersEndCol: 'T' }
    ]
  },
  Tuesday: {
    headerRow: 14,
    directory: { namesRange: 'A3:A10', emailsRange: 'H3:H10' },
    parts: [
      // 7 scheduled-player columns: B:H
      { firstRow: 15, lastRow: 33, dateCol: 'A', playersStartCol: 'B', playersEndCol: 'H' },
      // Right pane: date in K, players L:R
      { firstRow: 15, lastRow: 32, dateCol: 'K', playersStartCol: 'L', playersEndCol: 'R' }
    ]
  }
};

/***** OPTIONAL QUICK RUNNER *****/
function CreateReminderSaturday() {
  createReminderDraftForNextDate_All_TwoPane('Saturday');
}

function CreateReminderTuesday() {
  createReminderDraftForNextDate_All_TwoPane('Tuesday');
}

/***** ENTRY POINT *****/
/**
 * Run with 'Saturday' or 'Tuesday' to draft for that sheet,
 * or with no argument to pick the earliest of the two.
 */
function createReminderDraftForNextDate_All_TwoPane(whichDay) {
  if (!whichDay) {
    const candSat = findNextTwoPane_('Saturday');
    const candTue = findNextTwoPane_('Tuesday');
    const pick = pickEarlier_(candSat, candTue);
    if (!pick) { console.log('No upcoming matches.'); return; }
    buildAndCreateDraft_(pick);
    return;
  }
  const cand = findNextTwoPane_(whichDay);
  if (!cand) { console.log(`No upcoming matches on ${whichDay}.`); return; }
  buildAndCreateDraft_(cand);
}

/***** CORE LOGIC (dates read directly) *****/
function findNextTwoPane_(sheetName) {
  const spec = TWO_PANE_SPECS[sheetName];
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!spec || !sh) throw new Error(`Spec missing for "${sheetName}"`);

  const dir = readDirectoryFromSheet_(sh, spec.directory.namesRange, spec.directory.emailsRange);
  const candidates = [];

  for (const part of spec.parts) {
    const headers = sh.getRange(
      spec.headerRow,
      colA1ToNum_(part.playersStartCol),
      1,
      colA1ToNum_(part.playersEndCol) - colA1ToNum_(part.playersStartCol) + 1
    ).getValues()[0].map(s => String(s || '').trim());

    const numRows = part.lastRow - part.firstRow + 1;
    const dateVals = sh.getRange(part.firstRow, colA1ToNum_(part.dateCol), numRows, 1).getValues();
    const playerVals = sh.getRange(
      part.firstRow,
      colA1ToNum_(part.playersStartCol),
      numRows,
      headers.length
    ).getValues();

    for (let i = 0; i < numRows; i++) {
      const d = coerceDate_(dateVals[i][0]);
      if (!d) continue;

      const rowVals = playerVals[i];
      const scheduled = [];
      const subs = [];

      // Cell semantics: Play → scheduled; Ball → scheduled + balls; x → not available; '' → available (sub)
      for (let c = 0; c < rowVals.length; c++) {
        const cell = String(rowVals[c] || '').trim();
        const headerLabel = headers[c] || `Player_${c}`;

        const mapped = mapHeaderToDirectory_(headerLabel, dir);
        const fullName = mapped.fullName || headerLabel;
        const email    = mapped.email || '';

        if (/^Ball$/i.test(cell)) {
          scheduled.push({ lastName: extractLastName_(fullName), fullName, email, bringBalls: true });
        } else if (/^Play$/i.test(cell)) {
          scheduled.push({ lastName: extractLastName_(fullName), fullName, email, bringBalls: false });
        } else if (/^x$/i.test(cell)) {
          // unavailable → ignore
        } else if (cell === '') {
          // available, not scheduled → candidate sub
          subs.push({ lastName: extractLastName_(fullName), fullName });
        }
      }

      if (scheduled.length > 0) {
        const start = buildStartTimeOnly_(d, START_TIME_BY_SHEET[sheetName]);
        candidates.push({ sheetName, dateObj: start, scheduled, subs });
      }
    }
  }

  const today = new Date(); today.setHours(0,0,0,0);
  const upcoming = candidates
    .filter(c => c.dateObj >= today)
    .sort((a,b) => a.dateObj - b.dateObj);
  return upcoming[0] || null;
}

function buildAndCreateDraft_(cand) {
  // Sort roster: balls first, then by last name
  cand.scheduled.sort((a,b) => {
    if (a.bringBalls && !b.bringBalls) return -1;
    if (!a.bringBalls && b.bringBalls) return 1;
    return a.lastName.localeCompare(b.lastName);
  });
  cand.subs.sort((a,b) => a.lastName.localeCompare(b.lastName));

  // Recipients
  const toList = Array.from(new Set(
    cand.scheduled.map(p => (p.email || '').trim().toLowerCase()).filter(Boolean)
  ));
  if (!toList.length) { console.log('No emails for scheduled players.'); return; }

  // Subject: short weekday, no year/time
  const subjectDateShort = Utilities.formatDate(cand.dateObj, Session.getScriptTimeZone(), 'EEE MMM d');
  const subject = `Tennis reminder – ${subjectDateShort} @ ${LOCATION_TWO}`;

  // Body
  const weekdayLong = Utilities.formatDate(cand.dateObj, Session.getScriptTimeZone(), 'EEEE, MMMM d');
  const timeLabel = START_TIME_BY_SHEET[cand.sheetName];

  const text = buildBodiesText_(weekdayLong, cand.scheduled, cand.subs, timeLabel);
  const html = buildBodiesHtml_(weekdayLong, cand.scheduled, cand.subs, timeLabel);

  const opts = EMAIL_FROM_NAME_TWO
    ? { name: EMAIL_FROM_NAME_TWO, htmlBody: html, cc: CC_ADDRESS }
    : { htmlBody: html, cc: CC_ADDRESS };

  GmailApp.createDraft(toList.join(', '), subject, text, opts);
  console.log(`Draft created → [${cand.sheetName}] ${subject}`);
}

/***** BODY BUILDERS *****/
function buildBodiesText_(weekdayLong, scheduled, subs, timeLabel) {
  const rosterText = scheduled.map(p => `- ${p.fullName}${p.bringBalls ? ' (balls)' : ''}`).join('\n');
  const subsText = subs.length ? '\n\nAvailable subs\n' + subs.map(s => `- ${s.fullName}`).join('\n') : '';
  return `Hello players,

This is a reminder that you are scheduled to play on ${weekdayLong} at ${timeLabel} at the Longfellow Club in Wayland on court ${COURT_NUMBER}.

Roster
${rosterText}${subsText}

Have a great match!

Full schedule: ${DOC_URL_TWO_PANE}`;
}

function buildBodiesHtml_(weekdayLong, scheduled, subs, timeLabel) {
  const rosterHtml = scheduled.map(p => `<li>${htmlEscape(p.fullName)}${p.bringBalls ? ' (balls)' : ''}</li>`).join('');
  const subsHtml = subs.length
    ? `<p><strong>Available Subs</strong></p><ul>${subs.map(s => `<li>${htmlEscape(s.fullName)}</li>`).join('')}</ul>`
    : '';
  return `<div>
  <p>Hello players,</p>
  <p>This is a reminder that you are scheduled to play on ${htmlEscape(weekdayLong)} at ${htmlEscape(timeLabel)} at the Longfellow Club in Wayland on court ${htmlEscape(COURT_NUMBER)}.</p>
  <p><strong>Main Roster</strong></p>
  <ul>${rosterHtml}</ul>
  ${subsHtml}
  <p>Have a great match!</p>
  <p><a href="${DOC_URL_TWO_PANE}">Click here for the full schedule</a></p>
</div>`;
}

/***** NAME/DIRECTORY MATCHING *****/
/**
 * Build lookup maps from Directory full names.
 * byLast        : LASTNAME → { fullName, email } (last assignment wins)
 * byFullLower   : "firstname lastname" (lower) → record
 * byFirst       : FIRSTNAME → [records...]  (for “Jim” only)
 * byFirstInit   : "firstname l" → record     (for “Jim B”)
 */
function readDirectoryFromSheet_(sh, namesA1, emailsA1) {
  const names = sh.getRange(namesA1).getValues().flat().map(v => String(v || '').trim());
  const emails = sh.getRange(emailsA1).getValues().flat().map(v => String(v || '').trim());

  const byLast = {};
  const byFullLower = {};
  const byFirst = {};
  const byFirstInit = {};

  for (let i = 0; i < Math.max(names.length, emails.length); i++) {
    const fullName = names[i] || '';
    const email = emails[i] || '';
    if (!fullName) continue;

    const parts = splitName_(fullName);
    const last = parts.last.toUpperCase();
    const first = parts.first.toLowerCase();
    const lastInit = parts.last ? parts.last[0].toLowerCase() : '';

    const rec = { fullName, email };

    byLast[last] = rec; // simple map; fine if unique or last-one-wins

    byFullLower[`${first} ${parts.last.toLowerCase()}`] = rec;

    if (!byFirst[first]) byFirst[first] = [];
    byFirst[first].push(rec);

    if (first && lastInit) {
      byFirstInit[`${first} ${lastInit}`] = rec;
    }
  }

  return { byLast, byFullLower, byFirst, byFirstInit };
}

/**
 * Map a header label like "Jim" or "Jim B" to a directory record.
 * Order:
 * 1) Custom alias → exact full name
 * 2) Exact full-name match
 * 3) First+initial match (Jim B)
 * 4) Unique first-name match (Jim with only one Jim)
 */
function mapHeaderToDirectory_(headerLabel, dir) {
  const raw = String(headerLabel || '').trim();
  if (!raw) return { fullName: '', email: '' };

  // 1) Custom alias
  const aliasFull = CUSTOM_NAME_ALIASES[raw];
  if (aliasFull) {
    const parts = splitName_(aliasFull);
    const rec = dir.byFullLower[`${parts.first.toLowerCase()} ${parts.last.toLowerCase()}`];
    if (rec) return rec;
    // If alias not found in directory, fall through to attempt heuristics
  }

  // 2) Exact full name (case-insensitive)
  const partsRaw = splitName_(raw);
  const exactKey = `${partsRaw.first.toLowerCase()} ${partsRaw.last.toLowerCase()}`;
  if (dir.byFullLower[exactKey]) return dir.byFullLower[exactKey];

  // Check if header is "First L" pattern
  const m = raw.match(/^([A-Za-z]+)\s+([A-Za-z])$/);
  if (m) {
    const first = m[1].toLowerCase();
    const lin   = m[2].toLowerCase();
    const keyFL = `${first} ${lin}`;
    if (dir.byFirstInit[keyFL]) return dir.byFirstInit[keyFL];
  }

  // 3) Only first name
  const firstOnly = partsRaw.first.toLowerCase();
  const candidates = dir.byFirst[firstOnly] || [];
  if (candidates.length === 1) return candidates[0];
  if (candidates.length > 1) {
    console.log(`Ambiguous header "${raw}" → multiple matches for first name "${partsRaw.first}". Add CUSTOM_NAME_ALIASES to disambiguate.`);
  } else {
    console.log(`No directory match for header "${raw}".`);
  }

  // Fallback: return empty to keep name as header but skip email
  return { fullName: '', email: '' };
}

function splitName_(fullName) {
  const s = String(fullName || '').trim();
  if (!s) return { first: '', last: '' };
  const parts = s.split(/\s+/);
  if (parts.length === 1) return { first: parts[0], last: '' };
  return { first: parts[0], last: parts[parts.length - 1] };
}

/***** HELPERS *****/
function extractLastName_(fullName) {
  const parts = String(fullName || '').trim().split(/\s+/);
  return parts.length ? parts[parts.length - 1] : '';
}

function coerceDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function withDay_(d) { const x = new Date(d); x.setHours(0,0,0,0); return x; }

function buildStartTimeOnly_(dateOnly, hhmmStr) {
  const d = new Date(dateOnly);
  const m = String(hhmmStr).trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (!m) { d.setHours(19,0,0,0); return d; }
  let h = parseInt(m[1], 10);
  const min = m[2] ? parseInt(m[2], 10) : 0;
  const ap = (m[3]||'').toUpperCase();
  if (ap === 'PM' && h < 12) h += 12;
  if (ap === 'AM' && h === 12) h = 0;
  d.setHours(h, min, 0, 0);
  return d;
}

function colA1ToNum_(a1) {
  let n = 0;
  for (let i = 0; i < a1.length; i++) n = n * 26 + (a1.charCodeAt(i) - 64);
  return n;
}

function htmlEscape(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function pickEarlier_(a, b) {
  if (a && b) return a.dateObj < b.dateObj ? a : b;
  return a || b;
}


/** Calendar.gs
 * Creates/updates Google Calendar events for Ivan on Tuesday/Saturday.
 * Supports "next event" and "ALL upcoming events" modes.
 * DEPENDS ON (from your main file):
 *   - TWO_PANE_SPECS, START_TIME_BY_SHEET, SPREADSHEET_TZ
 *   - LOCATION_TWO, COURT_NUMBER
 *   - readDirectoryFromSheet_(), mapHeaderToDirectory_()
 *   - buildStartTimeOnly_(), withDay_(), colA1ToNum_()
 */

/***** CONFIG (Ivan-specific) *****/

// Which calendar to write to (primary = default calendar)
const CALENDAR_ID_IVAN = 'primary';

// Ivan's full name exactly as it appears in the Directory (A3:A..)
const IVAN_FULL_NAME   = 'Ivan Loukovnikov';

// Event duration (minutes)
const DURATION_MIN_IVAN = 90;

// Optional: small delay (ms) between Calendar writes to be gentle on quotas
const OP_DELAY_MS_IVAN = 300;

// Event title
function titleForIvan_(bringBalls) {
  return bringBalls ? 'Tennis (balls)' : 'Tennis';
}

/***** PUBLIC ENTRY POINTS: NEXT EVENT (kept for convenience) *****/

function createIvanNextEvent_Saturday() {
  createIvanNextEventForSheet_('Saturday');
}

function createIvanNextEvent_Tuesday() {
  createIvanNextEventForSheet_('Tuesday');
}

function createIvanNextEventsBothDays() {
  const sat = findNextForIvan_('Saturday');
  const tue = findNextForIvan_('Tuesday');

  if (!sat && !tue) {
    console.log('No upcoming scheduled matches for Ivan on Tuesday or Saturday.');
    return;
  }
  const cal = CalendarApp.getCalendarById(CALENDAR_ID_IVAN);
  if (!cal) throw new Error(`Calendar "${CALENDAR_ID_IVAN}" not found.`);

  if (sat) { upsertIvanEvent_(cal, sat); if (OP_DELAY_MS_IVAN) Utilities.sleep(OP_DELAY_MS_IVAN); }
  if (tue) { upsertIvanEvent_(cal, tue); if (OP_DELAY_MS_IVAN) Utilities.sleep(OP_DELAY_MS_IVAN); }
  console.log('Done (next events).');
}

/***** PUBLIC ENTRY POINTS: ALL UPCOMING EVENTS *****/

/** Create/update ALL future Ivan events on BOTH sheets. */
function createIvanAllUpcomingEvents() {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID_IVAN);
  if (!cal) throw new Error(`Calendar "${CALENDAR_ID_IVAN}" not found.`);

  const allSat = enumerateAllForIvan_('Saturday');
  const allTue = enumerateAllForIvan_('Tuesday');

  const all = [...allSat, ...allTue].sort((a,b) => a.dateObj - b.dateObj);
  if (!all.length) {
    console.log('No upcoming scheduled matches for Ivan on either sheet.');
    return;
  }

  let created = 0, updated = 0;
  for (const cand of all) {
    const res = upsertIvanEvent_(cal, cand);
    if (res === 'created') created++;
    else if (res === 'updated') updated++;
    if (OP_DELAY_MS_IVAN) Utilities.sleep(OP_DELAY_MS_IVAN);
  }
  console.log(`Done (all upcoming): created ${created}, updated ${updated}, total ${all.length}.`);
}

/** Create/update ALL future Ivan events for a single sheet. */
function createIvanAllUpcomingEventsForSheet_(sheetName) {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID_IVAN);
  if (!cal) throw new Error(`Calendar "${CALENDAR_ID_IVAN}" not found.`);

  const items = enumerateAllForIvan_(sheetName);
  if (!items.length) {
    console.log(`No upcoming scheduled matches for Ivan on ${sheetName}.`);
    return;
  }

  let created = 0, updated = 0;
  for (const cand of items) {
    const res = upsertIvanEvent_(cal, cand);
    if (res === 'created') created++;
    else if (res === 'updated') updated++;
    if (OP_DELAY_MS_IVAN) Utilities.sleep(OP_DELAY_MS_IVAN);
  }
  console.log(`Done (${sheetName} all upcoming): created ${created}, updated ${updated}, total ${items.length}.`);
}

/***** CORE (per-sheet) *****/

function createIvanNextEventForSheet_(sheetName) {
  const cand = findNextForIvan_(sheetName);
  if (!cand) {
    console.log(`No upcoming scheduled matches for Ivan on ${sheetName}.`);
    return;
  }
  const cal = CalendarApp.getCalendarById(CALENDAR_ID_IVAN);
  if (!cal) throw new Error(`Calendar "${CALENDAR_ID_IVAN}" not found.`);
  upsertIvanEvent_(cal, cand);
}

/**
 * Find the next upcoming row on a given sheet where Ivan is scheduled (Play/Ball).
 * Returns:
 *   { sheetName, dateObj, bringBalls, scheduledNames[] }
 */
function findNextForIvan_(sheetName) {
  const items = enumerateAllForIvan_(sheetName);
  return items[0] || null;
}

/**
 * Enumerate ALL upcoming rows on a sheet where Ivan is scheduled.
 * Returns array of:
 *   { sheetName, dateObj, bringBalls, scheduledNames[] }
 */
function enumerateAllForIvan_(sheetName) {
  if (typeof TWO_PANE_SPECS === 'undefined') throw new Error('TWO_PANE_SPECS not found (define in main file).');
  const spec = TWO_PANE_SPECS[sheetName];
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!spec || !sh) throw new Error(`Spec or sheet missing for "${sheetName}"`);

  const dir = readDirectoryFromSheet_(sh, spec.directory.namesRange, spec.directory.emailsRange);
  const out = [];

  for (const part of spec.parts) {
    // Column headers (player labels per column)
    const headers = sh.getRange(
      spec.headerRow,
      colA1ToNum_(part.playersStartCol),
      1,
      colA1ToNum_(part.playersEndCol) - colA1ToNum_(part.playersStartCol) + 1
    ).getValues()[0].map(s => String(s || '').trim());

    const numRows = part.lastRow - part.firstRow + 1;
    const dateVals = sh.getRange(part.firstRow, colA1ToNum_(part.dateCol), numRows, 1).getValues();
    const playerVals = sh.getRange(
      part.firstRow,
      colA1ToNum_(part.playersStartCol),
      numRows,
      headers.length
    ).getValues();

    const mappedCols = headers.map(h => mapHeaderToDirectory_(h, dir));
    const ivanColIdx = mappedCols.findIndex(m => (m.fullName || '').toLowerCase() === IVAN_FULL_NAME.toLowerCase());
    if (ivanColIdx === -1) continue; // Ivan not present in this pane

    for (let i = 0; i < numRows; i++) {
      const dateOnly = coerceDateSheetTZ_(dateVals[i][0]);
      if (!dateOnly) continue;

      const row = playerVals[i];
      const ivanCell = String(row[ivanColIdx] || '').trim();

      const ivanBringBalls = /^Ball$/i.test(ivanCell);
      const ivanScheduled  = ivanBringBalls || /^Play$/i.test(ivanCell);
      if (!ivanScheduled) continue;

      const scheduledNames = [];
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] || '').trim();
        if (/^(Play|Ball)$/i.test(cell)) {
          const fn = mappedCols[c].fullName || headers[c] || `Player_${c}`;
          scheduledNames.push(fn);
        }
      }

      const start = buildStartTimeOnly_(dateOnly, START_TIME_BY_SHEET[sheetName]);
      out.push({
        sheetName,
        dateObj: start,
        bringBalls: ivanBringBalls,
        scheduledNames
      });
    }
  }

  const today = new Date(); today.setHours(0,0,0,0);
  return out
    .filter(c => withDay_(c.dateObj) >= today)
    .sort((a, b) => a.dateObj - b.dateObj);
}

/***** CALENDAR UPSERT *****/

/**
 * Upsert an event using a deterministic signature embedded in the description:
 *   Ref: [Ivan:<SheetName>:<YYYY-MM-DD>]
 * Returns 'updated' or 'created'.
 */
function upsertIvanEvent_(cal, cand) {
  const title = titleForIvan_(cand.bringBalls);
  const start = cand.dateObj;
  const end   = new Date(start.getTime() + DURATION_MIN_IVAN * 60 * 1000);

  const ymd     = Utilities.formatDate(start, SPREADSHEET_TZ || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(start, SPREADSHEET_TZ || Session.getScriptTimeZone(), 'h:mm a');

  const signature = `Ref: [Ivan:${cand.sheetName}:${ymd}]`;

  const notes = [
    `Date: ${ymd} ${timeStr}`,
    `Location: ${LOCATION_TWO}`,
    `Court: ${COURT_NUMBER}`,
    `Players: ${cand.scheduledNames.join(', ')}`,
    signature
  ].join('\n');

  // Search for existing event by signature over the whole day
  const dayStart = new Date(start); dayStart.setHours(0,0,0,0);
  const dayEnd   = new Date(start); dayEnd.setHours(23,59,59,999);

  let matches = cal.getEvents(dayStart, dayEnd, { search: signature }) || [];

  // If multiple, pick the closest to desired start
  let match = null;
  if (matches.length === 1) {
    match = matches[0];
  } else if (matches.length > 1) {
    matches.sort((a,b) =>
      Math.abs(a.getStartTime().getTime() - start.getTime()) -
      Math.abs(b.getStartTime().getTime() - start.getTime())
    );
    match = matches[0];
  }

  if (match) {
    match.setTitle(title);
    match.setTime(start, end);
    if (typeof LOCATION_TWO !== 'undefined' && LOCATION_TWO) match.setLocation(LOCATION_TWO);
    match.setDescription(notes);
    console.log(`Updated Ivan's event: "${title}" on ${ymd} (${cand.sheetName})`);
    return 'updated';
  }

  // Otherwise create new
  const ev = cal.createEvent(title, start, end, {
    description: notes,
    location: (typeof LOCATION_TWO !== 'undefined' && LOCATION_TWO) ? LOCATION_TWO : undefined
  });
  console.log(`Created Ivan's event: "${title}" on ${ymd} (${cand.sheetName}) → ${ev.getId()}`);
  return 'created';
}

/***** LOCAL HELPER (date coercion) *****/

/**
 * Normalize any sheet date to a true date-only in the spreadsheet's timezone.
 * Keep here for safety in case it's not present in the main file.
 */
function coerceDateSheetTZ_(v) {
  if (!v) return null;

  let d;
  if (v instanceof Date && !isNaN(v)) {
    d = v;
  } else {
    d = new Date(String(v).trim());
  }
  if (isNaN(d)) return null;

  const tz = (typeof SPREADSHEET_TZ !== 'undefined')
    ? SPREADSHEET_TZ
    : SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  const ymd = Utilities.formatDate(d, tz, 'yyyy-MM-dd').split('-');
  const y = +ymd[0], m = +ymd[1] - 1, day = +ymd[2];
  return new Date(y, m, day, 0, 0, 0, 0);
}
