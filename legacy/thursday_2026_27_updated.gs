/***** CONFIG *****/
// Add this near the other constants at the top of your script
const DOC_URL = 'https://docs.google.com/spreadsheets/d/1SpJCM26QPpT5NU8rNRPm33TA4mN_LRQYnDP2BJ4lI2o/edit';

const SHEET_NAME      = 'Sheet1';   // your tab name
const CALENDAR_ID     = 'primary';  // default Google Calendar
const TIMEZONE        = Session.getScriptTimeZone();

const FIRST_DATA_ROW  = 8;
const LAST_DATA_ROW   = 44;
const DATE_COL        = 2;   // B
const FIRST_PLAYER_COL= 3;   // C
const LAST_PLAYER_COL = 10;  // J
const YOUR_MARK_COL   = 6;   // F
const EVENT_ID_COL    = 11;  // K

// Event defaults
const START_TIME_HHMM = '7 PM';
const DURATION_MIN    = 90;
const LOCATION        = 'Longfellow Club';
const REMINDER_MIN    = null;

// Delay (ms) between event operations
const OP_DELAY_MS = 1000;

/***** TITLES / NOTES *****/
function titleTemplate(bringBalls) {
  return bringBalls ? 'Tennis (balls)' : 'Tennis';
}
function descriptionTemplate(dateStr, allPlayers) {
  let parts = [
    `Date: ${dateStr}`,
    `Location: ${LOCATION}`,
    `Players: ${allPlayers.join(', ')}`
  ];
  return parts.join('\n');
}

/***** MAIN *****/
function syncMyMatchesInColumnF() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

  const numRows = LAST_DATA_ROW - FIRST_DATA_ROW + 1;

  // player names from row 7 (C..J)
  const playerNames = sh.getRange(7, FIRST_PLAYER_COL, 1, LAST_PLAYER_COL - FIRST_PLAYER_COL + 1)
                        .getValues()[0].map(v => String(v || '').trim());

  const dates = sh.getRange(FIRST_DATA_ROW, DATE_COL, numRows, 1).getValues();
  const marks = sh.getRange(FIRST_DATA_ROW, YOUR_MARK_COL, numRows, 1).getValues();
  const idRange = sh.getRange(FIRST_DATA_ROW, EVENT_ID_COL, numRows, 1);
  const ids = idRange.getValues();

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error(`Calendar "${CALENDAR_ID}" not found.`);

  for (let i = 0; i < numRows; i++) {
    const dateCell = dates[i][0];
    const myMark = String(marks[i][0] || '').trim().toUpperCase();
    const hasX = /X/.test(myMark);
    const bringBalls = /B/.test(myMark);
    const existingId = String(ids[i][0] || '').trim();

    const dateObj = coerceDate(dateCell);
    if (!dateObj) {
      if (existingId) {
        const ev = getEventSafe(cal, existingId);
        if (ev) {
          ev.deleteEvent();
          console.log(`Deleted event on row ${i + FIRST_DATA_ROW}`);
          Utilities.sleep(OP_DELAY_MS);
        }
        ids[i][0] = '';
      }
      continue;
    }

    // gather all players for this row
    const rowMarks = sh.getRange(FIRST_DATA_ROW + i, FIRST_PLAYER_COL, 1, LAST_PLAYER_COL - FIRST_PLAYER_COL + 1).getValues()[0];
    let scheduledPlayers = [];
    for (let c = 0; c < rowMarks.length; c++) {
      const mark = String(rowMarks[c] || '').trim().toUpperCase();
      if (/X/.test(mark)) {
        scheduledPlayers.push(playerNames[c] || `Player_${c + FIRST_PLAYER_COL}`);
      }
    }

    const { start, end } = buildStartEnd(dateObj, START_TIME_HHMM, DURATION_MIN);
    const dateStr = Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd');
    const title = titleTemplate(bringBalls);
    const notes = descriptionTemplate(dateStr, scheduledPlayers);

    if (hasX) {
      if (existingId) {
        const ev = getEventSafe(cal, existingId);
        if (ev) {
          ev.setTitle(title);
          ev.setTime(start, end);
          if (LOCATION) ev.setLocation(LOCATION);
          ev.setDescription(notes);
          console.log(`Updated event: "${title}" on ${dateStr}, Players: ${scheduledPlayers.join(', ')}`);
          Utilities.sleep(OP_DELAY_MS);
        } else {
          const newEv = createEvent(cal, bringBalls, start, end, dateStr, scheduledPlayers);
          ids[i][0] = newEv.getId();
          console.log(`Re-created event: "${title}" on ${dateStr}, Players: ${scheduledPlayers.join(', ')}`);
          Utilities.sleep(OP_DELAY_MS);
        }
      } else {
        const newEv = createEvent(cal, bringBalls, start, end, dateStr, scheduledPlayers);
        ids[i][0] = newEv.getId();
        console.log(`Created new event: "${title}" on ${dateStr}, Players: ${scheduledPlayers.join(', ')}`);
        Utilities.sleep(OP_DELAY_MS);
      }
    } else {
      if (existingId) {
        const ev = getEventSafe(cal, existingId);
        if (ev) {
          ev.deleteEvent();
          console.log(`Deleted event on ${dateStr}`);
          Utilities.sleep(OP_DELAY_MS);
        }
        ids[i][0] = '';
      }
    }
  }

  const headerCell = sh.getRange(7, EVENT_ID_COL);
  if (!String(headerCell.getValue()).trim()) headerCell.setValue('Event ID (F)');

  idRange.setValues(ids);
}

/***** HELPERS *****/
function coerceDate(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function buildStartEnd(dateOnly, hhmm, durationMin) {
  const start = new Date(dateOnly);
  const m = String(hhmm || '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?$/i);

  let hour = 19; // fallback: 7 PM
  let minute = 0;

  if (m) {
    hour = parseInt(m[1], 10);
    minute = m[2] ? parseInt(m[2], 10) : 0;
    const ap = (m[3] || '').toUpperCase();

    if (ap === 'PM' && hour < 12) hour += 12;
    if (ap === 'AM' && hour === 12) hour = 0;
  }

  start.setHours(hour, minute, 0, 0);
  const end = new Date(start.getTime() + (Number(durationMin) || 90) * 60 * 1000);
  return { start, end };
}
function getEventSafe(cal, id) {
  try { return cal.getEventById(id) || null; } catch (e) { return null; }
}
function createEvent(cal, bringBalls, start, end, dateStr, players) {
  const ev = cal.createEvent(
    titleTemplate(bringBalls),
    start,
    end,
    {
      location: LOCATION || undefined,
      description: descriptionTemplate(dateStr, players)
    }
  );
  return ev;
}

/***** OPTIONAL: auto-sync hourly *****/
function createHourlyTrigger() {
  ScriptApp.newTrigger('syncMyMatchesInColumnF')
    .timeBased().everyHours(1).create();
}

/***** ONE-DRAFT-TO-ALL REMINDER *****/

// Optional: display name for the draft (leave '' to use your default Gmail name)
const EMAIL_FROM_NAME = ''; // e.g., 'Ivan'

function createReminderDraftForNextDate_All() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

  const directory = buildPlayerDirectory_(sh);
  const next = findNextAnyMatchRow_(sh);
  if (!next) {
    console.log('No upcoming rows with any scheduled players.');
    return;
  }
  const { rowIndex, dateObj } = next;

  // Player headers (C..J)
  const playerHeaders = sh.getRange(7, FIRST_PLAYER_COL, 1, LAST_PLAYER_COL - FIRST_PLAYER_COL + 1)
                          .getValues()[0].map(v => String(v || '').trim());

  // Row values + backgrounds
  const rowMarks   = sh.getRange(rowIndex, FIRST_PLAYER_COL, 1, LAST_PLAYER_COL - FIRST_PLAYER_COL + 1).getValues()[0];
  const rowColors  = sh.getRange(rowIndex, FIRST_PLAYER_COL, 1, LAST_PLAYER_COL - FIRST_PLAYER_COL + 1).getBackgrounds()[0];

  // Collect players
  let scheduled = [];
  let subs = [];

  for (let c = 0; c < rowMarks.length; c++) {
    const mark = String(rowMarks[c] || '').trim().toUpperCase();
    const lastName = playerHeaders[c] || `Player_${c + FIRST_PLAYER_COL}`;
    const dirEntry = directory[lastName.toUpperCase()] || null;
    const fullName = dirEntry ? dirEntry.fullName : lastName;
    const email    = dirEntry ? dirEntry.email : '';
    const bringBalls = /B/.test(mark);

    if (/X/.test(mark)) {
      scheduled.push({ lastName, fullName, email, bringBalls });
    } else {
      const bg = String(rowColors[c] || '').toLowerCase();
      const isWhite = !bg || bg === '#ffffff' || bg === 'white' || bg === 'rgb(255,255,255)';
      if (!mark && isWhite) {
        subs.push({ lastName, fullName });
      }
    }
  }

  if (scheduled.length === 0) {
    console.log(`No players scheduled on row ${rowIndex}.`);
    return;
  }

  // Sort: balls first, then by last name
  scheduled.sort((a, b) => {
    if (a.bringBalls && !b.bringBalls) return -1;
    if (!a.bringBalls && b.bringBalls) return 1;
    return a.lastName.localeCompare(b.lastName);
  });

  const recipients = Array.from(new Set(
    scheduled.map(p => (p.email || '').trim().toLowerCase()).filter(Boolean)
  ));
  if (recipients.length === 0) {
    console.log('No emails found for scheduled players. Draft not created.');
    return;
  }

    // Format pieces
  const weekdayDateShort = Utilities.formatDate(dateObj, TIMEZONE, 'EEE MMM d'); 
  const weekdayDate = Utilities.formatDate(dateObj, TIMEZONE, 'EEEE, MMMM d'); 
  // Short day (Mon, Tue, Wed, etc.), no year, no time
  const subject = `Tennis reminder – ${weekdayDateShort} @ ${LOCATION}`;

  const rosterText = scheduled
    .map(p => `- ${p.fullName}${p.bringBalls ? ' (balls)' : ''}`)
    .join('\n');

  let subsText = '';
  if (subs.length > 0) {
    subs.sort((a, b) => a.lastName.localeCompare(b.lastName));
    subsText = '\n\nPotential subs\n' + subs.map(s => `- ${s.fullName}`).join('\n');
  }

  const textBody =
`Hello players,

This is a reminder that you are scheduled to play this ${weekdayDate}, at ${START_TIME_HHMM} at Longfellow Club in Wayland on court 5.

Roster
${rosterText}${subsText}

Have a great match!

Full schedule: ${DOC_URL}`;

  const htmlRoster = scheduled
    .map(p => `<li>${htmlEscape(p.fullName)}${p.bringBalls ? ' (balls)' : ''}</li>`)
    .join('');

  let htmlSubs = '';
  if (subs.length > 0) {
    subs.sort((a, b) => a.lastName.localeCompare(b.lastName));
    htmlSubs = `
      <p><strong>Available Subs</strong></p>
      <ul>
        ${subs.map(s => `<li>${htmlEscape(s.fullName)}</li>`).join('')}
      </ul>`;
  }

  const htmlBody = `
    <div>
      <p>Hello players,</p>
      <p>This is a reminder that you are scheduled to play on ${htmlEscape(weekdayDate)} at ${htmlEscape(START_TIME_HHMM)} at the Longfellow Club in Wayland on court 5.</p>
      <p><strong>Main Roster</strong></p>
      <ul>${htmlRoster}</ul>
      ${htmlSubs}
      <p>Have a great match!</p>
      <p><a href="${DOC_URL}">Click here for the full schedule</a></p>
    </div>
  `;

  const toField = recipients.join(', ');
  const draftOpts = EMAIL_FROM_NAME 
      ? { name: EMAIL_FROM_NAME, htmlBody, cc: 'bahec700@gmail.com' } 
      : { htmlBody, cc: 'bahec700@gmail.com' };

  GmailApp.createDraft(toField, subject, textBody, draftOpts);

  console.log(`Draft created → To: ${toField} | ${subject}`);
}

/***** small HTML escape helper *****/
function htmlEscape(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/***** HELPERS USED ABOVE (reuse from your script) *****/
// Uses the same helper you already added:
function buildPlayerDirectory_(sh) {
  const startRow = 51, count = 8;
  const names = sh.getRange(startRow, 1, count, 1).getValues();  // A51:A58
  const emails = sh.getRange(startRow, 6, count, 1).getValues(); // F51:F58 (merged)
  const dir = {};
  for (let i = 0; i < count; i++) {
    const fullName = String(names[i][0] || '').trim();
    const email = String(emails[i][0] || '').trim();
    if (!fullName) continue;
    const lastName = extractLastName_(fullName);
    if (!lastName) continue;
    dir[lastName.toUpperCase()] = { fullName, email };
  }
  return dir;
}
function extractLastName_(fullName) {
  const parts = fullName.trim().split(/\s+/);
  return parts.length ? parts[parts.length - 1] : '';
}
/** Find the next row (>= today) where ANY player C..J has X/BX. */
function findNextAnyMatchRow_(sh) {
  const numRows = LAST_DATA_ROW - FIRST_DATA_ROW + 1;
  const today = new Date(); today.setHours(0,0,0,0);

  const dateVals = sh.getRange(FIRST_DATA_ROW, DATE_COL, numRows, 1).getValues(); // B8..B44

  let bestRow = null, bestDate = null;
  for (let i = 0; i < numRows; i++) {
    const d = coerceDate(dateVals[i][0]);
    if (!d) continue;

    // Any X/BX in C..J for this row?
    const marks = sh.getRange(FIRST_DATA_ROW + i, FIRST_PLAYER_COL, 1, LAST_PLAYER_COL - FIRST_PLAYER_COL + 1).getValues()[0];
    const anyScheduled = marks.some(m => /X/i.test(String(m || ''))); // matches X or BX

    if (!anyScheduled) continue;

    const day = new Date(d.getTime()); day.setHours(0,0,0,0);
    if (day > today && (!bestDate || day < bestDate)) {
      bestDate = day;
      bestRow = FIRST_DATA_ROW + i;
    }
  }

  if (!bestRow) return null;
  const { start } = buildStartEnd(bestDate, START_TIME_HHMM, DURATION_MIN); // 7pm per your config
  return { rowIndex: bestRow, dateObj: start };
}
