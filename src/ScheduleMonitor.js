const SCHEDULE_MONITOR_PREFIX =
  'SCHEDULE_MONITOR_SNAPSHOT:';


// ============================================================
// Public functions
// ============================================================

/**
 * Real daily monitor.
 *
 * Compares the current schedules with the previous saved snapshot.
 * If anything changed, sends one summary email to CONFIG.adminEmail.
 * The snapshot is updated only after the check completes successfully.
 *
 * First run only establishes the baseline and sends no email.
 */
function checkScheduleChanges() {
  const current =
    buildAllScheduleMonitorSnapshots_();

  const previous =
    loadAllScheduleMonitorSnapshots_();

  const missingGroups =
    Object.keys(current)
      .filter(groupName =>
        !previous[groupName]
      );

  // First run (or a newly added group): establish baseline.
  if (
    Object.keys(previous).length === 0
  ) {
    saveAllScheduleMonitorSnapshots_(
      current
    );

    Logger.log(
      'Schedule monitor initialized. ' +
      'Baseline saved; no email sent.'
    );

    return;
  }

  // If only one group is missing from an older baseline,
  // initialize that group but still compare the others.
  missingGroups.forEach(groupName => {
    previous[groupName] =
      current[groupName];

    Logger.log(
      'Initialized missing schedule-monitor baseline for ' +
      groupName +
      '.'
    );
  });

  const report =
    buildScheduleChangeReport_(
      previous,
      current
    );

  if (
    report.changedSessions.length === 0
  ) {
    saveAllScheduleMonitorSnapshots_(
      current
    );

    Logger.log(
      'Schedule monitor: no changes detected.'
    );

    return;
  }

  sendScheduleChangeAlert_(report);

  saveAllScheduleMonitorSnapshots_(
    current
  );

  Logger.log(
    'Schedule monitor: change alert sent. ' +
    report.changedSessions.length +
    ' changed session(s).'
  );
}


/**
 * Test mode.
 *
 * Performs the exact same comparison and validation as the real monitor.
 * Sends a REAL test email to CONFIG.adminEmail so the final formatting
 * can be reviewed, but does NOT update the saved baseline.
 *
 * Because the baseline is left unchanged, the same test can be repeated
 * safely. Use initializeScheduleMonitor() when you intentionally want the
 * current sheet state to become the new baseline.
 */
function testScheduleChanges() {
  const current =
    buildAllScheduleMonitorSnapshots_();

  const previous =
    loadAllScheduleMonitorSnapshots_();

  if (
    Object.keys(previous).length === 0
  ) {
    throw new Error(
      'No schedule-monitor baseline exists. ' +
      'Run initializeScheduleMonitor() first.'
    );
  }

  const report =
    buildScheduleChangeReport_(
      previous,
      current
    );

  sendScheduleChangeTestAlert_(
    report
  );

  Logger.log(
    'TEST MODE - test email sent to ' +
    CONFIG.adminEmail +
    '; snapshot was NOT updated.\n\n' +
    buildScheduleChangePlainText_(report)
  );
}


/**
 * Saves the current schedules as the monitor baseline.
 * Sends no email.
 *
 * Useful before beginning testing or after intentionally accepting
 * a large set of schedule changes.
 */
function initializeScheduleMonitor() {
  const current =
    buildAllScheduleMonitorSnapshots_();

  saveAllScheduleMonitorSnapshots_(
    current
  );

  Logger.log(
    'Schedule monitor baseline saved. No email sent.'
  );
}


/**
 * Installs the daily schedule-monitor trigger at approximately 8 AM
 * in the Apps Script project's configured timezone.
 *
 * Apps Script time-driven triggers run within the selected hour,
 * not necessarily at exactly 8:00:00.
 */
function createScheduleMonitorTrigger() {
  const functionName =
    'checkScheduleChanges';

  ScriptApp
    .getProjectTriggers()
    .filter(trigger =>
      trigger.getHandlerFunction() ===
      functionName
    )
    .forEach(trigger =>
      ScriptApp.deleteTrigger(trigger)
    );

  ScriptApp
    .newTrigger(functionName)
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log(
    'Daily schedule monitor trigger created for the 8 AM hour.'
  );
}


// ============================================================
// Snapshot creation
// ============================================================

function buildAllScheduleMonitorSnapshots_() {
  return {
    Tuesday:
      buildTueSatMonitorSnapshot_(
        'Tuesday'
      ),

    Thursday:
      buildThursdayMonitorSnapshot_(),

    Saturday:
      buildTueSatMonitorSnapshot_(
        'Saturday'
      )
  };
}


function buildTueSatMonitorSnapshot_(
  groupName
) {
  const cfg =
    CONFIG.groups[groupName];

  const layout =
    TUE_SAT_LAYOUT[groupName];

  if (
    !cfg ||
    !layout
  ) {
    throw new Error(
      'Missing monitor configuration for ' +
      groupName
    );
  }

  const ss =
    SpreadsheetApp.openById(
      cfg.spreadsheetId
    );

  const sheet =
    ss.getSheetByName(
      cfg.sheetName
    );

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' +
      cfg.sheetName
    );
  }

  const directory =
    readTueSatDirectory_(
      sheet,
      layout.directoryNamesRange,
      layout.directoryEmailsRange
    );

  const sessions = {};

  layout.parts.forEach(
    (part, partIndex) => {
      const firstPlayerCol =
        columnLetterToNumber_(
          part.playersStartCol
        );

      const lastPlayerCol =
        columnLetterToNumber_(
          part.playersEndCol
        );

      const dateCol =
        columnLetterToNumber_(
          part.dateCol
        );

      const numRows =
        part.lastRow -
        part.firstRow +
        1;

      const numPlayerCols =
        lastPlayerCol -
        firstPlayerCol +
        1;

      const headers = sheet
        .getRange(
          layout.headerRow,
          firstPlayerCol,
          1,
          numPlayerCols
        )
        .getValues()[0]
        .map(v =>
          String(v || '').trim()
        );

      const dates = sheet
        .getRange(
          part.firstRow,
          dateCol,
          numRows,
          1
        )
        .getValues();

      const values = sheet
        .getRange(
          part.firstRow,
          firstPlayerCol,
          numRows,
          numPlayerCols
        )
        .getValues();

      for (
        let r = 0;
        r < numRows;
        r++
      ) {
        const rawDate =
          dates[r][0];

        if (
          !(rawDate instanceof Date)
        ) {
          continue;
        }

        const date =
          normalizeMonitorDate_(
            rawDate
          );

        const dateKey =
          formatMonitorDateKey_(date);

        const row =
          part.firstRow + r;

        const players =
          headers.map(
            (header, c) => {
              const mapped =
                mapTueSatHeaderToDirectory_(
                  header,
                  directory
                );

              return {
                header:
                  header ||
                  `Player_${c + 1}`,

                fullName:
                  mapped.fullName ||
                  header ||
                  `Player_${c + 1}`,

                email:
                  String(
                    mapped.email || ''
                  ).trim(),

                col:
                  firstPlayerCol + c
              };
            }
          );

        sessions[dateKey] = {
          group:
            groupName,

          dateKey,

          dateLabel:
            formatMonitorDateLong_(
              date
            ),

          row,

          partIndex,

          values:
            values[r].map(
              normalizeMonitorCellValue_
            ),

          backgrounds: [],

          players
        };
      }
    }
  );

  return {
    group:
      groupName,

    spreadsheetId:
      cfg.spreadsheetId,

    sheetName:
      cfg.sheetName,

    sheetId:
      sheet.getSheetId(),

    readerType:
      'TUE_SAT',

    sessions
  };
}


function buildThursdayMonitorSnapshot_() {
  const groupName =
    'Thursday';

  const cfg =
    CONFIG.groups[groupName];

  const ss =
    SpreadsheetApp.openById(
      cfg.spreadsheetId
    );

  const sheet =
    ss.getSheetByName(
      cfg.sheetName
    );

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' +
      cfg.sheetName
    );
  }

  const FIRST_DATA_ROW = 8;
  const LAST_DATA_ROW = 44;

  const DATE_COL = 2;
  const FIRST_PLAYER_COL = 3;
  const LAST_PLAYER_COL = 10;
  const HEADER_ROW = 7;

  const directory =
    buildThursdayDirectory_(
      sheet,
      'A51:A58',
      'F51:F58'
    );

  const numRows =
    LAST_DATA_ROW -
    FIRST_DATA_ROW +
    1;

  const numPlayerCols =
    LAST_PLAYER_COL -
    FIRST_PLAYER_COL +
    1;

  const headers = sheet
    .getRange(
      HEADER_ROW,
      FIRST_PLAYER_COL,
      1,
      numPlayerCols
    )
    .getValues()[0]
    .map(v =>
      String(v || '').trim()
    );

  const dates = sheet
    .getRange(
      FIRST_DATA_ROW,
      DATE_COL,
      numRows,
      1
    )
    .getValues();

  const values = sheet
    .getRange(
      FIRST_DATA_ROW,
      FIRST_PLAYER_COL,
      numRows,
      numPlayerCols
    )
    .getValues();

  const backgrounds = sheet
    .getRange(
      FIRST_DATA_ROW,
      FIRST_PLAYER_COL,
      numRows,
      numPlayerCols
    )
    .getBackgrounds();

  const sessions = {};

  for (
    let r = 0;
    r < numRows;
    r++
  ) {
    const rawDate =
      dates[r][0];

    if (
      !(rawDate instanceof Date)
    ) {
      continue;
    }

    const date =
      normalizeMonitorDate_(
        rawDate
      );

    const dateKey =
      formatMonitorDateKey_(
        date
      );

    const players =
      headers.map(
        (header, c) => {
          const entry =
            directory[
              String(
                header || ''
              ).toUpperCase()
            ] || null;

          return {
            header:
              header ||
              `Player_${c + 1}`,

            fullName:
              entry
                ? entry.fullName
                : (
                    header ||
                    `Player_${c + 1}`
                  ),

            email:
              entry
                ? String(
                    entry.email || ''
                  ).trim()
                : '',

            col:
              FIRST_PLAYER_COL + c
          };
        }
      );

    sessions[dateKey] = {
      group:
        groupName,

      dateKey,

      dateLabel:
        formatMonitorDateLong_(
          date
        ),

      row:
        FIRST_DATA_ROW + r,

      partIndex: 0,

      values:
        values[r].map(
          normalizeMonitorCellValue_
        ),

      backgrounds:
        backgrounds[r].map(
          color =>
            String(
              color || ''
            ).toLowerCase()
        ),

      players
    };
  }

  return {
    group:
      groupName,

    spreadsheetId:
      cfg.spreadsheetId,

    sheetName:
      cfg.sheetName,

    sheetId:
      sheet.getSheetId(),

    readerType:
      'THURSDAY',

    sessions
  };
}


// ============================================================
// Comparison and validation
// ============================================================

function buildScheduleChangeReport_(
  previous,
  current
) {
  const changedSessions = [];

  [
    'Tuesday',
    'Thursday',
    'Saturday'
  ].forEach(groupName => {
    const oldGroup =
      previous[groupName];

    const newGroup =
      current[groupName];

    if (
      !oldGroup ||
      !newGroup
    ) {
      return;
    }

    const oldSessions =
      oldGroup.sessions || {};

    const newSessions =
      newGroup.sessions || {};

    const allDateKeys = [
      ...new Set([
        ...Object.keys(
          oldSessions
        ),
        ...Object.keys(
          newSessions
        )
      ])
    ].sort();

    allDateKeys.forEach(
      dateKey => {
        const oldSession =
          oldSessions[dateKey] ||
          null;

        const newSession =
          newSessions[dateKey] ||
          null;

        if (
          oldSession &&
          newSession &&
          monitorSessionsEqual_(
            oldSession,
            newSession
          )
        ) {
          return;
        }

        changedSessions.push(
          analyzeChangedSession_(
            oldGroup,
            newGroup,
            oldSession,
            newSession
          )
        );
      }
    );
  });

  return {
    generatedAt:
      new Date(),

    changedSessions
  };
}


function analyzeChangedSession_(
  oldGroup,
  newGroup,
  oldSession,
  newSession
) {
  const groupName =
    newGroup.group ||
    oldGroup.group;

  const session =
    newSession ||
    oldSession;

  const result = {
    group:
      groupName,

    dateKey:
      session.dateKey,

    dateLabel:
      session.dateLabel,

    scheduleUrl:
      buildMonitorScheduleUrl_(
        newGroup ||
        oldGroup
      ),

    status:
      'LOOKS GOOD',

    problems: [],

    summaryLines: [],

    changes: [],

    technicalChanges: []
  };

  if (
    !oldSession &&
    newSession
  ) {
    result.changes.push(
      'A new dated session was added to the schedule.'
    );

    result.technicalChanges.push(
      'New schedule row: ' +
      newSession.dateLabel
    );
  }

  if (
    oldSession &&
    !newSession
  ) {
    result.status =
      'PLEASE REVIEW';

    result.problems.push(
      'This dated session was removed from the schedule or its date was changed.'
    );

    result.changes.push(
      'The previously scheduled date is no longer present.'
    );

    result.technicalChanges.push(
      'Removed schedule row: ' +
      oldSession.dateLabel
    );

    return result;
  }

  if (
    newSession
  ) {
    const validation =
      validateMonitorSession_(
        newGroup,
        newSession
      );

    result.problems.push(
      ...validation.problems
    );

    result.summaryLines.push(
      ...validation.summaryLines
    );

    if (
      validation.problems.length > 0
    ) {
      result.status =
        'PLEASE REVIEW';
    }
  }

  if (
    oldSession &&
    newSession
  ) {
    const changes =
      describeMonitorCellChanges_(
        newGroup,
        oldSession,
        newSession
      );

    result.changes.push(
      ...changes.human
    );

    result.technicalChanges.push(
      ...changes.technical
    );
  }

  return result;
}


function validateMonitorSession_(
  groupSnapshot,
  session
) {
  const match =
    buildMonitorMatchObject_(
      groupSnapshot,
      session
    );

  const problems =
    validateReminderMatch_(
      match
    );

  const unknownValues =
    findMonitorUnknownValues_(
      groupSnapshot,
      session
    );

  unknownValues.forEach(
    issue => {
      problems.push(
        issue.player +
        ': unrecognized value "' +
        issue.value +
        '" in ' +
        issue.a1
      );
    }
  );

  const summaryLines = [
    match.players.length +
      ' scheduled player' +
      (
        match.players.length === 1
          ? ''
          : 's'
      ),

    match.ballPeople.length +
      ' player' +
      (
        match.ballPeople.length === 1
          ? ''
          : 's'
      ) +
      ' bringing balls'
  ];

  if (
    match.unresolvedPlayers.length === 0 &&
    match.recipients.length ===
      match.players.length
  ) {
    summaryLines.push(
      'All scheduled players have email addresses'
    );
  }

  if (
    unknownValues.length === 0
  ) {
    summaryLines.push(
      'No unrecognized schedule values'
    );
  }

  return {
    match,
    problems,
    summaryLines
  };
}


function buildMonitorMatchObject_(
  groupSnapshot,
  session
) {
  const players = [];
  const ballPeople = [];
  const recipients = [];
  const unresolvedPlayers = [];

  session.values.forEach(
    (value, index) => {
      const playerInfo =
        session.players[index];

      const state =
        getMonitorPlayerState_(
          groupSnapshot.readerType,
          value,
          session.backgrounds[index]
        );

      if (
        !state.scheduled
      ) {
        return;
      }

      players.push(
        playerInfo.fullName
      );

      if (
        state.balls
      ) {
        ballPeople.push(
          playerInfo.fullName
        );
      }

      if (
        playerInfo.email
      ) {
        recipients.push(
          playerInfo.email
        );
      } else {
        unresolvedPlayers.push(
          playerInfo.fullName
        );
      }
    }
  );

  const cfg =
    CONFIG.groups[
      groupSnapshot.group
    ];

  return createMatchObject({
    group:
      groupSnapshot.group,

    date:
      parseMonitorDateKey_(
        session.dateKey
      ),

    time:
      cfg.time,

    location:
      cfg.location || '',

    court:
      cfg.court || '',

    players,

    ballPerson:
      ballPeople.length === 1
        ? ballPeople[0]
        : (
            ballPeople.length
              ? ballPeople[
                  ballPeople.length - 1
                ]
              : ''
          ),

    ballPeople,

    availableSubs: [],

    availableSubRecipients: [],

    recipients,

    unresolvedPlayers
  });
}


function findMonitorUnknownValues_(
  groupSnapshot,
  session
) {
  const issues = [];

  session.values.forEach(
    (value, index) => {
      if (
        isMonitorAllowedValue_(
          groupSnapshot.readerType,
          value
        )
      ) {
        return;
      }

      issues.push({
        player:
          session.players[index]
            .fullName,

        value:
          value,

        a1:
          monitorCellA1_(
            session.players[index]
              .col,
            session.row
          )
      });
    }
  );

  return issues;
}


// ============================================================
// Human-readable change descriptions
// ============================================================

function describeMonitorCellChanges_(
  groupSnapshot,
  oldSession,
  newSession
) {
  const human = [];
  const technical = [];

  const count =
    Math.max(
      oldSession.values.length,
      newSession.values.length
    );

  for (
    let i = 0;
    i < count;
    i++
  ) {
    const oldValue =
      oldSession.values[i] || '';

    const newValue =
      newSession.values[i] || '';

    const oldBg =
      oldSession.backgrounds[i] || '';

    const newBg =
      newSession.backgrounds[i] || '';

    if (
      oldValue === newValue &&
      oldBg === newBg
    ) {
      continue;
    }

    const player =
      (
        newSession.players[i] ||
        oldSession.players[i]
      );

    if (!player) {
      continue;
    }

    const a1 =
      monitorCellA1_(
        player.col,
        newSession.row
      );

    const description =
      describeMonitorPlayerTransition_(
        groupSnapshot.readerType,
        player.fullName,
        oldValue,
        newValue,
        oldBg,
        newBg
      );

    human.push(
      description
    );

    let detail =
      a1 +
      ': ' +
      displayMonitorValue_(
        oldValue
      ) +
      ' → ' +
      displayMonitorValue_(
        newValue
      );

    if (
      oldValue === newValue &&
      oldBg !== newBg
    ) {
      detail +=
        ' (background changed ' +
        oldBg +
        ' → ' +
        newBg +
        ')';
    }

    technical.push(
      detail
    );
  }

  return {
    human,
    technical
  };
}


function describeMonitorPlayerTransition_(
  readerType,
  playerName,
  oldValue,
  newValue,
  oldBackground,
  newBackground
) {
  const oldState =
    getMonitorPlayerState_(
      readerType,
      oldValue,
      oldBackground
    );

  const newState =
    getMonitorPlayerState_(
      readerType,
      newValue,
      newBackground
    );

  if (
    oldState.scheduled &&
    !newState.scheduled
  ) {
    if (
      newState.unavailable
    ) {
      return (
        playerName +
        ': removed from roster; marked unavailable'
      );
    }

    if (
      newState.availableSub
    ) {
      return (
        playerName +
        ': removed from roster; now available as a sub'
      );
    }

    return (
      playerName +
      ': removed from roster'
    );
  }

  if (
    !oldState.scheduled &&
    newState.scheduled
  ) {
    return (
      playerName +
      ': added to roster' +
      (
        newState.balls
          ? ' and now bringing balls'
          : ''
      )
    );
  }

  if (
    oldState.scheduled &&
    newState.scheduled
  ) {
    if (
      !oldState.balls &&
      newState.balls
    ) {
      return (
        playerName +
        ': now bringing balls'
      );
    }

    if (
      oldState.balls &&
      !newState.balls
    ) {
      return (
        playerName +
        ': no longer bringing balls'
      );
    }
  }

  if (
    oldState.availableSub &&
    newState.unavailable
  ) {
    return (
      playerName +
      ': available sub → unavailable'
    );
  }

  if (
    oldState.unavailable &&
    newState.availableSub
  ) {
    return (
      playerName +
      ': unavailable → available as a sub'
    );
  }

  if (
    newState.unknown
  ) {
    return (
      playerName +
      ': changed to unrecognized value "' +
      newValue +
      '"'
    );
  }

  if (
    oldState.unknown
  ) {
    return (
      playerName +
      ': unrecognized value "' +
      oldValue +
      '" changed to ' +
      describeMonitorState_(
        newState
      )
    );
  }

  if (
    oldValue === newValue &&
    oldBackground !==
      newBackground
  ) {
    return (
      playerName +
      ': availability formatting changed'
    );
  }

  return (
    playerName +
    ': ' +
    displayMonitorValue_(
      oldValue
    ) +
    ' → ' +
    displayMonitorValue_(
      newValue
    )
  );
}


function describeMonitorState_(
  state
) {
  if (
    state.scheduled &&
    state.balls
  ) {
    return 'scheduled and bringing balls';
  }

  if (
    state.scheduled
  ) {
    return 'scheduled';
  }

  if (
    state.unavailable
  ) {
    return 'unavailable';
  }

  if (
    state.availableSub
  ) {
    return 'available as a sub';
  }

  if (
    state.unknown
  ) {
    return 'an unrecognized value';
  }

  return 'blank';
}


// ============================================================
// Email output
// ============================================================

function sendScheduleChangeTestAlert_(
  report
) {
  if (
    !CONFIG.adminEmail
  ) {
    throw new Error(
      'CONFIG.adminEmail is required for schedule-change test alerts.'
    );
  }

  const count =
    report.changedSessions.length;

  const subject =
    count === 0
      ? 'TEST – Tennis schedule monitor – no changes detected'
      : (
          count === 1
            ? (
                'TEST – Tennis schedule changed – ' +
                report.changedSessions[0]
                  .group +
                ' ' +
                formatMonitorDateShortFromKey_(
                  report.changedSessions[0]
                    .dateKey
                )
              )
            : (
                'TEST – Tennis schedules changed – ' +
                count +
                ' sessions'
              )
        );

  GmailApp.sendEmail(
    CONFIG.adminEmail,
    subject,
    buildScheduleChangePlainText_(
      report
    ),
    {
      htmlBody:
        '<p><strong>TEST MODE</strong> — this is a test notification. ' +
        'The saved schedule baseline was not changed.</p>' +
        buildScheduleChangeHtml_(
          report
        )
    }
  );
}


function sendScheduleChangeAlert_(
  report
) {
  if (
    !CONFIG.adminEmail
  ) {
    throw new Error(
      'CONFIG.adminEmail is required for schedule-change alerts.'
    );
  }

  const count =
    report.changedSessions.length;

  const subject =
    count === 1
      ? (
          'Tennis schedule changed – ' +
          report.changedSessions[0]
            .group +
          ' ' +
          formatMonitorDateShortFromKey_(
            report.changedSessions[0]
              .dateKey
          )
        )
      : (
          'Tennis schedules changed – ' +
          count +
          ' sessions'
        );

  GmailApp.sendEmail(
    CONFIG.adminEmail,
    subject,
    buildScheduleChangePlainText_(
      report
    ),
    {
      htmlBody:
        buildScheduleChangeHtml_(
          report
        )
    }
  );
}


function buildScheduleChangePlainText_(
  report
) {
  if (
    report.changedSessions.length === 0
  ) {
    return (
      'No schedule changes detected.'
    );
  }

  const lines = [];

  lines.push(
    'Tennis schedule changes detected.'
  );

  lines.push('');

  report.changedSessions.forEach(
    (session, index) => {
      lines.push(
        session.group +
        ' – ' +
        session.dateLabel
      );

      lines.push(
        'Status: ' +
        (
          session.status ===
          'LOOKS GOOD'
            ? 'Looks good'
            : 'PLEASE REVIEW'
        )
      );

      if (
        session.problems.length
      ) {
        lines.push('');
        lines.push('Problems:');

        session.problems.forEach(
          problem =>
            lines.push(
              '- ' + problem
            )
        );
      } else if (
        session.summaryLines.length
      ) {
        session.summaryLines.forEach(
          item =>
            lines.push(
              '- ' + item
            )
        );
      }

      if (
        session.changes.length
      ) {
        lines.push('');
        lines.push('Changes:');

        session.changes.forEach(
          change =>
            lines.push(
              '- ' + change
            )
        );
      }

      if (
        session.technicalChanges.length
      ) {
        lines.push('');
        lines.push(
          'Technical details:'
        );

        session.technicalChanges
          .forEach(
            change =>
              lines.push(
                '- ' + change
              )
          );
      }

      lines.push('');
      lines.push(
        'View schedule: ' +
        session.scheduleUrl
      );

      if (
        index <
        report.changedSessions.length -
        1
      ) {
        lines.push('');
        lines.push(
          '----------------------------------------'
        );
        lines.push('');
      }
    }
  );

  return lines.join('\n');
}


function buildScheduleChangeHtml_(
  report
) {
  if (
    report.changedSessions.length === 0
  ) {
    return (
      '<p>No schedule changes detected.</p>'
    );
  }

  const blocks =
    report.changedSessions.map(
      session => {
        const good =
          session.status ===
          'LOOKS GOOD';

        const problemsHtml =
          session.problems.length
            ? (
                '<p><strong>Problems:</strong></p>' +
                '<ul>' +
                session.problems
                  .map(
                    item =>
                      '<li>' +
                      htmlEscape_(item) +
                      '</li>'
                  )
                  .join('') +
                '</ul>'
              )
            : (
                '<ul>' +
                session.summaryLines
                  .map(
                    item =>
                      '<li>' +
                      htmlEscape_(item) +
                      '</li>'
                  )
                  .join('') +
                '</ul>'
              );

        const changesHtml =
          session.changes.length
            ? (
                '<p><strong>Changes:</strong></p>' +
                '<ul>' +
                session.changes
                  .map(
                    item =>
                      '<li>' +
                      htmlEscape_(item) +
                      '</li>'
                  )
                  .join('') +
                '</ul>'
              )
            : '';

        const technicalHtml =
          session.technicalChanges.length
            ? (
                '<p><strong>Technical details:</strong></p>' +
                '<ul>' +
                session.technicalChanges
                  .map(
                    item =>
                      '<li><code>' +
                      htmlEscape_(item) +
                      '</code></li>'
                  )
                  .join('') +
                '</ul>'
              )
            : '';

        return (
          '<div style="margin-bottom:24px;">' +
            '<p style="margin-bottom:4px;">' +
              '<strong>' +
                htmlEscape_(
                  session.group +
                  ' – ' +
                  session.dateLabel
                ) +
              '</strong>' +
            '</p>' +
            '<p style="margin-top:0;">' +
              '<strong>Status: ' +
                (
                  good
                    ? 'Looks good'
                    : 'PLEASE REVIEW'
                ) +
              '</strong>' +
            '</p>' +
            problemsHtml +
            changesHtml +
            technicalHtml +
            '<p>' +
              '<a href="' +
                htmlEscape_(
                  session.scheduleUrl
                ) +
              '">View schedule</a>' +
            '</p>' +
          '</div>'
        );
      }
    );

  return (
    '<div>' +
      '<p>Tennis schedule changes detected.</p>' +
      blocks.join(
        '<hr style="border:none;border-top:1px solid #ddd;">'
      ) +
    '</div>'
  );
}


// ============================================================
// State helpers
// ============================================================

function loadAllScheduleMonitorSnapshots_() {
  const properties =
    PropertiesService
      .getScriptProperties();

  const result = {};

  [
    'Tuesday',
    'Thursday',
    'Saturday'
  ].forEach(groupName => {
    const raw =
      properties.getProperty(
        SCHEDULE_MONITOR_PREFIX +
        groupName
      );

    if (!raw) {
      return;
    }

    try {
      result[groupName] =
        JSON.parse(raw);
    } catch (error) {
      throw new Error(
        'Unable to parse saved schedule-monitor snapshot for ' +
        groupName +
        ': ' +
        error.message
      );
    }
  });

  return result;
}


function saveAllScheduleMonitorSnapshots_(
  snapshots
) {
  const properties =
    PropertiesService
      .getScriptProperties();

  Object.keys(
    snapshots
  ).forEach(groupName => {
    properties.setProperty(
      SCHEDULE_MONITOR_PREFIX +
      groupName,
      JSON.stringify(
        compactMonitorSnapshot_(
          snapshots[groupName]
        )
      )
    );
  });
}


/**
 * PropertiesService values have size limits. Store only the fields needed
 * for tomorrow's comparison; reconstruct names/emails from the live sheet.
 */
function compactMonitorSnapshot_(
  groupSnapshot
) {
  const compactSessions = {};

  Object.keys(
    groupSnapshot.sessions
  ).forEach(dateKey => {
    const session =
      groupSnapshot.sessions[
        dateKey
      ];

    compactSessions[dateKey] = {
      group:
        session.group,

      dateKey:
        session.dateKey,

      dateLabel:
        session.dateLabel,

      row:
        session.row,

      partIndex:
        session.partIndex,

      values:
        session.values,

      backgrounds:
        session.backgrounds,

      // Keep player metadata because it lets tomorrow's report
      // identify the person even if the directory/header changes.
      players:
        session.players
    };
  });

  return {
    group:
      groupSnapshot.group,

    spreadsheetId:
      groupSnapshot.spreadsheetId,

    sheetName:
      groupSnapshot.sheetName,

    sheetId:
      groupSnapshot.sheetId,

    readerType:
      groupSnapshot.readerType,

    sessions:
      compactSessions
  };
}


// ============================================================
// Value/state helpers
// ============================================================

function getMonitorPlayerState_(
  readerType,
  rawValue,
  background
) {
  const value =
    String(
      rawValue || ''
    )
      .trim()
      .toUpperCase();

  if (
    readerType ===
    'TUE_SAT'
  ) {
    if (
      value === 'PLAY'
    ) {
      return {
        scheduled: true,
        balls: false,
        unavailable: false,
        availableSub: false,
        unknown: false
      };
    }

    if (
      value === 'BALL'
    ) {
      return {
        scheduled: true,
        balls: true,
        unavailable: false,
        availableSub: false,
        unknown: false
      };
    }

    if (
      value === 'X'
    ) {
      return {
        scheduled: false,
        balls: false,
        unavailable: true,
        availableSub: false,
        unknown: false
      };
    }

    if (
      value === ''
    ) {
      return {
        scheduled: false,
        balls: false,
        unavailable: false,
        availableSub: true,
        unknown: false
      };
    }

    return {
      scheduled: false,
      balls: false,
      unavailable: false,
      availableSub: false,
      unknown: true
    };
  }

  // Thursday format.
  if (
    value === 'X'
  ) {
    return {
      scheduled: true,
      balls: false,
      unavailable: false,
      availableSub: false,
      unknown: false
    };
  }

  if (
    value === 'BX'
  ) {
    return {
      scheduled: true,
      balls: true,
      unavailable: false,
      availableSub: false,
      unknown: false
    };
  }

  if (
    value === ''
  ) {
    return {
      scheduled: false,
      balls: false,
      unavailable: false,
      availableSub:
        isWhiteBackground(
          background
        ),
      unknown: false
    };
  }

  return {
    scheduled: false,
    balls: false,
    unavailable: false,
    availableSub: false,
    unknown: true
  };
}


function isMonitorAllowedValue_(
  readerType,
  rawValue
) {
  const value =
    String(
      rawValue || ''
    )
      .trim()
      .toUpperCase();

  if (
    readerType ===
    'TUE_SAT'
  ) {
    return (
      value === '' ||
      value === 'PLAY' ||
      value === 'BALL' ||
      value === 'X'
    );
  }

  return (
    value === '' ||
    value === 'X' ||
    value === 'BX'
  );
}


function monitorSessionsEqual_(
  oldSession,
  newSession
) {
  return (
    JSON.stringify(
      oldSession.values || []
    ) ===
      JSON.stringify(
        newSession.values || []
      ) &&
    JSON.stringify(
      oldSession.backgrounds || []
    ) ===
      JSON.stringify(
        newSession.backgrounds || []
      )
  );
}


function normalizeMonitorCellValue_(
  value
) {
  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value).trim();
}


function displayMonitorValue_(
  value
) {
  const text =
    String(
      value || ''
    );

  return text
    ? '`' + text + '`'
    : 'blank';
}


// ============================================================
// Date / URL / cell helpers
// ============================================================

function normalizeMonitorDate_(
  value
) {
  const date =
    new Date(value);

  date.setHours(
    0,
    0,
    0,
    0
  );

  return date;
}


function formatMonitorDateKey_(
  date
) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function formatMonitorDateLong_(
  date
) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'EEEE, MMMM d, yyyy'
  );
}


function formatMonitorDateShortFromKey_(
  dateKey
) {
  return Utilities.formatDate(
    parseMonitorDateKey_(
      dateKey
    ),
    Session.getScriptTimeZone(),
    'EEE MMM d'
  );
}


function parseMonitorDateKey_(
  dateKey
) {
  const parts =
    String(dateKey)
      .split('-')
      .map(Number);

  return new Date(
    parts[0],
    parts[1] - 1,
    parts[2]
  );
}


function buildMonitorScheduleUrl_(
  groupSnapshot
) {
  return (
    'https://docs.google.com/spreadsheets/d/' +
    groupSnapshot.spreadsheetId +
    '/edit#gid=' +
    groupSnapshot.sheetId
  );
}


function monitorCellA1_(
  columnNumber,
  row
) {
  return (
    monitorColumnNumberToLetter_(
      columnNumber
    ) +
    row
  );
}


function monitorColumnNumberToLetter_(
  columnNumber
) {
  let number =
    Number(columnNumber);

  let result = '';

  while (
    number > 0
  ) {
    const remainder =
      (number - 1) % 26;

    result =
      String.fromCharCode(
        65 + remainder
      ) +
      result;

    number =
      Math.floor(
        (number - 1) / 26
      );
  }

  return result;
}
