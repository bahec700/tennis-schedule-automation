function readNextThursdayMatch() {
  const cfg = CONFIG.groups.Thursday;

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = ss.getSheetByName(cfg.sheetName);

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' + cfg.sheetName
    );
  }

  const FIRST_DATA_ROW = 8;
  const LAST_DATA_ROW = 44;

  const DATE_COL = 2;          // B
  const FIRST_PLAYER_COL = 3;  // C
  const LAST_PLAYER_COL = 10;  // J

  const HEADER_ROW = 7;

  const DIRECTORY_NAME_RANGE = 'A51:A58';
  const DIRECTORY_EMAIL_RANGE = 'F51:F58';

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // ---------------------------------
  // Build directory:
  // LASTNAME -> { fullName, email }
  // ---------------------------------

  const directory = buildThursdayDirectory_(
    sheet,
    DIRECTORY_NAME_RANGE,
    DIRECTORY_EMAIL_RANGE
  );

  // ---------------------------------
  // Player headers C:J
  // ---------------------------------

  const playerHeaders = sheet
    .getRange(
      HEADER_ROW,
      FIRST_PLAYER_COL,
      1,
      LAST_PLAYER_COL - FIRST_PLAYER_COL + 1
    )
    .getValues()[0]
    .map(v => String(v || '').trim());

  // ---------------------------------
  // Schedule data
  // ---------------------------------

  const numRows =
    LAST_DATA_ROW - FIRST_DATA_ROW + 1;

  const dates = sheet
    .getRange(
      FIRST_DATA_ROW,
      DATE_COL,
      numRows,
      1
    )
    .getValues();

  const marks = sheet
    .getRange(
      FIRST_DATA_ROW,
      FIRST_PLAYER_COL,
      numRows,
      LAST_PLAYER_COL - FIRST_PLAYER_COL + 1
    )
    .getValues();

  const backgrounds = sheet
    .getRange(
      FIRST_DATA_ROW,
      FIRST_PLAYER_COL,
      numRows,
      LAST_PLAYER_COL - FIRST_PLAYER_COL + 1
    )
    .getBackgrounds();

  // ---------------------------------
  // Find the next Thursday session
  // ---------------------------------

  for (let r = 0; r < numRows; r++) {
    const rawDate = dates[r][0];

    if (!(rawDate instanceof Date)) {
      continue;
    }

    const matchDate = new Date(rawDate);
    matchDate.setHours(0, 0, 0, 0);

    if (matchDate < today) {
      continue;
    }

    const players = [];
    const availableSubs = [];

    let ballPerson = '';

    for (let c = 0; c < playerHeaders.length; c++) {
      const headerName =
        playerHeaders[c] ||
        `Player_${c + FIRST_PLAYER_COL}`;

      const mark =
        String(marks[r][c] || '')
          .trim()
          .toUpperCase();

      const directoryEntry =
        directory[headerName.toUpperCase()] ||
        null;

      const fullName =
        directoryEntry
          ? directoryEntry.fullName
          : headerName;

      const bringBalls =
        /B/.test(mark);

      // X or BX = scheduled
      if (/X/.test(mark)) {
        players.push(fullName);

        if (bringBalls) {
          ballPerson = fullName;
        }

        continue;
      }

      // Blank + white = available sub
      const background =
        String(backgrounds[r][c] || '')
          .toLowerCase();

      if (
        !mark &&
        isWhiteBackground(background)
      ) {
        availableSubs.push(fullName);
      }
    }

    // Skip rows that have no scheduled players at all.
    if (players.length === 0) {
      continue;
    }

    // ---------------------------------
    // Scheduled-player recipients
    // ---------------------------------

    const recipients = players
      .map(fullName => {
        const lastName =
          getLastName(fullName).toUpperCase();

        const entry =
          directory[lastName];

        return entry
          ? String(entry.email || '').trim()
          : '';
      })
      .filter(Boolean);

    return createMatchObject({
      group: 'Thursday',
      date: matchDate,
      time: cfg.time,
      location: cfg.location,
      court: cfg.court,
      players,
      ballPerson,
      availableSubs,
      recipients
    });
  }

  return null;
}


// ---------------------------------
// Thursday directory helper
// ---------------------------------

function buildThursdayDirectory_(
  sheet,
  namesRange,
  emailsRange
) {
  const names = sheet
    .getRange(namesRange)
    .getValues()
    .flat();

  const emails = sheet
    .getRange(emailsRange)
    .getValues()
    .flat();

  const directory = {};

  for (
    let i = 0;
    i < Math.max(names.length, emails.length);
    i++
  ) {
    const fullName =
      String(names[i] || '').trim();

    const email =
      String(emails[i] || '').trim();

    if (!fullName) {
      continue;
    }

    const lastName =
      getLastName(fullName);

    if (!lastName) {
      continue;
    }

    directory[lastName.toUpperCase()] = {
      fullName,
      email
    };
  }

  return directory;
}