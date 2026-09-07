const TUE_SAT_LAYOUT = {
  Tuesday: {
    headerRow: 14,

    directoryNamesRange: 'A3:A10',
    directoryEmailsRange: 'H3:H10',

    parts: [
      {
        firstRow: 15,
        lastRow: 33,
        dateCol: 'A',
        playersStartCol: 'B',
        playersEndCol: 'H'
      },
      {
        firstRow: 15,
        lastRow: 32,
        dateCol: 'K',
        playersStartCol: 'L',
        playersEndCol: 'R'
      }
    ]
  },

  Saturday: {
    headerRow: 14,

    directoryNamesRange: 'A3:A11',
    directoryEmailsRange: 'H3:H11',

    parts: [
      {
        firstRow: 15,
        lastRow: 33,
        dateCol: 'A',
        playersStartCol: 'B',
        playersEndCol: 'I'
      },
      {
        firstRow: 15,
        lastRow: 32,
        dateCol: 'L',
        playersStartCol: 'M',
        playersEndCol: 'T'
      }
    ]
  }
};


// ---------------------------------
// Explicit aliases from old script
// ---------------------------------

const TUE_SAT_NAME_ALIASES = {
  'Jim B': 'Jim Bryant'
};


// ---------------------------------
// Public reader functions
// ---------------------------------

function readNextTuesdayMatch() {
  return readNextTueSatMatch('Tuesday');
}


function readNextSaturdayMatch() {
  return readNextTueSatMatch('Saturday');
}


// ---------------------------------
// Main reader
// ---------------------------------

function readNextTueSatMatch(groupName) {
  const cfg =
    CONFIG.groups[groupName];

  const layout =
    TUE_SAT_LAYOUT[groupName];

  if (!cfg) {
    throw new Error(
      'Missing CONFIG for group: ' +
      groupName
    );
  }

  if (!layout) {
    throw new Error(
      'Missing layout for group: ' +
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

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let bestMatch = null;

  for (const part of layout.parts) {
    const candidate =
      readTueSatPart_(
        sheet,
        groupName,
        cfg,
        layout,
        part,
        directory,
        today
      );

    if (!candidate) {
      continue;
    }

    if (
      !bestMatch ||
      candidate.date.getTime() <
      bestMatch.date.getTime()
    ) {
      bestMatch = candidate;
    }
  }

  return bestMatch;
}


// ---------------------------------
// Read one left/right schedule pane
// ---------------------------------

function readTueSatPart_(
  sheet,
  groupName,
  cfg,
  layout,
  part,
  directory,
  today
) {
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

  // ---------------------------------
  // Headers: Mike, Jim H, Ivan, etc.
  // ---------------------------------

  const playerHeaders = sheet
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

  // ---------------------------------
  // Dates and marks
  // ---------------------------------

  const dates = sheet
    .getRange(
      part.firstRow,
      dateCol,
      numRows,
      1
    )
    .getValues();

  const marks = sheet
    .getRange(
      part.firstRow,
      firstPlayerCol,
      numRows,
      numPlayerCols
    )
    .getValues();

  for (let r = 0; r < numRows; r++) {
    const rawDate =
      dates[r][0];

    if (!(rawDate instanceof Date)) {
      continue;
    }

    const matchDate =
      new Date(rawDate);

    matchDate.setHours(
      0,
      0,
      0,
      0
    );

    if (matchDate < today) {
      continue;
    }

    const players = [];
    const availableSubs = [];
    const unresolvedNames = [];

    let ballPerson = '';

    // ---------------------------------
    // Interpret row
    //
    // Play = scheduled
    // Ball = scheduled + balls
    // x    = unavailable
    // blank = potential sub
    // ---------------------------------

    for (
      let c = 0;
      c < playerHeaders.length;
      c++
    ) {
      const headerLabel =
        playerHeaders[c] ||
        `Player_${c}`;

      const cell =
        String(marks[r][c] || '')
          .trim();

      const mapped =
        mapTueSatHeaderToDirectory_(
          headerLabel,
          directory
        );

      const fullName =
        mapped.fullName ||
        headerLabel;

      if (/^Ball$/i.test(cell)) {
        players.push(fullName);
        ballPerson = fullName;

        if (!mapped.email) {
          unresolvedNames.push(fullName);
        }

        continue;
      }

      if (/^Play$/i.test(cell)) {
        players.push(fullName);

        if (!mapped.email) {
          unresolvedNames.push(fullName);
        }

        continue;
      }

      if (/^x$/i.test(cell)) {
        continue;
      }

      if (cell === '') {
        availableSubs.push(fullName);
      }
    }

    // Skip rows that have no scheduled players.
    if (players.length === 0) {
      continue;
    }

    if (unresolvedNames.length) {
      Logger.log(
        groupName +
        ': could not find an email for scheduled player(s): ' +
        unresolvedNames.join(', ')
      );
    }

    // ---------------------------------
    // Available-sub recipient addresses
    // ---------------------------------

    const availableSubRecipients = availableSubs
      .map(fullName => {
        const parts =
          splitTueSatName_(
            fullName
          );

        const key =
          `${parts.first.toLowerCase()} ` +
          `${parts.last.toLowerCase()}`;

        const entry =
          directory.byFullLower[key];

        return entry
          ? String(entry.email || '').trim()
          : '';
      })
      .filter(Boolean);

    // ---------------------------------
    // Recipient addresses
    // ---------------------------------

    const recipients = players
      .map(fullName => {
        const parts =
          splitTueSatName_(
            fullName
          );

        const key =
          `${parts.first.toLowerCase()} ` +
          `${parts.last.toLowerCase()}`;

        const entry =
          directory.byFullLower[key];

        return entry
          ? String(entry.email || '').trim()
          : '';
      })
      .filter(Boolean);

    return createMatchObject({
      group: groupName,
      date: matchDate,
      time: cfg.time,
      location:
        cfg.location || '',
      court:
        cfg.court || '',
      players,
      ballPerson,
      availableSubs,
      availableSubRecipients,
      recipients,
      unresolvedPlayers: unresolvedNames
    });
  }

  return null;
}


// ---------------------------------
// Build directory lookup maps
// ---------------------------------

function readTueSatDirectory_(
  sheet,
  namesA1,
  emailsA1
) {
  const names = sheet
    .getRange(namesA1)
    .getValues()
    .flat()
    .map(v =>
      String(v || '').trim()
    );

  const emails = sheet
    .getRange(emailsA1)
    .getValues()
    .flat()
    .map(v =>
      String(v || '').trim()
    );

  const byLast = {};
  const byFullLower = {};
  const byFirst = {};
  const byFirstInit = {};

  for (
    let i = 0;
    i < Math.max(
      names.length,
      emails.length
    );
    i++
  ) {
    const fullName =
      names[i] || '';

    const email =
      emails[i] || '';

    if (!fullName) {
      continue;
    }

    const parts =
      splitTueSatName_(
        fullName
      );

    const last =
      parts.last.toUpperCase();

    const first =
      parts.first.toLowerCase();

    const lastInit =
      parts.last
        ? parts.last[0]
            .toLowerCase()
        : '';

    const record = {
      fullName,
      email
    };

    if (last) {
      byLast[last] =
        record;
    }

    byFullLower[
      `${first} ` +
      `${parts.last.toLowerCase()}`
    ] = record;

    if (!byFirst[first]) {
      byFirst[first] = [];
    }

    byFirst[first]
      .push(record);

    if (
      first &&
      lastInit
    ) {
      byFirstInit[
        `${first} ${lastInit}`
      ] = record;
    }
  }

  return {
    byLast,
    byFullLower,
    byFirst,
    byFirstInit
  };
}


// ---------------------------------
// Map schedule header to directory
// ---------------------------------

function mapTueSatHeaderToDirectory_(
  headerLabel,
  directory
) {
  const raw =
    String(headerLabel || '')
      .trim();

  if (!raw) {
    return {
      fullName: '',
      email: ''
    };
  }

  // 1. Explicit alias
  const aliasFull =
    TUE_SAT_NAME_ALIASES[raw];

  if (aliasFull) {
    const parts =
      splitTueSatName_(
        aliasFull
      );

    const key =
      `${parts.first.toLowerCase()} ` +
      `${parts.last.toLowerCase()}`;

    const match =
      directory.byFullLower[key];

    if (match) {
      return match;
    }
  }

  // 2. Exact full-name match
  const rawParts =
    splitTueSatName_(raw);

  const exactKey =
    `${rawParts.first.toLowerCase()} ` +
    `${rawParts.last.toLowerCase()}`;

  if (
    directory.byFullLower[
      exactKey
    ]
  ) {
    return directory
      .byFullLower[
        exactKey
      ];
  }

  // 3. First + last initial
  const initialMatch =
    raw.match(
      /^([A-Za-z]+)\s+([A-Za-z])$/
    );

  if (initialMatch) {
    const first =
      initialMatch[1]
        .toLowerCase();

    const lastInitial =
      initialMatch[2]
        .toLowerCase();

    const key =
      `${first} ${lastInitial}`;

    if (
      directory.byFirstInit[
        key
      ]
    ) {
      return directory
        .byFirstInit[
          key
        ];
    }
  }

  // 4. Unique first-name match
  const firstOnly =
    rawParts.first
      .toLowerCase();

  const candidates =
    directory.byFirst[
      firstOnly
    ] || [];

  if (
    candidates.length === 1
  ) {
    return candidates[0];
  }

  if (
    candidates.length > 1
  ) {
    Logger.log(
      `Ambiguous header "${raw}" -> multiple directory entries have first name "${rawParts.first}".`
    );
  } else {
    Logger.log(
      `No directory match for header "${raw}".`
    );
  }

  return {
    fullName: '',
    email: ''
  };
}


// ---------------------------------
// Name helper
// ---------------------------------

function splitTueSatName_(
  fullName
) {
  const text =
    String(fullName || '')
      .trim();

  if (!text) {
    return {
      first: '',
      last: ''
    };
  }

  const parts =
    text.split(/\s+/);

  if (
    parts.length === 1
  ) {
    return {
      first: parts[0],
      last: ''
    };
  }

  return {
    first: parts[0],
    last:
      parts[
        parts.length - 1
      ]
  };
}


// ---------------------------------
// Column helper
// ---------------------------------

function columnLetterToNumber_(
  columnLetter
) {
  const text =
    String(columnLetter || '')
      .trim()
      .toUpperCase();

  let result = 0;

  for (
    let i = 0;
    i < text.length;
    i++
  ) {
    result =
      result * 26 +
      (
        text.charCodeAt(i) -
        64
      );
  }

  return result;
}