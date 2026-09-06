function testNormalizedObject() {
  const match = createMatchObject({
    group: 'Thursday',
    date: new Date(2026, 8, 10),
    time: '7:00 PM',
    location: 'Longfellow Club',
    court: '5',
    players: [
      'Ivan',
      'Andy',
      'Cenk',
      'Dan'
    ],
    ballPerson: 'Andy',
    availableSubs: [
      'Craig'
    ],
    recipients: [
      'example1@gmail.com',
      'example2@gmail.com'
    ]
  });

  Logger.log(
    JSON.stringify(match, null, 2)
  );
}


function testThursdayReader() {
  const match = readNextThursdayMatch();

  if (!match) {
    Logger.log(
      'No upcoming Thursday match found for Ivan.'
    );
    return;
  }

  Logger.log(
    JSON.stringify(match, null, 2)
  );
}


function testTimeParser() {
  const tests = [
    '7 PM',
    '7:00 PM',
    '7:30 PM',
    '9:30 AM'
  ];

  tests.forEach(time => {
    Logger.log(
      time + ' -> ' +
      JSON.stringify(parseTime(time))
    );
  });
}


function testThursdayCalendar() {
  syncThursdayCalendar();
}

function testTuesdayReader() {
  const match = readNextTuesdayMatch();

  if (!match) {
    Logger.log(
      'No upcoming Tuesday match found for Ivan.'
    );
    return;
  }

  Logger.log(
    JSON.stringify(match, null, 2)
  );
}


function testSaturdayReader() {
  const match = readNextSaturdayMatch();

  if (!match) {
    Logger.log(
      'No upcoming Saturday match found for Ivan.'
    );
    return;
  }

  Logger.log(
    JSON.stringify(match, null, 2)
  );
}

function testTuesdayReminderBuilder() {
  const match = readNextTuesdayMatch();
  const email = buildReminderEmail(match);

  Logger.log(
    'SUBJECT:\n' +
    email.subject +
    '\n\nBODY:\n' +
    email.body
  );
}


function testThursdayReminderBuilder() {
  const match = readNextThursdayMatch();
  const email = buildReminderEmail(match);

  Logger.log(
    'SUBJECT:\n' +
    email.subject +
    '\n\nBODY:\n' +
    email.body
  );
}


function testSaturdayReminderBuilder() {
  const match = readNextSaturdayMatch();
  const email = buildReminderEmail(match);

  Logger.log(
    'SUBJECT:\n' +
    email.subject +
    '\n\nBODY:\n' +
    email.body
  );
}

function testTuesdayDraft() {
  createTuesdayReminderDraft();
}

function testThursdayDraft() {
  createThursdayReminderDraft();
}

function testSaturdayDraft() {
  createSaturdayReminderDraft();
}

function testReminderState() {
  const match = readNextThursdayMatch();

  Logger.log(
    'Key: ' +
    getReminderKey(match)
  );

  Logger.log(
    'Already processed before test: ' +
    wasReminderProcessed(match)
  );

  markReminderProcessed(
    match,
    'TEST'
  );

  Logger.log(
    'Already processed after marking: ' +
    wasReminderProcessed(match)
  );

  Logger.log(
    'Stored state: ' +
    JSON.stringify(
      getReminderState(match),
      null,
      2
    )
  );

  // Clean up after the test so it does not
  // interfere with the real reminder.
  clearReminderState(match);

  Logger.log(
    'Already processed after cleanup: ' +
    wasReminderProcessed(match)
  );
}
function testScheduler() {
  runReminderScheduler();
}

function testDueReminder_NoDuplicate() {
  const match = readNextThursdayMatch();

  // Make sure we're starting clean.
  clearReminderState(match);

  Logger.log('--- FIRST RUN ---');

  processReminderMatchForTest_(match);

  Logger.log(
    'State after first run: ' +
    JSON.stringify(
      getReminderState(match),
      null,
      2
    )
  );

  Logger.log('--- SECOND RUN ---');

  processReminderMatchForTest_(match);

  Logger.log(
    'State after second run: ' +
    JSON.stringify(
      getReminderState(match),
      null,
      2
    )
  );
}


function processReminderMatchForTest_(match) {
  if (!match) {
    throw new Error('Match is required.');
  }

  const state = getReminderState(match);

  if (state) {
    Logger.log(
      match.group +
      ': reminder already processed. ' +
      JSON.stringify(state)
    );

    return;
  }

  createReminderDraft(match);

  if (CONFIG.mode === 'DRAFT') {
    markReminderProcessed(
      match,
      'DRAFT_CREATED'
    );
  } else if (
    CONFIG.mode === 'SEND'
  ) {
    markReminderProcessed(
      match,
      'SENT'
    );
  }

  Logger.log(
    match.group +
    ': reminder processed successfully.'
  );
}

function clearThursdayTestState() {
  const match = readNextThursdayMatch();
  clearReminderState(match);
}