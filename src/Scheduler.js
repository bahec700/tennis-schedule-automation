function runReminderScheduler() {
  Logger.log('Reminder scheduler started.');

  const groups = [
    ['Tuesday', readNextTuesdayMatch],
    ['Thursday', readNextThursdayMatch],
    ['Saturday', readNextSaturdayMatch]
  ];

  // Each group is isolated: a failure in one (e.g. a sheet layout
  // problem or a missing recipient) should not prevent the others
  // from being processed.
  groups.forEach(([groupName, readerFunction]) => {
    try {
      processReminderGroup_(groupName, readerFunction);
    } catch (error) {
      Logger.log(
        groupName +
        ': FAILED, continuing with remaining groups - ' +
        error.message
      );
    }
  });

  Logger.log('Reminder scheduler completed.');
}


function processReminderGroup_(
  groupName,
  readerFunction
) {
  const cfg = CONFIG.groups[groupName];

  if (!cfg) {
    throw new Error(
      'Missing CONFIG for group: ' +
      groupName
    );
  }

  if (!cfg.enabled) {
    Logger.log(
      groupName +
      ': skipped because group is disabled.'
    );
    return;
  }

  if (CONFIG.mode === 'DISABLED') {
    Logger.log(
      groupName +
      ': skipped because reminder system is disabled.'
    );
    return;
  }

  const match = readerFunction();

  if (!match) {
    Logger.log(
      groupName +
      ': no upcoming session found.'
    );
    return;
  }

  const daysAway =
    getDaysUntilMatch_(match.date);

  Logger.log(
    groupName +
    ': next session = ' +
    formatDateForLog_(match.date) +
    ', days away = ' +
    daysAway
  );

  // Act on or within CONFIG.reminderLeadDays of the session (not only
  // on the exact day), so a reminder still goes out even if the
  // scheduler didn't run on the precise lead day (e.g. a missed
  // trigger). The state check below still prevents duplicate sends.
  if (
    daysAway < 0 ||
    daysAway > CONFIG.reminderLeadDays
  ) {
    Logger.log(
      groupName +
      ': no reminder due today.'
    );
    return;
  }

  const state =
    getReminderState(match);

  if (state) {
    Logger.log(
      groupName +
      ': reminder already processed. ' +
      JSON.stringify(state)
    );
    return;
  }

  try {
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
      groupName +
      ': reminder processed successfully.'
    );

  } catch (error) {
    Logger.log(
      groupName +
      ': ERROR - ' +
      error.message
    );

    throw error;
  }
}


function getDaysUntilMatch_(matchDate) {
  const today = new Date();

  today.setHours(
    0,
    0,
    0,
    0
  );

  const target = new Date(
    matchDate
  );

  target.setHours(
    0,
    0,
    0,
    0
  );

  const millisecondsPerDay =
    24 * 60 * 60 * 1000;

  return Math.round(
    (
      target.getTime() -
      today.getTime()
    ) /
    millisecondsPerDay
  );
}


function formatDateForLog_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}