function createReminderDraft(match) {
  if (!match) {
    throw new Error('Match is required.');
  }

  if (CONFIG.mode === 'DISABLED') {
    Logger.log('Reminder system is disabled.');
    return null;
  }

  const email = buildReminderEmail(match);

  if (!match.recipients || match.recipients.length === 0) {
    throw new Error(
      'No recipients found for ' +
      match.group +
      ' ' +
      Utilities.formatDate(
        match.date,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      )
    );
  }

  const to = match.recipients.join(',');

  const options = {
    htmlBody: email.htmlBody
  };

  // Send a hidden copy to the administrator.
  if (CONFIG.adminEmail) {
    options.bcc = CONFIG.adminEmail;
  }

  if (CONFIG.mode === 'DRAFT') {
    const draft = GmailApp.createDraft(
      to,
      email.subject,
      email.body,
      options
    );

    Logger.log(
      'Draft created: ' +
      email.subject +
      ' -> ' +
      to +
      (CONFIG.adminEmail
        ? ' | BCC: ' + CONFIG.adminEmail
        : '')
    );

    return draft;
  }

  if (CONFIG.mode === 'SEND') {
    GmailApp.sendEmail(
      to,
      email.subject,
      email.body,
      options
    );

    Logger.log(
      'Email sent: ' +
      email.subject +
      ' -> ' +
      to +
      (CONFIG.adminEmail
        ? ' | BCC: ' + CONFIG.adminEmail
        : '')
    );

    return null;
  }

  throw new Error(
    'Unknown CONFIG.mode: ' +
    CONFIG.mode
  );
}


function createTuesdayReminderDraft() {
  return createReminderDraft(
    readNextTuesdayMatch()
  );
}


function createThursdayReminderDraft() {
  return createReminderDraft(
    readNextThursdayMatch()
  );
}


function createSaturdayReminderDraft() {
  return createReminderDraft(
    readNextSaturdayMatch()
  );
}