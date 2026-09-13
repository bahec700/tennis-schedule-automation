function createReminderDraft(match) {
  if (!match) {
    throw new Error('Match is required.');
  }

  if (CONFIG.mode === 'DISABLED') {
    Logger.log('Reminder system is disabled.');
    return null;
  }

  const validationErrors =
    validateReminderMatch_(match);

  if (validationErrors.length > 0) {
    sendReminderValidationAlert_(
      match,
      validationErrors
    );

    throw new Error(
      'Reminder validation failed: ' +
      validationErrors.join(' | ')
    );
  }

  // If this match previously had a validation problem
  // and is now fixed, clear the old alert marker.
  clearValidationAlertState_(match);

  const email = buildReminderEmail(match);

  const to = match.recipients.join(',');

  const options = {
    htmlBody: email.htmlBody,
    inlineImages: {
      ballIcon: getBallIconBlob_()
    }
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



function validateReminderMatch_(match) {
  const errors = [];

  const players =
    match.players || [];

  const ballPeople =
    match.ballPeople || [];

  const recipients =
    match.recipients || [];

  const unresolved =
    match.unresolvedPlayers || [];

  // 1. Exactly four scheduled players.
  if (players.length !== 4) {
    errors.push(
      'Expected exactly 4 scheduled players, found ' +
      players.length +
      (
        players.length
          ? ': ' + players.join(', ')
          : ''
      )
    );
  }

  // 2. Exactly one scheduled player brings balls.
  if (ballPeople.length !== 1) {
    errors.push(
      'Expected exactly 1 player bringing balls, found ' +
      ballPeople.length +
      (
        ballPeople.length
          ? ': ' + ballPeople.join(', ')
          : ''
      )
    );
  }

  // 3. Every scheduled player must have one unique email address.
  const normalizedRecipients =
    recipients
      .map(email =>
        String(email || '')
          .trim()
          .toLowerCase()
      )
      .filter(Boolean);

  const uniqueRecipients = [
    ...new Set(normalizedRecipients)
  ];

  if (
    unresolved.length > 0 ||
    recipients.length !== players.length ||
    uniqueRecipients.length !== players.length
  ) {
    let detail =
      'Expected a valid unique email for every scheduled player. ' +
      'Players: ' + players.length +
      ', recipient emails: ' + recipients.length +
      ', unique recipient emails: ' + uniqueRecipients.length;

    if (unresolved.length > 0) {
      detail +=
        '. Missing/unresolved: ' +
        unresolved.join(', ');
    }

    errors.push(detail);
  }

  return errors;
}


function sendReminderValidationAlert_(
  match,
  errors
) {
  if (!CONFIG.adminEmail) {
    Logger.log(
      'Validation failed, but CONFIG.adminEmail is empty.'
    );
    return;
  }

  const dateText =
    Utilities.formatDate(
      match.date,
      Session.getScriptTimeZone(),
      'EEE MMM d, yyyy'
    );

  const alertKey =
    getValidationAlertKey_(match);

  const signature =
    errors.join(' || ');

  const properties =
    PropertiesService
      .getScriptProperties();

  const previousSignature =
    properties.getProperty(alertKey);

  // Avoid sending the same error email every time
  // the scheduler runs with unchanged bad data.
  if (previousSignature === signature) {
    Logger.log(
      'Validation alert already sent: ' +
      alertKey
    );
    return;
  }

  const subject =
    'Tennis reminder ERROR – ' +
    match.group +
    ' ' +
    dateText;

  const bodyLines = [
    'The tennis reminder was NOT sent because the schedule failed validation.',
    '',
    'Group: ' + match.group,
    'Date: ' + dateText,
    'Time: ' + match.time,
    '',
    'Problem(s):'
  ];

  errors.forEach(error => {
    bodyLines.push(
      '- ' + error
    );
  });

  bodyLines.push(
    '',
    'Scheduled players:'
  );

  if (
    match.players &&
    match.players.length
  ) {
    match.players.forEach(
      player =>
        bodyLines.push(
          '- ' + player
        )
    );
  } else {
    bodyLines.push('- none');
  }

  bodyLines.push(
    '',
    'Players marked for balls:'
  );

  if (
    match.ballPeople &&
    match.ballPeople.length
  ) {
    match.ballPeople.forEach(
      player =>
        bodyLines.push(
          '- ' + player
        )
    );
  } else {
    bodyLines.push('- none');
  }

  bodyLines.push(
    '',
    'Recipient emails:'
  );

  if (
    match.recipients &&
    match.recipients.length
  ) {
    match.recipients.forEach(
      email =>
        bodyLines.push(
          '- ' + email
        )
    );
  } else {
    bodyLines.push('- none');
  }

  bodyLines.push(
    '',
    'Please correct the schedule. The normal reminder can run after the data is fixed.'
  );

  GmailApp.sendEmail(
    CONFIG.adminEmail,
    subject,
    bodyLines.join('\n')
  );

  properties.setProperty(
    alertKey,
    signature
  );

  Logger.log(
    'Validation alert sent to ' +
    CONFIG.adminEmail +
    ': ' +
    subject
  );
}


function getValidationAlertKey_(match) {
  const dateKey =
    Utilities.formatDate(
      match.date,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  return (
    'VALIDATION_ALERT:' +
    match.group +
    ':' +
    dateKey
  );
}


function clearValidationAlertState_(match) {
  PropertiesService
    .getScriptProperties()
    .deleteProperty(
      getValidationAlertKey_(match)
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

function getBallIconBlob_() {
  const base64 =
    'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFbUlEQVR4AdRVa2wUVRT+ZmZ3+wC27bYUqEkhLRQL1AQTQTEaJEQUFbRKTLTKKy5GfPyQqJjUGPAVMUQSQKQYQBujCLVogmJUIBgJ6i+LtNRAgcLSbbfb7bb7nJ0Zz5lOp7vtLtstPwybc+4995zz3fvde8/cFfE//25aAiV0cAtJqw1lm300zEwyOYGqIoe0Kz9PcpdNtTY9vGTCIedKxz5WttnHMc4hClWko5LRECiZMsnSSAsc37Rx0rqW0zOKW/+scDR+UerYuWXKBFa22ccxzuFcxhCDtKeSjkB1nl1srt1QvJwXcK4sAO2Q5k0uHOMczq19tXh5QZ7USpl8TdQll+sReK5smrXutyNldp40OTy117mqAMcPlI6rLLPWiyKcqTJTEaimxT/4/cdyR2VFVipsWv/subk4un1Kzsxp1m2UnPQkkhEooWPfe7h+qqPIIRHuxmTyrFzs3+DILsyX6mmmETUxggAVz873ayfbb2TntJApYq6EyluzUVtjz5leat1nBgxjOIGqnGzhnuveuZAFIWsRhNx1EMe/qSvb7KOAMW1iJ+VbsXrJONgsuJciCZ9oAgE68vUbXpzooKSkIuSuhVTQCMF2H6D5oMmndWWbfXqMcoaDLXkW3eVcOj5rUrH1JX1gNAkEYgoeq37EboTiOukWiHk7IIiFUHyroPbVQgt9DS3yy4CSzT6OcQ7ngjDmDDZBN5fdlQMo6gp9YDSi0XNX4sgXLUXDC0+wQ5ywiRY6AbX/Q0B1c25ypRjnaJETOgaE5UTRNrBMoV2EI0/MJl8JqS4DEd1ExSwqlgFzqBXHrYcWPQUtfGDImcbiXMYwllM1ReNO19lTbfxpVegDauIJOEomW63kM0WwzAQsc6AF95i+0Ro6hrA8hxZVTVhRvsgEzDoTzUgyw3Y37f7nZJFR+bQoYWkONTREYDgwnoDX1SHL8QmCVA7EWuJdmdmE5TkU39C0Hp+i0CReUl3iCbSebQnrTrMRxlPR+c1hxoZKWJpD6YmZ0LOXZD4O/pPSffEEXF6fGvN4maAeAzQiKk40BmPoCKsGujBYhN1+Fd29Ku/SNThbPAFYJHzb8D2xNqJa7BwE6xxjNIaOilDuOGMCvzsVgmQRD5oOMhII0O53fLS9i7ZNERItehKwLQbo+UXGvyxAWIRYx68Y/O0+0h9xueVtg2PuEwiQoykU1k7u3t9DJolyFYgeg5jrpEFmoslrILt+ghps14F7jwYQjGrHaNBEaspwArjmjr2wcXOHv7k1oiepgZ2ApRJCzlP6OG1Db44SoNeWMJHWrXr6uXYZ79T7QxevxNbqjrhGjLMHTVevX129vOZSD10J+WR6+9+GYFtAJJ6mcWpRQwqUvhUQcxYg9PdrgCqDC2/lFm+IPr8aQprFR7YuyQhwoOHiZfmNhQ+eD5z5ox+a7CYSb0HIfpSK8jaOm6qpGmJeGeHWIKKXZ0AqelxfXAu7wTt/qLYr1Nwmv0KABtIRkooAVBW7W9rkZxc/cyW8491rCP/TjtjVz+iTWkYLhRE5H0S4OYDgX35E/g1C6ZFhLX0S0bY90CIe8J0v3egJN1+Qa2jVOtKkkpKAkd3g6VHKtx0OHJ23+mrkk61fojtQAflaBDGPDMVPDwzduZELHypR9/k3uPNld+jjxr4fPL0KPaVIunMYv3QEOM3V1h594OwF+Y73vvLUVd0+z0cLhNdu7ZFf39Orsa4hez4tWjV3nnfzfvendOTzqeCWEnjEnZMvQUZDYBDQ5O6UnZ3dSgEtUH7wWOD+XYf7nmA9RDZd1/ROr1JIhfs8ARI+NRqnlEwIxE/COztODj5eVrbZR67MZKwERr1KusT/AAAA//8Z6SRnAAAABklEQVQDAOu4FV+wMKiOAAAAAElFTkSuQmCC';

  return Utilities.newBlob(
    Utilities.base64Decode(base64),
    'image/png',
    'tennis-ball.png'
  );
}
