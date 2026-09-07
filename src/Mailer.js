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
