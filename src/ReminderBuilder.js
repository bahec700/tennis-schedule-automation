function buildReminderEmail(match) {
  if (!match) {
    throw new Error('Match is required.');
  }

  const weekdayLong = Utilities.formatDate(
    match.date,
    Session.getScriptTimeZone(),
    'EEEE, MMMM d'
  );

  const weekdayShort = Utilities.formatDate(
    match.date,
    Session.getScriptTimeZone(),
    'EEE MMM d'
  );

  const scheduleUrl = getScheduleUrl_(match);

  // Match last year's ordering:
  // ball person first, then everyone else alphabetically
  const roster = match.players
    .map(name => ({
      name: name,
      bringBalls:
        match.ballPerson &&
        name.toLowerCase() === match.ballPerson.toLowerCase()
    }))
    .sort((a, b) => {
      if (a.bringBalls && !b.bringBalls) return -1;
      if (!a.bringBalls && b.bringBalls) return 1;

      return a.name.localeCompare(b.name);
    });

  const subs = [...(match.availableSubs || [])]
    .sort((a, b) => a.localeCompare(b));

  const unresolved = match.unresolvedPlayers || [];

  // Build the location/court clause from the match itself instead of
  // assuming every group plays at the same place.
  const locationClause = match.location
    ? ` at ${match.location}` +
      (match.court ? ` on court ${match.court}` : '')
    : '';

  const subject = match.location
    ? `Tennis reminder – ${weekdayShort} @ ${match.location}`
    : `Tennis reminder – ${weekdayShort}`;

  // ---------------------------------
  // Plain-text fallback
  // ---------------------------------

  const rosterText = roster
    .map(p =>
      `- ${p.name}${p.bringBalls ? ' (balls)' : ''}`
    )
    .join('\n');

  const subsText = subs.length
    ? '\n\nAvailable Subs\n' +
      subs.map(s => `- ${s}`).join('\n')
    : '';

  const warningBlock = unresolved.length
    ? `⚠️ Could not find an email address for: ${unresolved.join(', ')}. ` +
      `They will NOT receive this reminder automatically — please notify them directly.`
    : '';

  const body =
`Hello players,

This is a reminder that you are scheduled to play on ${weekdayLong} at ${match.time}${locationClause}.
${warningBlock}
Main Roster
${rosterText}${subsText}

Have a great match!

Click here for the full schedule:
${scheduleUrl}`;

  // ---------------------------------
  // Authoritative HTML version
  // ---------------------------------

  const rosterHtml = roster
    .map(p =>
      `<li>${htmlEscape_(p.name)}${p.bringBalls ? ' (balls)' : ''}</li>`
    )
    .join('');

  const subsHtml = subs.length
    ? `
      <p><strong>Available Subs</strong></p>
      <ul>
        ${subs.map(s => `<li>${htmlEscape_(s)}</li>`).join('')}
      </ul>
    `
    : '';

  const locationClauseHtml = match.location
    ? ` at ${htmlEscape_(match.location)}` +
      (match.court ? ` on court ${htmlEscape_(match.court)}` : '')
    : '';

  const warningHtml = unresolved.length
    ? `<p><strong>⚠️ Could not find an email for:</strong> ` +
      `${htmlEscape_(unresolved.join(', '))}. They will NOT receive this ` +
      `reminder automatically — please notify them directly.</p>`
    : '';

  const htmlBody =
`<div>
  <p>Hello players,</p>

  <p>
    This is a reminder that you are scheduled to play on
    ${htmlEscape_(weekdayLong)} at ${htmlEscape_(match.time)}${locationClauseHtml}.
  </p>

  ${warningHtml}

  <p><strong>Main Roster</strong></p>

  <ul>
    ${rosterHtml}
  </ul>

  ${subsHtml}

  <p>Have a great match!</p>

  <p>
    <a href="${htmlEscape_(scheduleUrl)}">
      Click here for the full schedule
    </a>
  </p>
</div>`;

  return {
    subject,
    body,
    htmlBody
  };
}


function getScheduleUrl_(match) {
  const cfg = CONFIG.groups[match.group];

  if (!cfg) {
    throw new Error(
      'Missing CONFIG for group: ' + match.group
    );
  }

  const ss = SpreadsheetApp.openById(
    cfg.spreadsheetId
  );

  const sheet = ss.getSheetByName(
    cfg.sheetName
  );

  if (!sheet) {
    throw new Error(
      'Sheet not found for schedule URL: ' +
      cfg.sheetName
    );
  }

  return (
    'https://docs.google.com/spreadsheets/d/' +
    cfg.spreadsheetId +
    '/edit#gid=' +
    sheet.getSheetId()
  );
}


function htmlEscape_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}