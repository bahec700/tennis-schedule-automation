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

  const requestSubUrl = buildRequestSubUrl_(
    match,
    weekdayLong,
    weekdayShort
  );

  // Ball person first, then everyone else alphabetically.
  const roster = match.players
    .map(name => ({
      name: name,
      bringBalls:
        match.ballPerson &&
        name.toLowerCase() ===
          match.ballPerson.toLowerCase()
    }))
    .sort((a, b) => {
      if (a.bringBalls && !b.bringBalls) return -1;
      if (!a.bringBalls && b.bringBalls) return 1;

      return a.name.localeCompare(b.name);
    });

  const subs = [
    ...(match.availableSubs || [])
  ].sort((a, b) =>
    a.localeCompare(b)
  );

  const unresolved =
    match.unresolvedPlayers || [];

  // Build location/court clause from match configuration.
  const locationClause = match.location
    ? ` at ${match.location}` +
      (
        match.court
          ? `, court ${match.court}`
          : ''
      )
    : '';

  const subject = match.location
    ? `Tennis reminder – ${weekdayShort} @ ${match.location}`
    : `Tennis reminder – ${weekdayShort}`;

  // ---------------------------------
  // Plain-text fallback
  // ---------------------------------

  const rosterText = roster
    .map(p =>
      `- ${p.name}${
        p.bringBalls
          ? ' (balls)'
          : ''
      }`
    )
    .join('\n');

  const subsText = subs.length
    ? '\n\nSubs\n' +
      subs
        .map(s => `- ${s}`)
        .join('\n')
    : '';

  const warningBlock = unresolved.length
    ? `⚠️ Could not find an email address for: ${unresolved.join(', ')}. ` +
      `They will NOT receive this reminder automatically — please notify them directly.`
    : '';

  const requestSubText =
    requestSubUrl
      ? `\n\nRequest a sub:\n${requestSubUrl}`
      : '';

  const body =
`Hello players,

You're scheduled to play ${weekdayLong} at ${match.time}${locationClause}.
${warningBlock}
Main Roster
${rosterText}${subsText}${requestSubText}

View schedule:
${scheduleUrl}`;

  // ---------------------------------
  // Authoritative HTML version
  // ---------------------------------

  const rosterHtml = roster
    .map(p =>
      `<li>${htmlEscape_(p.name)}${
        p.bringBalls
          ? ' <img src="cid:ballIcon" width="16" height="16" style="vertical-align:-2px;" alt="ball">'
          : ''
      }</li>`
    )
    .join('');

  const subsHtml = subs.length
    ? `
      <p style="margin:16px 0 4px 0;">
        <strong>Subs</strong>
      </p>

      <ul style="margin:0 0 8px 24px; padding:0;">
        ${
          subs
            .map(
              s =>
                `<li>${htmlEscape_(s)}</li>`
            )
            .join('')
        }
      </ul>

      ${
        requestSubUrl
          ? `<p style="margin:4px 0 16px 0;">
               <a href="${htmlEscape_(requestSubUrl)}">
                 Request a sub
               </a>
             </p>`
          : ''
      }
    `
    : '';

  const locationClauseHtml = match.location
    ? ` at ${htmlEscape_(match.location)}` +
      (
        match.court
          ? `, court ${htmlEscape_(match.court)}`
          : ''
      )
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
    You're scheduled to play
    ${htmlEscape_(weekdayLong)} at ${htmlEscape_(match.time)}${locationClauseHtml}.
  </p>

  ${warningHtml}

  <p style="margin:16px 0 4px 0;">
    <strong>Main Roster</strong>
  </p>

  <ul style="margin:0 0 12px 24px; padding:0;">
    ${rosterHtml}
  </ul>

  ${subsHtml}

  <p style="margin:8px 0 0 0;">
    <a href="${htmlEscape_(scheduleUrl)}">
      View schedule
    </a>
  </p>
</div>`;

  return {
    subject,
    body,
    htmlBody
  };
}


function buildRequestSubUrl_(
  match,
  weekdayLong,
  weekdayShort
) {
  const recipients = [
    ...new Set(
      (match.availableSubRecipients || [])
        .map(
          email =>
            String(email || '').trim()
        )
        .filter(Boolean)
    )
  ];

  if (recipients.length === 0) {
    return '';
  }

  const locationClause = match.location
    ? ` at ${match.location}` +
      (
        match.court
          ? `, court ${match.court}`
          : ''
      )
    : '';

  const subject =
    `Sub needed – ${weekdayShort} tennis`;

  const body =
`Hi,

Can anyone sub for tennis on ${weekdayLong} at ${match.time}${locationClause}?

Thanks`;

  const params = [
    'subject=' +
      encodeURIComponent(subject),

    'body=' +
      encodeURIComponent(body)
  ];

  if (CONFIG.adminEmail) {
    params.push(
      'bcc=' +
      encodeURIComponent(
        CONFIG.adminEmail
      )
    );
  }

  return (
    'mailto:' +
    recipients
      .map(
        encodeURIComponent
      )
      .join(',') +
    '?' +
    params.join('&')
  );
}


function getScheduleUrl_(match) {
  const cfg =
    CONFIG.groups[match.group];

  if (!cfg) {
    throw new Error(
      'Missing CONFIG for group: ' +
      match.group
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