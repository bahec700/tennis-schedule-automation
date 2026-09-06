function syncThursdayCalendar() {
  if (!CONFIG.calendar.enabled) {
    Logger.log('Calendar sync is disabled.');
    return;
  }

  const match = readNextThursdayMatch();

  if (!match) {
    Logger.log('No upcoming Thursday match found for Ivan.');
    return;
  }

  syncMatchToCalendar(match);
}


function syncMatchToCalendar(match) {
  const cfg = CONFIG.groups[match.group];

  const calendar = CalendarApp.getCalendarById(
    CONFIG.calendar.id
  );

  if (!calendar) {
    throw new Error(
      'Calendar not found: ' + CONFIG.calendar.id
    );
  }

  const { start, end } = buildStartEnd(
    match.date,
    match.time,
    cfg.durationMinutes
  );

  const signature = buildCalendarSignature(match);

  const title = buildCalendarEventTitle(match);

  const descriptionLines = [
    signature,
    '',
    'Players: ' + match.players.join(', ')
  ];

  if (match.ballPerson) {
    descriptionLines.push(
      'Balls: ' + match.ballPerson
    );
  }

  const description = descriptionLines.join('\n');

  // Look around the match date for an existing event
  const searchStart = new Date(start);
  searchStart.setHours(0, 0, 0, 0);

  const searchEnd = new Date(searchStart);
  searchEnd.setDate(searchEnd.getDate() + 1);

  const events = calendar.getEvents(
    searchStart,
    searchEnd
  );

  const existing = events.find(event =>
    String(event.getDescription() || '')
      .includes(signature)
  );

  if (existing) {
    existing.setTitle(title);
    existing.setTime(start, end);
    existing.setDescription(description);

    if (match.location) {
      existing.setLocation(match.location);
    }

    Logger.log(
      'Updated calendar event: ' + title
    );

    return;
  }

  const options = {
    description: description
  };

  if (match.location) {
    options.location = match.location;
  }

  calendar.createEvent(
    title,
    start,
    end,
    options
  );

  Logger.log(
    'Created calendar event: ' + title
  );
}


function buildCalendarEventTitle(match) {
  return 'Tennis - ' + match.group;
}


function buildCalendarSignature(match) {
  const dateKey = Utilities.formatDate(
    match.date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  return `Ref: [Ivan:${match.group}:${dateKey}]`;
}


function buildStartEnd(date, timeText, durationMinutes) {
  const start = new Date(date);

  const parsed = parseTime(timeText);

  start.setHours(
    parsed.hours,
    parsed.minutes,
    0,
    0
  );

  const end = new Date(
    start.getTime() +
    durationMinutes * 60 * 1000
  );

  return {
    start,
    end
  };
}


function parseTime(timeText) {
  const text = String(timeText || '')
    .trim()
    .toUpperCase();

  const match = text.match(
    /^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?$/
  );

  if (!match) {
    throw new Error(
      'Unable to parse time: ' + timeText
    );
  }

  let hours = Number(match[1]);
  const minutes = Number(match[2] || 0);
  const ampm = match[3];

  if (minutes < 0 || minutes > 59) {
    throw new Error(
      'Invalid minutes in time: ' + timeText
    );
  }

  if (ampm) {
    if (hours < 1 || hours > 12) {
      throw new Error(
        'Invalid hour in time: ' + timeText
      );
    }

    if (ampm === 'AM' && hours === 12) {
      hours = 0;
    }

    if (ampm === 'PM' && hours !== 12) {
      hours += 12;
    }
  } else {
    if (hours < 0 || hours > 23) {
      throw new Error(
        'Invalid hour in time: ' + timeText
      );
    }
  }

  return {
    hours,
    minutes
  };
}