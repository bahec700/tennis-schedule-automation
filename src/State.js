function getReminderKey(match) {
  if (!match) {
    throw new Error('Match is required.');
  }

  const dateKey = Utilities.formatDate(
    match.date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  return `REMINDER:${match.group}:${dateKey}`;
}


function wasReminderProcessed(match) {
  const key = getReminderKey(match);

  const value = PropertiesService
    .getScriptProperties()
    .getProperty(key);

  return Boolean(value);
}


function markReminderProcessed(match, status) {
  const key = getReminderKey(match);

  const value = JSON.stringify({
    status: status || 'processed',
    timestamp: new Date().toISOString()
  });

  PropertiesService
    .getScriptProperties()
    .setProperty(key, value);

  Logger.log(
    'Reminder state saved: ' +
    key +
    ' -> ' +
    value
  );
}


function getReminderState(match) {
  const key = getReminderKey(match);

  const value = PropertiesService
    .getScriptProperties()
    .getProperty(key);

  if (!value) {
    return null;
  }

  try {
    return JSON.parse(value);
  } catch (e) {
    return {
      status: value
    };
  }
}


function clearReminderState(match) {
  const key = getReminderKey(match);

  PropertiesService
    .getScriptProperties()
    .deleteProperty(key);

  Logger.log(
    'Reminder state cleared: ' +
    key
  );
}
