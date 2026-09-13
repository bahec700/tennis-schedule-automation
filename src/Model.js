function createMatchObject({
  group,
  date,
  time,
  location = '',
  court = '',
  players = [],
  ballPerson = '',
  ballPeople = [],
  availableSubs = [],
  availableSubRecipients = [],
  recipients = [],
  unresolvedPlayers = []
}) {
  return {
    group,
    date,
    time,
    location,
    court,
    players,
    ballPerson,
    ballPeople,
    availableSubs,
    availableSubRecipients,
    recipients,
    unresolvedPlayers
  };
}


function getLastName(name) {
  const parts = String(name || '')
    .trim()
    .split(/\s+/);

  return parts.length ? parts[parts.length - 1] : '';
}


function isWhiteBackground(color) {
  const c = String(color || '').toLowerCase();

  return (
    c === '#ffffff' ||
    c === '#fff' ||
    c === 'white'
  );
}