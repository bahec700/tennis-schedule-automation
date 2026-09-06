const CONFIG = {
  mode: 'DRAFT', // DRAFT | SEND | DISABLED

  adminEmail: 'bahec700@gmail.com',

  reminderLeadDays: 3,

  calendar: {
    enabled: true,
    id: 'primary'
  },

  groups: {
    Tuesday: {
      spreadsheetId: '12oZlgXc5SZuxytPnUX6jeXjGan0UvHM0SqjZ0dgAfao',
      sheetName: 'Tuesday',
      reader: 'TUE_SAT',
      dayName: 'Tuesday',
      time: '7:30 PM',
      durationMinutes: 90,
      location: 'Longfellow Club',
      court: '1',
      enabled: true
    },

    Thursday: {
      spreadsheetId: '1SpJCM26QPpT5NU8rNRPm33TA4mN_LRQYnDP2BJ4lI2o',
      sheetName: 'Sheet1',
      reader: 'THURSDAY',
      dayName: 'Thursday',
      time: '7:00 PM',
      durationMinutes: 90,
      location: 'Longfellow Club',
      court: '5',
      enabled: true
    },

    Saturday: {
      spreadsheetId: '12oZlgXc5SZuxytPnUX6jeXjGan0UvHM0SqjZ0dgAfao',
      sheetName: 'Saturday',
      reader: 'TUE_SAT',
      dayName: 'Saturday',
      time: '9:30 AM',
      durationMinutes: 90,
      location: 'Longfellow Club',
      court: '1',
      enabled: true
    }
  }
};