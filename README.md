# Tennis Schedule Automation

Google Apps Script used for Longfellow Club tennis schedule reminders and Ivan's Google Calendar sync.

## Files

- `tue-sat-2026-27.gs` — Tuesday/Saturday reminder drafts and calendar sync
- `thursday-2026-27.gs` — Thursday reminder drafts and calendar sync

## 2026–27 source sheets

- Tuesday/Saturday workbook: `2026-27🎾 Tues/Sat Tennis Schedule`
- Thursday workbook: `2026-2027 Thursday TENNIS Sch`

## Main functions

### Tuesday / Saturday

- `CreateReminderTuesday()`
- `CreateReminderSaturday()`
- `createIvanAllUpcomingEvents()`

### Thursday

- `createReminderDraftForNextDate_All()`
- `syncMyMatchesInColumnF()`

## Next-season checklist

Before reusing these scripts for a new season, compare the new Google Sheet layout and update as needed:

- spreadsheet URL
- sheet/tab names
- header and data rows
- date column(s)
- player start/end columns
- player directory name/email ranges
- Ivan's column and calendar event-ID column
- start time and court if changed

These 2026–27 versions were checked against the live schedules on September 6, 2026.
