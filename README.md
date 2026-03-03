# SlackTrack Google Script

Google Apps Script backend for SlackTrack, a Slack-based attendance bot that writes attendance records to Google Sheets.

## Purpose

- Send attendance reminders in Slack (morning and evening).
- Let users update attendance using buttons (`WFO`, `WFH`, `Leave (-1)`, `Half Day (-0.5)`).
- Auto-create monthly sheets when missing.
- Keep reminders and updates synchronized with a Google Sheet attendance ledger.
- Parse chat messages using Gemini to support natural language updates (for example, leave date ranges).

## How It Works

### 1) Slack Webhook Entry (`doPost`)

The script receives Slack requests through a deployed Apps Script Web App URL.

- Handles Slack URL verification (`url_verification`) for Events API setup.
- Handles interactive button payloads from Slack messages.
- Handles optional slash command (`/attendance`) to send update buttons in DM.
- Quickly acknowledges Events API callbacks and queues chat events to avoid Slack timeout limits.

### 2) Attendance Update Flow

- Maps Slack button action IDs to attendance values.
- Resolves or creates the correct month tab in the spreadsheet.
- Finds (or creates) the user row by `SlackUserID`.
- Finds today's date column.
- Writes the attendance value and updates `Last Updated at`.

### 3) Monthly Sheet Management

- Uses `getOrCreateMonthSheet_` to fetch the current month tab (`MMMM yyyy`).
- If missing, clones the latest month template when available.
- If no template exists, initializes a new month sheet with base columns and day columns.

### 4) Reminders

- `sendMorningReminders`: sends attendance buttons to configured Slack users.
- `sendEveningReminders`: sends reminders only to users who have not yet updated attendance for the day.

### 5) Chat Intent Parsing (Gemini)

- DM messages are queued in script properties.
- `processQueuedChatEvents` consumes queued messages on a time trigger.
- Gemini extracts intent/date/value from natural language.
- Script updates one or more dates in the sheet and sends confirmation back in Slack.

## Required Script Properties

- `SLACK_BOT_TOKEN`
- `SPREADSHEET_ID`
- `TIMEZONE`
- `SLACK_USER_IDS` (comma-separated Slack user IDs)
- `GEMINI_API_KEY` (for chat parsing)
- `SLACK_TEAM_ID` (optional safety check)

## Required Triggers

- Time-driven trigger for `sendMorningReminders`
- Time-driven trigger for `sendEveningReminders`
- Time-driven trigger for `processQueuedChatEvents`

## Notes

- Secrets are read from Script Properties and should not be hardcoded.
- Dedupe uses Apps Script cache to reduce duplicate event processing.
- Slack App Home message input can be restricted by workspace policy; slash commands and button flows can still be used.
