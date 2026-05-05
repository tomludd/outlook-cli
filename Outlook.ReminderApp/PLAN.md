# Outlook.ReminderApp Plan

## Goal
Build a small Windows infobox app that stays near the taskbar bottom-left, shows upcoming/ongoing meetings with countdowns, supports dismiss/join actions, and auto-opens eligible Teams meetings at meeting start.

## Functional scope
- Show meetings that are ongoing or starting within 5 minutes.
- Show overlapping meetings together in one infobox list.
- Highlight ongoing meetings clearly.
- Show meeting name, meeting room, and countdown.
- Countdown text is red when less than 5 minutes remain.
- Dismiss action prevents auto-open for that meeting.
- Declined meetings must never auto-open.
- Join button is available for meetings with Teams link.
- Eligible Teams meetings auto-open once at meeting start if not dismissed.

## UI behavior
- Borderless always-on-top infobox anchored bottom-left above the taskbar.
- Infobox grows upward as more meetings are listed.
- Internal scrolling appears after a maximum height.

## Technical approach
- Reuse `Outlook.COM.OutlookCalendarService` for event retrieval.
- Poll Outlook on an interval and update countdown every second.
- Resolve Teams URLs from event body/location via regex.
- Track per-meeting state (dismissed/opened) in memory.

## Validation
- Verify declined or dismissed meetings do not auto-open.
- Verify overlapping and ongoing meetings are listed.
- Verify Join button only appears for Teams meetings.
- Verify auto-open triggers once at start for eligible meetings.
