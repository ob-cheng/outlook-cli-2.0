---
name: calendar
description: Manage Outlook calendar. Use when user asks about "calendar", "meetings", "events", "schedule", or "appointments".
---

# Calendar Skill

List, view, create, and delete calendar events in Outlook.

## How to run

```bash
cd "./"
python outlook.py cal <command> [options]
```

## List events

```bash
python outlook.py cal list
python outlook.py cal list --start 2026-05-15 --end 2026-05-20
python outlook.py cal list --subject "standup"
python outlook.py cal list --recurring
```

## View event details

```bash
python outlook.py cal read <event-id>
```

## Create event

```bash
python outlook.py cal create \
  --subject "Team Sync" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 15:00" \
  --location "Room A" \
  --body "Weekly sync"
```

Create without reminder:

```bash
python outlook.py cal create \
  --subject "Focus time" \
  --start "2026-05-15 09:00" \
  --end "2026-05-15 12:00" \
  --no-reminder
```

## Delete event

```bash
python outlook.py cal delete <event-id>
```

## List options

- `--start DATE` - Start date (YYYY-MM-DD, default: today)
- `--end DATE` - End date (YYYY-MM-DD, default: 7 days from start)
- `--subject TEXT` - Filter by subject
- `--location TEXT` - Filter by location
- `--organizer EMAIL` - Filter by organizer
- `--all-day` - All-day events only
- `--recurring` - Recurring events only

## Create options

- `--subject TEXT` - Event subject (required)
- `--start DATETIME` - Start "YYYY-MM-DD HH:MM" (required)
- `--end DATETIME` - End "YYYY-MM-DD HH:MM" (required)
- `--location TEXT` - Location
- `--body TEXT` - Description
- `--required EMAILS` - Required attendees (comma-separated)
- `--optional EMAILS` - Optional attendees (comma-separated)
- `--reminder N` - Reminder minutes (default: 15)
- `--no-reminder` - No reminder