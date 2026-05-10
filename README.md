# Outlook CLI 2.0

A command line interface for Outlook that works through COM automation instead of the Graph API. You can search emails, send messages, and manage calendar events from your terminal.

## Why this exists

The other [outlook-cli](https://github.com/mhattingpete/outlook-cli) requires Azure app registration, OAuth setup, and admin approval. At some companies, that takes weeks. At others, it never happens.

This tool talks directly to Outlook on your Windows machine. If you can open Outlook, you can use this. Install it and start working.

## What it does

- Search and filter emails by folder, sender, date range, read status
- Send emails with CC, BCC, attachments, HTML formatting
- Reply and forward messages
- Save drafts for review before sending
- Export email threads to Obsidian markdown
- List, create, and delete calendar events
- Works offline since everything is local

## Install

```bash
git clone https://github.com/ob-cheng/outlook-cli-2.0.git
cd outlook-cli-2.0
pip install -r requirements.txt
```

You need Windows 10 or 11 with Outlook desktop installed. Python 3.8 or higher.

## Quick start

```bash
# See unread messages
python outlook.py search --unread

# What's on your calendar this week
python outlook.py cal list

# Send an email
python outlook.py send \
  --to someone@company.com \
  --subject "Quick question" \
  --body "Can we sync tomorrow?"

# Save as draft instead of sending
python outlook.py send \
  --to client@company.com \
  --subject "Project update" \
  --body "Here's where we stand..." \
  --draft

# Export recent threads to markdown
python outlook.py export --output ~/Obsidian/Emails --days 30
```

## Commands

### Search

```bash
# Default shows last 7 days
python outlook.py search

# Unread only
python outlook.py search --unread

# From specific person
python outlook.py search --filter-email john@company.com

# With keyword
python outlook.py search --keyword "meeting" --days 14

# Combine filters
python outlook.py search \
  --folder Inbox \
  --unread \
  --filter-email boss@company.com \
  --keyword "urgent" \
  --days 7

# Date range
python outlook.py search --from-date 2026-05-01 --to-date 2026-05-31
```

### Export to markdown

```bash
# Basic export
python outlook.py export --output ./emails --days 30

# With filters
python outlook.py export --output ./project \
  --filter-email client@vendor.com \
  --keyword "contract" \
  --days 90

# Multiple folders
python outlook.py export --output ./all \
  --folder Inbox --folder "Sent Items" --folder Archive
```

### Send email

```bash
# Basic
python outlook.py send \
  --to user@example.com \
  --subject "Update" \
  --body "Here's the status"

# With CC, BCC, attachments
python outlook.py send \
  --to person1@example.com \
  --cc manager@example.com \
  --bcc archive@example.com \
  --subject "Report" \
  --body "See attached" \
  --attach report.pdf \
  --attach data.xlsx

# HTML body
python outlook.py send \
  --to user@example.com \
  --subject "Announcement" \
  --body "<h1>New feature</h1><p>Live now</p>" \
  --html

# Draft mode
python outlook.py send \
  --to user@example.com \
  --subject "Draft message" \
  --body "Draft me first" \
  --draft
```

### Reply and forward

```bash
# Reply to sender
python outlook.py reply <message-id> --body "Thanks!"

# Reply all
python outlook.py reply <message-id> --body "Team, here's my take" --all

# Reply with attachment
python outlook.py reply <message-id> \
  --body "See attached analysis" \
  --attach analysis.pdf

# Draft reply
python outlook.py reply <message-id> \
  --body "My response" \
  --draft

# Forward
python outlook.py forward <message-id> \
  --to colleague@example.com \
  --body "FYI"
```

### Calendar

```bash
# List upcoming events
python outlook.py cal list

# Date range
python outlook.py cal list --start 2026-05-01 --end 2026-05-31

# Filter by subject
python outlook.py cal list --subject "standup"

# Show recurring events
python outlook.py cal list --recurring

# View event details
python outlook.py cal read <event-id>

# Create event
python outlook.py cal create \
  --subject "Sync" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 15:00" \
  --location "Room A" \
  --body "Quarterly planning"

# Delete event
python outlook.py cal delete <event-id>
```

### Folders

```bash
python outlook.py folders
```

Lists all folders including Inbox, Sent Items, Archive, and any custom folders.

## Export format

The export produces markdown files with Obsidian frontmatter:

```markdown
---
title: "Project Update"
date: 2026-05-08
message_count: 3
participants:
  - "Alice Smith"
  - "Bob Jones"
tags:
  - email/thread
  - email/conversation
---

> [!info] Thread Summary
> **3 message(s)** from [[2026-05-08]] to [[2026-05-10]]
> **Participants:** Alice Smith, Bob Jones

## Message 1 (RECEIVED) ^msg-1

**From:** Alice Smith  
**Date:** [[2026-05-08]] 09:15

Email content here. Quoted replies and signatures are stripped.

---

## Message 2 (SENT) ^msg-2

Your reply here.
```

Threads are grouped by subject. Each message is marked as sent or received.

## Using with agents

### Claude Code

```bash
# Get context before responding
python outlook.py search --keyword "ProjectX" --days 7

# Send a follow-up
python outlook.py send \
  --to client@company.com \
  --subject "Next steps" \
  --body "$(cat next-steps.txt)"

# Check calendar before proposing time
python outlook.py cal list --start 2026-05-15 --end 2026-05-20
```

### Hermes Agent

```json
{
  "tools": [
    {
      "name": "search_email",
      "command": "python outlook.py search --keyword {query} --days {days}",
      "description": "Find emails"
    },
    {
      "name": "send_email",
      "command": "python outlook.py send --to {recipient} --subject {subject} --body {body}",
      "description": "Send email"
    },
    {
      "name": "check_calendar",
      "command": "python outlook.py cal list --start {date1} --end {date2}",
      "description": "List calendar events"
    }
  ]
}
```

### OpenClaw

```yaml
actions:
  - name: monitor_inbox
    command: python outlook.py search --unread --folder Inbox
    schedule: "0 * * * *"

  - name: daily_archive
    command: python outlook.py export --output ./archive --days 1
    schedule: "0 20 * * *"

  - name: send_eod_reminder
    command: |
      python outlook.py send \
        --to team@company.com \
        --subject "EOD: Submit status" \
        --body "End-of-day reminder"
    schedule: "0 17 * * 1-5"
```

## What gets cleaned during export

The export strips tracking URLs, signature tables, "CAUTION: external email" warnings, quoted reply chains, and tracking pixels. Names in signatures and legitimate tables are preserved.

## Limitations

- Windows only (COM automation)
- Needs Outlook desktop app installed
- No shared mailbox support yet

## License

MIT

[https://github.com/ob-cheng/outlook-cli-2.0](https://github.com/ob-cheng/outlook-cli-2.0)
