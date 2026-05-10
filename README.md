# Outlook CLI 2.0

Stop clicking through Outlook. Use your terminal instead.

This is a command-line tool for Microsoft Outlook that lets you search, view, send emails and manage your calendar without touching the UI. It's built for people who live in their terminal, folks automating email workflows, and AI agents that need to work with email.

**Actually works with:**
- Searching and filtering emails across folders
- Sending emails with attachments and CC/BCC
- Replying and forwarding
- Exporting threads to Obsidian markdown
- Listing and creating calendar events
- Integrating into AI agent workflows (Claude Code, Hermes, OpenClaw)

## Install

```bash
git clone https://github.com/ob-cheng/outlook-cli-2.0.git
cd outlook-cli-2.0
pip install -r requirements.txt
```

## Quick Start

```bash
# List unread emails from this week
python outlook.py search --unread --days 7

# Search for emails from a specific person
python outlook.py search --filter-email boss@company.com --days 30

# Send an email
python outlook.py send \
  --to recipient@example.com \
  --subject "Quick update" \
  --body "Here's what I've done"

# Export threads to markdown for your Obsidian vault
python outlook.py export --output ~/Obsidian/Emails --days 30

# Check calendar for the next week
python outlook.py cal list
```

## Commands

### Search and Display

Find emails and see them in a clean table format:

```bash
# Default: next 7 days from Inbox
python outlook.py search

# Unread only
python outlook.py search --unread

# Search by person
python outlook.py search --filter-email john@company.com

# Search by keyword
python outlook.py search --keyword "meeting notes"

# Combine filters
python outlook.py search --folder Inbox --unread --keyword "urgent" --days 3

# Custom date range
python outlook.py search --from-date 2026-05-01 --to-date 2026-05-31
```

### Export to Markdown

Save emails to markdown files for archival or knowledge bases:

```bash
# Basic export
python outlook.py export --output ./emails --days 30

# With filters
python outlook.py export --output ./project-emails \
  --filter-email client@vendor.com \
  --keyword "contract"

# Multiple folders at once
python outlook.py export --output ./archive \
  --folder Inbox --folder "Sent Items" --folder Archive
```

### Read a Single Email

View full details of one message:

```bash
python outlook.py read <message-id>
```

### Send Email

Compose and send a message:

```bash
# Basic
python outlook.py send \
  --to user@example.com \
  --subject "Your subject here" \
  --body "Message body"

# With CC, BCC, attachments
python outlook.py send \
  --to person1@example.com \
  --cc manager@example.com \
  --bcc archive@example.com \
  --subject "Report" \
  --body "See attached" \
  --attach report.pdf

# Save as draft instead of sending
python outlook.py send \
  --to user@example.com \
  --subject "Draft message" \
  --body "Not ready yet" \
  --draft

# HTML email
python outlook.py send \
  --to user@example.com \
  --subject "Announcement" \
  --body "<h1>Hello</h1><p>New feature live</p>" \
  --html
```

### Reply and Forward

```bash
# Reply to sender
python outlook.py reply <message-id> --body "Thanks for the update"

# Reply to everyone
python outlook.py reply <message-id> --body "Team, here's my take" --all

# Reply with attachment as draft
python outlook.py reply <message-id> \
  --body "See attached" \
  --attach response.pdf \
  --draft

# Forward to someone
python outlook.py forward <message-id> \
  --to colleague@example.com \
  --body "FYI on this"
```

### Calendar

List events with filters, create new ones, delete old ones:

```bash
# What's coming up
python outlook.py cal list

# Next month
python outlook.py cal list --start 2026-05-01 --end 2026-05-31

# Find all "standup" meetings
python outlook.py cal list --subject "standup"

# Recurring events only
python outlook.py cal list --recurring

# View event details
python outlook.py cal read <event-id>

# Create a meeting
python outlook.py cal create \
  --subject "Quick sync" \
  --start "2026-05-12 14:00" \
  --end "2026-05-12 15:00" \
  --location "Conference Room A" \
  --body "Quarterly planning"

# Delete an event
python outlook.py cal delete <event-id>
```

### Folders

See what's available to search:

```bash
python outlook.py folders
```

This lists your Inbox, Sent Items, Archive, custom folders, everything.

## Output Format

When you export, files look like this:

```markdown
---
title: "Project Update"
date: 2026-05-08
date_end: 2026-05-10
message_count: 5
participants:
  - "Alice Smith"
  - "Bob Jones"
tags:
  - email/thread
  - email/conversation
---

> [!info] Thread Summary
> **5 message(s)** from [[2026-05-08]] to [[2026-05-10]]
> **Participants:** Alice Smith, Bob Jones

## Message 1 (RECEIVED) ^msg-1

**From:** Alice Smith  
**Date:** [[2026-05-08]] 09:15

Here's the actual email body, stripped of noise, quoted replies, and signatures.

---

## Message 2 (SENT) ^msg-2

**From:** You  
**Date:** [[2026-05-08]] 14:30

Your response.
```

Threads are grouped by subject (RE:/FW: headers stripped), and each message is marked as sent or received. Perfect for dropping into Obsidian.

## Using with AI

### Claude Code

Add this to your Claude Code commands:

```bash
# Get context before responding
python outlook.py search --keyword "ProjectX" --days 7 --export ./context

# Check what's urgent
python outlook.py search --unread

# Send a follow-up
python outlook.py send --to client@company.com --subject "Next steps" --body "..."
```

Claude can search your emails, read threads, and draft responses. Just tell it to run these commands.

### Hermes Agent

Configure Hermes to use Outlook as a tool:

```json
{
  "tools": [
    {
      "name": "search_inbox",
      "command": "python outlook.py search --keyword {query} --days {days}",
      "description": "Find emails by keyword"
    },
    {
      "name": "send_email",
      "command": "python outlook.py send --to {recipient} --subject {subject} --body {body}",
      "description": "Send an email"
    },
    {
      "name": "check_calendar",
      "command": "python outlook.py cal list --start {date1} --end {date2}",
      "description": "List calendar events"
    }
  ]
}
```

Your agent can now check your calendar before scheduling, search for context, and send reports.

### OpenClaw

Set up recurring tasks:

```yaml
# .openclaw/config.yml
actions:
  - name: check_inbox_every_hour
    command: python outlook.py search --folder Inbox --unread
    schedule: "0 * * * *"

  - name: daily_email_archive
    command: python outlook.py export --output ./archive --days 1
    schedule: "0 20 * * *"

  - name: send_eod_reminder
    command: |
      python outlook.py send \
        --to team@company.com \
        --subject "EOD reminder" \
        --body "Submit your status updates"
    schedule: "0 17 * * 1-5"
```

## What Gets Cleaned

When exporting to markdown, we remove the noise:

**Gone:** SafeLinks tracking URLs, signature tables, "CAUTION: external email" warnings, quoted replies, tracking pixels, empty cells, random HTML junk.

**Stays:** The actual email content, sender names in signatures, legitimate tables.

## What You Need

- Windows 10 or 11
- Outlook Classic (the desktop app, not "new Outlook")
- Python 3.8 or higher

That's it. No internet required. Works offline.

## Coming From 1.0?

The old `outlook_to_markdown.py` script is still there, but this is better. Here's the translation:

**Old:** `python outlook_to_markdown.py --full --days 30 --output ./emails`  
**New:** `python outlook.py export --output ./emails --days 30`

See CHANGELOG.md for what changed.

## License

MIT
