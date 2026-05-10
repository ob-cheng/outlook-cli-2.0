# Outlook CLI 2.0

> The complete command-line interface for Microsoft Outlook — search, view, export, send emails and manage calendar events, all from your terminal.

A next-generation CLI that gives you full control over Outlook without touching the UI. Search and filter emails with rich terminal tables, export to Obsidian-flavored markdown, compose and send messages with attachments, manage your calendar, and integrate seamlessly with AI agents like Claude Code, Hermes, and OpenClaw.

**What's New in 2.0:**

- 🎨 Rich terminal UI with styled tables and panels
- ✉️ Full email operations: send, reply, forward, draft
- 📅 Complete calendar management
- 🔍 Advanced filtering: folders, dates, unread, keywords
- 🤖 AI agent integration ready
- 🏗️ Modular service architecture

## Highlights

- **Rich Terminal UI** — Styled tables, panels, and formatted output via Rich library
- **Full Email Operations** — Search, read, send, reply, forward, draft
- **Calendar Management** — List, create, read, and delete calendar events
- **Thread Consolidation** — Groups RE:/FW: chains into single markdown files
- **Smart Filtering** — By person, domain, keyword, date range, unread status, folder
- **Obsidian Export** — YAML frontmatter, wikilinks, callouts, and tags
- **Multiple Folders** — Search across Inbox, Sent, Archive, or custom folders

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Search emails in terminal
python outlook.py search --folder Inbox --days 7

# Export to Obsidian vault
python outlook.py export --output "C:\Users\You\Obsidian\Emails" --days 30

# Send an email
python outlook.py send --to user@example.com --subject "Hello" --body "Test message"

# List upcoming calendar events
python outlook.py cal list
```

## Commands

### Email Operations

#### Search & View

Display emails in a rich terminal table:

```bash
# Recent emails from Inbox
python outlook.py search --folder Inbox --days 7

# Unread messages only
python outlook.py search --unread

# Filter by keyword
python outlook.py search --keyword "project update"

# Filter by sender
python outlook.py search --filter-email boss@company.com

# Combine filters
python outlook.py search --folder Inbox --unread --keyword "urgent" --days 3

# Search and export to markdown
python outlook.py search --folder Inbox --export ./markdown
```

#### Export to Markdown

Export emails to Obsidian-flavored markdown:

```bash
# Basic export
python outlook.py export --output ./emails --days 30

# With filters
python outlook.py export --output ./emails \
  --filter-email client@vendor.com \
  --keyword "contract" \
  --days 90

# Date range export
python outlook.py export --output ./archive \
  --from-date 2026-01-01 --to-date 2026-03-31

# Multiple folders
python outlook.py export --output ./all-mail \
  --folder Inbox --folder "Sent Items" --folder Archive
```

#### Read Single Email

View email details in the terminal:

```bash
python outlook.py read <message-id>
```

#### Send Email

Compose and send new emails:

```bash
# Simple send
python outlook.py send \
  --to recipient@example.com \
  --subject "Project Update" \
  --body "Here's the latest status..."

# With CC, BCC, attachments
python outlook.py send \
  --to person1@example.com,person2@example.com \
  --cc manager@example.com \
  --bcc archive@example.com \
  --subject "Monthly Report" \
  --body "Please review the attached report." \
  --attach report.pdf \
  --attach data.xlsx

# HTML email
python outlook.py send \
  --to user@example.com \
  --subject "Announcement" \
  --body "<h1>Important</h1><p>New <b>feature</b> released!</p>" \
  --html

# Save as draft (don't send)
python outlook.py send \
  --to user@example.com \
  --subject "Draft Message" \
  --body "Not ready yet..." \
  --draft
```

#### Reply

Reply to an existing email:

```bash
# Reply to sender
python outlook.py reply <message-id> --body "Thanks for the update!"

# Reply all
python outlook.py reply <message-id> --body "Team, noted" --all

# With attachment, save as draft
python outlook.py reply <message-id> \
  --body "See attached" \
  --attach response.pdf \
  --draft
```

#### Forward

Forward an email:

```bash
# Forward to someone
python outlook.py forward <message-id> \
  --to colleague@example.com \
  --body "FYI - see below"

# Forward with additional context
python outlook.py forward <message-id> \
  --to team@example.com \
  --cc manager@example.com \
  --body "Team, please review this request" \
  --attach additional-info.pdf
```

### Calendar Operations

#### List Events

Display calendar events in a rich terminal table:

```bash
# Next 7 days (default)
python outlook.py cal list

# Custom date range
python outlook.py cal list --start 2026-05-01 --end 2026-05-31

# Filter by subject
python outlook.py cal list --subject "standup"

# Filter by location
python outlook.py cal list --location "Conference Room"

# Filter by organizer
python outlook.py cal list --organizer manager@example.com

# All-day events only
python outlook.py cal list --all-day

# Recurring events only
python outlook.py cal list --recurring
```

#### Read Event Details

View event details including attendees and recurrence:

```bash
python outlook.py cal read <event-id>
```

#### Create Event

Schedule a new calendar event:

```bash
# Simple meeting
python outlook.py cal create \
  --subject "Team Standup" \
  --start "2026-05-12 09:00" \
  --end "2026-05-12 09:30"

# With all details
python outlook.py cal create \
  --subject "Quarterly Review" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 16:00" \
  --location "Conference Room A" \
  --body "Q1 performance review and Q2 planning" \
  --required alice@example.com,bob@example.com \
  --optional manager@example.com \
  --reminder 30
```

#### Delete Event

```bash
python outlook.py cal delete <event-id>
```

### Folder Management

```bash
# List all available folders
python outlook.py folders

# Search specific folder
python outlook.py search --folder "Project X"

# Search multiple folders
python outlook.py search --folder Inbox --folder Archive
```

## Output Format

Markdown files are named for chronological sorting:

```text
2026-05-08 1727 - Project Update.md
```

Each file contains:

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
#email/received

Here's the update on the project...

---

## Message 2 (SENT) ^msg-2

**From:** Bob Jones  
**Date:** [[2026-05-08]] 14:30  
#email/sent

Thanks Alice, looks good...
```

## Using with AI Agents

### Claude Code

Add Outlook CLI to your Claude Code workspace for email and calendar access:

```bash
# Search recent emails
python outlook.py search --folder Inbox --days 3

# Find emails from a specific person
python outlook.py search --filter-email vendor@company.com --days 30

# Export project emails for context
python outlook.py export --output ./context \
  --keyword "ProjectX" --days 60

# Send status update
python outlook.py send \
  --to team@company.com \
  --subject "Weekly Update" \
  --body "$(cat weekly-summary.txt)"

# Check calendar
python outlook.py cal list --start 2026-05-12 --end 2026-05-19
```

**Workflow Example:**
1. Claude Code searches your emails: `python outlook.py search --keyword "bug report"`
2. Exports relevant threads: `python outlook.py export --output ./bug-context`
3. Analyzes markdown files to summarize issues
4. Drafts response: `python outlook.py send --draft --to reporter@company.com`

### Hermes Agent

Integrate Outlook CLI as a tool in your Hermes agent configuration:

```json
{
  "tools": [
    {
      "name": "search_email",
      "command": "python outlook.py search --folder Inbox --keyword {query} --days {days}",
      "description": "Search emails by keyword"
    },
    {
      "name": "send_email",
      "command": "python outlook.py send --to {recipient} --subject {subject} --body {body}",
      "description": "Send an email"
    },
    {
      "name": "check_calendar",
      "command": "python outlook.py cal list --start {start_date} --end {end_date}",
      "description": "List calendar events in date range"
    }
  ]
}
```

**Use Cases:**
- Agent checks calendar before scheduling meetings
- Searches emails for context before responding
- Sends automated status reports
- Exports email threads for RAG/knowledge base

### OpenClaw

Configure OpenClaw to use Outlook CLI for email automation:

```yaml
# .openclaw/config.yml
actions:
  - name: check_inbox
    command: python outlook.py search --folder Inbox --unread
    schedule: "*/15 * * * *"  # Every 15 minutes
    
  - name: export_daily_emails
    command: python outlook.py export --output ./archive --days 1
    schedule: "0 20 * * *"  # 8 PM daily
    
  - name: send_reminder
    command: |
      python outlook.py send \
        --to team@company.com \
        --subject "EOD Reminder" \
        --body "Please submit your status updates"
    schedule: "0 17 * * 1-5"  # 5 PM weekdays
```

**Automation Examples:**
- Monitor inbox for urgent emails
- Daily email archival to markdown
- Scheduled reminder emails
- Calendar sync to external systems

## CLI Reference

### Search/Export Options

| Option              | Short | Description                                   |
| ------------------- | ----- | --------------------------------------------- |
| `--folder`          | `-F`  | Folder to search (can specify multiple)       |
| `--days`            | `-d`  | Days to look back (default: 7)                |
| `--from-date`       |       | Start date (YYYY-MM-DD)                       |
| `--to-date`         |       | End date (YYYY-MM-DD)                         |
| `--unread`          | `-u`  | Only unread messages                          |
| `--filter-email`    | `-f`  | Email address to filter (can specify multiple)|
| `--filter-domain`   | `-D`  | Domain to filter (can specify multiple)       |
| `--keyword`         | `-k`  | Keyword in subject/body                       |

### Send/Reply/Forward Options

| Option         | Short | Description                             |
| -------------- | ----- | --------------------------------------- |
| `--to`         | `-t`  | Recipient(s) (comma-separated)          |
| `--subject`    | `-s`  | Email subject                           |
| `--body`       | `-b`  | Email body                              |
| `--cc`         |       | CC recipient(s) (comma-separated)       |
| `--bcc`        |       | BCC recipient(s) (comma-separated)      |
| `--attach`     | `-a`  | File to attach (can specify multiple)   |
| `--html`       |       | Body is HTML format                     |
| `--draft`      |       | Save as draft instead of sending        |
| `--all`        |       | Reply all (reply command only)          |

### Calendar Options

| Option         | Short | Description                             |
| -------------- | ----- | --------------------------------------- |
| `--start`      |       | Start date (YYYY-MM-DD)                 |
| `--end`        |       | End date (YYYY-MM-DD)                   |
| `--subject`    | `-s`  | Event subject                           |
| `--location`   | `-l`  | Event location                          |
| `--body`       | `-b`  | Event description                       |
| `--required`   |       | Required attendees (comma-separated)    |
| `--optional`   |       | Optional attendees (comma-separated)    |
| `--reminder`   |       | Reminder minutes before event           |
| `--organizer`  |       | Filter by organizer email               |
| `--all-day`    |       | All-day events only                     |
| `--recurring`  |       | Recurring events only                   |

## Architecture

```text
outlook_emails/
├── core/
│   ├── connection.py    # Outlook COM connection
│   ├── models.py        # Email dataclass
│   └── folders.py       # Folder management
├── services/
│   ├── search.py        # Email search/filtering
│   ├── viewer.py        # Terminal display (Rich)
│   ├── export.py        # Markdown conversion
│   ├── compose.py       # Send/reply/forward
│   └── calendar.py      # Calendar operations
└── utils/
    └── formatting.py    # Parsing utilities
```

**Service Separation:**

- `SearchService` — Returns `list[Email]` objects from Outlook
- `ViewerService` — Displays emails/events in styled terminal tables
- `ExportService` — Converts emails to Obsidian markdown files
- `ComposeService` — Sends, replies, forwards emails
- `CalendarService` — Manages calendar events

## What Gets Cleaned (Export)

The exporter removes noise while preserving content:

- **Removed:** SafeLinks URLs, signature tables, contact info blocks, social links, tracking pixels, empty table rows, external email warnings, quoted replies
- **Preserved:** Message content, sender names, legitimate data tables, attachments metadata

## Requirements

- **Windows 10/11**
- **Outlook Classic** (desktop app, not "new Outlook")
- **Python 3.8+**
- Dependencies: `pywin32`, `markdownify`, `beautifulsoup4`, `rich`

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/outlook-cli-2.0.git
cd outlook-cli-2.0
pip install -r requirements.txt
```

Or install dependencies directly:

```bash
pip install pywin32 markdownify beautifulsoup4 rich
```

## Use Cases

| Scenario                   | Command                                                        |
| -------------------------- | -------------------------------------------------------------- |
| Archive vendor emails      | `export --filter-domain "vendor.com" -o vendor/`               |
| Project documentation      | `export --keyword "ProjectX" -o projectx/`                     |
| Team correspondence        | `export --filter-email "alice@co.com,bob@co.com" -o team/`     |
| AI/RAG pipeline            | `export` for clean markdown embeddings                         |
| Terminal email triage      | `search --unread --folder Inbox`                               |
| Automated status emails    | `send --to team@co.com --subject "Status" --body "..."`        |
| Calendar-aware scheduling  | `cal list --start 2026-05-15 --end 2026-05-20`                 |
| Send with AI-generated msg | `send --to client@co.com --body "$(cat ai-response.txt)"`      |

## Upgrading from 1.0

If you're coming from the original `outlook_to_markdown.py` script:

**Old (v1.0):**

```bash
python outlook_to_markdown.py --full --days 30 --output ./emails
```

**New (v2.0):**

```bash
python outlook.py export --output ./emails --days 30
```

The v1.0 script (`outlook_to_markdown.py`) is still available for backward compatibility but is no longer maintained. Migrate to the new CLI for access to terminal viewing, sending, and calendar features.

## Limitations

- Windows only (uses COM automation via `pywin32`)
- Requires Outlook Classic desktop app
- Calendar recurring instances may be slow on large calendars

## License

MIT
