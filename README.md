# Outlook CLI 2.0

A Python command-line interface (CLI) for Microsoft Outlook automation on Windows using COM automation. Search emails, send messages, manage calendar events, and export to Obsidian markdown - all from your terminal without Graph API or Azure setup.

## Why this exists

There's a good [outlook-cli](https://github.com/mhattingpete/outlook-cli) that uses Microsoft's Graph API. It works well if you have Azure app registration set up.

This one takes a different approach: COM automation on your local Windows machine. If you already have Outlook desktop running, you can use this without any cloud setup. Useful when you need something working quickly, or when your company is slow to approve Azure app registrations (or won't allow them at all).

## What it does

- Search and filter emails by folder, sender, date range, read status
- Send emails with CC, BCC, attachments, HTML formatting
- Reply and forward messages
- Save drafts for review before sending
- Export email threads to Obsidian markdown or JSON (with incremental sync)
- JSON export with `--stdout` for direct AI agent ingestion
- List, create, and delete calendar events
- Manage tasks/todos (list, create, complete, delete)
- JSON output on all commands for agent/automation integration
- Works offline since everything is local

## Install

**One-liner (auto-installs AI agent skills):**

```powershell
# Windows (PowerShell)
irm https://raw.githubusercontent.com/ob-cheng/outlook-cli-2.0/main/install.ps1 | iex
```

```bash
# macOS/Linux (for WSL or development)
curl -fsSL https://raw.githubusercontent.com/ob-cheng/outlook-cli-2.0/main/install.sh | bash
```

**Manual install:**

```bash
git clone https://github.com/ob-cheng/outlook-cli-2.0.git
cd outlook-cli-2.0
pip install -r requirements.txt
```

Requirements: Windows 10/11 with Outlook desktop installed, Python 3.8+

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

# Export recent threads to Obsidian
python outlook.py export --output ~/Obsidian/Emails --days 30

# Incremental sync (only new emails since last run)
python outlook.py export --output ~/Obsidian/Emails --incremental

# Export as JSON for AI processing (direct to stdout)
python outlook.py export --output . --stdout --days 7
```

## Command reference

### Email commands

| Command | Description |
|---------|-------------|
| `search` | Search and display emails in terminal |
| `read` | Read full email content by message ID |
| `export` | Export emails to markdown or JSON |
| `send` | Send a new email |
| `reply` | Reply to an email |
| `forward` | Forward an email |
| `folders` | List all available folders |

### Calendar commands

| Command | Description |
|---------|-------------|
| `cal list` | List calendar events |
| `cal read` | View event details |
| `cal create` | Create a calendar event |
| `cal delete` | Delete a calendar event |

### Task commands

| Command | Description |
|---------|-------------|
| `tasks list` | List tasks/todos |
| `tasks read` | View task details |
| `tasks create` | Create a new task |
| `tasks complete` | Mark task as complete |
| `tasks delete` | Delete a task |

### Global options

All commands support `--json` for structured output.

### search

```bash
python outlook.py search [options]
```

| Option | Description |
|--------|-------------|
| `--folder`, `-F` | Folder to search (can specify multiple) |
| `--days`, `-d` | Days to look back (default: 7) |
| `--from-date` | Start date (YYYY-MM-DD) |
| `--to-date` | End date (YYYY-MM-DD) |
| `--unread`, `-u` | Only unread messages |
| `--filter-email`, `-f` | Filter by email address (can specify multiple) |
| `--filter-domain`, `-D` | Filter by domain (can specify multiple) |
| `--keyword`, `-k` | Search in subject/body |
| `--export`, `-e` | Also export results to directory |
| `--no-view` | Skip terminal display (use with --export) |
| `--json` | Output as JSON |

### read

```bash
python outlook.py read <message-id> [<message-id>...]
```

| Option | Description |
|--------|-------------|
| `--json` | Output as JSON |

### export

```bash
python outlook.py export --output DIR [options]
```

| Option | Description |
|--------|-------------|
| `--output`, `-o` | Output directory (required) |
| `--days`, `-d` | Days to look back (default: 7) |
| `--folder`, `-F` | Folder to export (can specify multiple) |
| `--filter-email`, `-f` | Filter by email address |
| `--filter-domain`, `-D` | Filter by domain |
| `--keyword`, `-k` | Filter by keyword |
| `--format` | Output format: `markdown` (default) or `json` |
| `--batch` | For JSON: combine all emails into single file |
| `--stdout` | Output JSON to terminal (no files, for agents/pipelines) |
| `--no-threads` | Export each email separately |
| `--no-overwrite` | Skip existing files |
| `--incremental` | Only export since last run (saves state) |
| `--json` | Output as JSON |

### send

```bash
python outlook.py send --to ADDRESS --subject TEXT --body TEXT [options]
```

| Option | Description |
|--------|-------------|
| `--to`, `-t` | Recipient (required, comma-separated) |
| `--subject`, `-s` | Subject line (required) |
| `--body`, `-b` | Message body (required) |
| `--cc` | CC recipients |
| `--bcc` | BCC recipients |
| `--attach`, `-a` | File attachment (can specify multiple) |
| `--html` | Body is HTML formatted |
| `--draft` | Save as draft instead of sending |
| `--json` | Output as JSON |

### reply

```bash
python outlook.py reply <message-id> --body TEXT [options]
```

| Option | Description |
|--------|-------------|
| `--body`, `-b` | Reply message (required) |
| `--all` | Reply to all recipients |
| `--attach`, `-a` | Add attachment |
| `--html` | Body is HTML |
| `--draft` | Save as draft |
| `--json` | Output as JSON |

### forward

```bash
python outlook.py forward <message-id> --to ADDRESS [options]
```

| Option | Description |
|--------|-------------|
| `--to`, `-t` | Forward recipient (required) |
| `--body`, `-b` | Optional message to add |
| `--cc` | CC recipients |
| `--bcc` | BCC recipients |
| `--attach`, `-a` | Additional attachment |
| `--html` | Body is HTML |
| `--draft` | Save as draft |
| `--json` | Output as JSON |

### cal list

```bash
python outlook.py cal list [options]
```

| Option | Description |
|--------|-------------|
| `--start` | Start date (YYYY-MM-DD, default: today) |
| `--end` | End date (YYYY-MM-DD, default: 7 days) |
| `--subject` | Filter by subject |
| `--location` | Filter by location |
| `--organizer` | Filter by organizer email |
| `--all-day` | All-day events only |
| `--recurring` | Recurring events only |
| `--json` | Output as JSON |

### cal create

```bash
python outlook.py cal create --subject TEXT --start DATETIME --end DATETIME [options]
```

| Option | Description |
|--------|-------------|
| `--subject`, `-s` | Event subject (required) |
| `--start` | Start "YYYY-MM-DD HH:MM" (required) |
| `--end` | End "YYYY-MM-DD HH:MM" (required) |
| `--location`, `-l` | Location |
| `--body`, `-b` | Description |
| `--required` | Required attendees (comma-separated) |
| `--optional` | Optional attendees |
| `--reminder` | Reminder minutes (default: 15) |
| `--no-reminder` | No reminder |
| `--json` | Output as JSON |

### tasks list

```bash
python outlook.py tasks list [options]
```

| Option | Description |
|--------|-------------|
| `--status` | Filter: `not_started`, `in_progress`, `completed`, `waiting`, `deferred` |
| `--all` | Include completed tasks |
| `--due-before` | Tasks due before date (YYYY-MM-DD) |
| `--due-after` | Tasks due after date (YYYY-MM-DD) |
| `--priority` | Filter: `low`, `normal`, `high` |
| `--category` | Filter by category name |
| `--json` | Output as JSON |

### tasks create

```bash
python outlook.py tasks create --subject TEXT [options]
```

| Option | Description |
|--------|-------------|
| `--subject`, `-s` | Task subject (required) |
| `--due` | Due date (YYYY-MM-DD) |
| `--start` | Start date (YYYY-MM-DD) |
| `--priority` | Priority: `low`, `normal` (default), `high` |
| `--body`, `-b` | Task description |
| `--category` | Category name |
| `--reminder` | Reminder date/time (YYYY-MM-DD HH:MM) |
| `--json` | Output as JSON |

### tasks complete

```bash
python outlook.py tasks complete <task-id>
```

### tasks delete

```bash
python outlook.py tasks delete <task-id>
```

## Use cases

### Incremental email sync to Obsidian

Keep your Obsidian vault in sync with Outlook without re-exporting everything:

```bash
# First run - exports last 7 days, saves timestamp
python outlook.py export --output ~/Obsidian/Emails --incremental

# Next run - only exports emails since last run
python outlook.py export --output ~/Obsidian/Emails --incremental
```

State is saved to `extraction_state.json` in the output directory.

### Triage inbox

```bash
# Find unread emails from a specific person
python outlook.py search --unread --filter-email boss@company.com

# Read the full content
python outlook.py read <message-id>

# Reply
python outlook.py reply <message-id> --body "I'll handle this today"
```

### Schedule a meeting

```bash
# Check availability first
python outlook.py cal list --start 2026-05-15 --end 2026-05-20

# Create the event
python outlook.py cal create \
  --subject "Team Sync" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 15:00" \
  --required "alice@co.com,bob@co.com"
```

### Manage tasks

```bash
# See pending tasks
python outlook.py tasks list

# Create a task with due date
python outlook.py tasks create --subject "Review PR #42" --due 2026-05-15 --priority high

# Mark complete
python outlook.py tasks complete <task-id>

# Get tasks as JSON for automation
python outlook.py tasks list --json
```

## Agent integration

This tool integrates with AI agents via **direct CLI calls** (not MCP server). The install script auto-detects and installs skill files to:

- Claude Code (`~/.claude/skills/`)
- Cursor (`~/.cursor/skills/`)
- Windsurf (`~/.windsurf/skills/`)
- GitHub Copilot (`~/.copilot/skills/`)
- Hermes (`~/.hermes/skills/`)
- OpenClaw (`~/.openclaw/skills/`)

Agents invoke commands directly via shell. All commands support `--json` for structured output that's easy to parse.

The skill definition is in `SKILL.md` at the repo root.

## JSON output

All commands support `--json` for structured, parseable output:

```bash
python outlook.py search --unread --json
```

Returns:

```json
{
  "success": true,
  "count": 3,
  "emails": [
    {
      "message_id": "00000000ABC...",
      "subject": "Meeting tomorrow",
      "sender_clean": "Alice Smith",
      "date": "2026-05-11T09:30:00",
      "is_read": false
    }
  ]
}
```

Errors return:

```json
{
  "success": false,
  "error": "Message not found",
  "code": "not_found"
}
```

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

Threads are grouped by subject. Each message is marked as sent or received. Tracking URLs, signature tables, "CAUTION: external email" warnings, and quoted reply chains are stripped.

### JSON export (for AI ingestion)

For AI/LLM workflows, JSON format is more token-efficient:

```bash
# One JSON file per thread
python outlook.py export --output ~/data --format json

# All emails in single file (most efficient for bulk processing)
python outlook.py export --output ~/data --format json --batch

# Direct to terminal (no files written) - ideal for agents
python outlook.py export --output . --stdout --days 7
```

JSON structure:

```json
{
  "export_date": "2026-05-11T10:30:00",
  "thread_count": 5,
  "email_count": 12,
  "threads": [
    {
      "subject": "Project Update",
      "message_count": 3,
      "date_start": "2026-05-08T09:00:00",
      "date_end": "2026-05-10T14:30:00",
      "participants": ["Alice Smith", "Bob Jones"],
      "messages": [
        {
          "id": "00000000ABC...",
          "subject": "Project Update",
          "from": "Alice Smith",
          "from_email": "alice@company.com",
          "to": ["bob@company.com"],
          "cc": null,
          "date": "2026-05-08T09:00:00",
          "direction": "received",
          "body": "Clean text content without HTML or quoted replies"
        }
      ]
    }
  ]
}
```

Benefits over markdown for AI:

- ~15-20% fewer tokens (no formatting overhead)
- Single file for batch processing
- Structured fields for direct extraction
- Same content cleaning (signatures, quoted replies stripped)

## Documentation

- [Examples & Use Cases](docs/EXAMPLES.md)
- [Comparison with Graph API & other tools](docs/COMPARISON.md)
- [Skill Definition](SKILL.md)

## Limitations

- Windows only (COM automation)
- Needs Outlook desktop app installed
- No shared mailbox support yet
- No MCP server (direct CLI integration only)

## License

MIT

[https://github.com/ob-cheng/outlook-cli-2.0](https://github.com/ob-cheng/outlook-cli-2.0)
