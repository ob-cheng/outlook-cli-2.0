---
name: outlook-cli
description: Search, send, and manage Outlook emails, calendar, and tasks from the command line. Use when the user wants to search emails, send messages, reply, forward, export to markdown/JSON, manage calendar events, or manage tasks/todos. Supports --stdout for direct JSON output to AI agents.
---

# outlook-cli

AI-friendly CLI for Microsoft Outlook. Works via COM automation - no Azure setup, no OAuth, no API keys needed.

## Install

```bash
# Clone and install dependencies
git clone https://github.com/ob-cheng/outlook-cli-2.0.git
cd outlook-cli-2.0
pip install -r requirements.txt
```

Requirements: Windows 10/11, Outlook desktop app, Python 3.8+

Verify with `python outlook.py --help`.

---

## Strategy

**Search → Read → Act**. Find emails first, then read details, then reply/forward/export. Add `--json` for structured output.

---

## Quick Start

```bash
# Search recent emails
python outlook.py search --days 7

# Read specific email
python outlook.py read <message-id>

# Send email
python outlook.py send --to user@example.com --subject "Hello" --body "Message here"

# Check calendar
python outlook.py cal list

# Export to Obsidian markdown
python outlook.py export --output ./emails --days 30
```

---

## Commands

### search - Find emails

```bash
python outlook.py search [options]
```

Options:
- `--folder NAME` - Folder to search (can specify multiple)
- `--days N` - Days to look back (default: 7)
- `--unread` - Only unread messages
- `--filter-email ADDRESS` - Filter by email address (can specify multiple)
- `--filter-domain DOMAIN` - Filter by domain (can specify multiple)
- `--keyword TEXT` - Search in subject/body
- `--from-date DATE` - Start date (YYYY-MM-DD)
- `--to-date DATE` - End date (YYYY-MM-DD)
- `--export DIR` - Also export results to markdown
- `--no-view` - Skip terminal display
- `--json` - Output as JSON

Examples:
```bash
python outlook.py search --unread
python outlook.py search --filter-email boss@company.com --days 14
python outlook.py search --keyword "urgent" --folder Inbox --json
```

---

### read - View email content

```bash
python outlook.py read <message-id> [<message-id>...]
```

Options:
- `--json` - Output as JSON

Get message IDs from search results. Can read multiple emails at once.

---

### send - Compose and send email

```bash
python outlook.py send --to ADDRESS --subject TEXT --body TEXT [options]
```

Options:
- `--to ADDRESS` - Recipient (required, comma-separated for multiple)
- `--subject TEXT` - Subject line (required)
- `--body TEXT` - Message body (required)
- `--cc ADDRESS` - CC recipients
- `--bcc ADDRESS` - BCC recipients
- `--attach PATH` - File attachment (can specify multiple)
- `--html` - Body is HTML formatted
- `--draft` - Save as draft instead of sending
- `--json` - Output as JSON

Examples:
```bash
python outlook.py send --to user@example.com --subject "Update" --body "Here's the status"
python outlook.py send --to client@co.com --subject "Report" --body "See attached" --attach report.pdf --draft
```

---

### reply - Reply to email

```bash
python outlook.py reply <message-id> --body TEXT [options]
```

Options:
- `--body TEXT` - Reply message (required)
- `--all` - Reply to all recipients
- `--attach PATH` - Add attachment
- `--html` - Body is HTML
- `--draft` - Save as draft
- `--json` - Output as JSON

---

### forward - Forward email

```bash
python outlook.py forward <message-id> --to ADDRESS [options]
```

Options:
- `--to ADDRESS` - Forward recipient (required)
- `--body TEXT` - Optional message to add
- `--cc ADDRESS` - CC recipients
- `--attach PATH` - Additional attachment
- `--html` - Body is HTML
- `--draft` - Save as draft
- `--json` - Output as JSON

---

### export - Export emails to markdown or JSON

```bash
python outlook.py export --output DIR [options]
```

Options:
- `--output DIR` - Output directory (required, use `.` with --stdout)
- `--days N` - Days to look back (default: 7)
- `--folder NAME` - Folder to export (can specify multiple)
- `--filter-email ADDRESS` - Filter by participant
- `--filter-domain DOMAIN` - Filter by domain
- `--keyword TEXT` - Filter by keyword
- `--format FORMAT` - Output format: `markdown` (default) or `json`
- `--batch` - For JSON: combine all emails into single file
- `--stdout` - Output JSON to terminal (no files, best for AI ingestion)
- `--no-threads` - Export each email separately
- `--no-overwrite` - Skip existing files
- `--incremental` - Only export emails since last run (saves state)
- `--json` - Output summary as JSON

Examples:
```bash
# Markdown export (default)
python outlook.py export --output ./emails --days 30

# JSON export to files
python outlook.py export --output ./data --format json --batch

# JSON direct to stdout (best for AI processing)
python outlook.py export --output . --stdout --days 7

# Filtered export
python outlook.py export --output ./project --filter-email client@vendor.com --keyword "contract"
```

---

### cal list - List calendar events

```bash
python outlook.py cal list [options]
```

Options:
- `--start DATE` - Start date (YYYY-MM-DD, default: today)
- `--end DATE` - End date (YYYY-MM-DD, default: 7 days)
- `--subject TEXT` - Filter by subject
- `--location TEXT` - Filter by location
- `--organizer EMAIL` - Filter by organizer
- `--all-day` - All-day events only
- `--recurring` - Recurring events only
- `--json` - Output as JSON

---

### cal read - View event details

```bash
python outlook.py cal read <event-id>
```

Options:
- `--json` - Output as JSON

---

### cal create - Create calendar event

```bash
python outlook.py cal create --subject TEXT --start DATETIME --end DATETIME [options]
```

Options:
- `--subject TEXT` - Event subject (required)
- `--start DATETIME` - Start "YYYY-MM-DD HH:MM" (required)
- `--end DATETIME` - End "YYYY-MM-DD HH:MM" (required)
- `--location TEXT` - Location
- `--body TEXT` - Description
- `--required EMAILS` - Required attendees (comma-separated)
- `--optional EMAILS` - Optional attendees
- `--reminder N` - Reminder minutes (default: 15)
- `--no-reminder` - No reminder
- `--json` - Output as JSON

---

### cal delete - Delete calendar event

```bash
python outlook.py cal delete <event-id>
```

Options:
- `--json` - Output as JSON

---

### tasks list - List tasks/todos

```bash
python outlook.py tasks list [options]
```

Options:
- `--status STATUS` - Filter: `not_started`, `in_progress`, `completed`, `waiting`, `deferred`
- `--all` - Include completed tasks
- `--due-before DATE` - Tasks due before date (YYYY-MM-DD)
- `--due-after DATE` - Tasks due after date (YYYY-MM-DD)
- `--priority PRIORITY` - Filter: `low`, `normal`, `high`
- `--category NAME` - Filter by category
- `--json` - Output as JSON

---

### tasks read - View task details

```bash
python outlook.py tasks read <task-id>
```

Options:
- `--json` - Output as JSON

---

### tasks create - Create a new task

```bash
python outlook.py tasks create --subject TEXT [options]
```

Options:
- `--subject TEXT` - Task subject (required)
- `--due DATE` - Due date (YYYY-MM-DD)
- `--start DATE` - Start date (YYYY-MM-DD)
- `--priority PRIORITY` - Priority: `low`, `normal` (default), `high`
- `--body TEXT` - Task description
- `--category NAME` - Category name
- `--reminder DATETIME` - Reminder (YYYY-MM-DD HH:MM)
- `--json` - Output as JSON

Examples:
```bash
python outlook.py tasks create --subject "Review PR" --due 2026-05-15 --priority high
python outlook.py tasks create --subject "Follow up" --due 2026-05-20 --body "Check status with team"
```

---

### tasks complete - Mark task as complete

```bash
python outlook.py tasks complete <task-id>
```

Options:
- `--json` - Output as JSON

---

### tasks delete - Delete a task

```bash
python outlook.py tasks delete <task-id>
```

Options:
- `--json` - Output as JSON

---

### folders - List all folders

```bash
python outlook.py folders
```

Options:
- `--json` - Output as JSON

---

## JSON Output

All commands support `--json` for structured output. Example:

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
      "sender": "alice@company.com",
      "date": "2026-05-11T09:30:00",
      "is_read": false,
      "has_attachments": true
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

---

## Common Workflows

**Triage inbox:**
```bash
python outlook.py search --unread --json | jq '.emails[] | {subject, sender}'
```

**Reply to all unread from boss:**
```bash
# First search
python outlook.py search --unread --filter-email boss@company.com

# Then reply to specific message
python outlook.py reply <message-id> --body "I'll handle this today"
```

**Export project emails:**
```bash
python outlook.py export --output ./project-archive \
  --filter-email client@company.com \
  --keyword "ProjectX" \
  --days 90
```

**Schedule meeting:**
```bash
python outlook.py cal create \
  --subject "Team Sync" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 15:00" \
  --location "Room A" \
  --required "alice@co.com,bob@co.com"
```

**AI-friendly email export (direct JSON):**
```bash
# Get recent emails as JSON for AI processing - no files written
python outlook.py export --output . --stdout --days 7

# Filter and pipe to processing
python outlook.py export --output . --stdout --filter-email client@co.com | your-ai-tool
```

**Task management:**
```bash
# List pending tasks
python outlook.py tasks list

# Create task with due date
python outlook.py tasks create --subject "Review PR #42" --due 2026-05-15 --priority high

# Mark complete
python outlook.py tasks complete <task-id>

# Get tasks as JSON for automation
python outlook.py tasks list --json
```
