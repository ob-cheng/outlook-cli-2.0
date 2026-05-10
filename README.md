# Outlook CLI 2.0

The agent-friendly Outlook CLI. No Azure setup. No OAuth dance. Just your terminal and Outlook.

**Other tools require Azure app registration, admin approval, and OAuth setup.** We don't. This is for people and agents that need to work with Outlook *right now*, without waiting for IT or infrastructure teams.

## Why This Exists

There's already an excellent [outlook-cli](https://github.com/mhattingpete/outlook-cli) out there. It uses Microsoft's Graph API. It's well-built.

But it requires:
- Azure app registration (your company's IT has to allow it)
- OAuth setup and approval flow
- API rate limits
- Network connectivity (cloud-based)
- Waiting weeks for enterprise security review

Not every company moves fast. Not every agent can wait.

**This tool** runs locally on your Windows machine. It talks directly to Outlook via COM automation. No approval needed. No infrastructure setup. Just `pip install` and go.

## What This Can Do

Full read/write access to Outlook:

- Search and filter emails across folders
- Send emails with CC, BCC, attachments
- Reply and forward messages
- Draft emails for review
- Export threads to Obsidian markdown
- List, create, and delete calendar events
- Works offline, no API rate limits, zero infrastructure overhead

## Use Cases This Enables

**Agents that need to draft and send emails without human intervention:**

```bash
python outlook.py send \
  --to client@company.com \
  --subject "Report ready for review" \
  --body "$(cat report-summary.txt)"
```

**Agents that manage your calendar:**

```bash
# Check what you're doing before scheduling
python outlook.py cal list --start 2026-05-15 --end 2026-05-20

# Create a meeting slot
python outlook.py cal create \
  --subject "Team sync" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 15:00"
```

**Agents that triage your inbox without you:**

```bash
# Find urgent stuff
python outlook.py search --unread --keyword "urgent" \
  --filter-email boss@company.com

# Forward to someone else
python outlook.py forward <message-id> \
  --to delegated@company.com \
  --body "Can you handle this?"
```

**Rapid agent iteration (no waiting for approvals):**

- Deploy at 9am
- Agent works with Outlook immediately
- No "pending infrastructure review"
- No "waiting for Azure app approval"

## Install

```bash
git clone https://github.com/ob-cheng/outlook-cli-2.0.git
cd outlook-cli-2.0
pip install -r requirements.txt
```

Requires:

- Windows 10/11
- Outlook Classic (desktop app)
- Python 3.8+

No Azure account. No OAuth flow. No IT approval needed.

## Quick Start

```bash
# What's in your inbox
python outlook.py search --unread

# What you're doing this week
python outlook.py cal list

# Send an email
python outlook.py send \
  --to someone@company.com \
  --subject "Quick question" \
  --body "Can we sync tomorrow?"

# Draft something for review (don't send yet)
python outlook.py send \
  --to client@company.com \
  --subject "Project update" \
  --body "Here's where we stand..." \
  --draft

# Export threads to markdown
python outlook.py export --output ~/Obsidian/Emails --days 30
```

## Commands

### Search and Display

```bash
# Default: next 7 days
python outlook.py search

# Unread only
python outlook.py search --unread

# From a specific person
python outlook.py search --filter-email john@company.com

# With keyword
python outlook.py search --keyword "meeting" --days 14

# Complex query
python outlook.py search \
  --folder Inbox \
  --unread \
  --filter-email boss@company.com \
  --keyword "urgent" \
  --days 7

# Date range
python outlook.py search --from-date 2026-05-01 --to-date 2026-05-31
```

### Export to Markdown

```bash
# Basic
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

### Send Email

```bash
# Basic send
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

# HTML
python outlook.py send \
  --to user@example.com \
  --subject "Announcement" \
  --body "<h1>New feature</h1><p>Live now</p>" \
  --html

# Draft (for review before sending)
python outlook.py send \
  --to user@example.com \
  --subject "Draft message" \
  --body "Draft me first" \
  --draft
```

### Reply and Forward

```bash
# Reply to sender
python outlook.py reply <message-id> --body "Thanks!"

# Reply to all
python outlook.py reply <message-id> --body "Team, here's my take" --all

# Reply with attachment
python outlook.py reply <message-id> \
  --body "See attached analysis" \
  --attach analysis.pdf

# Draft reply for review
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
# What's coming
python outlook.py cal list

# Next month
python outlook.py cal list --start 2026-05-01 --end 2026-05-31

# Find all standups
python outlook.py cal list --subject "standup"

# Recurring events
python outlook.py cal list --recurring

# View details
python outlook.py cal read <event-id>

# Create a meeting
python outlook.py cal create \
  --subject "Sync" \
  --start "2026-05-15 14:00" \
  --end "2026-05-15 15:00" \
  --location "Room A" \
  --body "Quarterly planning"

# Delete
python outlook.py cal delete <event-id>
```

### Folders

```bash
python outlook.py folders
```

Lists everything: Inbox, Sent, Archive, custom folders, all of it.

## Output Format

Export produces clean Obsidian-ready markdown:

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

Here's the email content. Quoted replies stripped. Signatures cleaned. Just the meat.

---

## Message 2 (SENT) ^msg-2

Your reply here.
```

Threads grouped by subject. Each message marked sent/received. Perfect for knowledge bases and AI context.

## Using with Agents

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

Claude can run these commands directly. No setup. No waiting.

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

Your agent has full email and calendar access. No Azure. No OAuth. Deploy now.

### OpenClaw

```yaml
actions:
  - name: monitor_inbox
    command: python outlook.py search --unread --folder Inbox
    schedule: "0 * * * *"  # Hourly

  - name: daily_archive
    command: python outlook.py export --output ./archive --days 1
    schedule: "0 20 * * *"  # 8pm daily

  - name: send_eod_reminder
    command: |
      python outlook.py send \
        --to team@company.com \
        --subject "EOD: Submit status" \
        --body "End-of-day reminder"
    schedule: "0 17 * * 1-5"  # 5pm weekdays
```

Automate email workflows without infrastructure overhead.

## What Gets Cleaned

When exporting to markdown, we strip the noise:

**Gone:** Tracking URLs, signature tables, "CAUTION: external email" warnings, quoted replies, tracking pixels.

**Stays:** Actual email content, names in signatures, legitimate tables.

## Why Local? Why COM?

**Local (no cloud):**
- Works offline
- No API rate limits
- No network latency
- Your emails stay on your machine
- Zero infrastructure cost

**COM automation (direct to Outlook):**
- No OAuth flow
- No app registration
- No Azure dependency
- Faster than REST API
- Deployed and working in minutes

This is why agents can use it immediately. No waiting for approvals, no infrastructure reviews, no "IT will get back to you next quarter."

## Limitations

- Windows only (uses COM)
- Requires Outlook Classic desktop app
- No shared mailbox support (yet)

That's it. Everything else works.

## License

MIT

---

**Deploy this morning. Your agent has email by afternoon.** No Azure. No OAuth. No IT tickets.

[https://github.com/ob-cheng/outlook-cli-2.0](https://github.com/ob-cheng/outlook-cli-2.0)
