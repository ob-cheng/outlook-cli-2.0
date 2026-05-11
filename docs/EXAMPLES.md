# Outlook CLI Examples

Common workflows and real-world examples for Outlook CLI automation.

## Email Management

### Daily inbox triage
```bash
# See what's urgent
python outlook.py search --unread --days 1

# Check emails from your manager
python outlook.py search --filter-email manager@company.com --days 7
```

### Bulk export for backup
```bash
# Export last year of emails to markdown
python outlook.py export --output ~/Backups/Emails --days 365
```

### Find all emails about a project
```bash
python outlook.py search --keyword "Project Phoenix" --days 90 --export ~/Projects/Phoenix/emails
```

## AI Agent Workflows

### Automated email summaries with Claude Code
```bash
# Export recent emails and let Claude summarize
python outlook.py export --output ~/temp/emails --days 3 --incremental
# Then in Claude Code: "Summarize the emails in ~/temp/emails"
```

### Token-efficient export for LLMs

```bash
# JSON batch export - all emails in single file, optimized for AI ingestion
python outlook.py export --output ~/data --format json --batch --days 7

# Result: emails_20260511_103000.json with all threads in one file
# ~15-20% fewer tokens than markdown equivalent
```

### Direct stdout for agents (no files)

```bash
# Output JSON directly to terminal - agent can read instantly
python outlook.py export --output . --stdout --days 3

# Pipe to another tool
python outlook.py export --output . --stdout --filter-email boss@co.com | jq '.threads[0]'
```

### Export for RAG pipelines

```bash
# JSON export with threads for embedding
python outlook.py export --output ~/rag/emails --format json --days 30

# Each thread becomes a separate JSON file, good for chunking by conversation
```

### Schedule meetings from natural language
With Claude Code skill installed:
```
User: "Schedule a sync with alice@co.com next Tuesday at 2pm"
Claude: [Uses outlook-cli to create calendar event]
```

## Calendar Automation

### Weekly schedule review
```bash
# See next week's meetings
python outlook.py cal list --start 2026-05-12 --end 2026-05-19 --json
```

### Find free time
```bash
# Check for gaps in your calendar
python outlook.py cal list --start 2026-05-15 --end 2026-05-15
```

## Obsidian Integration

### Continuous email sync
```bash
# Set up daily cron job (Windows Task Scheduler)
python outlook.py export \
  --output "C:/Users/You/Obsidian/Emails" \
  --incremental \
  --filter-domain company.com
```

This creates searchable, linked notes in Obsidian with proper frontmatter and tags.

## Batch Operations

### Send meeting invites to a team
```bash
# Create multiple calendar events from a script
for attendee in alice@co.com bob@co.com charlie@co.com; do
  python outlook.py cal create \
    --subject "Quarterly Review" \
    --start "2026-06-01 10:00" \
    --end "2026-06-01 11:00" \
    --required "$attendee"
done
```

### Forward multiple emails
```bash
# Forward all unread emails from a domain to your assistant
python outlook.py search --filter-domain vendor.com --unread --json | \
  jq -r '.[].message_id' | \
  xargs -I {} python outlook.py forward {} --to assistant@company.com
```

## JSON Output for Scripting

### Parse unread count
```bash
# Get count for status bar
python outlook.py search --unread --json | jq '.count'
```

### Extract calendar as CSV
```bash
python outlook.py cal list --json | \
  jq -r '.events[] | [.subject, .start, .location] | @csv' > calendar.csv
```

## Keywords

outlook cli examples, outlook automation scripts, outlook python examples, send email command line, outlook calendar automation, outlook to obsidian workflow, email export markdown, pywin32 examples
