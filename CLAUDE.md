# Outlook CLI 2.0

Command-line interface for Microsoft Outlook using COM automation.

## Quick reference

| Task | Command |
|------|---------|
| Search emails | `python outlook.py search [options]` |
| Read email | `python outlook.py read <message-id>` |
| Export to markdown | `python outlook.py export --output DIR` |
| Send email | `python outlook.py send --to X --subject Y --body Z` |
| Reply | `python outlook.py reply <id> --body "text"` |
| Forward | `python outlook.py forward <id> --to X` |
| List calendar | `python outlook.py cal list` |
| Create event | `python outlook.py cal create --subject X --start Y --end Z` |
| List folders | `python outlook.py folders` |

## Working directory

Always run from: `./`

## Default output

Obsidian vault: `~/Obsidian/Emails`

## Architecture

- `outlook.py` - Entry point
- `outlook_emails/cli.py` - CLI argument parsing
- `outlook_emails/services/` - Business logic (search, export, compose, calendar, viewer)
- `outlook_emails/core/` - Connection, models, folder utilities
- `outlook_emails/utils/` - Formatting helpers
