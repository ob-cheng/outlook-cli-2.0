# Outlook CLI 2.0

Command-line interface for Microsoft Outlook using COM automation.

## Quick reference

| Task | Command |
|------|---------|
| Search emails | `python outlook.py search [options]` |
| Read email | `python outlook.py read <message-id>` |
| Export to markdown | `python outlook.py export --output DIR` |
| Export to JSON | `python outlook.py export --output DIR --format json` |
| Export batch JSON | `python outlook.py export --output DIR --format json --batch` |
| Export to stdout | `python outlook.py export --output . --stdout` |
| Send email | `python outlook.py send --to X --subject Y --body Z` |
| Reply | `python outlook.py reply <id> --body "text"` |
| Forward | `python outlook.py forward <id> --to X` |
| List calendar | `python outlook.py cal list` |
| Create event | `python outlook.py cal create --subject X --start Y --end Z` |
| List tasks | `python outlook.py tasks list` |
| Create task | `python outlook.py tasks create --subject X --due Y` |
| Complete task | `python outlook.py tasks complete <task-id>` |
| List folders | `python outlook.py folders` |

## Architecture

- `outlook.py` - Entry point
- `outlook_cli/cli.py` - CLI argument parsing
- `outlook_cli/services/` - Business logic (search, export, compose, calendar, viewer)
- `outlook_cli/core/` - Connection, models, folder utilities
- `outlook_cli/utils/` - Formatting helpers
