---
name: extract-emails
description: Export Outlook emails to Obsidian markdown. Use when user asks to "extract emails", "export emails", or "get emails to markdown".
---

# Email Export Skill

Export emails from Outlook to Obsidian-flavored markdown files.

## How to run

```bash
cd "./"
python outlook.py export --output "~/Obsidian/Emails" --days 7
```

## Options

- `--output DIR` - Output directory (required)
- `--days N` - Days to look back (default: 7)
- `--folder NAME` - Folder to export from (can specify multiple)
- `--filter-email ADDRESS` - Filter by participant email
- `--filter-domain DOMAIN` - Filter by email domain
- `--keyword TEXT` - Filter by keyword in subject/body
- `--no-threads` - Export each email as separate file
- `--no-overwrite` - Skip files that already exist

## Examples

Last 7 days:

```bash
python outlook.py export --output "~/Obsidian/Emails"
```

Last 30 days from Inbox only:

```bash
python outlook.py export --output "..." --folder Inbox --days 30
```

Filter by sender:

```bash
python outlook.py export --output "..." --filter-email client@company.com --days 14
```

Incremental export (skip existing):

```bash
python outlook.py export --output "..." --days 30 --no-overwrite
```