---
name: search-emails
description: Search Outlook emails. Use when user asks to "search emails", "find emails", "look for emails", or "check inbox".
---

# Email Search Skill

Search and display Outlook emails in the terminal.

## How to run

```bash
cd "./"
python outlook.py search [options]
```

## Common searches

Unread emails:

```bash
python outlook.py search --unread
```

From specific person:

```bash
python outlook.py search --filter-email person@company.com
```

By keyword:

```bash
python outlook.py search --keyword "project update" --days 14
```

From specific domain:

```bash
python outlook.py search --filter-domain vendor.com
```

## Options

- `--folder NAME` - Folder to search (can specify multiple)
- `--days N` - Days to look back (default: 7)
- `--unread` - Only unread messages
- `--filter-email ADDRESS` - Filter by email address (can specify multiple)
- `--filter-domain DOMAIN` - Filter by domain (can specify multiple)
- `--keyword TEXT` - Search in subject/body
- `--from-date DATE` - Start date (YYYY-MM-DD)
- `--to-date DATE` - End date (YYYY-MM-DD)
- `--export DIR` - Also export results to markdown
- `--no-view` - Skip terminal display (use with --export)

## Output

Returns a table with message IDs, dates, senders, subjects, and status.
Use `python outlook.py read <id>` to view full content.