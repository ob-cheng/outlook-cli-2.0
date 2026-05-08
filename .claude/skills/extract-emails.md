---
name: extract-emails
description: Extract Outlook emails to Obsidian markdown. Use when user asks to "extract emails", "do an extraction", or "get emails".
---

# Email Extraction Skill

Extract emails from Outlook (Inbox and Sent Items) and convert them to Obsidian-flavored markdown.

## How to run

```bash
cd "./"
python outlook_to_markdown.py --output "~/Obsidian/Emails"
```

## Options

- **Incremental extraction** (default): Only extracts emails since the last run
- **Full extraction**: Add `--full` flag to extract last 7 days

## Examples

Incremental:
```bash
python outlook_to_markdown.py --output "~/Obsidian/Emails"
```

Full 7-day extraction:
```bash
python outlook_to_markdown.py --full --output "~/Obsidian/Emails"
```

## Output

- Markdown files are saved to the Obsidian vault at `00_Emails Inbox`
- Emails are grouped by thread (conversation)
- State is tracked in `extraction_state.json`
