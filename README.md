# Outlook to Markdown

Convert Microsoft Outlook emails to clean, readable Markdown files. Supports both `.msg` file conversion and live Outlook extraction with Obsidian-flavored output.

## Features

- **Thread consolidation** - Groups related emails (RE:/FW:) into single files
- **Quote stripping** - Removes quoted replies, keeping only new content
- **Obsidian-ready** - YAML frontmatter, wikilinks to daily notes, callouts, and tags
- **Incremental extraction** - Tracks last run to avoid re-processing
- **Auto-reply filtering** - Skips Out of Office and automatic responses
- **Clean output** - Removes tracking pixels, signature images, and HTML cruft

## Scripts

### `outlook_to_markdown.py`

Connects directly to Outlook Classic via COM and extracts emails from Inbox and Sent Items.

```bash
# Incremental extraction (emails since last run)
python outlook_to_markdown.py

# Full extraction (last 7 days)
python outlook_to_markdown.py --full

# Specify custom output directory
python outlook_to_markdown.py --output "C:\path\to\obsidian\vault"

# Combine flags
python outlook_to_markdown.py --full --output "C:\path\to\obsidian\vault"
```

**Requirements:** Windows with Outlook Classic installed and configured.

### `convert_msg_to_md.py`

Converts `.msg` files in the current directory to Markdown.

```bash
# Place .msg files in the directory, then run:
python convert_msg_to_md.py
```

## Installation

```bash
# Clone the repository
git clone https://github.com/ob-cheng/email-to-obsidian.git
cd email-to-obsidian

# Install dependencies
pip install -r requirements.txt
```

## Output Format

Emails are converted to Markdown with:

- **YAML frontmatter** with title, date, participants, and tags
- **Thread summary** callout showing message count and date range
- **Individual messages** with sender, date (as wikilink), and cleaned content
- **AI processing hints** (hidden) for downstream LLM processing

Example output:

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
  - contact/internal
---

> [!info] Thread Summary
> **3 message(s)** from [[2026-05-06]] to [[2026-05-08]]
> **Participants:** Alice Smith, Bob Jones

## Message 1 (RECEIVED) ^msg-1

**From:** Alice Smith  
**Date:** [[2026-05-06]] 09:15  
#email/received

Here's the update on the project...

---
```

## Configuration

**Command-line options:**

- `--full` - Extract last 7 days instead of incremental
- `--output`, `-o` - Specify output directory (overrides default)

**Script constants** (edit `outlook_to_markdown.py`):

- `DEFAULT_LOOKBACK_DAYS` - Days to look back on full extraction (default: 7)
- `STATE_FILE` - JSON file tracking extraction state
- `OUTPUT_DIR` - Default output directory for Markdown files

## Use Cases

- **Obsidian integration** - Import emails into your knowledge base with proper linking
- **Email archival** - Create searchable, version-controlled email archives
- **AI/LLM processing** - Clean markdown is ideal for RAG and summarization
- **Cross-platform access** - Read your emails anywhere without Outlook

## Limitations

- Windows only (uses COM automation)
- Requires Outlook Classic (not the new Outlook)
- Only processes Inbox and Sent Items by default

## License

MIT
