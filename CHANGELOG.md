# Changelog

All notable changes to Outlook CLI 2.0 will be documented in this file.

## [2.2.0] - 2026-05-11

### ✨ Added

**Task/Todo Management:**

- `tasks list` - List tasks with filters (status, priority, due date, category)
- `tasks read` - View task details
- `tasks create` - Create tasks with due date, priority, reminder
- `tasks complete` - Mark tasks as complete
- `tasks delete` - Delete tasks
- Full JSON output support for all task commands
- Task model and TaskService for programmatic access

### 📚 Documentation

- Updated README with task commands and examples
- Updated SKILL.md with task management for AI agents
- Updated CLAUDE.md quick reference

---

## [2.1.0] - 2026-05-11

### ✨ Added

**JSON Export for AI Ingestion:**

- `--format json` option for export command - output as JSON instead of markdown
- `--batch` option - combine all emails into a single JSON file (most token-efficient)
- `--stdout` option - output JSON directly to terminal for AI agent pipelines
- Lean JSON structure optimized for LLM consumption (~15-20% fewer tokens than markdown)
- `to_json_data()` method in ExportService for programmatic access

### 📚 Documentation

- Updated README with JSON export section and examples
- Updated SKILL.md with new export options for AI agents
- Updated EXAMPLES.md with AI workflow examples
- Updated CLAUDE.md quick reference

---

## [2.0.0] - 2026-05-11

### 🎉 Major Rewrite - Outlook CLI 2.0

Complete architectural overhaul with new command structure and service-based design.

### ✨ Added

**Terminal UI:**
- Rich terminal tables with styled output for emails and calendar events
- Formatted panels for detailed email/event viewing
- Color-coded status indicators (unread, importance, attachments)
- Summary panels with statistics

**Email Sending & Composing:**
- `send` command - Send new emails with To/CC/BCC
- `reply` command - Reply or Reply All to emails
- `forward` command - Forward emails to others
- Draft mode - Save emails as drafts instead of sending
- Attachment support - Attach multiple files to emails
- HTML email support - Send rich HTML-formatted messages

**Calendar Management:**
- `cal list` - List calendar events with filtering
- `cal read` - View detailed event information
- `cal create` - Create new calendar events
- `cal delete` - Delete calendar events
- Filter by: subject, location, organizer, all-day, recurring
- Support for attendees (required/optional) and reminders

**Advanced Search & Filtering:**
- `--folder` - Search specific folders (can specify multiple)
- `--unread` - Filter unread messages only
- `--from-date` / `--to-date` - Arbitrary date ranges
- `--filter-email` - Filter by participant email (can specify multiple)
- `--filter-domain` - Filter by email domain (can specify multiple)
- `folders` command - List all available Outlook folders

**Service Architecture:**
- Modular design with separate services:
  - `SearchService` - Email search and filtering
  - `ViewerService` - Terminal display with Rich
  - `ExportService` - Markdown conversion
  - `ComposeService` - Email sending operations
  - `CalendarService` - Calendar management
- Clean `Email` and `CalendarEvent` data models
- Reusable utility functions for parsing and formatting

**AI Agent Integration:**
- Documented usage with Claude Code
- Hermes Agent configuration examples
- OpenClaw automation examples
- Workflow examples for AI-assisted email management

### 🔄 Changed

**Breaking Changes:**
- New command structure using subcommands
  - Old: `python outlook_to_markdown.py --full`
  - New: `python outlook.py export --output ./emails`
- Entry point renamed from `outlook_to_markdown.py` to `outlook.py`
- `--output` is now required for export command

**Improvements:**
- Better error handling and user feedback
- Consistent date parsing across all commands
- UTF-8 support for Windows console
- Performance optimizations for calendar queries
- Cleaner markdown output with better HTML stripping

### 📚 Documentation

- Complete README rewrite with comprehensive examples
- AI agent integration guide (Claude Code, Hermes, OpenClaw)
- Migration guide from v1.0
- Full CLI reference tables
- Architecture documentation

### 🐛 Fixed

- Unicode encoding issues on Windows console
- Date filter edge cases in email search
- Calendar recurring event performance issues
- Thread consolidation with special characters
- Signature stripping false positives

### 🔧 Deprecated

- `outlook_to_markdown.py` - Still available but no longer maintained
  - Use `python outlook.py export` instead

---

## [1.0.0] - 2026-05-10

### Initial Release

- Basic email extraction from Outlook
- Export to Obsidian-flavored markdown
- Thread consolidation by subject
- Filtering by email, domain, and keyword
- Incremental extraction with state tracking
