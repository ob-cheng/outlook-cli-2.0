# Changelog

All notable changes to Outlook CLI will be documented in this file.

## [1.0.1] - 2026-05-11

### ✨ Added

**JSON Export for AI Ingestion:**

- `--format json` option for export command - output as JSON instead of markdown
- `--batch` option - combine all emails into a single JSON file (most token-efficient)
- `--stdout` option - output JSON directly to terminal for AI agent pipelines
- Lean JSON structure optimized for LLM consumption (~15-20% fewer tokens than markdown)
- `to_json_data()` method in ExportService for programmatic access

**Task/Todo Management:**

- `tasks list` - List tasks with filters (status, priority, due date, category)
- `tasks read` - View task details
- `tasks create` - Create tasks with due date, priority, reminder
- `tasks complete` - Mark tasks as complete
- `tasks delete` - Delete tasks
- Full JSON output support for all task commands
- Task model and TaskService for programmatic access

### 🐛 Fixed

- Date handling for unset dates (4501-01-01 now shows as null)
- Timezone issue for task due dates (set at noon to avoid boundary issues)

### 📚 Documentation

- Updated README with JSON export and task commands
- Updated SKILL.md with new options for AI agents
- Updated EXAMPLES.md with AI workflow examples
- Updated CLAUDE.md quick reference

---

## [1.0.0] - 2026-05-10

### Initial Release

- Search and filter emails by folder, sender, date range, read status
- Send emails with CC, BCC, attachments, HTML formatting
- Reply and forward messages
- Save drafts for review before sending
- Export email threads to Obsidian markdown (with incremental sync)
- List, create, and delete calendar events
- JSON output for agent/automation integration
- Rich terminal UI with styled tables and panels
