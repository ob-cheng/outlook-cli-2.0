---
name: send-email
description: Send emails from Outlook. Use when user asks to "send email", "compose email", "reply to email", or "forward email".
---

# Send Email Skill

Compose and send emails, reply, or forward through Outlook.

## Send new email

```bash
cd "./"
python outlook.py send --to recipient@email.com --subject "Subject" --body "Message"
```

## Examples

Basic email:

```bash
python outlook.py send \
  --to user@example.com \
  --subject "Quick update" \
  --body "Here is the update."
```

With CC and attachment:

```bash
python outlook.py send \
  --to person@example.com \
  --cc manager@example.com \
  --subject "Report" \
  --body "Please see attached." \
  --attach report.pdf
```

Save as draft:

```bash
python outlook.py send \
  --to client@company.com \
  --subject "Proposal" \
  --body "Draft content..." \
  --draft
```

HTML formatted:

```bash
python outlook.py send \
  --to team@company.com \
  --subject "Announcement" \
  --body "<h1>News</h1><p>Update here.</p>" \
  --html
```

## Reply

```bash
python outlook.py reply <message-id> --body "Thanks!"
python outlook.py reply <message-id> --body "Team response" --all
python outlook.py reply <message-id> --body "Draft" --draft
```

## Forward

```bash
python outlook.py forward <message-id> --to colleague@example.com --body "FYI"
```

## Options

- `--to ADDRESS` - Recipient (required, comma-separated for multiple)
- `--subject TEXT` - Subject line (required for send)
- `--body TEXT` - Message body (required)
- `--cc ADDRESS` - CC recipients
- `--bcc ADDRESS` - BCC recipients
- `--attach PATH` - File attachment (can specify multiple)
- `--html` - Body is HTML formatted
- `--draft` - Save as draft instead of sending
- `--all` - Reply to all (for reply command)