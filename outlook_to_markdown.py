#!/usr/bin/env python
"""
Automated Outlook Email Extraction & Obsidian Markdown Conversion

Connects to Outlook Classic, extracts emails from Inbox and Sent Items,
consolidates threads, and converts to Obsidian-flavored markdown format.

Features:
- Incremental extraction (tracks last run time)
- Thread consolidation (groups emails by subject)
- Strips quoted replies, images, and attachments
- Outputs Obsidian-flavored markdown with frontmatter, callouts, tags, and wikilinks

Usage:
    python outlook_to_markdown.py                       # Incremental extraction
    python outlook_to_markdown.py --full                # Full re-extraction (last 7 days)
    python outlook_to_markdown.py --output /path/to/dir # Custom output directory
    python outlook_to_markdown.py --full -o /path/to/dir
"""
import os
import re
import json
import hashlib
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

import win32com.client as win32
from markdownify import markdownify as md
from bs4 import BeautifulSoup


# Configuration
DEFAULT_LOOKBACK_DAYS = 7
STATE_FILE = "extraction_state.json"
OUTPUT_DIR = "markdown"


def load_extraction_state(state_path):
    """Load the last extraction state from JSON file."""
    if state_path.exists():
        with open(state_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def save_extraction_state(state_path, last_extraction, emails_processed, threads_created):
    """Save extraction state to JSON file."""
    state = {
        "last_extraction": last_extraction.isoformat(),
        "emails_processed": emails_processed,
        "threads_created": threads_created,
        "updated_at": datetime.now().isoformat()
    }
    with open(state_path, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=2)
    return state


def connect_to_outlook():
    """Establish connection to Outlook via COM."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        print(f"[ERROR] Failed to connect to Outlook: {e}")
        print("Make sure Outlook is installed and configured.")
        raise


def format_outlook_date(dt):
    """Format datetime for Outlook Restrict filter."""
    return dt.strftime("%m/%d/%Y %H:%M %p")


def extract_email_address(full_address):
    """Extract just the email address from a full address string."""
    if not full_address:
        return None
    # Match email pattern
    match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', str(full_address))
    return match.group(0).lower() if match else None


def extract_display_name(full_address):
    """Extract display name from address, cleaning up Exchange paths."""
    if not full_address:
        return None
    name = str(full_address)
    # Remove Exchange path if present
    if '/CN=' in name.upper():
        # Try to get the last CN part
        parts = name.split('/')
        for part in reversed(parts):
            if part.upper().startswith('CN='):
                name = part[3:]
                break
    # Clean up common patterns
    name = re.sub(r'\s*<[^>]+>\s*', '', name)  # Remove <email>
    name = re.sub(r'\([^)]*\)', '', name)  # Remove (EXT) etc
    name = name.replace(';', ',').strip()
    return name if name else None


def extract_recipients(recipients_obj):
    """Extract recipient information from Outlook Recipients object."""
    if not recipients_obj:
        return "N/A", []
    try:
        recipients = []
        names = []
        for i in range(1, recipients_obj.Count + 1):
            recip = recipients_obj.Item(i)
            name = recip.Name
            email = recip.Address
            if name:
                clean_name = extract_display_name(name)
                if clean_name:
                    names.append(clean_name)
            if name and email:
                recipients.append(f"{name} <{email}>")
            elif email:
                recipients.append(email)
            elif name:
                recipients.append(name)
        return "; ".join(recipients) if recipients else "N/A", names
    except Exception:
        return "N/A", []


def extract_email_data(mail_item, is_sent=False):
    """Extract data from an Outlook MailItem."""
    try:
        # Basic properties
        subject = mail_item.Subject or "No Subject"

        # Sender info
        try:
            sender_name = mail_item.SenderName or ""
            sender_email = mail_item.SenderEmailAddress or ""
            sender_clean = extract_display_name(sender_name) or sender_email
            sender_domain = extract_email_address(sender_email)
            sender_domain = sender_domain.split('@')[1] if sender_domain and '@' in sender_domain else None
            if sender_name and sender_email:
                sender = f"{sender_name} <{sender_email}>"
            else:
                sender = sender_email or sender_name or "Unknown"
        except Exception:
            sender = "Unknown"
            sender_clean = "Unknown"
            sender_domain = None

        # Recipients
        try:
            to, to_names = extract_recipients(mail_item.Recipients)
        except Exception:
            to = "N/A"
            to_names = []

        # CC
        try:
            cc = mail_item.CC if mail_item.CC else None
        except Exception:
            cc = None

        # Date
        try:
            if is_sent:
                date = mail_item.SentOn
            else:
                date = mail_item.ReceivedTime
        except Exception:
            date = None

        # Body - prefer HTML
        try:
            html_body = mail_item.HTMLBody
        except Exception:
            html_body = None

        try:
            text_body = mail_item.Body
        except Exception:
            text_body = None

        return {
            'subject': subject,
            'sender': sender,
            'sender_clean': sender_clean,
            'sender_domain': sender_domain,
            'to': to,
            'to_names': to_names,
            'cc': cc,
            'date': date,
            'html_body': html_body,
            'text_body': text_body,
            'is_sent': is_sent
        }
    except Exception as e:
        return {'error': str(e), 'subject': 'Unknown'}


def extract_emails_from_folder(folder, since_date, is_sent=False):
    """Extract emails from an Outlook folder since a given date."""
    emails = []

    try:
        items = folder.Items

        # Apply date filter
        if is_sent:
            date_filter = f"[SentOn] > '{format_outlook_date(since_date)}'"
        else:
            date_filter = f"[ReceivedTime] > '{format_outlook_date(since_date)}'"

        filtered_items = items.Restrict(date_filter)

        # Sort by date
        if is_sent:
            filtered_items.Sort("[SentOn]", False)
        else:
            filtered_items.Sort("[ReceivedTime]", False)

        count = filtered_items.Count
        print(f"  Found {count} email(s)")

        for i in range(1, count + 1):
            try:
                mail = filtered_items.Item(i)
                # Only process mail items (not meeting requests, etc.)
                if mail.Class == 43:  # 43 = olMail
                    email_data = extract_email_data(mail, is_sent)
                    if 'error' not in email_data:
                        # Skip automatic replies
                        if is_auto_reply(email_data.get('subject', '')):
                            continue
                        emails.append(email_data)
            except Exception as e:
                print(f"  [WARN] Skipping email {i}: {e}")
                continue

    except Exception as e:
        print(f"  [ERROR] Failed to extract from folder: {e}")

    return emails


def is_auto_reply(subject):
    """Check if email is an automatic reply that should be skipped."""
    if not subject:
        return False
    auto_patterns = [
        r'^Automatic reply:',
        r'^Out of Office:',
        r'^Out of Office Re:',
        r'^Auto:',
        r'^Autoreply:',
        r'^Auto-reply:',
        r'^OOO:',
        r'^Absence:',
    ]
    for pattern in auto_patterns:
        if re.match(pattern, subject, re.IGNORECASE):
            return True
    return False


def normalize_subject(subject):
    """Normalize subject by removing RE:/FW: prefixes for thread grouping."""
    if not subject:
        return "No Subject"
    normalized = re.sub(r'^(RE|FW|Fwd|Re|Fw):\s*', '', subject, flags=re.IGNORECASE)
    normalized = re.sub(r'^(RE|FW|Fwd|Re|Fw):\s*', '', normalized, flags=re.IGNORECASE)
    return normalized.strip() or "No Subject"


def sanitize_filename(filename, max_length=50):
    """Sanitize filename by removing invalid characters and limiting length."""
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    filename = re.sub(r'\s+', ' ', filename)
    filename = filename.strip('. ')
    if not filename:
        return "untitled"
    if len(filename) > max_length:
        filename = filename[:max_length].strip()
    return filename


def extract_new_content(html_body, text_body):
    """Extract only new content from email, stripping quoted replies."""
    if html_body:
        soup = BeautifulSoup(html_body, 'html.parser')

        # Remove script, style, meta, head tags
        for tag in soup(['script', 'style', 'meta', 'head']):
            tag.decompose()

        # Remove blockquotes (quoted replies)
        for blockquote in soup.find_all('blockquote'):
            blockquote.decompose()

        # Convert to markdown first
        content = md(str(soup), heading_style="ATX")

        # Now strip quoted content using text patterns
        # Look for Outlook-style quote headers: "From: ... Sent: ... Subject:"
        patterns = [
            r'\n---\s*\n\*\*From:\*\*.*?\*\*Subject:\*\*.*?(?=\n|$).*',
            r'\n\*\*From:\*\*.*?\*\*Sent:\*\*.*?\*\*Subject:\*\*.*',
            r'\nFrom:.*?\nSent:.*?\nSubject:.*',
            r'\n-{3,}\s*\nFrom:.*',
            r'\nOn .+wrote:.*',
        ]

        for pattern in patterns:
            match = re.search(pattern, content, re.IGNORECASE | re.DOTALL)
            if match:
                content = content[:match.start()]
                break

    elif text_body:
        # For plain text, strip quoted lines
        lines = text_body.split('\n')
        new_lines = []
        for line in lines:
            if line.strip().startswith('>'):
                break
            if re.match(r'^On .+ wrote:$', line.strip()):
                break
            if re.match(r'^From:', line.strip()):
                break
            new_lines.append(line)
        content = '\n'.join(new_lines)
    else:
        content = "*No content available*"

    # Remove image markdown (logos, signature images, tracking pixels)
    content = re.sub(r'\[!\[[^\]]*\]\([^\)]*\)\]\([^\)]*\)', '', content)
    content = re.sub(r'!\[[^\]]*\]\([^\)]*\)', '', content)

    # Clean up
    content = re.sub(r'\n{3,}', '\n\n', content)
    return content.strip()


def format_date_wikilink(date_obj):
    """Format date as Obsidian wikilink for daily notes."""
    if not date_obj:
        return "Unknown"
    try:
        return f"[[{date_obj.strftime('%Y-%m-%d')}]]"
    except Exception:
        return "Unknown"


def determine_thread_tags(thread_emails):
    """Determine appropriate tags for a thread based on its content."""
    tags = ["email/thread"]

    has_sent = any(e.get('is_sent') for e in thread_emails)
    has_received = any(not e.get('is_sent') for e in thread_emails)

    if has_sent and has_received:
        tags.append("email/conversation")
    elif has_sent:
        tags.append("email/sent")
    else:
        tags.append("email/received")

    # Check for external contacts (customize INTERNAL_DOMAINS for your organization)
    # Example: ['yourcompany.com', 'yourcompany.org']
    internal_domains = []  # Add your organization's domains here
    if internal_domains:
        has_external = False
        for email in thread_emails:
            domain = email.get('sender_domain')
            if domain and domain.lower() not in internal_domains:
                has_external = True
                break

        if has_external:
            tags.append("contact/external")
        else:
            tags.append("contact/internal")

    return tags


def collect_participants(thread_emails):
    """Collect unique participants from a thread."""
    participants = set()
    for email in thread_emails:
        sender = email.get('sender_clean')
        if sender and sender != "Unknown":
            participants.add(sender)
        for name in email.get('to_names', []):
            if name:
                participants.add(name)
    return sorted(list(participants))


def create_thread_markdown(thread_emails, thread_subject):
    """Create Obsidian-flavored markdown content for an email thread."""
    # Sort emails by date
    sorted_emails = sorted(
        [e for e in thread_emails if e.get('date')],
        key=lambda x: x['date'] if x['date'] else datetime.min
    )

    if not sorted_emails:
        return None

    # Extract metadata for frontmatter
    first_date = sorted_emails[0]['date']
    last_date = sorted_emails[-1]['date']
    participants = collect_participants(sorted_emails)
    tags = determine_thread_tags(sorted_emails)
    message_count = len(sorted_emails)

    # Build YAML frontmatter
    safe_title = thread_subject.replace('"', "'")
    markdown = "---\n"
    markdown += f'title: "{safe_title}"\n'

    if first_date:
        markdown += f"date: {first_date.strftime('%Y-%m-%d')}\n"

    if first_date and last_date and first_date.date() != last_date.date():
        markdown += f"date_end: {last_date.strftime('%Y-%m-%d')}\n"

    markdown += f"message_count: {message_count}\n"

    if participants:
        markdown += "participants:\n"
        for p in participants[:10]:  # Limit to 10 participants
            safe_name = p.replace('"', "'")
            markdown += f'  - "{safe_name}"\n'

    markdown += "tags:\n"
    for tag in tags:
        markdown += f"  - {tag}\n"

    markdown += "---\n\n"

    # Thread summary callout
    date_range = format_date_wikilink(first_date)
    if first_date and last_date and first_date.date() != last_date.date():
        date_range = f"{format_date_wikilink(first_date)} to {format_date_wikilink(last_date)}"

    markdown += f"> [!info] Thread Summary\n"
    markdown += f"> **{message_count} message(s)** from {date_range}\n"
    markdown += f"> **Participants:** {', '.join(participants[:5])}"
    if len(participants) > 5:
        markdown += f" +{len(participants) - 5} more"
    markdown += "\n\n"

    # Process each message
    for i, email in enumerate(sorted_emails, 1):
        try:
            date_str = email['date'].strftime("%Y-%m-%d %H:%M") if email['date'] else "Unknown"
            date_link = format_date_wikilink(email['date'])
        except Exception:
            date_str = "Unknown"
            date_link = "Unknown"

        direction = "SENT" if email.get('is_sent') else "RECEIVED"
        direction_tag = "#email/sent" if email.get('is_sent') else "#email/received"
        sender_clean = email.get('sender_clean', 'Unknown')

        markdown += f"## Message {i} ({direction}) ^msg-{i}\n\n"
        markdown += f"**From:** {sender_clean}  \n"
        markdown += f"**Date:** {date_link} {email['date'].strftime('%H:%M') if email.get('date') else ''}  \n"
        markdown += f"{direction_tag}\n\n"

        # Check for external email warning in content
        content = extract_new_content(email.get('html_body'), email.get('text_body'))

        # Extract and format CAUTION warning as callout
        caution_match = re.search(
            r'CAUTION:?\s*This email originated from outside.*?(?=\n\n|\n[A-Z]|$)',
            content,
            re.IGNORECASE | re.DOTALL
        )
        if caution_match:
            markdown += "> [!warning] External Email\n"
            markdown += "> This email originated from outside of your organization.\n\n"
            # Remove the caution from content
            content = content[:caution_match.start()] + content[caution_match.end():]
            content = re.sub(r'_+\s*', '', content)  # Remove underscores separator

        # Clean up content
        content = content.strip()
        if content:
            markdown += content + "\n"
        else:
            markdown += "*No content*\n"

        markdown += "\n---\n\n"

    # Add hidden AI processing hints
    markdown += "%%\n"
    markdown += "ai-hints:\n"
    markdown += f"  thread_type: email_conversation\n"
    markdown += f"  message_count: {message_count}\n"
    markdown += f"  has_external: {'yes' if 'contact/external' in tags else 'no'}\n"
    markdown += f"  direction: {'mixed' if 'email/conversation' in tags else ('outbound' if 'email/sent' in tags else 'inbound')}\n"
    markdown += "%%\n"

    return markdown


def main():
    parser = argparse.ArgumentParser(description='Extract Outlook emails to Obsidian Markdown')
    parser.add_argument('--full', action='store_true',
                        help=f'Full extraction (last {DEFAULT_LOOKBACK_DAYS} days)')
    parser.add_argument('--output', '-o', type=str,
                        help='Output directory for markdown files')
    args = parser.parse_args()

    # Setup paths
    script_dir = Path(__file__).parent
    state_path = script_dir / STATE_FILE
    output_dir = Path(args.output) if args.output else script_dir / OUTPUT_DIR
    output_dir.mkdir(exist_ok=True)

    print("=" * 60)
    print("Outlook Email to Obsidian Markdown Conversion")
    print("=" * 60)

    # Determine extraction start date
    state = load_extraction_state(state_path)

    if args.full or state is None:
        since_date = datetime.now() - timedelta(days=DEFAULT_LOOKBACK_DAYS)
        print(f"\nMode: {'Full extraction' if args.full else 'First run'}")
        print(f"Extracting emails from last {DEFAULT_LOOKBACK_DAYS} days")
    else:
        since_date = datetime.fromisoformat(state['last_extraction'])
        print(f"\nMode: Incremental extraction")
        print(f"Last extraction: {state['last_extraction']}")

    print(f"Since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")

    # Connect to Outlook
    print("\nConnecting to Outlook...")
    outlook, namespace = connect_to_outlook()
    print("  Connected successfully")

    # Extract from Inbox
    print("\nExtracting from Inbox...")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    inbox_emails = extract_emails_from_folder(inbox, since_date, is_sent=False)

    # Extract from Sent Items
    print("\nExtracting from Sent Items...")
    sent_items = namespace.GetDefaultFolder(5)  # 5 = Sent Items
    sent_emails = extract_emails_from_folder(sent_items, since_date, is_sent=True)

    # Combine all emails
    all_emails = inbox_emails + sent_emails
    total_count = len(all_emails)

    if total_count == 0:
        print("\nNo new emails found.")
        save_extraction_state(state_path, datetime.now(), 0, 0)
        return

    print(f"\nTotal emails extracted: {total_count}")

    # Group by thread
    print("\nGrouping emails into threads...")
    threads = defaultdict(list)
    for email in all_emails:
        thread_key = normalize_subject(email['subject'])
        threads[thread_key].append(email)

    print(f"Found {len(threads)} unique thread(s)")

    # Create/update markdown files
    print("\nCreating Obsidian markdown files...")
    success_count = 0

    for thread_subject, emails in threads.items():
        filename = sanitize_filename(thread_subject) + ".md"
        output_path = output_dir / filename

        # Handle duplicate filenames
        if output_path.exists():
            file_hash = hashlib.md5(thread_subject.encode()).hexdigest()[:8]
            filename = sanitize_filename(thread_subject) + f"_{file_hash}.md"
            output_path = output_dir / filename

        markdown_content = create_thread_markdown(emails, thread_subject)

        if markdown_content:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            print(f"  [OK] {filename} ({len(emails)} email(s))")
            success_count += 1

    # Save extraction state
    extraction_time = datetime.now()
    save_extraction_state(state_path, extraction_time, total_count, success_count)

    # Summary
    print(f"\n{'=' * 60}")
    print("Extraction complete!")
    print(f"  Emails processed: {total_count}")
    print(f"  Threads created: {success_count}")
    print(f"  Output directory: {output_dir}")
    print(f"  State saved to: {state_path}")
    print(f"\nNext run will extract emails after: {extraction_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)


if __name__ == "__main__":
    main()
