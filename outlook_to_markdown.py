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
    python outlook_to_markdown.py --full --days 14      # Full re-extraction (last 14 days)
    python outlook_to_markdown.py --output /path/to/dir # Custom output directory
    python outlook_to_markdown.py --full -d 30 -o /path/to/dir
    python outlook_to_markdown.py --unread              # Unread messages only
    python outlook_to_markdown.py --folder Inbox        # Specific folder
    python outlook_to_markdown.py --folder "Project X"  # Custom folder by name
    python outlook_to_markdown.py --list-folders        # List all available folders
    python outlook_to_markdown.py --from-date 2024-01-01 --to-date 2024-01-31  # Date range
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


# Folder type constants for GetDefaultFolder
DEFAULT_FOLDERS = {
    'inbox': 6,
    'sent': 5,
    'sent items': 5,
    'drafts': 16,
    'deleted': 3,
    'deleted items': 3,
    'outbox': 4,
    'junk': 23,
    'junk email': 23,
    'archive': 0,  # Special handling needed
    'calendar': 9,
    'contacts': 10,
    'tasks': 13,
    'notes': 12,
}


def list_all_folders(namespace, indent=0):
    """Recursively list all folders in the mailbox."""
    folders = []

    def traverse_folders(folder_collection, path="", level=0):
        for i in range(1, folder_collection.Count + 1):
            try:
                folder = folder_collection.Item(i)
                folder_path = f"{path}/{folder.Name}" if path else folder.Name
                folders.append({
                    'name': folder.Name,
                    'path': folder_path,
                    'level': level,
                    'count': folder.Items.Count if hasattr(folder, 'Items') else 0
                })
                # Recurse into subfolders
                if folder.Folders.Count > 0:
                    traverse_folders(folder.Folders, folder_path, level + 1)
            except Exception:
                continue

    # Traverse all accounts/stores
    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            folders.append({
                'name': store.Name,
                'path': store.Name,
                'level': 0,
                'count': 0,
                'is_store': True
            })
            traverse_folders(store.Folders, store.Name, 1)
        except Exception:
            continue

    return folders


def find_folder_by_name(namespace, folder_name):
    """Find a folder by name (case-insensitive). Returns the folder object."""
    folder_lower = folder_name.lower().strip()

    # Check if it's a default folder
    if folder_lower in DEFAULT_FOLDERS:
        folder_id = DEFAULT_FOLDERS[folder_lower]
        if folder_id > 0:
            try:
                return namespace.GetDefaultFolder(folder_id)
            except Exception:
                pass

    # Search by name in all folders
    def search_folders(folder_collection, target_name):
        for i in range(1, folder_collection.Count + 1):
            try:
                folder = folder_collection.Item(i)
                if folder.Name.lower() == target_name:
                    return folder
                # Search subfolders
                if folder.Folders.Count > 0:
                    result = search_folders(folder.Folders, target_name)
                    if result:
                        return result
            except Exception:
                continue
        return None

    # Search all stores
    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            if store.Name.lower() == folder_lower:
                return store
            result = search_folders(store.Folders, folder_lower)
            if result:
                return result
        except Exception:
            continue

    return None


def find_folder_by_path(namespace, folder_path):
    """Find a folder by full path (e.g., 'Account/Inbox/Subfolder')."""
    parts = [p.strip() for p in folder_path.split('/') if p.strip()]
    if not parts:
        return None

    # Start from the store (first part)
    current = None
    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            if store.Name.lower() == parts[0].lower():
                current = store
                break
        except Exception:
            continue

    if not current:
        # Try finding just by name
        return find_folder_by_name(namespace, parts[-1])

    # Navigate through the path
    for part in parts[1:]:
        found = False
        for j in range(1, current.Folders.Count + 1):
            try:
                subfolder = current.Folders.Item(j)
                if subfolder.Name.lower() == part.lower():
                    current = subfolder
                    found = True
                    break
            except Exception:
                continue
        if not found:
            return None

    return current


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


def extract_all_email_addresses(field_string):
    """Extract all email addresses from a To/CC field string."""
    if not field_string:
        return []
    matches = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', str(field_string))
    return [m.lower() for m in matches]


def matches_keyword_filter(email_data, filter_keyword):
    """Check if email subject or body contains the keyword (case-insensitive)."""
    if not filter_keyword:
        return True
    keyword_lower = filter_keyword.lower()
    subject = (email_data.get('subject') or '').lower()
    text_body = (email_data.get('text_body') or '').lower()
    return keyword_lower in subject or keyword_lower in text_body


def matches_email_filter(email_data, filter_emails=None, filter_domains=None, max_recipients=20, debug=False):
    """Check if email involves any of the filtered email addresses or domains in From/To/CC.

    Args:
        email_data: Email data dictionary
        filter_emails: List of email addresses to filter by
        filter_domains: List of email domains to filter by (e.g., ['accessinfinity.com'])
        max_recipients: Maximum total recipients (To + CC) to include. Emails with more
                        recipients are considered mass distribution and excluded.
        debug: Print debug output
    """
    if not filter_emails and not filter_domains:
        return True

    # Skip mass distribution emails
    to_emails = email_data.get('to_emails', [])
    cc_emails = email_data.get('cc_emails', [])
    total_recipients = len(to_emails) + len(cc_emails)

    if total_recipients > max_recipients:
        if debug:
            print(f"    [DEBUG] Skipped: {total_recipients} recipients (mass distribution)")
        return False

    filter_email_set = {e.lower().strip() for e in filter_emails} if filter_emails else set()
    filter_domain_set = {d.lower().strip().lstrip('@') for d in filter_domains} if filter_domains else set()

    def matches_domain(email_addr):
        """Check if email address matches any filter domain."""
        if not email_addr or '@' not in email_addr:
            return False
        domain = email_addr.lower().split('@')[1]
        return domain in filter_domain_set

    # Check sender SMTP email
    sender_smtp = email_data.get('sender_smtp')
    if sender_smtp:
        if sender_smtp.lower() in filter_email_set:
            if debug:
                print(f"    [DEBUG] Matched sender email: {sender_smtp}")
            return True
        if matches_domain(sender_smtp):
            if debug:
                print(f"    [DEBUG] Matched sender domain: {sender_smtp}")
            return True

    # Check To recipient emails
    for addr in to_emails:
        if addr:
            if addr.lower() in filter_email_set:
                if debug:
                    print(f"    [DEBUG] Matched TO email: {addr}")
                return True
            if matches_domain(addr):
                if debug:
                    print(f"    [DEBUG] Matched TO domain: {addr}")
                return True

    # Check CC recipient emails
    for addr in cc_emails:
        if addr:
            if addr.lower() in filter_email_set:
                if debug:
                    print(f"    [DEBUG] Matched CC email: {addr}")
                return True
            if matches_domain(addr):
                if debug:
                    print(f"    [DEBUG] Matched CC domain: {addr}")
                return True

    return False


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
    name = name.strip()

    # Convert "Last, First" to "First Last"
    if ',' in name:
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2 and parts[0] and parts[1]:
            # Check if it looks like "Last, First" (not "Title, Name" etc.)
            # Both parts should start with capital letter and be short (name-like)
            if (parts[0][0].isupper() and parts[1][0].isupper() and
                len(parts[0]) < 20 and len(parts[1]) < 20):
                name = f"{parts[1]} {parts[0]}"

    return name if name else None


def get_smtp_address(recip):
    """Get SMTP email address from recipient, handling Exchange format."""
    try:
        # Try to get SMTP address from Exchange user
        addr_entry = recip.AddressEntry
        if addr_entry.Type == "EX":
            # Exchange user - get SMTP address
            exch_user = addr_entry.GetExchangeUser()
            if exch_user:
                return exch_user.PrimarySmtpAddress
        # Fallback to direct address
        return recip.Address
    except Exception:
        return recip.Address


def extract_recipients(recipients_obj):
    """Extract recipient information from Outlook Recipients object."""
    if not recipients_obj:
        return "N/A", [], []
    try:
        recipients = []
        names = []
        emails = []
        for i in range(1, recipients_obj.Count + 1):
            recip = recipients_obj.Item(i)
            name = recip.Name
            email = get_smtp_address(recip)
            if name:
                clean_name = extract_display_name(name)
                if clean_name:
                    names.append(clean_name)
            if email:
                smtp_email = extract_email_address(email)
                if smtp_email:
                    emails.append(smtp_email)
            if name and email:
                recipients.append(f"{name} <{email}>")
            elif email:
                recipients.append(email)
            elif name:
                recipients.append(name)
        return "; ".join(recipients) if recipients else "N/A", names, emails
    except Exception:
        return "N/A", [], []


def get_sender_smtp_address(mail_item):
    """Get SMTP email address of sender, handling Exchange format."""
    try:
        # Try to get from sender's AddressEntry
        sender = mail_item.Sender
        if sender and sender.Type == "EX":
            exch_user = sender.GetExchangeUser()
            if exch_user:
                return exch_user.PrimarySmtpAddress
        # Fallback to direct address
        return mail_item.SenderEmailAddress
    except Exception:
        return mail_item.SenderEmailAddress


def extract_email_data(mail_item, is_sent=False):
    """Extract data from an Outlook MailItem."""
    try:
        # Basic properties
        subject = mail_item.Subject or "No Subject"

        # Sender info
        try:
            sender_name = mail_item.SenderName or ""
            sender_email = get_sender_smtp_address(mail_item) or ""
            sender_smtp = extract_email_address(sender_email)
            sender_clean = extract_display_name(sender_name) or sender_email
            sender_domain = sender_smtp.split('@')[1] if sender_smtp and '@' in sender_smtp else None
            if sender_name and sender_email:
                sender = f"{sender_name} <{sender_email}>"
            else:
                sender = sender_email or sender_name or "Unknown"
        except Exception:
            sender = "Unknown"
            sender_clean = "Unknown"
            sender_domain = None
            sender_smtp = None

        # Recipients
        try:
            to, to_names, to_emails = extract_recipients(mail_item.Recipients)
        except Exception:
            to = "N/A"
            to_names = []
            to_emails = []

        # CC - need to extract emails from CC recipients too
        try:
            cc = mail_item.CC if mail_item.CC else None
            cc_emails = []
            # Try to get CC recipients with SMTP addresses
            for i in range(1, mail_item.Recipients.Count + 1):
                recip = mail_item.Recipients.Item(i)
                if recip.Type == 2:  # 2 = CC recipient
                    smtp = get_smtp_address(recip)
                    email = extract_email_address(smtp)
                    if email:
                        cc_emails.append(email)
        except Exception:
            cc = None
            cc_emails = []

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
            'sender_smtp': sender_smtp,
            'to': to,
            'to_names': to_names,
            'to_emails': to_emails,
            'cc': cc,
            'cc_emails': cc_emails,
            'date': date,
            'html_body': html_body,
            'text_body': text_body,
            'is_sent': is_sent
        }
    except Exception as e:
        return {'error': str(e), 'subject': 'Unknown'}


def extract_emails_from_folder(folder, since_date, until_date=None, is_sent=False,
                                filter_emails=None, filter_domains=None, filter_keyword=None,
                                unread_only=False):
    """Extract emails from an Outlook folder within a date range.

    Args:
        folder: Outlook folder object
        since_date: Start date (extract emails after this date)
        until_date: End date (extract emails before this date), optional
        is_sent: Whether this is a sent items folder
        filter_emails: List of email addresses to filter by
        filter_domains: List of domains to filter by
        filter_keyword: Keyword to search in subject/body
        unread_only: If True, only extract unread messages
    """
    emails = []

    try:
        items = folder.Items

        # Build date filter
        date_field = "[SentOn]" if is_sent else "[ReceivedTime]"
        date_filter = f"{date_field} > '{format_outlook_date(since_date)}'"

        if until_date:
            date_filter += f" AND {date_field} < '{format_outlook_date(until_date)}'"

        # Add unread filter if requested
        if unread_only:
            date_filter += " AND [UnRead] = True"

        filtered_items = items.Restrict(date_filter)

        # Sort by date
        if is_sent:
            filtered_items.Sort("[SentOn]", False)
        else:
            filtered_items.Sort("[ReceivedTime]", False)

        count = filtered_items.Count
        print(f"  Found {count} email(s)")

        matched_count = 0
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
                        # Apply participant email filter
                        if (filter_emails or filter_domains) and not matches_email_filter(email_data, filter_emails, filter_domains):
                            continue
                        # Apply keyword filter
                        if filter_keyword and not matches_keyword_filter(email_data, filter_keyword):
                            continue
                        emails.append(email_data)
                        matched_count += 1
            except Exception as e:
                print(f"  [WARN] Skipping email {i}: {e}")
                continue

        if filter_emails or filter_domains or filter_keyword:
            print(f"  Matched {matched_count} email(s) with filter")

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


def strip_signature(content):
    """Remove email signature noise (logos, contact tables) while keeping names."""
    lines = content.split('\n')
    cleaned_lines = []

    for line in lines:
        # Skip lines containing SafeLinks URLs (Microsoft URL protection)
        if 'safelinks.protection.outlook.com' in line:
            continue

        # Skip signature table lines with contact info patterns (W:, E:, LinkedIn, etc.)
        if '|' in line and re.search(r'\b(W:|E:|P:|M:|LinkedIn|Twitter|Facebook|www\.|\.com)\b', line, re.IGNORECASE):
            continue

        # Extract name from signature table lines FIRST (before other table filters)
        # This catches lines like: | | --- | | **Sriram Iyer** | | Head of Engineering | |
        if '|' in line and '**' in line:
            # Extract bold text
            name_match = re.search(r'\*\*([^*]+)\*\*', line)
            if name_match:
                name = name_match.group(1)
                # Clean Unicode chars (zero-width space, non-breaking space)
                name = name.replace('​', '').replace(' ', ' ').strip()
                # Skip if it looks like a title
                if name and not re.search(r'(Head|Director|Manager|Engineer|President|CEO|CTO|VP|Chief|Senior|Lead|of)', name, re.IGNORECASE):
                    cleaned_lines.append(f'**{name}**')
            continue
        # Skip other table lines with job titles
        if '|' in line and re.search(r'\b(Head|Director|Manager|Engineer|President|CEO|CTO|VP|Chief|Senior|Lead|Principal)\b', line, re.IGNORECASE):
            continue

        # Skip lines that are mostly empty table cells or malformed tables
        if re.match(r'^[\s|]+$', line):
            continue

        # Skip lines with escaped underscores (horizontal rules in signatures)
        if re.match(r'^[\s|_\\]+$', line) or '\\_' * 5 in line:
            continue

        # Skip table separator lines (| --- | --- |)
        if re.match(r'^\|[\s\-|]+\|', line):
            continue

        # Skip empty table rows (|  |  | or | |)
        if line.startswith('|') and re.sub(r'[\s|]', '', line) == '':
            continue

        cleaned_lines.append(line)

    content = '\n'.join(cleaned_lines)

    # Clean up mailto: links to just email address
    content = re.sub(r'\[([^\]]+@[^\]]+)\]\(mailto:[^)]+\)', r'\1', content)

    return content


def extract_new_content(html_body, text_body):
    """Extract only new content from email, stripping quoted replies and signatures."""
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

        # Strip signatures
        content = strip_signature(content)

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
        content = strip_signature(content)
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

        content = extract_new_content(email.get('html_body'), email.get('text_body'))

        # Remove external email CAUTION warning (noise)
        content = re.sub(
            r'CAUTION:?\s*This email originated from outside.*?(?=\n\n|\n[A-Z]|$)',
            '',
            content,
            flags=re.IGNORECASE | re.DOTALL
        )
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


def parse_date(date_str):
    """Parse a date string in various formats."""
    formats = ['%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y']
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {date_str}. Use YYYY-MM-DD format.")


def main():
    parser = argparse.ArgumentParser(description='Extract Outlook emails to Obsidian Markdown')
    parser.add_argument('--full', action='store_true',
                        help=f'Full extraction (last {DEFAULT_LOOKBACK_DAYS} days by default)')
    parser.add_argument('--days', '-d', type=int, default=DEFAULT_LOOKBACK_DAYS,
                        help=f'Number of days to look back (default: {DEFAULT_LOOKBACK_DAYS})')
    parser.add_argument('--output', '-o', type=str,
                        help='Output directory for markdown files')
    parser.add_argument('--filter-emails', '-f', type=str,
                        help='Comma-separated email addresses to filter by (matches From/To/CC)')
    parser.add_argument('--filter-domains', '-D', type=str,
                        help='Comma-separated email domains to filter by (e.g., accessinfinity.com)')
    parser.add_argument('--filter-keyword', '-k', type=str,
                        help='Keyword to filter by (matches subject or body, case-insensitive)')
    parser.add_argument('--no-overwrite', action='store_true',
                        help='Skip creating files that already exist')

    # New features
    parser.add_argument('--unread', '-u', action='store_true',
                        help='Extract only unread messages')
    parser.add_argument('--folder', type=str, action='append',
                        help='Folder to extract from (can specify multiple). Use --list-folders to see available folders.')
    parser.add_argument('--list-folders', action='store_true',
                        help='List all available Outlook folders and exit')
    parser.add_argument('--from-date', type=str,
                        help='Start date for extraction (YYYY-MM-DD). Overrides --days.')
    parser.add_argument('--to-date', type=str,
                        help='End date for extraction (YYYY-MM-DD). Defaults to now.')

    args = parser.parse_args()

    # Setup paths
    script_dir = Path(__file__).parent
    state_path = script_dir / STATE_FILE
    output_dir = Path(args.output) if args.output else script_dir / OUTPUT_DIR

    print("=" * 60)
    print("Outlook Email to Obsidian Markdown Conversion")
    print("=" * 60)

    # Connect to Outlook first (needed for --list-folders)
    print("\nConnecting to Outlook...")
    outlook, namespace = connect_to_outlook()
    print("  Connected successfully")

    # Handle --list-folders
    if args.list_folders:
        print("\nAvailable folders:")
        print("-" * 60)
        folders = list_all_folders(namespace)
        for f in folders:
            indent = "  " * f['level']
            count_str = f" ({f['count']} items)" if f.get('count', 0) > 0 else ""
            store_marker = " [ACCOUNT]" if f.get('is_store') else ""
            # Handle encoding issues with emojis/special chars
            folder_name = f['name'].encode('ascii', 'replace').decode()
            print(f"{indent}{folder_name}{count_str}{store_marker}")
        print("-" * 60)
        print("\nUse --folder <name> to extract from a specific folder.")
        print("Examples:")
        print("  --folder Inbox")
        print("  --folder 'Sent Items'")
        print("  --folder Archive")
        print("  --folder 'My Custom Folder'")
        return

    output_dir.mkdir(exist_ok=True)

    # Determine extraction date range
    state = load_extraction_state(state_path)
    until_date = None

    # Handle explicit date range
    if args.from_date:
        since_date = parse_date(args.from_date)
        print(f"\nMode: Date range extraction")
        print(f"From: {since_date.strftime('%Y-%m-%d')}")
        if args.to_date:
            until_date = parse_date(args.to_date)
            # Set to end of day
            until_date = until_date.replace(hour=23, minute=59, second=59)
            print(f"To: {until_date.strftime('%Y-%m-%d')}")
        else:
            print(f"To: now")
    elif args.full or state is None:
        since_date = datetime.now() - timedelta(days=args.days)
        print(f"\nMode: {'Full extraction' if args.full else 'First run'}")
        print(f"Extracting emails from last {args.days} days")
    else:
        since_date = datetime.fromisoformat(state['last_extraction'])
        print(f"\nMode: Incremental extraction")
        print(f"Last extraction: {state['last_extraction']}")

    print(f"Since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")

    # Parse email filter if provided
    filter_emails = None
    if args.filter_emails:
        filter_emails = [e.strip() for e in args.filter_emails.split(',') if e.strip()]
        print(f"\nFiltering by email addresses:")
        for email in filter_emails:
            print(f"  - {email}")

    # Parse domain filter if provided
    filter_domains = None
    if args.filter_domains:
        filter_domains = [d.strip().lstrip('@') for d in args.filter_domains.split(',') if d.strip()]
        print(f"\nFiltering by email domains:")
        for domain in filter_domains:
            print(f"  - @{domain}")

    # Parse keyword filter if provided
    filter_keyword = args.filter_keyword
    if filter_keyword:
        print(f"\nFiltering by keyword: \"{filter_keyword}\"")

    # Unread filter
    if args.unread:
        print(f"\nFiltering: Unread messages only")

    # Determine which folders to extract from
    all_emails = []

    if args.folder:
        # Extract from specified folder(s)
        for folder_name in args.folder:
            print(f"\nExtracting from folder: {folder_name}...")

            # Try to find by path first, then by name
            if '/' in folder_name:
                folder = find_folder_by_path(namespace, folder_name)
            else:
                folder = find_folder_by_name(namespace, folder_name)

            if folder:
                # Determine if this is a sent folder
                is_sent = folder_name.lower() in ['sent', 'sent items', 'outbox']
                folder_emails = extract_emails_from_folder(
                    folder, since_date, until_date=until_date, is_sent=is_sent,
                    filter_emails=filter_emails, filter_domains=filter_domains,
                    filter_keyword=filter_keyword, unread_only=args.unread
                )
                all_emails.extend(folder_emails)
            else:
                print(f"  [ERROR] Folder not found: {folder_name}")
                print(f"  Use --list-folders to see available folders")
    else:
        # Default: Extract from Inbox and Sent Items
        print("\nExtracting from Inbox...")
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        inbox_emails = extract_emails_from_folder(
            inbox, since_date, until_date=until_date, is_sent=False,
            filter_emails=filter_emails, filter_domains=filter_domains,
            filter_keyword=filter_keyword, unread_only=args.unread
        )
        all_emails.extend(inbox_emails)

        print("\nExtracting from Sent Items...")
        sent_items = namespace.GetDefaultFolder(5)  # 5 = Sent Items
        sent_emails = extract_emails_from_folder(
            sent_items, since_date, until_date=until_date, is_sent=True,
            filter_emails=filter_emails, filter_domains=filter_domains,
            filter_keyword=filter_keyword, unread_only=args.unread
        )
        all_emails.extend(sent_emails)
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
        # Get the latest date from the thread for sorting
        thread_dates = [e['date'] for e in emails if e.get('date')]
        if thread_dates:
            latest_date = max(thread_dates)
            date_prefix = latest_date.strftime("%Y-%m-%d %H%M")
        else:
            date_prefix = "0000-00-00 0000"

        filename = f"{date_prefix} - {sanitize_filename(thread_subject)}.md"
        output_path = output_dir / filename

        # Handle duplicate filenames or skip if no-overwrite
        if output_path.exists():
            if args.no_overwrite:
                print(f"  [SKIP] {filename.encode('ascii', 'replace').decode()} (already exists)")
                continue
            file_hash = hashlib.md5(thread_subject.encode()).hexdigest()[:8]
            filename = f"{date_prefix} {sanitize_filename(thread_subject)}_{file_hash}.md"
            output_path = output_dir / filename
            if output_path.exists() and args.no_overwrite:
                print(f"  [SKIP] {filename.encode('ascii', 'replace').decode()} (already exists)")
                continue

        markdown_content = create_thread_markdown(emails, thread_subject)

        if markdown_content:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            print(f"  [OK] {filename.encode('ascii', 'replace').decode()} ({len(emails)} email(s))")
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
