#!/usr/bin/env python
"""
Convert Microsoft Outlook .msg files to Markdown format.
Consolidates email threads into single files.
"""
import os
import re
import hashlib
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import extract_msg
from markdownify import markdownify as md
from bs4 import BeautifulSoup


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


def normalize_subject(subject):
    """Normalize subject by removing RE:, FW:, Fwd: prefixes for thread grouping."""
    if not subject:
        return "No Subject"
    normalized = re.sub(r'^(RE|FW|Fwd|Re|Fw):\s*', '', subject, flags=re.IGNORECASE)
    normalized = re.sub(r'^(RE|FW|Fwd|Re|Fw):\s*', '', normalized, flags=re.IGNORECASE)
    return normalized.strip()


def format_email_address(address):
    """Format email address with name and email if available."""
    if hasattr(address, 'name') and hasattr(address, 'email'):
        if address.name and address.email:
            return f"{address.name} <{address.email}>"
        elif address.email:
            return address.email
        elif address.name:
            return address.name
    return str(address)


def format_recipients(recipients):
    """Format list of recipients."""
    if not recipients:
        return "N/A"
    if isinstance(recipients, (list, tuple)):
        return ", ".join(format_email_address(r) for r in recipients)
    return format_email_address(recipients)


def strip_quoted_content(html_content):
    """
    Remove quoted reply content from HTML email.
    Looks for common quote patterns and removes them.
    """
    if not html_content:
        return html_content

    soup = BeautifulSoup(html_content, 'html.parser')

    # Remove script, style, meta, head tags
    for tag in soup(['script', 'style', 'meta', 'head']):
        tag.decompose()

    # Common patterns for quoted content in HTML emails
    # 1. Elements with class containing 'gmail_quote', 'moz-cite', 'OutlookMessageHeader', etc.
    quote_classes = ['gmail_quote', 'gmail_extra', 'moz-cite-prefix', 'OutlookMessageHeader',
                     'MsoNormal', 'WordSection1']

    # 2. Blockquote elements (common for quoted replies)
    for blockquote in soup.find_all('blockquote'):
        blockquote.decompose()

    # 3. Divs that look like forwarded/replied content
    for div in soup.find_all('div'):
        div_text = div.get_text(strip=True)[:100] if div.get_text(strip=True) else ""
        # Check for "From:", "Sent:", "On ... wrote:" patterns indicating quoted content
        if re.search(r'^(From:|Sent:|Date:|Subject:|On\s+.+wrote:)', div_text, re.IGNORECASE):
            # Check if this is a header for quoted content
            next_sibling = div.find_next_sibling()
            if next_sibling and next_sibling.name == 'blockquote':
                div.decompose()

    # 4. Look for "On ... wrote:" pattern and remove everything after it
    body_text = str(soup)

    # Pattern for "On [date] [person] wrote:" which typically starts quoted content
    patterns = [
        r'On\s+\w+,\s+\w+\s+\d+,\s+\d+\s+at\s+\d+:\d+\s*[AP]M.*?wrote:.*',  # Gmail style
        r'On\s+\d+/\d+/\d+.*?wrote:.*',  # Short date style
        r'-{3,}\s*Original Message\s*-{3,}.*',  # Outlook style
        r'_{3,}\s*$.*',  # Underline separator
    ]

    return str(soup)


def extract_new_content(html_content):
    """
    Extract only the new content from an email, removing quoted replies.
    """
    if not html_content:
        return ""

    soup = BeautifulSoup(html_content, 'html.parser')

    # Remove script, style, meta, head tags
    for tag in soup(['script', 'style', 'meta', 'head']):
        tag.decompose()

    # Remove blockquotes (quoted replies)
    for blockquote in soup.find_all('blockquote'):
        blockquote.decompose()

    # Try to find the main content before quoted replies
    # Look for dividers that typically separate new content from quoted content
    html_str = str(soup)

    # Common divider patterns
    dividers = [
        r'<div[^>]*>\s*-{3,}\s*Original Message\s*-{3,}.*',
        r'<div[^>]*>\s*From:.*?Sent:.*?Subject:.*',
        r'On\s+\w{3},\s+\w{3}\s+\d+,\s+\d{4}\s+at\s+\d+:\d+.*?wrote:',
    ]

    for pattern in dividers:
        match = re.search(pattern, html_str, re.IGNORECASE | re.DOTALL)
        if match:
            html_str = html_str[:match.start()]
            break

    # Parse again and convert
    soup = BeautifulSoup(html_str, 'html.parser')
    return md(str(soup), heading_style="ATX")


def parse_msg_file(msg_path):
    """Parse a single .msg file and return email data."""
    try:
        msg = extract_msg.Message(msg_path)

        subject = msg.subject or "No Subject"
        sender = format_email_address(msg.sender) if msg.sender else "Unknown"
        to = format_recipients(msg.to)
        cc = format_recipients(msg.cc) if msg.cc else None
        date = msg.date if msg.date else None

        # Get body - prefer HTML
        body_html = msg.htmlBody
        body_text = msg.body

        # Extract new content only (strip quoted replies)
        if body_html:
            new_content = extract_new_content(body_html)
        elif body_text:
            # For plain text, try to strip quoted lines (lines starting with >)
            lines = body_text.split('\n')
            new_lines = []
            for line in lines:
                if line.strip().startswith('>'):
                    break  # Stop at first quoted line
                if re.match(r'^On .+ wrote:$', line.strip()):
                    break  # Stop at "On ... wrote:" line
                new_lines.append(line)
            new_content = '\n'.join(new_lines)
        else:
            new_content = "*No content available*"

        # Clean up content
        new_content = re.sub(r'\n{3,}', '\n\n', new_content)
        new_content = new_content.strip()

        # Handle attachments
        attachments = []
        if msg.attachments:
            for att in msg.attachments:
                att_name = att.longFilename or att.shortFilename or "unknown"
                att_size = getattr(att, 'size', None)
                if isinstance(att_size, int):
                    if att_size < 1024:
                        att_size = f"{att_size} B"
                    elif att_size < 1024 * 1024:
                        att_size = f"{att_size / 1024:.1f} KB"
                    else:
                        att_size = f"{att_size / (1024 * 1024):.1f} MB"
                else:
                    att_size = "unknown size"
                attachments.append(f"{att_name} ({att_size})")

        msg.close()

        return {
            'subject': subject,
            'normalized_subject': normalize_subject(subject),
            'sender': sender,
            'to': to,
            'cc': cc,
            'date': date,
            'content': new_content,
            'attachments': attachments,
            'source_file': msg_path.name
        }
    except Exception as e:
        return {'error': str(e), 'source_file': msg_path.name}


def create_thread_markdown(thread_emails, thread_subject):
    """Create a single markdown file for an email thread."""
    # Sort emails by date
    sorted_emails = sorted(
        [e for e in thread_emails if 'error' not in e and e.get('date')],
        key=lambda x: x['date'] if x['date'] else datetime.min
    )

    if not sorted_emails:
        return None

    # Build markdown content
    markdown = f"# {thread_subject}\n\n"
    markdown += f"*Thread with {len(sorted_emails)} email(s)*\n\n"
    markdown += "---\n\n"

    all_attachments = []

    for i, email in enumerate(sorted_emails, 1):
        date_str = email['date'].strftime("%Y-%m-%d %H:%M:%S") if email['date'] else "Unknown"

        markdown += f"## Message {i}\n\n"
        markdown += f"**From:** {email['sender']}\n"
        markdown += f"**To:** {email['to']}\n"
        markdown += f"**Date:** {date_str}\n"
        if email.get('cc'):
            markdown += f"**CC:** {email['cc']}\n"
        markdown += "\n"

        # Add content
        content = email.get('content', '').strip()
        if content:
            markdown += content + "\n"
        else:
            markdown += "*No new content*\n"

        markdown += "\n---\n\n"

        # Collect attachments
        if email.get('attachments'):
            all_attachments.extend(email['attachments'])

    # Add attachments section at the end
    if all_attachments:
        # Deduplicate attachments
        unique_attachments = list(dict.fromkeys(all_attachments))
        markdown += "**Attachments (all messages):**\n"
        for att in unique_attachments:
            markdown += f"- {att}\n"

    return markdown


def main():
    """Main function to convert and consolidate .msg files."""
    current_dir = Path.cwd()
    output_dir = current_dir / "markdown"
    output_dir.mkdir(exist_ok=True)

    # Find all .msg files
    msg_files = list(current_dir.glob("*.msg"))

    if not msg_files:
        print("No .msg files found in current directory.")
        return

    print(f"Found {len(msg_files)} .msg files to process.")
    print(f"Output directory: {output_dir}\n")

    # Parse all emails and group by thread
    threads = defaultdict(list)
    errors = []

    print("Parsing emails...")
    for msg_file in msg_files:
        print(f"  Parsing: {msg_file.name}")
        email_data = parse_msg_file(msg_file)

        if 'error' in email_data:
            errors.append((msg_file.name, email_data['error']))
        else:
            thread_key = email_data['normalized_subject']
            threads[thread_key].append(email_data)

    print(f"\nFound {len(threads)} unique thread(s).")
    print("Creating markdown files...\n")

    # Create markdown files for each thread
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

    # Summary
    print(f"\n{'='*60}")
    print(f"Conversion complete!")
    print(f"  Threads created: {success_count}")
    print(f"  Emails processed: {len(msg_files)}")
    print(f"  Errors: {len(errors)}")
    if errors:
        print("\nFailed files:")
        for name, error in errors:
            print(f"  - {name}: {error}")
    print(f"\nOutput directory: {output_dir}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
