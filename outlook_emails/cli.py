"""Command-line interface for Outlook Email CLI."""

import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path

from . import __version__
from .core.connection import connect_to_outlook
from .services.search import SearchService
from .services.viewer import ViewerService
from .services.export import ExportService
from .services.compose import ComposeService
from .services.calendar import CalendarService
from .utils.formatting import parse_date


DEFAULT_LOOKBACK_DAYS = 7


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser with subcommands."""
    parser = argparse.ArgumentParser(
        prog='outlook',
        description='Outlook Email CLI - Search, View, and Export emails',
    )
    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # =========================================================================
    # folders command
    # =========================================================================
    folders_parser = subparsers.add_parser(
        'folders',
        help='List all available Outlook folders',
    )

    # =========================================================================
    # search command
    # =========================================================================
    search_parser = subparsers.add_parser(
        'search',
        help='Search emails and display in terminal',
    )
    _add_search_args(search_parser)
    search_parser.add_argument(
        '--export', '-e',
        type=str,
        metavar='DIR',
        help='Also export results to markdown files in DIR',
    )
    search_parser.add_argument(
        '--no-view',
        action='store_true',
        help='Skip terminal display (useful with --export)',
    )

    # =========================================================================
    # export command
    # =========================================================================
    export_parser = subparsers.add_parser(
        'export',
        help='Export emails to Obsidian markdown files',
    )
    _add_search_args(export_parser)
    export_parser.add_argument(
        '--output', '-o',
        type=str,
        required=True,
        help='Output directory for markdown files',
    )
    export_parser.add_argument(
        '--no-threads',
        action='store_true',
        help='Export each email as separate file (don\'t group threads)',
    )
    export_parser.add_argument(
        '--no-overwrite',
        action='store_true',
        help='Skip files that already exist',
    )

    # =========================================================================
    # read command
    # =========================================================================
    read_parser = subparsers.add_parser(
        'read',
        help='Read emails by message ID',
    )
    read_parser.add_argument(
        'message_ids',
        nargs='+',
        help='One or more message IDs (EntryID) to read',
    )

    # =========================================================================
    # send command
    # =========================================================================
    send_parser = subparsers.add_parser(
        'send',
        help='Send a new email',
    )
    send_parser.add_argument(
        '--to', '-t',
        type=str,
        required=True,
        help='Recipient email addresses (comma-separated)',
    )
    send_parser.add_argument(
        '--subject', '-s',
        type=str,
        required=True,
        help='Email subject',
    )
    send_parser.add_argument(
        '--body', '-b',
        type=str,
        required=True,
        help='Email body text',
    )
    send_parser.add_argument(
        '--cc',
        type=str,
        help='CC email addresses (comma-separated)',
    )
    send_parser.add_argument(
        '--bcc',
        type=str,
        help='BCC email addresses (comma-separated)',
    )
    send_parser.add_argument(
        '--attach', '-a',
        type=str,
        action='append',
        help='File path to attach (can specify multiple)',
    )
    send_parser.add_argument(
        '--html',
        action='store_true',
        help='Body is HTML format',
    )
    send_parser.add_argument(
        '--draft',
        action='store_true',
        help='Save as draft instead of sending',
    )

    # =========================================================================
    # reply command
    # =========================================================================
    reply_parser = subparsers.add_parser(
        'reply',
        help='Reply to an email',
    )
    reply_parser.add_argument(
        'message_id',
        help='The message ID (EntryID) to reply to',
    )
    reply_parser.add_argument(
        '--body', '-b',
        type=str,
        required=True,
        help='Reply message body',
    )
    reply_parser.add_argument(
        '--all',
        action='store_true',
        help='Reply to all recipients',
    )
    reply_parser.add_argument(
        '--attach', '-a',
        type=str,
        action='append',
        help='File path to attach (can specify multiple)',
    )
    reply_parser.add_argument(
        '--html',
        action='store_true',
        help='Body is HTML format',
    )
    reply_parser.add_argument(
        '--draft',
        action='store_true',
        help='Save as draft instead of sending',
    )

    # =========================================================================
    # forward command
    # =========================================================================
    forward_parser = subparsers.add_parser(
        'forward',
        help='Forward an email',
    )
    forward_parser.add_argument(
        'message_id',
        help='The message ID (EntryID) to forward',
    )
    forward_parser.add_argument(
        '--to', '-t',
        type=str,
        required=True,
        help='Recipient email addresses (comma-separated)',
    )
    forward_parser.add_argument(
        '--body', '-b',
        type=str,
        help='Optional message to add',
    )
    forward_parser.add_argument(
        '--cc',
        type=str,
        help='CC email addresses (comma-separated)',
    )
    forward_parser.add_argument(
        '--bcc',
        type=str,
        help='BCC email addresses (comma-separated)',
    )
    forward_parser.add_argument(
        '--attach', '-a',
        type=str,
        action='append',
        help='Additional file paths to attach (can specify multiple)',
    )
    forward_parser.add_argument(
        '--html',
        action='store_true',
        help='Body is HTML format',
    )
    forward_parser.add_argument(
        '--draft',
        action='store_true',
        help='Save as draft instead of sending',
    )

    # =========================================================================
    # cal command (calendar)
    # =========================================================================
    cal_parser = subparsers.add_parser('cal', help='Calendar management')
    cal_subparsers = cal_parser.add_subparsers(dest='cal_command', help='Calendar commands')

    # cal list
    cal_list = cal_subparsers.add_parser('list', help='List calendar events')
    cal_list.add_argument('--start', type=str, help='Start date (YYYY-MM-DD, default: today)')
    cal_list.add_argument('--end', type=str, help='End date (YYYY-MM-DD, default: 7 days from start)')
    cal_list.add_argument('--subject', type=str, help='Filter by subject (substring match)')
    cal_list.add_argument('--location', type=str, help='Filter by location (substring match)')
    cal_list.add_argument('--organizer', type=str, help='Filter by organizer email')
    cal_list.add_argument('--all-day', action='store_true', help='All-day events only')
    cal_list.add_argument('--recurring', action='store_true', help='Recurring events only')

    # cal read
    cal_read = cal_subparsers.add_parser('read', help='Read event details')
    cal_read.add_argument('event_id', help='Event ID (EntryID)')

    # cal create
    cal_create = cal_subparsers.add_parser('create', help='Create a new event')
    cal_create.add_argument('--subject', '-s', type=str, required=True, help='Event subject')
    cal_create.add_argument('--start', type=str, required=True, help='Start date/time (YYYY-MM-DD HH:MM)')
    cal_create.add_argument('--end', type=str, required=True, help='End date/time (YYYY-MM-DD HH:MM)')
    cal_create.add_argument('--location', '-l', type=str, help='Event location')
    cal_create.add_argument('--body', '-b', type=str, help='Event description')
    cal_create.add_argument('--required', type=str, help='Required attendees (comma-separated)')
    cal_create.add_argument('--optional', type=str, help='Optional attendees (comma-separated)')
    cal_create.add_argument('--reminder', type=int, help='Reminder minutes before event (default: 15)')
    cal_create.add_argument('--no-reminder', action='store_true', help='No reminder')

    # cal delete
    cal_delete = cal_subparsers.add_parser('delete', help='Delete an event')
    cal_delete.add_argument('event_id', help='Event ID (EntryID)')

    return parser


def _add_search_args(parser: argparse.ArgumentParser) -> None:
    """Add common search/filter arguments to a parser."""
    parser.add_argument(
        '--folder', '-F',
        type=str,
        action='append',
        help='Folder to search (can specify multiple). Default: Inbox + Sent Items',
    )
    parser.add_argument(
        '--days', '-d',
        type=int,
        default=DEFAULT_LOOKBACK_DAYS,
        help=f'Number of days to look back (default: {DEFAULT_LOOKBACK_DAYS})',
    )
    parser.add_argument(
        '--from-date',
        type=str,
        help='Start date (YYYY-MM-DD). Overrides --days.',
    )
    parser.add_argument(
        '--to-date',
        type=str,
        help='End date (YYYY-MM-DD). Default: now.',
    )
    parser.add_argument(
        '--unread', '-u',
        action='store_true',
        help='Only include unread messages',
    )
    parser.add_argument(
        '--filter-email', '-f',
        type=str,
        action='append',
        help='Filter by email address (From/To/CC). Can specify multiple.',
    )
    parser.add_argument(
        '--filter-domain', '-D',
        type=str,
        action='append',
        help='Filter by email domain. Can specify multiple.',
    )
    parser.add_argument(
        '--keyword', '-k',
        type=str,
        help='Search for keyword in subject/body',
    )


def _get_date_range(args) -> tuple[datetime | None, datetime | None]:
    """Parse date range from arguments."""
    since_date = None
    until_date = None

    if hasattr(args, 'from_date') and args.from_date:
        since_date = parse_date(args.from_date)
        if hasattr(args, 'to_date') and args.to_date:
            until_date = parse_date(args.to_date)
            until_date = until_date.replace(hour=23, minute=59, second=59)
    elif hasattr(args, 'days'):
        since_date = datetime.now() - timedelta(days=args.days)

    return since_date, until_date


def cmd_folders(args) -> int:
    """Handle 'folders' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    viewer = ViewerService()
    viewer.print_folders(namespace)

    print("\nUse: outlook search --folder <name>")
    return 0


def cmd_search(args) -> int:
    """Handle 'search' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    since_date, until_date = _get_date_range(args)

    # Print search parameters
    print(f"\nSearching emails...")
    if since_date:
        print(f"  From: {since_date.strftime('%Y-%m-%d')}")
    if until_date:
        print(f"  To: {until_date.strftime('%Y-%m-%d')}")
    if args.folder:
        print(f"  Folders: {', '.join(args.folder)}")
    if args.unread:
        print(f"  Unread only: Yes")
    if args.filter_email:
        print(f"  Filter emails: {', '.join(args.filter_email)}")
    if args.filter_domain:
        print(f"  Filter domains: {', '.join(args.filter_domain)}")
    if args.keyword:
        print(f"  Keyword: {args.keyword}")

    # Search
    search = SearchService(namespace)
    emails = search.search(
        folders=args.folder,
        since_date=since_date,
        until_date=until_date,
        unread_only=args.unread,
        filter_emails=args.filter_email,
        filter_domains=args.filter_domain,
        filter_keyword=args.keyword,
    )

    print(f"\nFound {len(emails)} email(s)")

    # Display results
    if not args.no_view:
        viewer = ViewerService()
        viewer.print_email_table(emails)
        viewer.print_summary(emails)

    # Export if requested
    if args.export:
        print(f"\nExporting to {args.export}...")
        exporter = ExportService(args.export)
        result = exporter.export_emails(emails)
        print(f"  Created {result['files_created']} file(s)")

    return 0


def cmd_export(args) -> int:
    """Handle 'export' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    since_date, until_date = _get_date_range(args)

    # Print parameters
    print(f"\nExporting emails to: {args.output}")
    if since_date:
        print(f"  From: {since_date.strftime('%Y-%m-%d')}")
    if until_date:
        print(f"  To: {until_date.strftime('%Y-%m-%d')}")

    # Search
    search = SearchService(namespace)
    emails = search.search(
        folders=args.folder,
        since_date=since_date,
        until_date=until_date,
        unread_only=args.unread,
        filter_emails=args.filter_email,
        filter_domains=args.filter_domain,
        filter_keyword=args.keyword,
    )

    print(f"Found {len(emails)} email(s)")

    if not emails:
        print("No emails to export.")
        return 0

    # Export
    exporter = ExportService(args.output)
    result = exporter.export_emails(
        emails,
        group_threads=not args.no_threads,
        no_overwrite=args.no_overwrite,
    )

    print(f"\nExport complete:")
    print(f"  Files created: {result['files_created']}")
    if result['files_skipped'] > 0:
        print(f"  Files skipped: {result['files_skipped']}")
    print(f"  Output directory: {args.output}")

    return 0


def cmd_read(args) -> int:
    """Handle 'read' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    search = SearchService(namespace)
    viewer = ViewerService()

    not_found = []
    for i, message_id in enumerate(args.message_ids):
        email = search.get_message_by_id(message_id)

        if not email:
            not_found.append(message_id)
            continue

        if i > 0:
            print("\n" + "=" * 80 + "\n")

        viewer.print_email_detail(email)

    if not_found:
        print(f"\nMessages not found: {len(not_found)}")
        for mid in not_found:
            print(f"  - {mid}")
        return 1

    return 0


def cmd_send(args) -> int:
    """Handle 'send' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    # Parse recipients
    to = [e.strip() for e in args.to.split(',')]
    cc = [e.strip() for e in args.cc.split(',')] if args.cc else None
    bcc = [e.strip() for e in args.bcc.split(',')] if args.bcc else None

    # Send
    compose = ComposeService(namespace)
    success, message = compose.send_email(
        to=to,
        subject=args.subject,
        body=args.body,
        cc=cc,
        bcc=bcc,
        attachments=args.attach,
        html=args.html,
        send_immediately=not args.draft,
    )

    if success:
        print(f"✓ {message}")
        return 0
    else:
        print(f"✗ {message}")
        return 1


def cmd_reply(args) -> int:
    """Handle 'reply' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    compose = ComposeService(namespace)
    success, message = compose.reply(
        message_id=args.message_id,
        body=args.body,
        reply_all=args.all,
        attachments=args.attach,
        html=args.html,
        send_immediately=not args.draft,
    )

    if success:
        print(f"✓ {message}")
        return 0
    else:
        print(f"✗ {message}")
        return 1


def cmd_forward(args) -> int:
    """Handle 'forward' command."""
    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    # Parse recipients
    to = [e.strip() for e in args.to.split(',')]
    cc = [e.strip() for e in args.cc.split(',')] if args.cc else None
    bcc = [e.strip() for e in args.bcc.split(',')] if args.bcc else None

    compose = ComposeService(namespace)
    success, message = compose.forward(
        message_id=args.message_id,
        to=to,
        body=args.body,
        cc=cc,
        bcc=bcc,
        attachments=args.attach,
        html=args.html,
        send_immediately=not args.draft,
    )

    if success:
        print(f"✓ {message}")
        return 0
    else:
        print(f"✗ {message}")
        return 1


def _parse_datetime(dt_str: str) -> datetime:
    """Parse datetime string in format YYYY-MM-DD HH:MM."""
    return datetime.strptime(dt_str, "%Y-%m-%d %H:%M")


def cmd_cal(args) -> int:
    """Handle 'cal' command."""
    if not args.cal_command:
        print("Usage: outlook cal {list,read,create,delete}")
        return 1

    print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    calendar = CalendarService(namespace)
    viewer = ViewerService()

    if args.cal_command == 'list':
        # Parse dates
        start_date = parse_date(args.start) if args.start else None
        end_date = parse_date(args.end) if args.end else None
        if end_date:
            end_date = end_date.replace(hour=23, minute=59, second=59)

        print("\nSearching calendar events...")
        if start_date:
            print(f"  From: {start_date.strftime('%Y-%m-%d')}")
        if end_date:
            print(f"  To: {end_date.strftime('%Y-%m-%d')}")

        events = calendar.list_events(
            start_date=start_date,
            end_date=end_date,
            subject_filter=args.subject,
            location_filter=args.location,
            organizer_filter=args.organizer,
            all_day_only=args.all_day,
            recurring_only=args.recurring,
        )

        print(f"\nFound {len(events)} event(s)\n")
        viewer.print_events_table(events)

    elif args.cal_command == 'read':
        event = calendar.get_event(args.event_id)
        if not event:
            print(f"Event not found: {args.event_id}")
            return 1
        viewer.print_event_detail(event)

    elif args.cal_command == 'create':
        start = _parse_datetime(args.start)
        end = _parse_datetime(args.end)

        required = [e.strip() for e in args.required.split(',')] if args.required else None
        optional = [e.strip() for e in args.optional.split(',')] if args.optional else None

        reminder = None if args.no_reminder else (args.reminder if args.reminder else 15)

        success, message = calendar.create_event(
            subject=args.subject,
            start=start,
            end=end,
            location=args.location,
            body=args.body,
            required_attendees=required,
            optional_attendees=optional,
            reminder_minutes=reminder,
        )

        if success:
            print(f"✓ Event created (ID: {message})")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    elif args.cal_command == 'delete':
        success, message = calendar.delete_event(args.event_id)
        if success:
            print(f"✓ {message}")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    return 0


def main() -> int:
    """Main entry point."""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 0

    try:
        if args.command == 'folders':
            return cmd_folders(args)
        elif args.command == 'search':
            return cmd_search(args)
        elif args.command == 'export':
            return cmd_export(args)
        elif args.command == 'read':
            return cmd_read(args)
        elif args.command == 'send':
            return cmd_send(args)
        elif args.command == 'reply':
            return cmd_reply(args)
        elif args.command == 'forward':
            return cmd_forward(args)
        elif args.command == 'cal':
            return cmd_cal(args)
        else:
            parser.print_help()
            return 1
    except KeyboardInterrupt:
        print("\nAborted.")
        return 130
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == '__main__':
    sys.exit(main())
