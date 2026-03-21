#!/usr/bin/env python3
"""
Email Archiver Script
Archives emails from any IMAP-compatible email server to local storage.

Supports Gmail, Outlook, Yahoo, and any email provider with IMAP access.
"""

import imaplib
import ssl
import email
from email.header import decode_header
import os
import sys
import time
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from getpass import getpass
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


class EmailArchiver:
    def __init__(self, email, password, imap_server, imap_port=993, use_ssl=True):
        self.email = email
        self.password = password
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.use_ssl = use_ssl
        self.mail = None
        self.stats = {
            "processed": 0,
            "archived": 0,
            "deleted": 0,
            "errors": 0,
            "saved_bytes": 0
        }

    def connect(self):
        """Connect to IMAP server."""
        try:
            if self.use_ssl:
                context = ssl.create_default_context()
                self.mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port, ssl_context=context)
            else:
                self.mail = imaplib.IMAP4(self.imap_server, self.imap_port)
            
            self.mail.login(self.email, self.password)
            print(f"✓ Connected to {self.imap_server}:{self.imap_port}")
            return True
        except Exception as e:
            print(f"✗ Connection failed: {e}")
            return False

    def disconnect(self):
        """Disconnect from IMAP server."""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
                print("✓ Disconnected from server")
            except:
                pass

    def get_folders(self):
        """List all folders in the mailbox."""
        _, folders = self.mail.list()
        folder_list = []
        for folder in folders:
            # Parse folder name from IMAP list response
            parts = folder.decode().split(' "/" ')
            if len(parts) >= 2:
                folder_name = parts[-1].strip('"')
                folder_list.append(folder_name)
        return folder_list

    def archive_folder(self, folder_name, archive_dir, cutoff_date=None, delete_after_archive=False, dry_run=True):
        """Archive emails from a specific folder."""
        print(f"\nProcessing folder: {folder_name}")
        
        try:
            # Select folder - properly quote folder names with spaces/special chars
            # IMAP requires quotes around mailbox names containing spaces
            quoted_folder = f'"{folder_name}"' if ' ' in folder_name else folder_name
            status, response = self.mail.select(quoted_folder, readonly=False)
            if status != "OK":
                error_msg = response[0].decode() if response else "Unknown error"
                print(f"  ✗ Could not select folder: {folder_name}")
                print(f"    Error: {error_msg}")
                return

            # Build search criteria
            if cutoff_date:
                # IMAP date format: DD-Month-YYYY (e.g., 01-Jan-2024)
                before_date = cutoff_date.strftime("%d-%b-%Y")
                search_criteria = f'(BEFORE "{before_date}")'
                print(f"  Archiving emails before: {before_date}")
            else:
                search_criteria = "ALL"
                print(f"  Archiving all emails")

            # Search for emails
            status, message_numbers = self.mail.search(None, search_criteria)
            if status != "OK" or not message_numbers[0]:
                print(f"  No emails found matching criteria")
                return

            message_ids = message_numbers[0].split()
            print(f"  Found {len(message_ids)} emails to process")

            if dry_run:
                print(f"  [DRY RUN - no actual changes will be made]")

            # Create archive directory for this folder
            safe_folder_name = folder_name.replace('/', '_').replace('\\', '_')
            folder_archive_path = Path(archive_dir) / safe_folder_name
            
            if not dry_run:
                folder_archive_path.mkdir(parents=True, exist_ok=True)

            # Process each email
            total_emails = len(message_ids)
            for idx, msg_id in enumerate(message_ids, 1):
                self.stats["processed"] += 1
                
                try:
                    # Fetch email
                    status, msg_data = self.mail.fetch(msg_id, "(RFC822)")
                    if status != "OK":
                        print(f"    [{idx}/{total_emails}] ✗ Failed to fetch email {msg_id.decode()}")
                        self.stats["errors"] += 1
                        continue

                    # Parse email
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    
                    # Decode MIME-encoded subject header
                    subject_raw = email_message.get("Subject", "No Subject")
                    if subject_raw:
                        decoded_parts = decode_header(subject_raw)
                        subject_parts = []
                        for part, charset in decoded_parts:
                            if isinstance(part, bytes):
                                try:
                                    subject_parts.append(part.decode(charset or 'utf-8', errors='replace'))
                                except:
                                    subject_parts.append(part.decode('utf-8', errors='replace'))
                            else:
                                subject_parts.append(part)
                        subject = ''.join(subject_parts)
                    else:
                        subject = "No Subject"
                    
                    # Sanitize subject for Windows filename
                    # Remove invalid chars: < > : " / \ | ? * [ ]
                    invalid_chars = '<>:"/\\|?*[]'
                    subject_clean = ''.join(c for c in subject if c not in invalid_chars)
                    subject_clean = subject_clean[:50].strip()
                    if not subject_clean:
                        subject_clean = "No_Subject"
                    
                    # For display (truncate to 40 chars)
                    subject_display = subject_clean[:40] if subject_clean else "No Subject"
                    
                    date_str = email_message.get("Date", "")
                    try:
                        parsed_date = email.utils.parsedate_to_datetime(date_str)
                        date_prefix = parsed_date.strftime("%Y-%m-%d")
                    except:
                        date_prefix = datetime.now().strftime("%Y-%m-%d")

                    filename = f"{date_prefix}_{subject_clean}_{msg_id.decode()}.eml"
                    filename = filename.replace(" ", "_")
                    filepath = folder_archive_path / filename

                    # Skip if already archived
                    if filepath.exists():
                        self.stats["processed"] += 1
                        print(f"    [{idx}/{total_emails}] ⏭ Skipped (already archived): {subject_display}")
                        continue

                    # Save email to file
                    file_written = False
                    file_size_mb = len(raw_email) / (1024 * 1024)
                    if not dry_run:
                        try:
                            with open(filepath, "wb") as f:
                                f.write(raw_email)
                            
                            # Verify file was written correctly
                            if filepath.exists() and filepath.stat().st_size == len(raw_email):
                                file_written = True
                                self.stats["saved_bytes"] += len(raw_email)
                                self.stats["archived"] += 1
                                print(f"    [{idx}/{total_emails}] ✓ {date_prefix} | {subject_display} ({file_size_mb:.2f} MB)")
                            else:
                                print(f"    [{idx}/{total_emails}] ✗ Verification failed: {subject_display}")
                                self.stats["errors"] += 1
                                # Remove partial/corrupt file if it exists
                                if filepath.exists():
                                    filepath.unlink()
                        except Exception as write_error:
                            print(f"    [{idx}/{total_emails}] ✗ Error writing: {subject_display} - {write_error}")
                            self.stats["errors"] += 1
                    else:
                        print(f"    [{idx}/{total_emails}] [DRY RUN] Would save: {date_prefix} | {subject_display}")

                    # Delete from server if requested AND file was verified written
                    if delete_after_archive and not dry_run:
                        if file_written:
                            self.mail.store(msg_id, "+FLAGS", "\\Deleted")
                            self.stats["deleted"] += 1
                        else:
                            print(f"  ⚠️  Skipping deletion - file not verified: {msg_id.decode()}")
                    elif delete_after_archive and dry_run:
                        print(f"  [DRY RUN] Would delete from server: {msg_id.decode()}")
                    
                    # Rate limiting: small delay between emails to avoid overwhelming server
                    time.sleep(0.1)

                except Exception as e:
                    print(f"  ✗ Error processing email {msg_id.decode()}: {e}")
                    self.stats["errors"] += 1

            # Expunge deleted messages
            if delete_after_archive and not dry_run:
                self.mail.expunge()
                print(f"  Deleted emails removed from server")
            
            # Rate limiting: delay between folders to avoid connection drops
            time.sleep(0.5)

        except Exception as e:
            print(f"  ✗ Error processing folder {folder_name}: {e}")
            self.stats["errors"] += 1

    def print_stats(self):
        """Print archiving statistics."""
        print("\n" + "="*50)
        print("ARCHIVING STATISTICS")
        print("="*50)
        print(f"Emails processed:     {self.stats['processed']}")
        print(f"Emails archived:      {self.stats['archived']}")
        print(f"Emails deleted:       {self.stats['deleted']}")
        print(f"Errors encountered:   {self.stats['errors']}")
        
        if self.stats['saved_bytes'] > 0:
            mb_saved = self.stats['saved_bytes'] / (1024 * 1024)
            print(f"Total data archived:  {mb_saved:.2f} MB")


def main():
    parser = argparse.ArgumentParser(
        description="Archive emails from any IMAP server to local storage",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Using .env file (recommended)
  python archive_emails.py --dry-run

  # Gmail - requires app password if 2FA is enabled
  python archive_emails.py --email user@gmail.com --imap-server imap.gmail.com --days-old 365 --dry-run

  # Outlook/Hotmail
  python archive_emails.py --email user@outlook.com --imap-server outlook.office365.com --days-old 180 --delete

  # Yahoo Mail
  python archive_emails.py --email user@yahoo.com --imap-server imap.mail.yahoo.com --days-old 90

  # Custom IMAP server
  python archive_emails.py --email user@example.com --imap-server mail.example.com --imap-port 993 --days-old 365

For detailed IMAP settings for your provider, see README.md
        """
    )
    
    parser.add_argument("--email", default=os.getenv("EMAIL_ADDRESS"),
                        help="Email address (or set EMAIL_ADDRESS in .env)")
    parser.add_argument("--password", 
                        help="Email password (or set EMAIL_PASSWORD in .env, or will prompt securely)")
    parser.add_argument("--imap-server", default=os.getenv("IMAP_SERVER"),
                        help="IMAP server hostname (or set IMAP_SERVER in .env)")
    parser.add_argument("--imap-port", type=int, default=int(os.getenv("IMAP_PORT", 993)),
                        help="IMAP port (default: 993, or set IMAP_PORT in .env)")
    parser.add_argument("--no-ssl", action="store_true",
                        help="Don't use SSL (not recommended)")
    parser.add_argument("--days-old", type=int, default=int(os.getenv("DAYS_OLD", 365)),
                        help="Archive emails older than N days (default: 365)")
    parser.add_argument("--folder",
                        help="Specific folder to archive (default: all folders except system)")
    parser.add_argument("--archive-dir", default=os.getenv("ARCHIVE_DIR", "./EmailArchive"),
                        help="Directory to save archived emails (default: ./EmailArchive)")
    parser.add_argument("--delete", action="store_true",
                        help="Delete emails from server after archiving (DANGEROUS!)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would be archived without making changes")
    parser.add_argument("--list-folders", action="store_true",
                        help="List all available folders and exit")

    args = parser.parse_args()

    # Validate required parameters
    if not args.email:
        print("Error: Email address is required. Provide via --email or EMAIL_ADDRESS in .env file.")
        sys.exit(1)
    
    if not args.imap_server:
        print("Error: IMAP server is required. Provide via --imap-server or IMAP_SERVER in .env file.")
        print("\nCommon IMAP servers:")
        print("  Gmail:          imap.gmail.com")
        print("  Outlook/365:    outlook.office365.com")
        print("  Yahoo Mail:     imap.mail.yahoo.com")
        sys.exit(1)

    # Get password from args, env var, or prompt
    password = args.password or os.getenv("EMAIL_PASSWORD")
    if not password:
        password = getpass(f"Enter password for {args.email}: ")

    # Create archiver instance
    archiver = EmailArchiver(
        email=args.email,
        password=password,
        imap_server=args.imap_server,
        imap_port=args.imap_port,
        use_ssl=not args.no_ssl
    )

    # Connect to server
    if not archiver.connect():
        sys.exit(1)

    try:
        # List folders if requested
        if args.list_folders:
            print("\nAvailable folders:")
            folders = archiver.get_folders()
            for folder in folders:
                print(f"  - {folder}")
            archiver.disconnect()
            return

        # Calculate cutoff date
        if args.days_old:
            cutoff_date = datetime.now() - timedelta(days=args.days_old)
        else:
            cutoff_date = None

        # Create archive directory
        archive_path = Path(args.archive_dir)
        if not args.dry_run:
            archive_path.mkdir(parents=True, exist_ok=True)

        print(f"\nArchive directory: {archive_path.absolute()}")
        if args.dry_run:
            print("MODE: DRY RUN (no changes will be made)")
        if args.delete:
            print("MODE: DELETE AFTER ARCHIVE")
            print("⚠️  WARNING: Emails will be removed from server after archiving!")

        # Archive emails
        if args.folder:
            # Archive specific folder
            archiver.archive_folder(
                args.folder, 
                archive_path, 
                cutoff_date, 
                args.delete, 
                args.dry_run
            )
        else:
            # Archive all folders
            folders = archiver.get_folders()
            print(f"\nFound {len(folders)} folders to process")
            
            # Skip special folders usually
            skip_folders = ["[Gmail]", "[Google Mail]", "Trash", "Spam", "Junk"]
            
            for folder in folders:
                # Skip system folders
                if any(skip in folder for skip in skip_folders):
                    print(f"\nSkipping system folder: {folder}")
                    continue
                
                archiver.archive_folder(
                    folder, 
                    archive_path, 
                    cutoff_date, 
                    args.delete, 
                    args.dry_run
                )

        # Print statistics
        archiver.print_stats()

        print(f"\n✓ Archiving complete!")
        if not args.dry_run:
            print(f"Emails saved to: {archive_path.absolute()}")

    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    finally:
        archiver.disconnect()


if __name__ == "__main__":
    main()
