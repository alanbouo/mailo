#!/usr/bin/env python3
"""
Email Archiver Script
Archives emails from IMAP server to local storage to reduce mailbox size.

Based on configuration:
- Email: alandji.bouorakima@netc.fr
- IMAP Server: mail.mailo.com:993 (SSL)
- SMTP Server: mail.mailo.com:465 (SSL)
"""

import imaplib
import ssl
import email
import os
import sys
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from getpass import getpass


class EmailArchiver:
    def __init__(self, email, password, imap_server="mail.mailo.com", imap_port=993, use_ssl=True):
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
            # Select folder
            status, _ = self.mail.select(folder_name, readonly=False)
            if status != "OK":
                print(f"  ✗ Could not select folder: {folder_name}")
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
            for msg_id in message_ids:
                self.stats["processed"] += 1
                
                try:
                    # Fetch email
                    status, msg_data = self.mail.fetch(msg_id, "(RFC822)")
                    if status != "OK":
                        self.stats["errors"] += 1
                        continue

                    # Parse email
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    
                    # Generate filename from date and subject
                    subject = email_message.get("Subject", "No Subject")
                    subject = subject[:50] if subject else "No Subject"
                    subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).strip()
                    
                    date_str = email_message.get("Date", "")
                    try:
                        parsed_date = email.utils.parsedate_to_datetime(date_str)
                        date_prefix = parsed_date.strftime("%Y-%m-%d")
                    except:
                        date_prefix = datetime.now().strftime("%Y-%m-%d")

                    filename = f"{date_prefix}_{subject}_{msg_id.decode()}.eml"
                    filename = filename.replace(" ", "_")
                    filepath = folder_archive_path / filename

                    # Save email to file
                    if not dry_run:
                        with open(filepath, "wb") as f:
                            f.write(raw_email)
                        
                        self.stats["saved_bytes"] += len(raw_email)
                        self.stats["archived"] += 1
                    else:
                        print(f"  [DRY RUN] Would save: {filepath}")

                    # Delete from server if requested
                    if delete_after_archive and not dry_run:
                        self.mail.store(msg_id, "+FLAGS", "\\Deleted")
                        self.stats["deleted"] += 1
                    elif delete_after_archive and dry_run:
                        print(f"  [DRY RUN] Would delete from server: {msg_id.decode()}")

                except Exception as e:
                    print(f"  ✗ Error processing email {msg_id.decode()}: {e}")
                    self.stats["errors"] += 1

            # Expunge deleted messages
            if delete_after_archive and not dry_run:
                self.mail.expunge()
                print(f"  Deleted emails removed from server")

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
        description="Archive emails from IMAP server to reduce mailbox storage",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Dry run - see what would be archived (emails older than 1 year)
  python archive_emails.py --email alandji.bouorakima@netc.fr --days-old 365 --dry-run

  # Archive emails older than 6 months and delete from server
  python archive_emails.py --email alandji.bouorakima@netc.fr --days-old 180 --delete --archive-dir ~/EmailArchive

  # Archive specific folder only
  python archive_emails.py --email alandji.bouorakima@netc.fr --folder "INBOX" --days-old 90
        """
    )
    
    parser.add_argument("--email", default="alandji.bouorakima@netc.fr",
                        help="Email address (default: alandji.bouorakima@netc.fr)")
    parser.add_argument("--password", 
                        help="Email password (if not provided, will prompt securely)")
    parser.add_argument("--imap-server", default="mail.mailo.com",
                        help="IMAP server (default: mail.mailo.com)")
    parser.add_argument("--imap-port", type=int, default=993,
                        help="IMAP port (default: 993 for SSL)")
    parser.add_argument("--no-ssl", action="store_true",
                        help="Don't use SSL (not recommended)")
    parser.add_argument("--days-old", type=int, default=365,
                        help="Archive emails older than N days (default: 365)")
    parser.add_argument("--folder",
                        help="Specific folder to archive (default: all folders)")
    parser.add_argument("--archive-dir", default="./EmailArchive",
                        help="Directory to save archived emails (default: ./EmailArchive)")
    parser.add_argument("--delete", action="store_true",
                        help="Delete emails from server after archiving (DANGEROUS!)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would be archived without making changes")
    parser.add_argument("--list-folders", action="store_true",
                        help="List all available folders and exit")

    args = parser.parse_args()

    # Get password if not provided
    password = args.password
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
