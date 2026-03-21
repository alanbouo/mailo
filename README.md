# Email Archiver

A Python script to archive emails from any IMAP-compatible email server to local storage. Works with Gmail, Outlook, Yahoo, iCloud, and any email provider supporting IMAP access.

## Features

- **IMAP SSL Connection**: Secure connection to IMAP servers
- **Date-based Archiving**: Archive emails older than a specified number of days
- **Folder Selection**: Archive all folders or specific folders only
- **Dry Run Mode**: Preview what would be archived without making changes
- **Safe Deletion**: Optionally delete archived emails from server (with confirmation)
- **Progress Statistics**: Detailed reporting of processed, archived, and deleted emails

## Environment Variables

The script supports configuration via environment variables in a `.env` file. Command line arguments override `.env` values.

| Variable | Description | Required |
|----------|-------------|----------|
| `EMAIL_ADDRESS` | Email address | Yes |
| `EMAIL_PASSWORD` | Email password (prompts if not provided) | Yes |
| `IMAP_SERVER` | IMAP server hostname | Yes |
| `IMAP_PORT` | IMAP port | No (default: 993) |
| `USE_SSL` | Use SSL connection | No (default: true) |
| `DAYS_OLD` | Archive emails older than N days | No (default: 365) |
| `ARCHIVE_DIR` | Directory to save archived emails | No (default: ./EmailArchive) |
| `FOLDER` | Specific folder to archive (optional) | No |
| `DELETE_AFTER_ARCHIVE` | Delete from server after archiving | No (default: false) |

⚠️ **Security Note**: Never commit your `.env` file to version control. It's already added to `.gitignore`.

## Requirements

- Python 3.6+
- `python-dotenv` for environment variable support (see [Installation](#installation))

## Installation

1. Install the required dependency:
```bash
pip install -r requirements.txt
```

2. Configure your environment variables by copying `.env.example` to `.env`:
```bash
cp .env.example .env
```

3. Edit `.env` with your email provider settings:

```bash
# Email Configuration
EMAIL_ADDRESS=your_email@example.com
EMAIL_PASSWORD=your_password

# IMAP Server Configuration - Common settings:
# Gmail:        imap.gmail.com
# Outlook/365:  outlook.office365.com  
# Yahoo:        imap.mail.yahoo.com
# iCloud:       imap.mail.me.com
IMAP_SERVER=imap.gmail.com
IMAP_PORT=993
USE_SSL=true

# Archiving Options
DAYS_OLD=365
ARCHIVE_DIR=./EmailArchive
```

## Usage

### Basic Usage

```bash
# Dry run - see what would be archived (emails older than 1 year)
python archive_emails.py --dry-run

# Archive emails older than 6 months and delete from server
python archive_emails.py --days-old 180 --delete --archive-dir ~/EmailArchive

# Archive specific folder only
python archive_emails.py --folder "INBOX" --days-old 90
```

### Command Line Options

| Option | Description | Required |
|--------|-------------|----------|
| `--email` | Email address | Yes (or set in .env) |
| `--password` | Email password (prompts if not provided) | Yes (or set in .env) |
| `--imap-server` | IMAP server hostname | Yes (or set in .env) |
| `--imap-port` | IMAP port | No (default: 993) |
| `--no-ssl` | Disable SSL (not recommended) | No |
| `--days-old` | Archive emails older than N days | No (default: 365) |
| `--folder` | Specific folder to archive | No |
| `--archive-dir` | Directory to save archived emails | No (default: ./EmailArchive) |
| `--delete` | Delete emails from server after archiving | No |
| `--dry-run` | Preview without making changes | No |
| `--list-folders` | List available folders and exit | No |

### Examples

**List all folders in the mailbox:**
```bash
python archive_emails.py --list-folders
```

**Dry run for emails older than 1 year:**
```bash
python archive_emails.py --days-old 365 --dry-run
```

**Archive emails older than 6 months to a custom directory:**
```bash
python archive_emails.py --days-old 180 --archive-dir ~/Backups/Emails
```

**Archive and delete from server (⚠️ use with caution):**
```bash
python archive_emails.py --days-old 365 --delete
```

**Archive specific folder only:**
```bash
python archive_emails.py --folder "Sent" --days-old 90
```

## Output Format

Archived emails are saved as `.eml` files in the following structure:

```
EmailArchive/
├── INBOX/
│   ├── 2023-01-15_Meeting_Notes_12345.eml
│   ├── 2023-02-20_Project_Update_12346.eml
│   └── ...
├── Sent/
│   ├── 2023-03-10_Re_Proposal_12347.eml
│   └── ...
└── ...
```

Files are named using the pattern: `YYYY-MM-DD_Subject_MessageID.eml`

## Safety Features

- **Dry Run Mode**: Always test with `--dry-run` first to see what would be archived
- **Password Prompt**: Password is securely prompted if not provided via command line
- **SSL by Default**: All connections use SSL/TLS encryption by default
- **System Folder Skip**: Automatically skips `[Gmail]`, `Trash`, `Spam`, `Junk` folders
- **Verify Before Delete**: Emails are only deleted from server after confirming the local file was written correctly (size verification)
- **Partial File Cleanup**: Corrupt/partial files are automatically removed if verification fails
- **Expunge on Delete**: Properly expunges deleted messages when using `--delete`

## Email Provider Setup

### Gmail
- **IMAP Server**: `imap.gmail.com`
- **Port**: `993`
- **SSL**: Yes
- **Notes**: If you have 2-Factor Authentication enabled, you must use an [App Password](https://myaccount.google.com/apppasswords) instead of your regular password.

### Outlook / Microsoft 365
- **IMAP Server**: `outlook.office365.com`
- **Port**: `993`
- **SSL**: Yes
- **Notes**: If you have MFA enabled, use an [App Password](https://account.live.com/proofs/AppPassword).


### Yahoo Mail
- **IMAP Server**: `imap.mail.yahoo.com`
- **Port**: `993`
- **SSL**: Yes

### iCloud
- **IMAP Server**: `imap.mail.me.com`
- **Port**: `993`
- **SSL**: Yes
- **Notes**: Generate an app-specific password from [Apple ID settings](https://appleid.apple.com).

### Custom/Private Servers
- **IMAP Server**: Your server's hostname (e.g., `mail.yourdomain.com`)
- **Port**: Usually `993` for SSL, sometimes `143` for non-SSL
- **SSL**: Depends on your server configuration

## License

MIT License - See project repository for details.
