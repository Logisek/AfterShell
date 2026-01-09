# Outlook Data Exporter (C#)

This is a C# console application that exports contacts and email recipients from Microsoft Outlook to CSV or JSON format.

## Features

- **Contacts Export**: Export all contacts from Outlook's contact folders
- **Email Recipients Export**: Extract unique email addresses from email messages (senders and recipients)
- **Contact Frequency Tracking**: Track how many times each contact was contacted and their latest contact date
- **Smart Sorting**: Results sorted by most frequently contacted (recipients mode)
- **Multi-Account Support**: Process all configured email accounts or specific ones
- **Folder Selection**: Choose specific contact or mail folders to process
- **Limit Control**: Process only a specific number of emails for faster extraction
- **Exchange Support**: Resolves Exchange/EX addresses to SMTP format
- **JSON Export**: Export data to JSON format
- **Matrix Display**: Display results as a formatted table in the terminal
- **Simultaneous Output**: Export to CSV, JSON, and display matrix all at once

## Requirements

- .NET 8.0 SDK or later (can be changed to .NET 6.0+ in the .csproj file)
- Microsoft Outlook installed on Windows
- Windows operating system

## Building the Application

### Option 1: Using .NET CLI (Recommended)

```bash
# Build the project
dotnet build

# Run directly
dotnet run

# Build and publish as a standalone executable
dotnet publish -c Release -r win-x64 --self-contained false
```

### Option 2: Using Visual Studio

1. Open the `OutlookExporter.csproj` file in Visual Studio
2. Build the solution (Ctrl+Shift+B)
3. Run the application (F5)

### Option 3: Using Build Scripts

```bash
# Build Debug and Release versions
build.bat

# Build and run with arguments
build-and-run.bat --help
```

## Usage

### Contacts Mode (Default)

Export all contacts to a default timestamped CSV file:
```bash
OutlookExporter.exe
```

Export contacts to JSON format only:
```bash
OutlookExporter.exe --json
```

Export to both CSV and JSON simultaneously:
```bash
OutlookExporter.exe --csv --json
```

Export to CSV, JSON and display matrix (all at once):
```bash
OutlookExporter.exe --csv --json --matrix
```

Display contacts as a matrix table in terminal only:
```bash
OutlookExporter.exe --matrix
```

Export to CSV and also display matrix:
```bash
OutlookExporter.exe --csv --matrix
```

Export contacts to a specific base filename (extensions added automatically):
```bash
OutlookExporter.exe --output my_contacts --csv --json
```

Export contacts from a specific Outlook folder:
```bash
OutlookExporter.exe --folder "Contacts/Work Contacts"
```

List all available contact folders in Outlook:
```bash
OutlookExporter.exe --list
```

### Email Recipients Mode

Export recipients from first 100 emails in Inbox (CSV by default):
```bash
OutlookExporter.exe --recipients --limit 100
```

Export recipients to both CSV and JSON:
```bash
OutlookExporter.exe --recipients --limit 100 --csv --json
```

Export to CSV, JSON and display matrix simultaneously:
```bash
OutlookExporter.exe --recipients --limit 100 --csv --json --matrix
```

Display recipients as a matrix table in terminal only:
```bash
OutlookExporter.exe --recipients --limit 100 --matrix
```

Export to JSON and also display matrix:
```bash
OutlookExporter.exe --recipients --limit 100 --json --matrix
```

Export recipients from Sent Items folder:
```bash
OutlookExporter.exe --recipients --mailfolder "Sent Items" --limit 200
```

Export recipients from all accounts in all formats:
```bash
OutlookExporter.exe --recipients --limit 100 --all-accounts --csv --json --matrix
```

Export from Sent Items across all accounts:
```bash
OutlookExporter.exe --recipients --mailfolder "Sent Items" --limit 100 --all-accounts
```

Export from specific account only:
```bash
OutlookExporter.exe --recipients --account "work@company.com" --limit 100
```

List all configured email accounts:
```bash
OutlookExporter.exe --list-accounts
```

## Command Line Options

### Contacts Mode
| Option | Description |
|--------|-------------|
| `-o, --output <file>` | Output file path (default: outlook_contacts_YYYYMMDD_HHMMSS.csv/json) |
| `-f, --folder <path>` | Specific contacts folder path (e.g., "Contacts/MyFolder") |
| `-l, --list` | List all available contact folders and exit |

### Email Recipients Mode
| Option | Description |
|--------|-------------|
| `-r, --recipients` | Export email recipients instead of contacts |
| `-m, --mailfolder <name>` | Mail folder to process (default: Inbox). Options: Inbox, "Sent Items", Outbox, Drafts |
| `--limit <n>` | Limit number of emails to process per folder |
| `-a, --account <email>` | Process specific email account only |
| `--all-accounts` | Process all configured email accounts |
| `--list-accounts` | List all configured email accounts and exit |

### Output Format Options (can be combined)
| Option | Description |
|--------|-------------|
| `--csv` | Export to CSV format (default if no format specified) |
| `--json` | Export to JSON format |
| `--matrix` | Display results as a matrix table in terminal |

**Note:** All output format options can be used simultaneously. For example, `--csv --json --matrix` will export to both CSV and JSON files and display the matrix in the terminal.

### Common Options
| Option | Description |
|--------|-------------|
| `-o, --output <file>` | Output file path |
| `-h, --help` | Show help message |

## Exported Fields

### Contacts (CSV/JSON)
- Full Name, First Name, Last Name, Middle Name
- Email 1, Email 2, Email 3
- Company, Job Title, Department
- Business Phone, Home Phone, Mobile Phone, Business Fax
- Business Address, Home Address
- Birthday, Anniversary
- Notes, Categories

### Recipients (CSV/JSON)
- Email - SMTP email address (Exchange addresses are resolved)
- Name - Display name
- Type - Sender, To (Recipient), CC (Recipient), or BCC (Recipient)
- Account - Which email account this came from
- ContactCount - Number of times this contact was contacted within the processed emails
- LatestContactDate - Most recent contact date/time with this email address
- IsOwnAccount - Indicates if the email address belongs to one of the configured Outlook accounts ("Yes" or empty)

**Note:** Recipients are sorted by ContactCount (most contacted first), then by LatestContactDate.

### Matrix Display
The `--matrix` option displays a subset of fields in a formatted table:
- **Contacts Matrix**: Name, Email, Company, Job Title, Mobile, Business Phone
- **Recipients Matrix**: Email, Name, Type, Account, Times Contacted, Latest Contact, Own (sorted by most contacted)
  - Own accounts (emails configured in Outlook) are marked with `*` and highlighted in cyan
  - A legend at the bottom shows the count of own accounts used for export

## Notes

- The application works whether Outlook is open or closed
- COM objects are properly released to prevent memory leaks
- CSV files are encoded in UTF-8 with BOM for Excel compatibility
- JSON files are UTF-8 encoded with proper escaping
- Progress is shown during export
- Dates are formatted as YYYY-MM-DD (contacts) or YYYY-MM-DD HH:MM:SS (recipients)
- Email recipients are deduplicated and aggregated (each address appears only once with contact count)
- Recipients are sorted by contact frequency (most contacted first)
- Matrix display truncates long values and caps column widths for readability
- Output formats can be combined: `--csv --json --matrix` outputs all three simultaneously
- When using multiple file formats, extensions are automatically appended (e.g., `output.csv` and `output.json`)

## Troubleshooting

**"Failed to connect to Outlook"**
- Ensure Microsoft Outlook is installed on your system
- Make sure you have proper permissions to access Outlook

**"Error accessing contacts folder"**
- Verify the folder path is correct using the `--list` option
- Check that you have access to the specified folder

**"Could not find folder for account"**
- Use `--list-accounts` to see available accounts
- Folder names may vary by language (e.g., "Sent Items" vs "Verzonden items")

## License

Same as the parent project.
