# Outlook Data Exporter

A native C# console application that exports contacts and email recipients from Microsoft Outlook to CSV or JSON format. Part of the AfterShell post-exploitation toolkit.

> **✨ Standalone Executable**: This application builds as a single-file executable with no DLL dependencies. Just copy the .exe and run it anywhere on Windows x64 systems.

## 🚀 Quick Start

### Get Started in 30 Seconds

**1. Build the Application**

```bash
cd C:\Dev\AfterShell\exfil\OutlookExporter
build-outlookexporter.bat
```

**2. Export Your First Data**

```bash
# Export all contacts to CSV (default)
bin\Release\net8.0-windows\win-x64\publish\OutlookExporter.exe

# Output: outlook_contacts_YYYYMMDD_HHMMSS.csv in current directory
```

### Quick Reference

| Task | Command |
|------|---------|
| Export contacts to CSV | `OutlookExporter.exe` |
| Export to JSON | `OutlookExporter.exe --json` |
| Export to both formats | `OutlookExporter.exe --csv --json` |
| Display as matrix | `OutlookExporter.exe --matrix` |
| List contact folders | `OutlookExporter.exe --list` |
| Export recipients | `OutlookExporter.exe --recipients --limit 100` |
| List accounts | `OutlookExporter.exe --list-accounts` |

---

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

> **Note**: This project is configured to build as a single-file standalone executable. See the [Single-File Build Guide](../SINGLE_FILE_BUILD.md) for details.

### Option 1: Using .NET CLI (Recommended)

```bash
# Navigate to the project folder
cd exfil\OutlookExporter

# Publish as single-file executable (Debug)
dotnet publish -c Debug

# Publish as single-file executable (Release)
dotnet publish -c Release

# Output location:
# bin\Release\net8.0-windows\win-x64\publish\OutlookExporter.exe
```

### Option 2: Using Visual Studio

1. Open the `OutlookExporter.csproj` file in Visual Studio
2. Build the solution (Ctrl+Shift+B)
3. Run the application (F5)

### Option 3: Using Build Scripts (Easiest)

```bash
# Navigate to the project folder
cd exfil\OutlookExporter

# Build both Debug and Release versions
build-outlookexporter.bat

# Build and run with arguments
build-and-run-outlookexporter.bat --help
build-and-run-outlookexporter.bat --recipients --limit 100
```

### Output Locations

After building, find your standalone executable at:

```
bin\Debug\net8.0-windows\win-x64\publish\OutlookExporter.exe      (Debug build)
bin\Release\net8.0-windows\win-x64\publish\OutlookExporter.exe    (Release build)
```

> **✨ No DLLs required!** The executable includes all dependencies. Just copy and run anywhere on Windows x64.

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

## Real-World Examples

### Scenario 1: Building a Target List

Extract all contacts and their email addresses:

```bash
OutlookExporter.exe --csv --json
```

### Scenario 2: Mapping Communication Patterns

Identify most frequently contacted addresses:

```bash
OutlookExporter.exe --recipients --limit 500 --all-accounts --matrix
```

### Scenario 3: Quick Assessment

Display contacts in terminal without creating files:

```bash
OutlookExporter.exe --matrix
```

### Scenario 4: Comprehensive Data Export

Export everything from all accounts:

```bash
# Export contacts
OutlookExporter.exe --csv --json -o all_contacts

# Export recipients from Inbox
OutlookExporter.exe --recipients --all-accounts --limit 1000 --csv --json -o inbox_recipients

# Export recipients from Sent Items
OutlookExporter.exe --recipients --all-accounts --mailfolder "Sent Items" --limit 1000 --csv --json -o sent_recipients
```

### Scenario 5: Specific Account Analysis

Focus on a single email account:

```bash
OutlookExporter.exe --recipients --account "work@company.com" --limit 500 --csv --matrix
```

## Technical Details

### COM Interop

- Uses Microsoft Outlook COM automation
- Proper COM object cleanup to prevent memory leaks
- Works whether Outlook is running or not
- Compatible with all Outlook versions that support COM

### Data Processing

- Deduplicates email addresses (recipients mode)
- Aggregates contact frequency and latest contact date
- Resolves Exchange (EX) addresses to SMTP format
- Handles multiple email accounts in Outlook profile

### File Formats

- **CSV**: UTF-8 with BOM for Excel compatibility
- **JSON**: UTF-8 with proper escaping
- **Matrix**: Formatted terminal output with column sizing

### Performance

- Progress indicators for long operations
- Efficient memory usage with COM object cleanup
- Limit controls for processing large mailboxes
- Sorted output for meaningful results

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

### "Failed to connect to Outlook"
**Solution:**
- Ensure Microsoft Outlook is installed on your system
- Make sure you have proper permissions to access Outlook
- Try running as administrator if permission issues persist

### "Error accessing contacts folder"
**Solution:**
- Verify the folder path is correct using the `--list` option
- Check that you have access to the specified folder
- Folder names are case-sensitive

### "Could not find folder for account"
**Solution:**
- Use `--list-accounts` to see available accounts
- Folder names may vary by language (e.g., "Sent Items" vs "Verzonden items")
- Verify the account email address is correct

### No data exported
**Solution:**
- Check if the contacts folder or mailbox has any items
- Verify folder permissions
- Try without specifying a folder to use defaults

### Exchange addresses not resolved
**Solution:**
- This is expected for some Exchange environments
- The tool tries multiple methods to resolve EX to SMTP
- Exchange cached mode may affect resolution

## Security Considerations

⚠️ **This tool is designed for post-exploitation scenarios. Use responsibly and only on systems you own or have explicit permission to test.**

- Exported data may contain sensitive personal information
- CSV/JSON files are not encrypted
- Consider secure deletion of exported files after use
- Be aware of data protection regulations (GDPR, etc.)

## Integration with AfterShell

This utility is part of the AfterShell post-exploitation toolkit. It can be integrated with other AfterShell tools:

```bash
# Example: Combine with screen capture
OutlookExporter.exe --recipients --limit 100 -o contacts.csv
ScreenCapture.exe -c 5 -i 2 -o screens.zip
# ... additional exfil operations
```

## Contributing

This tool is part of the AfterShell project. Contributions, bug reports, and feature requests are welcome at:

**GitHub**: https://github.com/Logisek/AfterShell

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

Same as the parent project.

Copyright (C) 2025 Logisek

## Contact

- **Website**: https://logisek.com
- **Email**: info@logisek.com
- **Project**: https://github.com/Logisek/AfterShell

## Changelog

### Version 1.0.0 (2025)

- Initial release
- Contacts export to CSV/JSON
- Email recipients export with frequency tracking
- Multi-account support
- Exchange address resolution
- Matrix terminal display
- Simultaneous multiple output formats
