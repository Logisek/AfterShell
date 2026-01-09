# AfterShell

AfterShell is a collection of tools and utilities designed to support Windows post-exploitation activities after initial access. The toolkit helps extract valuable data, enumerate local resources, and facilitate lateral movement during red team engagements.

---

## Tools

### Outlook Data Exporter

Native C# console application for extracting contacts and email recipients from Microsoft Outlook. Perfect for building target lists and mapping communication patterns during post-exploitation.

**Key Features:**
- **Contacts Export** - Extract all contacts from Outlook's contact folders
- **Email Recipients Export** - Extract unique email addresses from email messages (senders and recipients)
- **Contact Frequency Tracking** - Track how many times each contact was contacted and their latest contact date
- **Smart Sorting** - Results sorted by most frequently contacted
- **Multi-Account Support** - Process all configured email accounts or specific ones
- **Folder Selection** - Choose specific contact or mail folders to process
- **Limit Control** - Process only a specific number of emails for faster extraction
- **Exchange Support** - Resolves Exchange/EX addresses to SMTP format
- **Multiple Export Formats** - CSV, JSON, and terminal matrix display
- **Simultaneous Output** - Export to CSV, JSON, and display matrix all at once

| Version | Documentation | File |
|---------|---------------|------|
| C# | [exfil/README.md](exfil/README.md) | `exfil/OutlookExporter.cs` |

---

## Quick Start

### Outlook Data Exporter (C#)

**Requirements:** .NET 8.0 SDK, Microsoft Outlook installed, Windows OS

```bash
# Navigate to exfil folder
cd exfil

# Build the project
dotnet build

# Or use the build script
build.bat
```

**Export Contacts:**

```bash
# Export all contacts to CSV (default)
OutlookExporter.exe

# Export contacts to both CSV and JSON
OutlookExporter.exe --csv --json

# Display contacts as matrix table
OutlookExporter.exe --matrix

# List available contact folders
OutlookExporter.exe --list
```

**Export Email Recipients:**

```bash
# Export recipients from first 100 emails in Inbox
OutlookExporter.exe --recipients --limit 100

# Export recipients to CSV, JSON and display matrix
OutlookExporter.exe --recipients --limit 100 --csv --json --matrix

# Export recipients from Sent Items folder
OutlookExporter.exe --recipients --mailfolder "Sent Items" --limit 200

# Export from all accounts
OutlookExporter.exe --recipients --limit 100 --all-accounts

# List all configured email accounts
OutlookExporter.exe --list-accounts
```

📖 **Full documentation:** [exfil/README.md](exfil/README.md)

---

## Documentation

| Document | Description |
|----------|-------------|
| [exfil/README.md](exfil/README.md) | Outlook Data Exporter documentation including all parameters, features, build instructions, and usage examples |

---

## Feature Overview

| Feature | Outlook Data Exporter |
|---------|----------------------|
| **Purpose** | Extract contacts and email recipients from Outlook |
| Contact Export | ✅ All contact fields |
| Email Recipient Extract | ✅ Senders, To, CC, BCC |
| Contact Frequency | ✅ Count and latest date |
| Multi-Account | ✅ All or specific accounts |
| Exchange Address Resolution | ✅ EX to SMTP |
| CSV Export | ✅ UTF-8 with BOM |
| JSON Export | ✅ Full data export |
| Matrix Display | ✅ Terminal table view |
| Folder Selection | ✅ Specific folders |
| Email Limit Control | ✅ Process N emails |
| Deduplication | ✅ Unique addresses |
| Stealth Mode | ❌ (planned) |
| **Best For** | Building target lists, mapping communication patterns |

---

## Installation

### .NET SDK

Download and install .NET 8.0 SDK from [Microsoft .NET Downloads](https://dotnet.microsoft.com/download).

Or install via winget:

```powershell
winget install Microsoft.DotNet.SDK.8
```

### Build from Source

```bash
cd exfil
dotnet build -c Release
```

Or use the included build scripts:

```bash
# Build Debug and Release versions
build.bat

# Build and run with arguments
build-and-run.bat --help
```

---

## Legal Disclaimer

This toolkit is intended for authorized security testing and research purposes only. Users are responsible for ensuring they have proper authorization before using these tools against any systems. The authors assume no liability for misuse of this software.

**Use responsibly and only on systems you have explicit permission to test.**

---

## License

This project is provided for educational and authorized security testing purposes. See [LICENSE](LICENSE) file for details.

---

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

---
