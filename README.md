# AfterShell

AfterShell is a collection of tools and utilities designed to support Windows post-exploitation activities after initial access. The toolkit helps extract valuable data, enumerate local resources, and facilitate lateral movement during red team engagements.

---

## Tools

### 1. Outlook Data Exporter

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

📖 **[Detailed Documentation](exfil/OutlookExporter/README.md)** | **[Quick Start Guide](exfil/README.md#outlook-data-exporter)**

---

### 2. Screen Capture Utility

Powerful and flexible screen capture utility that captures screenshots and packages them into ZIP archives. Ideal for gathering visual evidence during post-exploitation.

**Key Features:**
- **Flexible Capture** - Single or multiple screenshots with configurable intervals
- **Multi-Monitor Support** - Capture from all monitors, primary monitor, or specific monitors
- **Initial Delay** - Set a delay before starting capture sequence
- **Automatic ZIP Packaging** - All screenshots packaged into a single ZIP file
- **High Quality PNG** - Lossless PNG format for maximum quality
- **Timestamped Filenames** - Unique timestamp for each screenshot
- **Automatic Cleanup** - Temporary files cleaned up automatically
- **Professional Interface** - AfterShell-styled console interface

📖 **[Detailed Documentation](exfil/ScreenCapture/README.md)** | **[Quick Start Guide](exfil/README.md#screen-capture-utility)**

---

## Quick Start

### Outlook Data Exporter (C#)

**Requirements:** .NET 8.0 SDK, Microsoft Outlook installed, Windows OS

```bash
# Navigate to OutlookExporter folder
cd exfil\OutlookExporter

# Build the project
dotnet build OutlookExporter.csproj

# Or use the build script
build-outlookexporter.bat
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

📖 **[Full OutlookExporter Documentation](exfil/OutlookExporter/README.md)**

---

### Screen Capture Utility (C#)

**Requirements:** .NET 8.0 SDK, Windows OS

```bash
# Navigate to ScreenCapture folder
cd exfil\ScreenCapture

# Build the project
dotnet build ScreenCapture.csproj

# Or use the build script
build-screencapture.bat
```

**Capture Screenshots:**

```bash
# Capture a single screenshot (default)
ScreenCapture.exe

# Capture 5 screenshots
ScreenCapture.exe -c 5

# Capture 10 screenshots, 2 seconds apart
ScreenCapture.exe -c 10 -i 2

# Capture with 3 second initial delay
ScreenCapture.exe -c 5 -i 1 -d 3

# Capture from primary monitor only
ScreenCapture.exe -c 5 --monitor 1

# List all available monitors
ScreenCapture.exe --list-monitors
```

📖 **[Full ScreenCapture Documentation](exfil/ScreenCapture/README.md)**

---

## Documentation

| Document | Description |
|----------|-------------|
| [exfil/README.md](exfil/README.md) | Overview of all exfiltration tools with quick reference |
| [exfil/OutlookExporter/README.md](exfil/OutlookExporter/README.md) | Comprehensive OutlookExporter documentation |
| [exfil/ScreenCapture/README.md](exfil/ScreenCapture/README.md) | Comprehensive ScreenCapture documentation |
| [exfil/SINGLE_FILE_BUILD.md](exfil/SINGLE_FILE_BUILD.md) | Single-file executable build configuration guide |

---

## Feature Overview

### Outlook Data Exporter

| Feature | Status | Description |
|---------|--------|-------------|
| Contact Export | ✅ | All contact fields exported |
| Email Recipient Extract | ✅ | Senders, To, CC, BCC |
| Contact Frequency | ✅ | Count and latest date tracking |
| Multi-Account | ✅ | All or specific accounts |
| Exchange Address Resolution | ✅ | EX to SMTP conversion |
| CSV Export | ✅ | UTF-8 with BOM |
| JSON Export | ✅ | Full data export |
| Matrix Display | ✅ | Terminal table view |
| Folder Selection | ✅ | Specific folders |
| Email Limit Control | ✅ | Process N emails |
| Deduplication | ✅ | Unique addresses only |
| **Best For** | | Building target lists, mapping communication patterns |

### Screen Capture Utility

| Feature | Status | Description |
|---------|--------|-------------|
| Single Screenshot | ✅ | Quick single capture |
| Multiple Screenshots | ✅ | N screenshots in sequence |
| Configurable Interval | ✅ | Seconds between captures |
| Initial Delay | ✅ | Delay before starting |
| Multi-Monitor Support | ✅ | All, primary, or specific |
| Monitor Detection | ✅ | List available monitors |
| Automatic ZIP | ✅ | All screenshots packaged |
| PNG Format | ✅ | Lossless, high quality |
| Timestamped Files | ✅ | Unique filenames |
| Auto Cleanup | ✅ | Temp files removed |
| **Best For** | | Visual evidence, surveillance, time-lapse capture |

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
# Build OutlookExporter
cd exfil\OutlookExporter
dotnet publish OutlookExporter.csproj -c Release

# Build ScreenCapture
cd exfil\ScreenCapture
dotnet publish ScreenCapture.csproj -c Release
```

Or use the included build scripts:

```bash
# Build all projects at once
cd exfil
build-all.bat

# Build individual projects (Debug and Release)
cd OutlookExporter
build-outlookexporter.bat

cd ScreenCapture
build-screencapture.bat

# Build and run with arguments
build-and-run-outlookexporter.bat --help
build-and-run-screencapture.bat --help
```

> **✨ Standalone Executables**: All applications now build as single-file executables with no DLL dependencies. See [SINGLE_FILE_BUILD.md](exfil/SINGLE_FILE_BUILD.md) for details.

---

## Usage Examples

### Scenario 1: Quick Contact Assessment

```bash
# View contacts in terminal
OutlookExporter.exe --matrix

# Export to file
OutlookExporter.exe --csv -o contacts.csv
```

### Scenario 2: Comprehensive Outlook Exfiltration

```bash
# Export all contacts
OutlookExporter.exe --csv --json -o all_contacts

# Export recipients from all accounts
OutlookExporter.exe --recipients --all-accounts --limit 1000 --csv --json -o all_recipients

# Export from Sent Items
OutlookExporter.exe --recipients --mailfolder "Sent Items" --all-accounts --limit 500 --csv -o sent_recipients
```

### Scenario 3: Visual Evidence Collection

```bash
# Single screenshot
ScreenCapture.exe

# Multiple screenshots with interval
ScreenCapture.exe -c 10 -i 2 -o evidence.zip

# Delayed capture (switch windows first)
ScreenCapture.exe -c 3 -d 10 -o delayed.zip
```

### Scenario 4: Surveillance / Time-Lapse

```bash
# Capture every 30 seconds for 1 hour (120 screenshots)
ScreenCapture.exe -c 120 -i 30 -o surveillance.zip

# Time-lapse: 60 screenshots, one per minute
ScreenCapture.exe -c 60 -i 60 -o timelapse.zip
```

### Scenario 5: Multi-Monitor Capture

```bash
# List available monitors
ScreenCapture.exe --list-monitors

# Capture all monitors
ScreenCapture.exe -c 5 -i 2 --monitor 0 -o all_monitors.zip

# Capture primary monitor only
ScreenCapture.exe -c 5 -i 2 --monitor 1 -o primary.zip

# Capture specific monitor
ScreenCapture.exe -c 5 -i 2 --monitor 2 -o monitor2.zip
```

### Scenario 6: Combined Exfiltration

```bash
# Export Outlook data
OutlookExporter.exe --recipients --all-accounts --limit 500 --csv -o contacts.csv

# Capture screenshots
ScreenCapture.exe -c 10 -i 2 -o screens.zip

# Results: contacts.csv and screens.zip ready for exfiltration
```

---

## Legal Disclaimer

This toolkit is intended for authorized security testing and research purposes only. Users are responsible for ensuring they have proper authorization before using these tools against any systems. The authors assume no liability for misuse of this software.

**Use responsibly and only on systems you have explicit permission to test.**

---

## Security Considerations

⚠️ **These tools are designed for post-exploitation scenarios:**

- All tools should be used only in authorized engagements
- Exported data may contain sensitive information
- Output files are not encrypted by default
- Consider secure deletion after exfiltration
- Be aware of data protection regulations (GDPR, etc.)
- Understand detection risks in monitored environments
- Use secure channels for data transfer

---

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

**GitHub**: https://github.com/Logisek/AfterShell

---

## License

This project is provided for educational and authorized security testing purposes. See [LICENSE](LICENSE) file for details.

GNU General Public License v3.0

Copyright (C) 2025 Logisek.

---

## Contact

- **Website**: https://logisek.com
- **Email**: info@logisek.com
- **GitHub**: https://github.com/Logisek/AfterShell

---

**AfterShell** - Fast Windows post-exploitation wins after initial access.
