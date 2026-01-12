# Screen Capture Utility

A powerful and flexible screen capture utility for Windows that captures one or multiple screenshots and packages them into a ZIP archive. Part of the AfterShell post-exploitation toolkit.

> **✨ Standalone Executable**: This application builds as a single-file executable with no DLL dependencies. Just copy the .exe and run it anywhere on Windows x64 systems.

## 🚀 Quick Start

### Get Started in 30 Seconds

**1. Build the Application**

```batch
cd C:\Dev\AfterShell\exfil\ScreenCapture
build-screencapture.bat
```

**2. Run Your First Capture**

```batch
# Capture a single screenshot
bin\Release\net8.0-windows\win-x64\publish\ScreenCapture.exe

# Output: screenshots_YYYYMMDD_HHMMSS.zip in current directory
```

### Quick Reference

| Task | Command |
|------|---------|
| Single screenshot | `ScreenCapture.exe` |
| 5 screenshots | `ScreenCapture.exe -c 5` |
| 10 screenshots, 2s apart | `ScreenCapture.exe -c 10 -i 2` |
| Wait 5s, then capture | `ScreenCapture.exe -d 5` |
| Custom filename | `ScreenCapture.exe -o myfile.zip` |
| Primary monitor only | `ScreenCapture.exe --monitor 1` |
| List monitors | `ScreenCapture.exe --list-monitors` |
| Show help | `ScreenCapture.exe --help` |

---

## Features

- **Single or Multiple Captures**: Capture one screenshot or N screenshots in sequence
- **Configurable Intervals**: Set time intervals between consecutive captures
- **Initial Delay**: Add a delay before starting the capture sequence
- **Multi-Monitor Support**: Capture from all monitors, primary monitor, or specific monitors
- **Automatic ZIP Packaging**: All screenshots are automatically packaged into a ZIP file
- **High Quality PNG Format**: Screenshots saved in lossless PNG format
- **Timestamped Filenames**: Each screenshot has a unique timestamp
- **Automatic Cleanup**: Temporary files are automatically cleaned up

## Requirements

- **Operating System**: Windows 10 or later
- **.NET Runtime**: .NET 8.0 or later
- **Dependencies**: System.Drawing.Common (automatically installed)

## Build Instructions

> **Note**: This project is configured to build as a single-file standalone executable. See the [Single-File Build Guide](../SINGLE_FILE_BUILD.md) for details.

### Using the Build Script (Easiest)

```batch
# Navigate to the project folder
cd exfil\ScreenCapture

# Build both Debug and Release versions
build-screencapture.bat
```

### Using .NET CLI

```batch
# Navigate to the project folder
cd exfil\ScreenCapture

# Publish as single-file executable (Debug)
dotnet publish -c Debug

# Publish as single-file executable (Release)
dotnet publish -c Release

# Output location:
# bin\Release\net8.0-windows\win-x64\publish\ScreenCapture.exe
```

### Quick Build and Test

```batch
# Build and run with parameters
build-and-run-screencapture.bat --help
build-and-run-screencapture.bat -c 3 -i 1
build-and-run-screencapture.bat -c 3 -i 1 -o test.zip
```

### Output Locations

After building, find your standalone executable at:

```
bin\Debug\net8.0-windows\win-x64\publish\ScreenCapture.exe      (Debug build)
bin\Release\net8.0-windows\win-x64\publish\ScreenCapture.exe    (Release build)
```

> **✨ No DLLs required!** The executable includes all dependencies. Just copy and run anywhere on Windows x64.

## Usage

### Basic Usage

```batch
# Capture a single screenshot (default)
ScreenCapture.exe

# Capture 5 screenshots
ScreenCapture.exe -c 5

# Capture 10 screenshots, 2 seconds apart
ScreenCapture.exe -c 10 -i 2
```

### Advanced Usage

```batch
# Capture with initial delay
ScreenCapture.exe -c 5 -i 1 -d 3

# Custom output filename
ScreenCapture.exe -c 3 -o myscreen.zip

# Capture from primary monitor only
ScreenCapture.exe -c 5 --monitor 1

# Capture from second monitor
ScreenCapture.exe -c 3 --monitor 2

# List all available monitors
ScreenCapture.exe --list-monitors

# Continuous capture with long interval
ScreenCapture.exe -c 100 -i 60
```

## Command-Line Options

| Option | Short | Description | Default |
|--------|-------|-------------|---------|
| `--count` | `-c` | Number of screenshots to capture | 1 |
| `--interval` | `-i` | Interval in seconds between screenshots | 0 |
| `--delay` | `-d` | Initial delay in seconds before starting | 0 |
| `--output` | `-o` | Output ZIP file path | screenshots_YYYYMMDD_HHMMSS.zip |
| `--monitor` | `-m` | Monitor to capture (0=all, 1=primary, 2+=specific) | 0 |
| `--list-monitors` | | List all available monitors and exit | |
| `--help` | `-h` | Show help message | |

## Monitor Selection

- **Monitor 0**: Captures the entire virtual screen (all monitors combined)
- **Monitor 1**: Always captures the primary monitor
- **Monitor 2+**: Captures the specific monitor number (use `--list-monitors` to see available monitors)

### List Your Monitors

```batch
ScreenCapture.exe --list-monitors
```

**Output Example**:
```
Found 3 monitor(s):

Monitor #1 (Primary)
  Resolution: 2560x1440
  Position:   X=0, Y=0
  Device:     \\.\DISPLAY1

Monitor #2
  Resolution: 1920x1080
  Position:   X=2560, Y=0
  Device:     \\.\DISPLAY2
```

### Capture Options
- `--monitor 0` = All monitors (default)
- `--monitor 1` = Primary monitor
- `--monitor 2` = Second monitor
- And so on...

## Output Format

- **File Format**: PNG (lossless compression, high quality)
- **Filename Pattern**: `screenshot_YYYYMMDD_HHMMSS_fff.png`
- **ZIP Archive**: All screenshots automatically packaged
- **Default ZIP Name**: `screenshots_YYYYMMDD_HHMMSS.zip`

## Real-World Examples

### Scenario 1: Quick Single Capture

```batch
ScreenCapture.exe
```

**Result**: 1 screenshot in ZIP file  
**Output**: `screenshots_20260112_143022.zip` containing 1 screenshot

### Scenario 2: Screenshot Evidence Collection

```batch
# Capture 5 screenshots with 1 second between each
ScreenCapture.exe -c 5 -i 1 -o evidence.zip
```

**Result**: evidence.zip with 5 screenshots

### Scenario 3: Surveillance Mode

Capture screenshots every 30 seconds for 1 hour (120 screenshots):

```batch
ScreenCapture.exe -c 120 -i 30 -o surveillance.zip
```

### Scenario 4: Monitoring Session

```batch
# Capture every 30 seconds for 10 minutes (20 screenshots)
ScreenCapture.exe -c 20 -i 30 -o monitoring.zip
```

### Scenario 5: Multi-Monitor Setup

Capture from each monitor separately:

```batch
# Capture all monitors as one image
ScreenCapture.exe -c 3 -i 2 --monitor 0 -o all_monitors.zip

# Capture primary monitor only
ScreenCapture.exe -c 3 -i 2 --monitor 1 -o primary_only.zip

# Capture specific monitor
ScreenCapture.exe -c 1 --monitor 2 -o monitor2.zip
```

### Scenario 6: Time-Lapse Capture

Capture 60 screenshots, one per minute:

```batch
ScreenCapture.exe -c 60 -i 60 -o timelapse.zip
```

### Scenario 7: Delayed Capture

Wait 10 seconds before starting (time to switch windows):

```batch
ScreenCapture.exe -c 3 -d 10 -o delayed.zip
```

**Result**: Time to switch windows before capture

### Scenario 8: Delayed Document Capture

```batch
# Wait 10 seconds to open document, then capture
ScreenCapture.exe -d 10 -o document.zip
```

## Tips & Tricks

1. **Use Initial Delay** (`-d`) to switch windows before capture
2. **Monitor 0** captures the entire desktop (all monitors combined)
3. **ZIP files** are automatically created - no need to specify .zip extension
4. **Timestamps** prevent filename conflicts
5. **Temporary files** are automatically cleaned up

## Technical Details

### Screen Capture Method

- Uses native Windows GDI+ APIs for screen capture
- Leverages `Graphics.CopyFromScreen()` for optimal performance
- Supports multi-monitor configurations through `System.Windows.Forms.Screen`

### ZIP Compression

- Uses .NET built-in `System.IO.Compression`
- Optimal compression level for balance between size and speed
- In-memory processing for efficiency

### File Handling

- Creates temporary directory for screenshots
- Automatic cleanup of temporary files
- Directory creation for output path if needed
- Handles file conflicts by overwriting

### Performance

- Low CPU usage during intervals
- Minimal memory footprint
- Fast PNG encoding
- Efficient ZIP compression

## Error Handling

The utility handles various error conditions:

- **Monitor not found**: Validates monitor index before capture
- **File system errors**: Creates directories as needed
- **Access denied**: Reports permission issues
- **Capture failures**: Continues with remaining captures if one fails
- **Cleanup errors**: Ignores cleanup failures (non-critical)

## Troubleshooting

### "Monitor #X not found"

**Solution**: Run `ScreenCapture.exe --list-monitors` to see available monitors

### "Error capturing screen: Access denied"

**Solution**: Ensure you have proper permissions. Some protected content may require elevated privileges.

### ZIP file is very large

**Solution**: 
- Use monitor selection to capture specific screens only
- Reduce the number of captures
- PNG format is necessary for quality but creates larger files

### Build errors

**Solution**: 
- Ensure .NET 8.0 SDK is installed
- Run `dotnet restore` before building
- Check that System.Drawing.Common package is available

```batch
dotnet --version
```

### ZIP File Not Created

**Solution**: Check output directory permissions

## Security Considerations

⚠️ **This tool is designed for post-exploitation scenarios. Use responsibly and only on systems you own or have explicit permission to test.**

- Screenshots may contain sensitive information
- ZIP files are not encrypted by default
- Consider secure transfer and deletion of capture files
- Be aware of monitoring and detection tools

## Integration with AfterShell

This utility is part of the AfterShell post-exploitation toolkit. It can be integrated with other AfterShell tools for comprehensive data exfiltration:

```batch
# Example: Combine with other exfil tools
ScreenCapture.exe -c 5 -i 2 -o screens.zip
OutlookExporter.exe --recipients --limit 100 -o contacts.csv
# ... additional exfil operations
```

## Contributing

This tool is part of the AfterShell project. Contributions, bug reports, and feature requests are welcome at:

**GitHub**: https://github.com/Logisek/AfterShell

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

Copyright (C) 2025 Logisek

## Contact

- **Website**: https://logisek.com
- **Email**: info@logisek.com
- **Project**: https://github.com/Logisek/AfterShell

## Changelog

### Version 1.0.0 (2026-01-12)

- Initial release
- Single and multiple screenshot capture
- Multi-monitor support
- Automatic ZIP packaging
- Configurable intervals and delays
- AfterShell integration
