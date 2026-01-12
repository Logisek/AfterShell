# Single-File Executable Build Configuration

## Overview

Both **OutlookExporter** and **ScreenCapture** applications are now configured to build as **standalone single-file executables**. This means:

✅ **No DLL dependencies required**  
✅ **Includes the .NET 8.0 runtime**  
✅ **Just copy the .exe and run anywhere on Windows x64**  
✅ **No installation or setup needed on target machines**

## What Changed

### Project Files (.csproj)

Added a conditional property group to both `OutlookExporter.csproj` and `ScreenCapture.csproj`:

```xml
<!-- Single-file executable configuration (applied during publish) -->
<PropertyGroup Condition="'$(RuntimeIdentifier)' != ''">
  <PublishSingleFile>true</PublishSingleFile>
  <SelfContained>true</SelfContained>
  <IncludeNativeLibrariesForSelfExtract>true</IncludeNativeLibrariesForSelfExtract>
  <EnableCompressionInSingleFile>true</EnableCompressionInSingleFile>
</PropertyGroup>
```

#### Property Explanations:

- **Condition**: Only applies these settings when RuntimeIdentifier is specified (during publish)
- **PublishSingleFile**: Bundles all application files (assemblies, dependencies, native libraries) into a single executable
- **SelfContained**: Includes the .NET runtime in the output (no need for .NET to be installed on target machine)
- **IncludeNativeLibrariesForSelfExtract**: Ensures native DLLs are extracted and loaded correctly
- **EnableCompressionInSingleFile**: Compresses the bundled files to reduce executable size

> **Important**: RuntimeIdentifier is NOT specified in the project file. It's passed as a command-line parameter during publish. This allows normal `dotnet build` to work while `dotnet publish -r win-x64` creates single-file executables.

### Build Scripts

Changed from `dotnet build` to `dotnet publish`:

**Before:**
```batch
dotnet build OutlookExporter.csproj -c Release
```

**After:**
```batch
dotnet publish OutlookExporter.csproj -c Release -r win-x64 --self-contained true
```

### Output Locations

**Before:**
```
bin\Release\net8.0-windows\OutlookExporter.exe
bin\Release\net8.0-windows\ScreenCapture.exe
```

**After:**
```
bin\Release\net8.0-windows\win-x64\publish\OutlookExporter.exe
bin\Release\net8.0-windows\win-x64\publish\ScreenCapture.exe
```

## How to Build

### Individual Projects

```batch
# OutlookExporter
cd exfil\OutlookExporter
build-outlookexporter.bat

# ScreenCapture
cd exfil\ScreenCapture
build-screencapture.bat
```

### All Projects at Once

```batch
cd exfil
build-all.bat
```

## File Sizes

Single-file executables are larger than regular builds because they include the .NET runtime:

- **OutlookExporter.exe**: ~35 MB (compressed)
- **ScreenCapture.exe**: ~72 MB (compressed)

The ScreenCapture utility is larger because it includes Windows Forms and additional graphics libraries. These are reasonable sizes for the portability benefit of true standalone executables.

## Deployment

Simply copy the executable from the publish folder to any Windows x64 machine and run it. No installation, no dependencies, no .NET runtime required on the target machine.

```batch
# Copy OutlookExporter
copy OutlookExporter\bin\Release\net8.0-windows\win-x64\publish\OutlookExporter.exe C:\Tools\

# Copy ScreenCapture
copy ScreenCapture\bin\Release\net8.0-windows\win-x64\publish\ScreenCapture.exe C:\Tools\

# Run them anywhere
C:\Tools\OutlookExporter.exe --help
C:\Tools\ScreenCapture.exe --help
```

## Advanced Configuration

### Reduce File Size (Optional)

To reduce executable size further, you can enable trimming. Add this to the conditional PropertyGroup in the .csproj file:

```xml
<PropertyGroup Condition="'$(RuntimeIdentifier)' != ''">
  <PublishSingleFile>true</PublishSingleFile>
  <SelfContained>true</SelfContained>
  <IncludeNativeLibrariesForSelfExtract>true</IncludeNativeLibrariesForSelfExtract>
  <EnableCompressionInSingleFile>true</EnableCompressionInSingleFile>
  <!-- Add these for trimming -->
  <PublishTrimmed>true</PublishTrimmed>
  <TrimMode>link</TrimMode>
</PropertyGroup>
```

⚠️ **Warning**: Trimming may cause issues with applications that use COM Interop (like OutlookExporter) or reflection heavily. Test thoroughly if you enable this. **Not recommended for OutlookExporter**.

### Framework-Dependent Build (Smaller Size)

If you want smaller executables but are okay with requiring .NET 8.0 on target machines, change:

```xml
<SelfContained>false</SelfContained>
```

This will produce ~5-10 MB executables but require .NET 8.0 Runtime installed on target systems.

## Troubleshooting

### Issue: "The application failed to start"

**Cause**: Missing Visual C++ Redistributable  
**Solution**: Some native dependencies may require Visual C++ Redistributable. Most Windows systems have this installed.

### Issue: Executable is too large

**Solution**: 
1. Enable `PublishTrimmed` (see Advanced Configuration)
2. Switch to framework-dependent deployment
3. Use compression tools (UPX, etc.) - though this may trigger antivirus

### Issue: Application runs slower on first start

**Cause**: Self-extracting executables need to extract files to a temp folder on first run  
**Solution**: This is normal behavior. Subsequent runs will be faster.

## References

- [.NET Single-File Applications](https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file/overview)
- [dotnet publish command](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-publish)
- [Trim self-contained deployments and executables](https://learn.microsoft.com/en-us/dotnet/core/deploying/trimming/trim-self-contained)
