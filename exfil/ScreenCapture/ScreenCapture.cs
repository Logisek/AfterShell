/*
*    This file is part of the psot-exploitation toolkit AfterShell
*    Copyright (C) 2025 Logisek
*    https://github.com/Logisek/AfterShell
*
*    AfterShell - Fast Windows post-exploitation wins after initial access..
*
*    This program is free software: you can redistribute it and/or modify
*    it under the terms of the GNU General Public License as published by
*    the Free Software Foundation, either version 3 of the License, or
*    (at your option) any later version.
*
*    This program is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*    GNU General Public License for more details.
*
*    You should have received a copy of the GNU General Public License
*    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*
*    For more see the file 'LICENSE' for copying permission.
 */


/*
 * Screen Capture Utility
 * Captures one or multiple screenshots and packages them into a ZIP archive.
 * Supports multiple monitors, configurable intervals, and initial delays.
 * 
 * Requirements:
 *     - .NET 8.0 or later (Windows)
 *     - System.Drawing.Common package
 * 
 * Build:
 *     dotnet build
 * 
 * Usage:
 *     ScreenCapture.exe                                    # Capture 1 screenshot (default)
 *     ScreenCapture.exe -c 5                               # Capture 5 screenshots
 *     ScreenCapture.exe -c 10 -i 2                         # Capture 10 screenshots, 2 seconds apart
 *     ScreenCapture.exe -c 5 -i 1 -d 3                     # Wait 3 seconds, then capture 5 screenshots, 1 second apart
 *     ScreenCapture.exe -c 30 --fps 30                     # Capture 30 screenshots at 30 FPS (1 screenshot every 0.033s)
 *     ScreenCapture.exe -c 60 --fps 60                     # Capture 60 screenshots at 60 FPS (1 screenshot every 0.016s)
 *     ScreenCapture.exe -c 3 -o myscreen                   # Custom output filename
 *     ScreenCapture.exe -c 5 --monitor 1                   # Capture from primary monitor only
 *     ScreenCapture.exe -c 3 --monitor 2                   # Capture from second monitor only
 *     ScreenCapture.exe --list-monitors                    # List all available monitors
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Threading;

namespace ScreenCapture
{
    class Program
    {
        static int Main(string[] args)
        {
            ShowBanner();
            
            var options = ParseArguments(args);
            
            if (options == null)
            {
                PrintUsage();
                return 1;
            }

            var capturer = new ScreenCapturer();

            try
            {
                // Handle list monitors mode
                if (options.ListMonitors)
                {
                    capturer.ListMonitors();
                    return 0;
                }

                Console.WriteLine("\nScreen Capture Configuration:");
                Console.WriteLine($"  Screenshots:    {options.Count}");
                if (options.Count > 1)
                {
                    if (options.Fps > 0)
                    {
                        Console.WriteLine($"  FPS:            {options.Fps} ({options.Interval * 1000:F2} ms interval)");
                    }
                    else
                    {
                        Console.WriteLine($"  Interval:       {options.Interval} second(s)");
                    }
                }
                if (options.InitialDelay > 0)
                {
                    Console.WriteLine($"  Initial delay:  {options.InitialDelay} second(s)");
                }
                Console.WriteLine($"  Monitor:        {GetMonitorDescription(options.Monitor)}");
                Console.WriteLine($"  Output ZIP:     {options.OutputFile}");
                Console.WriteLine();

                // Initial delay if specified
                if (options.InitialDelay > 0)
                {
                    Console.WriteLine($"Waiting {options.InitialDelay} second(s) before starting...");
                    for (int i = options.InitialDelay; i > 0; i--)
                    {
                        Console.Write($"\r  {i}... ");
                        Thread.Sleep(1000);
                    }
                    Console.WriteLine("\r  Starting capture!          ");
                }

                // Create temporary directory for screenshots
                string tempDir = Path.Combine(Path.GetTempPath(), $"screencap_{Guid.NewGuid()}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    var capturedFiles = new List<string>();

                    // Capture screenshots
                    for (int i = 1; i <= options.Count; i++)
                    {
                        Console.Write($"\rCapturing screenshot {i}/{options.Count}... ");
                        
                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                        string fileName = $"screenshot_{timestamp}.png";
                        string filePath = Path.Combine(tempDir, fileName);

                        bool success = capturer.CaptureScreen(filePath, options.Monitor);
                        
                        if (success)
                        {
                            capturedFiles.Add(filePath);
                            Console.WriteLine("✓");
                        }
                        else
                        {
                            Console.WriteLine("✗ Failed");
                        }

                        // Wait for interval if not the last screenshot
                        if (i < options.Count && options.Interval > 0)
                        {
                            int sleepMs = (int)(options.Interval * 1000);
                            Thread.Sleep(sleepMs);
                        }
                    }

                    Console.WriteLine();

                    if (capturedFiles.Count == 0)
                    {
                        Console.WriteLine("✗ No screenshots were captured successfully.");
                        return 1;
                    }

                    // Create ZIP archive
                    Console.WriteLine($"Creating ZIP archive: {options.OutputFile}");
                    
                    // Ensure output directory exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(options.OutputFile));
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Delete existing file if present
                    if (File.Exists(options.OutputFile))
                    {
                        File.Delete(options.OutputFile);
                    }

                    using (var zip = ZipFile.Open(options.OutputFile, ZipArchiveMode.Create))
                    {
                        foreach (string file in capturedFiles)
                        {
                            string entryName = Path.GetFileName(file);
                            zip.CreateEntryFromFile(file, entryName, CompressionLevel.Optimal);
                        }
                    }

                    Console.WriteLine($"\n✓ Success!");
                    Console.WriteLine($"  Captured:  {capturedFiles.Count} screenshot(s)");
                    Console.WriteLine($"  ZIP file:  {Path.GetFullPath(options.OutputFile)}");
                    Console.WriteLine($"  Size:      {new FileInfo(options.OutputFile).Length / 1024} KB");
                    
                    return 0;
                }
                finally
                {
                    // Cleanup temporary directory
                    try
                    {
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                        }
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n✗ Error: {ex.Message}");
                return 1;
            }
        }

        static string GetMonitorDescription(int monitor)
        {
            return monitor switch
            {
                0 => "All monitors (virtual screen)",
                1 => "Primary monitor",
                _ => $"Monitor #{monitor}"
            };
        }

        static CommandLineOptions ParseArguments(string[] args)
        {
            var options = new CommandLineOptions();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-c":
                    case "--count":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int count) && count > 0)
                        {
                            options.Count = count;
                        }
                        else
                        {
                            Console.WriteLine("Error: --count requires a positive number");
                            return null;
                        }
                        break;

                    case "-i":
                    case "--interval":
                        if (i + 1 < args.Length)
                        {
                            if (double.TryParse(args[++i], out double interval) && interval >= 0)
                            {
                                options.Interval = interval;
                                options.Fps = 0; // Clear FPS if interval is explicitly set
                            }
                            else
                            {
                                Console.WriteLine("Error: --interval requires a non-negative number");
                                return null;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: --interval requires a value");
                            return null;
                        }
                        break;

                    case "--fps":
                        if (i + 1 < args.Length)
                        {
                            if (double.TryParse(args[++i], out double fps) && fps > 0)
                            {
                                options.Fps = fps;
                                options.Interval = 1.0 / fps; // Calculate interval from FPS
                            }
                            else
                            {
                                Console.WriteLine("Error: --fps requires a positive number");
                                return null;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: --fps requires a value");
                            return null;
                        }
                        break;

                    case "-d":
                    case "--delay":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int delay) && delay >= 0)
                        {
                            options.InitialDelay = delay;
                        }
                        else
                        {
                            Console.WriteLine("Error: --delay requires a non-negative number");
                            return null;
                        }
                        break;

                    case "-o":
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            options.OutputFile = args[++i];
                        }
                        else
                        {
                            Console.WriteLine("Error: --output requires a file path");
                            return null;
                        }
                        break;

                    case "-m":
                    case "--monitor":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int monitor) && monitor >= 0)
                        {
                            options.Monitor = monitor;
                        }
                        else
                        {
                            Console.WriteLine("Error: --monitor requires a non-negative number");
                            return null;
                        }
                        break;

                    case "--list-monitors":
                        options.ListMonitors = true;
                        break;

                    case "-h":
                    case "--help":
                        return null;

                    default:
                        Console.WriteLine($"Unknown option: {args[i]}");
                        return null;
                }
            }

            // Generate default output filename if not specified
            if (string.IsNullOrEmpty(options.OutputFile) && !options.ListMonitors)
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                options.OutputFile = $"screenshots_{timestamp}.zip";
            }

            // Ensure .zip extension
            if (!options.ListMonitors && !string.IsNullOrEmpty(options.OutputFile))
            {
                if (!options.OutputFile.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    options.OutputFile += ".zip";
                }
            }

            return options;
        }

        static void ShowBanner()
        {
            string asciiArt = @"
 █████╗ ███████╗████████╗███████╗██████╗ ███████╗██╗  ██╗███████╗██╗     ██╗     
██╔══██╗██╔════╝╚══██╔══╝██╔════╝██╔══██╗██╔════╝██║  ██║██╔════╝██║     ██║     
███████║█████╗     ██║   █████╗  ██████╔╝███████╗███████║█████╗  ██║     ██║     
██╔══██║██╔══╝     ██║   ██╔══╝  ██╔══██╗╚════██║██╔══██║██╔══╝  ██║     ██║     
██║  ██║██║        ██║   ███████╗██║  ██║███████║██║  ██║███████╗███████╗███████╗
╚═╝  ╚═╝╚═╝        ╚═╝   ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝
";

            try
            {
                // Try to use ANSI colors if supported
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine(asciiArt);
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("    Post-Exploitation Toolkit - Screen Capture");
                Console.ResetColor();
                Console.WriteLine("    https://logisek.com | info@logisek.com");
                Console.WriteLine("    AfterShell | https://github.com/Logisek/AfterShell");
                Console.WriteLine();
            }
            catch
            {
                // Fallback without colors
                Console.WriteLine(asciiArt);
                Console.WriteLine("    Post-Exploitation Toolkit - Screen Capture");
                Console.WriteLine("    https://logisek.com | info@logisek.com");
                Console.WriteLine("    AfterShell | https://github.com/Logisek/AfterShell");
                Console.WriteLine();
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine(@"
Screen Capture Utility
Capture one or multiple screenshots and package them into a ZIP archive

Usage:
  ScreenCapture.exe [options]

OPTIONS:
  -c, --count <n>         Number of screenshots to capture (default: 1)
  -i, --interval <n>      Interval in seconds between screenshots (default: 0, supports decimals)
  --fps <n>               Capture rate in frames per second (overrides --interval if specified)
  -d, --delay <n>         Initial delay in seconds before starting (default: 0)
  -o, --output <file>     Output ZIP file path (default: screenshots_YYYYMMDD_HHMMSS.zip)
  -m, --monitor <n>       Monitor to capture (0 = all, 1 = primary, 2+ = specific) (default: 0)
  --list-monitors         List all available monitors and exit
  -h, --help              Show this help message

EXAMPLES:
  # Capture a single screenshot (default)
  ScreenCapture.exe

  # Capture 5 screenshots
  ScreenCapture.exe -c 5

  # Capture 10 screenshots, 2 seconds apart
  ScreenCapture.exe -c 10 -i 2

  # Capture 5 screenshots with 1 second interval, after 3 second delay
  ScreenCapture.exe -c 5 -i 1 -d 3

  # Capture 30 screenshots at 30 FPS (1 screenshot every 0.033 seconds)
  ScreenCapture.exe -c 30 --fps 30

  # Capture 60 screenshots at 60 FPS (1 screenshot every 0.016 seconds)
  ScreenCapture.exe -c 60 --fps 60

  # Capture 120 screenshots at 120 FPS
  ScreenCapture.exe -c 120 --fps 120

  # Use fractional interval for custom FPS (e.g., 0.033 seconds = ~30 FPS)
  ScreenCapture.exe -c 30 -i 0.033

  # Capture 3 screenshots to custom output file
  ScreenCapture.exe -c 3 -o myscreen.zip

  # Capture 5 screenshots from primary monitor only
  ScreenCapture.exe -c 5 --monitor 1

  # Capture 3 screenshots from second monitor
  ScreenCapture.exe -c 3 --monitor 2

  # List all available monitors
  ScreenCapture.exe --list-monitors

  # Capture continuously with interval (use Ctrl+C to stop after desired count)
  ScreenCapture.exe -c 100 -i 5

NOTES:
  - Screenshots are saved in PNG format for best quality
  - All screenshots are automatically zipped for easy transfer
  - Monitor 0 captures the entire virtual screen (all monitors combined)
  - Monitor 1 is always the primary monitor
  - Temporary files are automatically cleaned up after ZIP creation
");
        }
    }

    class CommandLineOptions
    {
        public int Count { get; set; } = 1;
        public double Interval { get; set; } = 0;
        public double Fps { get; set; } = 0;  // 0 means not specified, use Interval instead
        public int InitialDelay { get; set; } = 0;
        public string OutputFile { get; set; }
        public int Monitor { get; set; } = 0;  // 0 = all screens, 1 = primary, 2+ = specific monitor
        public bool ListMonitors { get; set; }
    }

    class ScreenCapturer
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest,
            int nWidth, int nHeight, IntPtr hObjectSource,
            int nXSrc, int nYSrc, int dwRop);

        private const int SRCCOPY = 0x00CC0020;

        public void ListMonitors()
        {
            try
            {
                var screens = System.Windows.Forms.Screen.AllScreens;
                
                Console.WriteLine($"\nFound {screens.Length} monitor(s):\n");
                
                for (int i = 0; i < screens.Length; i++)
                {
                    var screen = screens[i];
                    int monitorNumber = i + 1;
                    
                    Console.WriteLine($"  Monitor #{monitorNumber}" + (screen.Primary ? " (Primary)" : ""));
                    Console.WriteLine($"    Resolution: {screen.Bounds.Width}x{screen.Bounds.Height}");
                    Console.WriteLine($"    Position:   X={screen.Bounds.X}, Y={screen.Bounds.Y}");
                    Console.WriteLine($"    Bits/Pixel: {screen.BitsPerPixel}");
                    Console.WriteLine($"    Device:     {screen.DeviceName}");
                    Console.WriteLine();
                }

                Console.WriteLine("  Virtual Screen (All Monitors Combined):");
                Console.WriteLine($"    Resolution: {System.Windows.Forms.SystemInformation.VirtualScreen.Width}x{System.Windows.Forms.SystemInformation.VirtualScreen.Height}");
                Console.WriteLine($"    Bounds:     {System.Windows.Forms.SystemInformation.VirtualScreen}");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing monitors: {ex.Message}");
            }
        }

        public bool CaptureScreen(string filePath, int monitorIndex = 0)
        {
            try
            {
                Rectangle bounds;

                if (monitorIndex == 0)
                {
                    // Capture entire virtual screen (all monitors)
                    bounds = System.Windows.Forms.SystemInformation.VirtualScreen;
                }
                else
                {
                    // Capture specific monitor
                    var screens = System.Windows.Forms.Screen.AllScreens;
                    
                    if (monitorIndex > screens.Length)
                    {
                        Console.WriteLine($"Error: Monitor #{monitorIndex} not found. Available monitors: 1-{screens.Length}");
                        return false;
                    }

                    var screen = screens[monitorIndex - 1];
                    bounds = screen.Bounds;
                }

                using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb))
                {
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
                    }

                    bitmap.Save(filePath, ImageFormat.Png);
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError capturing screen: {ex.Message}");
                return false;
            }
        }
    }
}
