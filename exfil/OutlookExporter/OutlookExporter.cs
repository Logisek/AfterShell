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
 * Outlook Data Exporter
 * Exports contacts or email recipients from Microsoft Outlook to CSV and/or JSON format.
 * Supports simultaneous output to CSV, JSON, and terminal matrix display.
 * Works whether Outlook is open or closed.
 * 
 * Requirements:
 *     - .NET 8.0 or later (Windows)
 *     - Microsoft Outlook installed
 * 
 * Build:
 *     dotnet build
 * 
 * Usage - Contacts:
 *     OutlookExporter.exe                              # Export to CSV (default)
 *     OutlookExporter.exe --json                       # Export to JSON only
 *     OutlookExporter.exe --csv --json                 # Export to both CSV and JSON
 *     OutlookExporter.exe --csv --json --matrix        # All three outputs at once
 *     OutlookExporter.exe --matrix                     # Display matrix only
 *     OutlookExporter.exe --output contacts --csv --json  # Custom filename
 *     OutlookExporter.exe --folder "Contacts/MyFolder"
 *     OutlookExporter.exe --list
 * 
 * Usage - Email Recipients:
 *     OutlookExporter.exe --recipients --limit 100
 *     OutlookExporter.exe --recipients --limit 100 --csv --json
 *     OutlookExporter.exe --recipients --limit 100 --csv --json --matrix
 *     OutlookExporter.exe --recipients --mailfolder "Sent Items" --limit 100
 *     OutlookExporter.exe --recipients --limit 100 --all-accounts
 *     OutlookExporter.exe --list-accounts
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace OutlookExporter
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

            var exporter = new OutlookDataExporter();

            try
            {
                if (!exporter.ConnectToOutlook())
                {
                    Console.WriteLine("Failed to connect to Outlook. Make sure Outlook is installed.");
                    return 1;
                }

                // Handle list accounts mode
                if (options.ListAccounts)
                {
                    exporter.ListEmailAccounts();
                    return 0;
                }

                // Handle list contact folders mode
                if (options.ListFolders)
                {
                    exporter.ListContactFolders();
                    return 0;
                }

                // Handle recipients export mode
                if (options.ExportRecipients)
                {
                    Console.WriteLine("\nProcessing email recipients...");
                    Console.WriteLine($"Mail folder: {options.MailFolder}");
                    if (options.Limit > 0)
                    {
                        Console.WriteLine($"Limit: {options.Limit} emails per folder");
                    }
                    if (options.AllAccounts)
                    {
                        Console.WriteLine("Mode: All accounts");
                    }
                    else if (!string.IsNullOrEmpty(options.Account))
                    {
                        Console.WriteLine($"Account: {options.Account}");
                    }

                    // Build list of output formats
                    var formats = new List<string>();
                    if (options.ExportCsv) formats.Add("CSV");
                    if (options.ExportJson) formats.Add("JSON");
                    if (options.ShowMatrix) formats.Add("Matrix");
                    Console.WriteLine($"Output: {string.Join(" + ", formats)}");

                    bool success = true;
                    int resultCount = 0;

                    // Export to CSV
                    if (options.ExportCsv)
                    {
                        string csvFile = GetOutputFilePath(options.OutputFile, ".csv");
                        Console.WriteLine($"\nExporting to CSV: {csvFile}");
                        int result = exporter.ExportEmailRecipientsToCSV(
                            csvFile,
                            options.MailFolder,
                            options.Limit,
                            options.Account,
                            options.AllAccounts);
                        if (result < 0) success = false;
                        else resultCount = result;
                    }

                    // Export to JSON
                    if (options.ExportJson)
                    {
                        string jsonFile = GetOutputFilePath(options.OutputFile, ".json");
                        Console.WriteLine($"\nExporting to JSON: {jsonFile}");
                        int result = exporter.ExportEmailRecipientsToJSON(
                            jsonFile,
                            options.MailFolder,
                            options.Limit,
                            options.Account,
                            options.AllAccounts);
                        if (result < 0) success = false;
                        else resultCount = result;
                    }

                    // Display matrix
                    if (options.ShowMatrix)
                    {
                        exporter.PrintRecipientsMatrix(
                            options.MailFolder,
                            options.Limit,
                            options.Account,
                            options.AllAccounts);
                    }

                    return success ? 0 : 1;
                }

                // Default: export contacts
                Console.WriteLine("\nProcessing contacts...");
                if (!string.IsNullOrEmpty(options.FolderPath))
                {
                    Console.WriteLine($"From folder: {options.FolderPath}");
                }

                // Build list of output formats
                var contactFormats = new List<string>();
                if (options.ExportCsv) contactFormats.Add("CSV");
                if (options.ExportJson) contactFormats.Add("JSON");
                if (options.ShowMatrix) contactFormats.Add("Matrix");
                Console.WriteLine($"Output: {string.Join(" + ", contactFormats)}");

                bool contactSuccess = true;
                int contactResultCount = 0;

                // Export to CSV
                if (options.ExportCsv)
                {
                    string csvFile = GetOutputFilePath(options.OutputFile, ".csv");
                    Console.WriteLine($"\nExporting to CSV: {csvFile}");
                    int result = exporter.ExportContactsToCSV(csvFile, options.FolderPath);
                    if (result < 0) contactSuccess = false;
                    else contactResultCount = result;
                }

                // Export to JSON
                if (options.ExportJson)
                {
                    string jsonFile = GetOutputFilePath(options.OutputFile, ".json");
                    Console.WriteLine($"\nExporting to JSON: {jsonFile}");
                    int result = exporter.ExportContactsToJSON(jsonFile, options.FolderPath);
                    if (result < 0) contactSuccess = false;
                    else contactResultCount = result;
                }

                // Display matrix
                if (options.ShowMatrix)
                {
                    exporter.PrintContactsMatrix(options.FolderPath);
                }

                if (contactSuccess && contactResultCount >= 0)
                {
                    Console.WriteLine($"\n✓ All operations completed successfully!");
                    return 0;
                }
                else
                {
                    Console.WriteLine("\n✗ Some operations failed");
                    return 1;
                }
            }
            finally
            {
                exporter.Cleanup();
            }
        }

        static CommandLineOptions ParseArguments(string[] args)
        {
            var options = new CommandLineOptions();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
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

                    case "-f":
                    case "--folder":
                        if (i + 1 < args.Length)
                        {
                            options.FolderPath = args[++i];
                        }
                        else
                        {
                            Console.WriteLine("Error: --folder requires a folder path");
                            return null;
                        }
                        break;

                    case "-l":
                    case "--list":
                        options.ListFolders = true;
                        break;

                    case "-r":
                    case "--recipients":
                        options.ExportRecipients = true;
                        break;

                    case "--limit":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int limit))
                        {
                            options.Limit = limit;
                        }
                        else
                        {
                            Console.WriteLine("Error: --limit requires a number");
                            return null;
                        }
                        break;

                    case "-m":
                    case "--mailfolder":
                        if (i + 1 < args.Length)
                        {
                            options.MailFolder = args[++i];
                        }
                        else
                        {
                            Console.WriteLine("Error: --mailfolder requires a folder name");
                            return null;
                        }
                        break;

                    case "-a":
                    case "--account":
                        if (i + 1 < args.Length)
                        {
                            options.Account = args[++i];
                        }
                        else
                        {
                            Console.WriteLine("Error: --account requires an email address");
                            return null;
                        }
                        break;

                    case "--all-accounts":
                        options.AllAccounts = true;
                        break;

                    case "--list-accounts":
                        options.ListAccounts = true;
                        break;

                    case "--csv":
                        options.ExportCsv = true;
                        break;

                    case "--json":
                        options.ExportJson = true;
                        break;

                    case "--matrix":
                        options.ShowMatrix = true;
                        break;

                    case "-h":
                    case "--help":
                        return null;

                    default:
                        Console.WriteLine($"Unknown option: {args[i]}");
                        return null;
                }
            }

            // If no output format specified, default to CSV
            if (!options.ExportCsv && !options.ExportJson && !options.ShowMatrix)
            {
                options.ExportCsv = true;
            }

            // Generate default output filename base if not specified and not listing
            if (string.IsNullOrEmpty(options.OutputFile) && !options.ListFolders && !options.ListAccounts)
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                if (options.ExportRecipients)
                {
                    options.OutputFile = $"outlook_recipients_{timestamp}";
                }
                else
                {
                    options.OutputFile = $"outlook_contacts_{timestamp}";
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
                Console.WriteLine("    Post-Exploitation Toolkit");
                Console.ResetColor();
                Console.WriteLine("    https://logisek.com | info@logisek.com");
                Console.WriteLine("    AfterShell | https://github.com/Logisek/AfterShell");
                Console.WriteLine();
            }
            catch
            {
                // Fallback without colors
                Console.WriteLine(asciiArt);
                Console.WriteLine("    Post-Exploitation Toolkit");
                Console.WriteLine("    https://logisek.com | info@logisek.com");
                Console.WriteLine("    AfterShell | https://github.com/Logisek/AfterShell");
                Console.WriteLine();
            }
        }

        static string GetOutputFilePath(string basePath, string extension)
        {
            if (string.IsNullOrEmpty(basePath))
            {
                return $"output{extension}";
            }

            // If the path already has an extension, replace it
            string existingExt = Path.GetExtension(basePath);
            if (!string.IsNullOrEmpty(existingExt))
            {
                return Path.ChangeExtension(basePath, extension);
            }

            // Otherwise, just append the extension
            return basePath + extension;
        }

        static void PrintUsage()
        {
            Console.WriteLine(@"
Outlook Data Exporter
Export contacts or email recipients from Microsoft Outlook to CSV/JSON format

Usage:
  OutlookExporter.exe [options]

CONTACTS MODE (default):
  -o, --output <file>     Output file path (default: outlook_contacts_YYYYMMDD_HHMMSS.csv/json)
  -f, --folder <path>     Specific contacts folder path (e.g., ""Contacts/MyFolder"")
  -l, --list              List all available contact folders and exit

EMAIL RECIPIENTS MODE:
  -r, --recipients        Export email recipients instead of contacts
  -m, --mailfolder <name> Mail folder to process (default: Inbox)
                          Options: Inbox, ""Sent Items"", Outbox, Drafts
  --limit <n>             Limit number of emails to process per folder
  -a, --account <email>   Process specific email account only
  --all-accounts          Process all configured email accounts
  --list-accounts         List all configured email accounts and exit

OUTPUT FORMAT OPTIONS (can be combined):
  --csv                   Export to CSV format (default if no format specified)
  --json                  Export to JSON format
  --matrix                Display results as a matrix table in terminal

COMMON OPTIONS:
  -o, --output <file>     Output file path
  -h, --help              Show this help message

EXAMPLES - Contacts:
  # Export all contacts to default CSV file
  OutlookExporter.exe

  # Export to JSON format only
  OutlookExporter.exe --json

  # Export to both CSV and JSON simultaneously
  OutlookExporter.exe --csv --json

  # Export to CSV, JSON and display matrix (all at once)
  OutlookExporter.exe --csv --json --matrix

  # Display contacts as matrix in terminal only
  OutlookExporter.exe --matrix

  # Export to CSV and also display matrix
  OutlookExporter.exe --csv --matrix

  # Export to specific base filename (extensions added automatically)
  OutlookExporter.exe --output my_contacts --csv --json

  # Export from specific folder
  OutlookExporter.exe --folder ""Contacts/Work Contacts""

  # List all contact folders
  OutlookExporter.exe --list

EXAMPLES - Email Recipients:
  # Export recipients from first 100 emails in Inbox (CSV by default)
  OutlookExporter.exe --recipients --limit 100

  # Export recipients to both CSV and JSON
  OutlookExporter.exe --recipients --limit 100 --csv --json

  # Export to CSV, JSON and display matrix simultaneously
  OutlookExporter.exe --recipients --limit 100 --csv --json --matrix

  # Display recipients as matrix in terminal only
  OutlookExporter.exe --recipients --limit 100 --matrix

  # Export to JSON and also display matrix
  OutlookExporter.exe --recipients --limit 100 --json --matrix

  # Export recipients from Sent Items folder
  OutlookExporter.exe --recipients --mailfolder ""Sent Items"" --limit 200

  # Export recipients from all accounts in all formats
  OutlookExporter.exe --recipients --limit 100 --all-accounts --csv --json --matrix

  # Export from specific account only
  OutlookExporter.exe --recipients --account ""work@company.com"" --limit 100

  # List all email accounts
  OutlookExporter.exe --list-accounts
");
        }
    }

    class CommandLineOptions
    {
        public string OutputFile { get; set; }
        public string FolderPath { get; set; }
        public bool ListFolders { get; set; }
        
        // Email recipients options
        public bool ExportRecipients { get; set; }
        public int Limit { get; set; } = 0;  // 0 = no limit
        public string MailFolder { get; set; } = "Inbox";
        public string Account { get; set; }
        public bool AllAccounts { get; set; }
        public bool ListAccounts { get; set; }
        
        // Output format options
        public bool ExportCsv { get; set; }
        public bool ExportJson { get; set; }
        public bool ShowMatrix { get; set; }
    }

    // Helper class to track recipient info with contact count and latest date
    class RecipientInfo
    {
        public string Email { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Account { get; set; }
        public int ContactCount { get; set; }
        public DateTime LatestContactDate { get; set; }
        public bool IsOwnAccount { get; set; }

        public RecipientInfo()
        {
            ContactCount = 0;
            LatestContactDate = DateTime.MinValue;
            IsOwnAccount = false;
        }

        public void UpdateContact(DateTime contactDate)
        {
            ContactCount++;
            if (contactDate > LatestContactDate)
            {
                LatestContactDate = contactDate;
            }
        }

        public string LatestContactDateString => 
            LatestContactDate > DateTime.MinValue ? LatestContactDate.ToString("yyyy-MM-dd HH:mm:ss") : "";
    }

    class OutlookDataExporter
    {
        private dynamic _outlook;
        private dynamic _namespace;

        // Outlook folder constants
        private const int olFolderContacts = 10;
        private const int olFolderInbox = 6;
        private const int olFolderSentMail = 5;
        private const int olFolderOutbox = 4;
        private const int olFolderDrafts = 16;
        
        // Outlook item class constants
        private const int olContactItem = 40;
        private const int olMailItem = 43;
        
        // Recipient type constants
        private const int olTo = 1;
        private const int olCC = 2;
        private const int olBCC = 3;

        public bool ConnectToOutlook()
        {
            try
            {
                Console.WriteLine("Connecting to Outlook...");
                
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    Console.WriteLine("Error: Outlook is not installed on this system.");
                    return false;
                }
                
                _outlook = Activator.CreateInstance(outlookType);
                _namespace = _outlook.GetNamespace("MAPI");
                
                Console.WriteLine("Successfully connected to Outlook");
                return true;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error connecting to Outlook: {ex.Message}");
                return false;
            }
        }

        public dynamic GetContactsFolder(string folderPath = null)
        {
            try
            {
                if (!string.IsNullOrEmpty(folderPath))
                {
                    // Navigate to specific folder
                    dynamic folder = _namespace.GetDefaultFolder(olFolderContacts);
                    string[] subfolderNames = folderPath.Split('/');

                    foreach (string subfolderName in subfolderNames)
                    {
                        if (subfolderName.ToLower() != "contacts")
                        {
                            folder = folder.Folders[subfolderName];
                        }
                    }
                    return folder;
                }
                else
                {
                    // Get default contacts folder
                    return _namespace.GetDefaultFolder(olFolderContacts);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error accessing contacts folder: {ex.Message}");
                return null;
            }
        }

        public int ExportContactsToCSV(string outputFile, string folderPath = null)
        {
            dynamic contactsFolder = GetContactsFolder(folderPath);

            if (contactsFolder == null)
            {
                Console.WriteLine("Failed to access contacts folder");
                return -1;
            }

            dynamic contacts = contactsFolder.Items;
            int contactCount = contacts.Count;

            Console.WriteLine($"Found {contactCount} items in contacts folder");

            if (contactCount == 0)
            {
                Console.WriteLine("No contacts to export");
                return 0;
            }

            // CSV headers
            string[] headers = new[]
            {
                "Full Name", "First Name", "Last Name", "Middle Name",
                "Email 1", "Email 2", "Email 3",
                "Company", "Job Title", "Department",
                "Business Phone", "Home Phone", "Mobile Phone", "Business Fax",
                "Business Address", "Home Address",
                "Birthday", "Anniversary",
                "Notes", "Categories"
            };

            try
            {
                using (var writer = new StreamWriter(outputFile, false, new UTF8Encoding(true)))
                {
                    // Write CSV headers
                    writer.WriteLine(string.Join(",", headers));

                    int exportedCount = 0;

                    for (int i = 1; i <= contactCount; i++)
                    {
                        dynamic item = null;
                        try
                        {
                            item = contacts[i];
                            
                            // Check if it's actually a contact item (Class = 40)
                            int itemClass = 0;
                            try { itemClass = item.Class; } catch { }
                            
                            if (itemClass == olContactItem)
                            {
                                var row = new[]
                                {
                                    CsvEscape(SafeGetProperty(item, "FullName")),
                                    CsvEscape(SafeGetProperty(item, "FirstName")),
                                    CsvEscape(SafeGetProperty(item, "LastName")),
                                    CsvEscape(SafeGetProperty(item, "MiddleName")),
                                    CsvEscape(SafeGetProperty(item, "Email1Address")),
                                    CsvEscape(SafeGetProperty(item, "Email2Address")),
                                    CsvEscape(SafeGetProperty(item, "Email3Address")),
                                    CsvEscape(SafeGetProperty(item, "CompanyName")),
                                    CsvEscape(SafeGetProperty(item, "JobTitle")),
                                    CsvEscape(SafeGetProperty(item, "Department")),
                                    CsvEscape(SafeGetProperty(item, "BusinessTelephoneNumber")),
                                    CsvEscape(SafeGetProperty(item, "HomeTelephoneNumber")),
                                    CsvEscape(SafeGetProperty(item, "MobileTelephoneNumber")),
                                    CsvEscape(SafeGetProperty(item, "BusinessFaxNumber")),
                                    CsvEscape(SafeGetProperty(item, "BusinessAddress")),
                                    CsvEscape(SafeGetProperty(item, "HomeAddress")),
                                    CsvEscape(SafeGetDateProperty(item, "Birthday")),
                                    CsvEscape(SafeGetDateProperty(item, "Anniversary")),
                                    CsvEscape(SafeGetProperty(item, "Body")),
                                    CsvEscape(SafeGetProperty(item, "Categories"))
                                };

                                writer.WriteLine(string.Join(",", row));
                                exportedCount++;

                                if (exportedCount % 10 == 0)
                                {
                                    Console.Write($"\rExported {exportedCount} contacts...");
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Console.WriteLine($"\nWarning: Error processing item {i}: {ex.Message}");
                            continue;
                        }
                        finally
                        {
                            if (item != null)
                            {
                                try { Marshal.ReleaseComObject(item); } catch { }
                            }
                        }
                    }

                    Console.WriteLine($"\nSuccessfully exported {exportedCount} contacts to {outputFile}");
                    
                    // Release COM objects
                    try { Marshal.ReleaseComObject(contacts); } catch { }
                    try { Marshal.ReleaseComObject(contactsFolder); } catch { }
                    
                    return exportedCount;
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error writing CSV file: {ex.Message}");
                return -1;
            }
        }

        public int ExportContactsToJSON(string outputFile, string folderPath = null)
        {
            dynamic contactsFolder = GetContactsFolder(folderPath);

            if (contactsFolder == null)
            {
                Console.WriteLine("Failed to access contacts folder");
                return -1;
            }

            dynamic contacts = contactsFolder.Items;
            int contactCount = contacts.Count;

            Console.WriteLine($"Found {contactCount} items in contacts folder");

            if (contactCount == 0)
            {
                Console.WriteLine("No contacts to export");
                return 0;
            }

            try
            {
                var contactsList = new List<Dictionary<string, string>>();
                int exportedCount = 0;

                for (int i = 1; i <= contactCount; i++)
                {
                    dynamic item = null;
                    try
                    {
                        item = contacts[i];
                        
                        int itemClass = 0;
                        try { itemClass = item.Class; } catch { }
                        
                        if (itemClass == olContactItem)
                        {
                            var contact = new Dictionary<string, string>
                            {
                                ["FullName"] = SafeGetProperty(item, "FullName"),
                                ["FirstName"] = SafeGetProperty(item, "FirstName"),
                                ["LastName"] = SafeGetProperty(item, "LastName"),
                                ["MiddleName"] = SafeGetProperty(item, "MiddleName"),
                                ["Email1"] = SafeGetProperty(item, "Email1Address"),
                                ["Email2"] = SafeGetProperty(item, "Email2Address"),
                                ["Email3"] = SafeGetProperty(item, "Email3Address"),
                                ["Company"] = SafeGetProperty(item, "CompanyName"),
                                ["JobTitle"] = SafeGetProperty(item, "JobTitle"),
                                ["Department"] = SafeGetProperty(item, "Department"),
                                ["BusinessPhone"] = SafeGetProperty(item, "BusinessTelephoneNumber"),
                                ["HomePhone"] = SafeGetProperty(item, "HomeTelephoneNumber"),
                                ["MobilePhone"] = SafeGetProperty(item, "MobileTelephoneNumber"),
                                ["BusinessFax"] = SafeGetProperty(item, "BusinessFaxNumber"),
                                ["BusinessAddress"] = SafeGetProperty(item, "BusinessAddress"),
                                ["HomeAddress"] = SafeGetProperty(item, "HomeAddress"),
                                ["Birthday"] = SafeGetDateProperty(item, "Birthday"),
                                ["Anniversary"] = SafeGetDateProperty(item, "Anniversary"),
                                ["Notes"] = SafeGetProperty(item, "Body"),
                                ["Categories"] = SafeGetProperty(item, "Categories")
                            };

                            contactsList.Add(contact);
                            exportedCount++;

                            if (exportedCount % 10 == 0)
                            {
                                Console.Write($"\rExported {exportedCount} contacts...");
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine($"\nWarning: Error processing item {i}: {ex.Message}");
                        continue;
                    }
                    finally
                    {
                        if (item != null)
                        {
                            try { Marshal.ReleaseComObject(item); } catch { }
                        }
                    }
                }

                // Write JSON
                using (var writer = new StreamWriter(outputFile, false, new UTF8Encoding(true)))
                {
                    writer.WriteLine(SerializeToJson(contactsList));
                }

                Console.WriteLine($"\nSuccessfully exported {exportedCount} contacts to {outputFile}");
                
                try { Marshal.ReleaseComObject(contacts); } catch { }
                try { Marshal.ReleaseComObject(contactsFolder); } catch { }
                
                return exportedCount;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error writing JSON file: {ex.Message}");
                return -1;
            }
        }

        public List<Dictionary<string, string>> GetContactsData(string folderPath = null)
        {
            var contactsList = new List<Dictionary<string, string>>();
            dynamic contactsFolder = GetContactsFolder(folderPath);

            if (contactsFolder == null)
            {
                return contactsList;
            }

            dynamic contacts = contactsFolder.Items;
            int contactCount = contacts.Count;

            for (int i = 1; i <= contactCount; i++)
            {
                dynamic item = null;
                try
                {
                    item = contacts[i];
                    
                    int itemClass = 0;
                    try { itemClass = item.Class; } catch { }
                    
                    if (itemClass == olContactItem)
                    {
                        var contact = new Dictionary<string, string>
                        {
                            ["FullName"] = SafeGetProperty(item, "FullName"),
                            ["Email1"] = SafeGetProperty(item, "Email1Address"),
                            ["Company"] = SafeGetProperty(item, "CompanyName"),
                            ["JobTitle"] = SafeGetProperty(item, "JobTitle"),
                            ["MobilePhone"] = SafeGetProperty(item, "MobileTelephoneNumber"),
                            ["BusinessPhone"] = SafeGetProperty(item, "BusinessTelephoneNumber")
                        };
                        contactsList.Add(contact);
                    }
                }
                catch { continue; }
                finally
                {
                    if (item != null) { try { Marshal.ReleaseComObject(item); } catch { } }
                }
            }

            try { Marshal.ReleaseComObject(contacts); } catch { }
            try { Marshal.ReleaseComObject(contactsFolder); } catch { }

            return contactsList;
        }

        public void PrintContactsMatrix(string folderPath = null)
        {
            var contacts = GetContactsData(folderPath);
            
            if (contacts.Count == 0)
            {
                Console.WriteLine("No contacts to display.");
                return;
            }

            // Define columns and their max widths
            var columns = new[] { "FullName", "Email1", "Company", "JobTitle", "MobilePhone", "BusinessPhone" };
            var headers = new[] { "Name", "Email", "Company", "Job Title", "Mobile", "Business Phone" };
            var widths = new int[columns.Length];

            // Calculate column widths
            for (int c = 0; c < columns.Length; c++)
            {
                widths[c] = headers[c].Length;
                foreach (var contact in contacts)
                {
                    if (contact.TryGetValue(columns[c], out string val) && !string.IsNullOrEmpty(val))
                    {
                        widths[c] = Math.Max(widths[c], Math.Min(val.Length, 40)); // Cap at 40 chars
                    }
                }
            }

            // Print header separator
            Console.WriteLine();
            PrintMatrixSeparator(widths);
            
            // Print header row
            Console.Write("|");
            for (int c = 0; c < headers.Length; c++)
            {
                Console.Write($" {TruncateString(headers[c], widths[c]).PadRight(widths[c])} |");
            }
            Console.WriteLine();
            
            PrintMatrixSeparator(widths);

            // Print data rows
            foreach (var contact in contacts)
            {
                Console.Write("|");
                for (int c = 0; c < columns.Length; c++)
                {
                    string val = contact.TryGetValue(columns[c], out string v) ? v : "";
                    Console.Write($" {TruncateString(val, widths[c]).PadRight(widths[c])} |");
                }
                Console.WriteLine();
            }

            PrintMatrixSeparator(widths);
            Console.WriteLine($"\nTotal: {contacts.Count} contacts");
        }

        private void PrintMatrixSeparator(int[] widths)
        {
            Console.Write("+");
            foreach (var w in widths)
            {
                Console.Write(new string('-', w + 2) + "+");
            }
            Console.WriteLine();
        }

        private string TruncateString(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Length <= maxLength) return value;
            return value.Substring(0, maxLength - 3) + "...";
        }

        public void ListContactFolders()
        {
            try
            {
                dynamic contactsFolder = _namespace.GetDefaultFolder(olFolderContacts);
                Console.WriteLine($"\nMain Contacts Folder: {contactsFolder.Name}");
                Console.WriteLine($"Number of items: {contactsFolder.Items.Count}");

                int folderCount = contactsFolder.Folders.Count;
                if (folderCount > 0)
                {
                    Console.WriteLine("\nSubfolders:");
                    ListSubfolders(contactsFolder, "  ");
                }

                Marshal.ReleaseComObject(contactsFolder);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error listing folders: {ex.Message}");
            }
        }

        private void ListSubfolders(dynamic folder, string indent)
        {
            try
            {
                foreach (dynamic subfolder in folder.Folders)
                {
                    Console.WriteLine($"{indent}- {subfolder.Name} ({subfolder.Items.Count} items)");
                    
                    int subCount = subfolder.Folders.Count;
                    if (subCount > 0)
                    {
                        ListSubfolders(subfolder, indent + "  ");
                    }
                    
                    Marshal.ReleaseComObject(subfolder);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"{indent}Error listing subfolder: {ex.Message}");
            }
        }

        public void ListEmailAccounts()
        {
            try
            {
                dynamic accounts = _namespace.Accounts;
                int accountCount = accounts.Count;
                
                Console.WriteLine($"\nFound {accountCount} email account(s):\n");
                
                for (int i = 1; i <= accountCount; i++)
                {
                    dynamic account = accounts[i];
                    try
                    {
                        string displayName = SafeGetProperty(account, "DisplayName");
                        string smtpAddress = SafeGetProperty(account, "SmtpAddress");
                        string accountType = SafeGetProperty(account, "AccountType");
                        
                        Console.WriteLine($"  {i}. {displayName}");
                        Console.WriteLine($"     Email: {smtpAddress}");
                        Console.WriteLine($"     Type:  {accountType}");
                        
                        // Try to get folder info
                        try
                        {
                            dynamic store = account.DeliveryStore;
                            if (store != null)
                            {
                                Console.WriteLine($"     Store: {store.DisplayName}");
                                Marshal.ReleaseComObject(store);
                            }
                        }
                        catch { }
                        
                        Console.WriteLine();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(account);
                    }
                }
                
                Marshal.ReleaseComObject(accounts);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error listing accounts: {ex.Message}");
            }
        }

        public List<(string SmtpAddress, string DisplayName, dynamic Store)> GetAllAccounts()
        {
            var result = new List<(string, string, dynamic)>();
            
            try
            {
                dynamic accounts = _namespace.Accounts;
                
                for (int i = 1; i <= accounts.Count; i++)
                {
                    dynamic account = accounts[i];
                    try
                    {
                        string smtpAddress = SafeGetProperty(account, "SmtpAddress");
                        string displayName = SafeGetProperty(account, "DisplayName");
                        dynamic store = null;
                        
                        try { store = account.DeliveryStore; } catch { }
                        
                        if (!string.IsNullOrEmpty(smtpAddress))
                        {
                            result.Add((smtpAddress, displayName, store));
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(account);
                    }
                }
                
                Marshal.ReleaseComObject(accounts);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error getting accounts: {ex.Message}");
            }
            
            return result;
        }

        public dynamic GetMailFolder(string folderName = "Inbox", string accountEmail = null)
        {
            try
            {
                if (!string.IsNullOrEmpty(accountEmail))
                {
                    return GetMailFolderForAccount(accountEmail, folderName);
                }
                
                // Get default folder based on name
                int folderId = folderName.ToLower() switch
                {
                    "inbox" => olFolderInbox,
                    "sent" or "sent items" or "sentmail" => olFolderSentMail,
                    "outbox" => olFolderOutbox,
                    "drafts" => olFolderDrafts,
                    _ => olFolderInbox
                };
                
                return _namespace.GetDefaultFolder(folderId);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error accessing mail folder: {ex.Message}");
                return null;
            }
        }

        public dynamic GetMailFolderForAccount(string accountEmail, string folderName)
        {
            try
            {
                dynamic accounts = _namespace.Accounts;
                
                for (int i = 1; i <= accounts.Count; i++)
                {
                    dynamic account = accounts[i];
                    string smtpAddress = SafeGetProperty(account, "SmtpAddress");
                    
                    if (smtpAddress.Equals(accountEmail, StringComparison.OrdinalIgnoreCase))
                    {
                        dynamic store = account.DeliveryStore;
                        if (store != null)
                        {
                            dynamic rootFolder = store.GetRootFolder();
                            
                            // Navigate to find the desired folder
                            foreach (dynamic folder in rootFolder.Folders)
                            {
                                string name = SafeGetProperty(folder, "Name");
                                
                                // Match common folder names
                                bool isMatch = folderName.ToLower() switch
                                {
                                    "inbox" => name.ToLower().Contains("inbox") || name.ToLower() == "postvak in",
                                    "sent" or "sent items" or "sentmail" => name.ToLower().Contains("sent") || name.ToLower().Contains("verzonden"),
                                    "outbox" => name.ToLower().Contains("outbox") || name.ToLower().Contains("postvak uit"),
                                    "drafts" => name.ToLower().Contains("draft") || name.ToLower().Contains("concepten"),
                                    _ => name.Equals(folderName, StringComparison.OrdinalIgnoreCase)
                                };
                                
                                if (isMatch)
                                {
                                    Marshal.ReleaseComObject(rootFolder);
                                    Marshal.ReleaseComObject(account);
                                    Marshal.ReleaseComObject(accounts);
                                    return folder;
                                }
                                
                                Marshal.ReleaseComObject(folder);
                            }
                            
                            Marshal.ReleaseComObject(rootFolder);
                            Marshal.ReleaseComObject(store);
                        }
                    }
                    
                    Marshal.ReleaseComObject(account);
                }
                
                Marshal.ReleaseComObject(accounts);
                Console.WriteLine($"Could not find folder '{folderName}' for account '{accountEmail}'");
                return null;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error accessing mail folder for account: {ex.Message}");
                return null;
            }
        }

        public dynamic GetMailFolderFromStore(dynamic store, string folderName)
        {
            try
            {
                if (store == null) return null;
                
                dynamic rootFolder = store.GetRootFolder();
                
                foreach (dynamic folder in rootFolder.Folders)
                {
                    string name = SafeGetProperty(folder, "Name");
                    
                    bool isMatch = folderName.ToLower() switch
                    {
                        "inbox" => name.ToLower().Contains("inbox") || name.ToLower() == "postvak in",
                        "sent" or "sent items" or "sentmail" => name.ToLower().Contains("sent") || name.ToLower().Contains("verzonden"),
                        "outbox" => name.ToLower().Contains("outbox") || name.ToLower().Contains("postvak uit"),
                        "drafts" => name.ToLower().Contains("draft") || name.ToLower().Contains("concepten"),
                        _ => name.Equals(folderName, StringComparison.OrdinalIgnoreCase)
                    };
                    
                    if (isMatch)
                    {
                        Marshal.ReleaseComObject(rootFolder);
                        return folder;
                    }
                    
                    Marshal.ReleaseComObject(folder);
                }
                
                Marshal.ReleaseComObject(rootFolder);
                return null;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error accessing folder from store: {ex.Message}");
                return null;
            }
        }

        public int ExportEmailRecipientsToCSV(string outputFile, string mailFolder, int limit, string accountEmail = null, bool allAccounts = false)
        {
            try
            {
                Console.WriteLine("\nCollecting recipient data...");
                var recipientsMap = GetAggregatedRecipientsData(mailFolder, limit, accountEmail, allAccounts);
                
                // Get all configured accounts to mark own accounts
                var ownAccounts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var accounts = GetAllAccounts();
                foreach (var (smtpAddress, displayName, store) in accounts)
                {
                    if (!string.IsNullOrEmpty(smtpAddress))
                    {
                        ownAccounts.Add(smtpAddress.ToLower());
                    }
                    if (store != null) { try { Marshal.ReleaseComObject(store); } catch { } }
                }
                
                // Sort by ContactCount descending, then by LatestContactDate descending
                var sortedRecipients = recipientsMap.Values
                    .OrderByDescending(r => r.ContactCount)
                    .ThenByDescending(r => r.LatestContactDate)
                    .ToList();
                
                using (var writer = new StreamWriter(outputFile, false, new UTF8Encoding(true)))
                {
                    // CSV headers - added ContactCount, LatestContactDate, and IsOwnAccount
                    writer.WriteLine("Email,Name,Type,Account,ContactCount,LatestContactDate,IsOwnAccount");
                    
                    foreach (var recipient in sortedRecipients)
                    {
                        bool isOwn = ownAccounts.Contains(recipient.Email.ToLower());
                        writer.WriteLine($"{CsvEscape(recipient.Email)},{CsvEscape(recipient.Name)},{recipient.Type},{CsvEscape(recipient.Account)},{recipient.ContactCount},{recipient.LatestContactDateString},{(isOwn ? "Yes" : "")}");
                    }
                }
                
                Console.WriteLine($"\n✓ Export completed!");
                Console.WriteLine($"  Unique recipients found: {sortedRecipients.Count}");
                Console.WriteLine($"  File: {Path.GetFullPath(outputFile)}");
                
                return sortedRecipients.Count;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error exporting recipients: {ex.Message}");
                return -1;
            }
        }

        public int ExportEmailRecipientsToJSON(string outputFile, string mailFolder, int limit, string accountEmail = null, bool allAccounts = false)
        {
            try
            {
                Console.WriteLine("\nCollecting recipient data...");
                var recipientsMap = GetAggregatedRecipientsData(mailFolder, limit, accountEmail, allAccounts);
                
                // Get all configured accounts to mark own accounts
                var ownAccounts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var accounts = GetAllAccounts();
                foreach (var (smtpAddress, displayName, store) in accounts)
                {
                    if (!string.IsNullOrEmpty(smtpAddress))
                    {
                        ownAccounts.Add(smtpAddress.ToLower());
                    }
                    if (store != null) { try { Marshal.ReleaseComObject(store); } catch { } }
                }
                
                // Sort by ContactCount descending, then by LatestContactDate descending
                var sortedRecipients = recipientsMap.Values
                    .OrderByDescending(r => r.ContactCount)
                    .ThenByDescending(r => r.LatestContactDate)
                    .ToList();
                
                // Convert to list of dictionaries for JSON serialization
                var recipientsList = new List<Dictionary<string, string>>();
                foreach (var recipient in sortedRecipients)
                {
                    bool isOwn = ownAccounts.Contains(recipient.Email.ToLower());
                    recipientsList.Add(new Dictionary<string, string>
                    {
                        ["Email"] = recipient.Email,
                        ["Name"] = recipient.Name,
                        ["Type"] = recipient.Type,
                        ["Account"] = recipient.Account,
                        ["ContactCount"] = recipient.ContactCount.ToString(),
                        ["LatestContactDate"] = recipient.LatestContactDateString,
                        ["IsOwnAccount"] = isOwn ? "Yes" : ""
                    });
                }
                
                // Write JSON
                using (var writer = new StreamWriter(outputFile, false, new UTF8Encoding(true)))
                {
                    writer.WriteLine(SerializeToJson(recipientsList));
                }
                
                Console.WriteLine($"\n✓ Export completed!");
                Console.WriteLine($"  Unique recipients found: {sortedRecipients.Count}");
                Console.WriteLine($"  File: {Path.GetFullPath(outputFile)}");
                
                return sortedRecipients.Count;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error exporting recipients to JSON: {ex.Message}");
                return -1;
            }
        }

        public List<Dictionary<string, string>> GetRecipientsData(string mailFolder, int limit, string accountEmail = null, bool allAccounts = false)
        {
            var aggregatedRecipients = GetAggregatedRecipientsData(mailFolder, limit, accountEmail, allAccounts);
            
            // Get all configured accounts to mark own accounts
            var ownAccounts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var accounts = GetAllAccounts();
            foreach (var (smtpAddress, displayName, store) in accounts)
            {
                if (!string.IsNullOrEmpty(smtpAddress))
                {
                    ownAccounts.Add(smtpAddress.ToLower());
                }
                if (store != null) { try { Marshal.ReleaseComObject(store); } catch { } }
            }
            
            // Mark own accounts and convert to list of dictionaries, sorted by ContactCount descending
            var recipientsList = new List<Dictionary<string, string>>();
            var sortedRecipients = aggregatedRecipients.Values
                .OrderByDescending(r => r.ContactCount)
                .ThenByDescending(r => r.LatestContactDate)
                .ToList();
            
            foreach (var recipient in sortedRecipients)
            {
                bool isOwn = ownAccounts.Contains(recipient.Email.ToLower());
                recipientsList.Add(new Dictionary<string, string>
                {
                    ["Email"] = recipient.Email,
                    ["Name"] = recipient.Name,
                    ["Type"] = recipient.Type,
                    ["Account"] = recipient.Account,
                    ["ContactCount"] = recipient.ContactCount.ToString(),
                    ["LatestContactDate"] = recipient.LatestContactDateString,
                    ["IsOwnAccount"] = isOwn ? "Yes" : ""
                });
            }
            
            return recipientsList;
        }

        public Dictionary<string, RecipientInfo> GetAggregatedRecipientsData(string mailFolder, int limit, string accountEmail = null, bool allAccounts = false)
        {
            var recipientsMap = new Dictionary<string, RecipientInfo>(StringComparer.OrdinalIgnoreCase);
            int totalProcessed = 0;
            
            if (allAccounts)
            {
                var accounts = GetAllAccounts();
                foreach (var (smtpAddress, displayName, store) in accounts)
                {
                    dynamic folder = GetMailFolderFromStore(store, mailFolder);
                    if (folder != null)
                    {
                        totalProcessed += ProcessMailFolderAggregated(folder, recipientsMap, limit, smtpAddress);
                        Marshal.ReleaseComObject(folder);
                    }
                    if (store != null) { try { Marshal.ReleaseComObject(store); } catch { } }
                }
            }
            else
            {
                dynamic folder = GetMailFolder(mailFolder, accountEmail);
                if (folder != null)
                {
                    string account = accountEmail ?? "default";
                    totalProcessed = ProcessMailFolderAggregated(folder, recipientsMap, limit, account);
                    Marshal.ReleaseComObject(folder);
                }
            }
            
            return recipientsMap;
        }

        private int ProcessMailFolderAggregated(dynamic folder, Dictionary<string, RecipientInfo> recipientsMap, int limit, string accountName)
        {
            try
            {
                dynamic items = folder.Items;
                
                try { items.Sort("[ReceivedTime]", true); } catch { }
                
                int itemCount = items.Count;
                int processLimit = limit > 0 ? Math.Min(limit, itemCount) : itemCount;
                
                int processed = 0;
                
                for (int i = 1; i <= processLimit; i++)
                {
                    dynamic item = null;
                    try
                    {
                        item = items[i];
                        
                        int itemClass = 0;
                        try { itemClass = item.Class; } catch { }
                        
                        if (itemClass == olMailItem)
                        {
                            processed++;
                            
                            DateTime emailDate = GetDateTimeProperty(item, "ReceivedTime");
                            
                            // Process sender
                            string senderEmail = GetSenderEmail(item);
                            string senderName = SafeGetProperty(item, "SenderName");
                            
                            if (!string.IsNullOrEmpty(senderEmail))
                            {
                                string emailKey = senderEmail.ToLower();
                                if (!recipientsMap.TryGetValue(emailKey, out var senderInfo))
                                {
                                    senderInfo = new RecipientInfo
                                    {
                                        Email = senderEmail,
                                        Name = senderName,
                                        Type = "Sender",
                                        Account = accountName
                                    };
                                    recipientsMap[emailKey] = senderInfo;
                                }
                                senderInfo.UpdateContact(emailDate);
                            }
                            
                            // Process all recipients
                            try
                            {
                                dynamic recipients = item.Recipients;
                                for (int r = 1; r <= recipients.Count; r++)
                                {
                                    dynamic recipient = null;
                                    try
                                    {
                                        recipient = recipients[r];
                                        string email = GetSmtpAddress(recipient);
                                        string name = SafeGetProperty(recipient, "Name");
                                        string type = GetRecipientType(recipient);
                                        
                                        if (!string.IsNullOrEmpty(email))
                                        {
                                            string emailKey = email.ToLower();
                                            if (!recipientsMap.TryGetValue(emailKey, out var recipientInfo))
                                            {
                                                recipientInfo = new RecipientInfo
                                                {
                                                    Email = email,
                                                    Name = name,
                                                    Type = type,
                                                    Account = accountName
                                                };
                                                recipientsMap[emailKey] = recipientInfo;
                                            }
                                            recipientInfo.UpdateContact(emailDate);
                                        }
                                    }
                                    finally
                                    {
                                        if (recipient != null)
                                        {
                                            try { Marshal.ReleaseComObject(recipient); } catch { }
                                        }
                                    }
                                }
                                Marshal.ReleaseComObject(recipients);
                            }
                            catch { }
                        }
                    }
                    catch { continue; }
                    finally
                    {
                        if (item != null)
                        {
                            try { Marshal.ReleaseComObject(item); } catch { }
                        }
                    }
                }
                
                try { Marshal.ReleaseComObject(items); } catch { }
                
                return processed;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error processing mail folder: {ex.Message}");
                return 0;
            }
        }

        private DateTime GetDateTimeProperty(dynamic obj, string propertyName)
        {
            try
            {
                object value = obj.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    obj,
                    null);

                if (value is DateTime dateValue)
                {
                    return dateValue;
                }
                return DateTime.MinValue;
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        public void PrintRecipientsMatrix(string mailFolder, int limit, string accountEmail = null, bool allAccounts = false)
        {
            Console.WriteLine("\nCollecting recipients data...");
            var recipients = GetRecipientsData(mailFolder, limit, accountEmail, allAccounts);
            
            if (recipients.Count == 0)
            {
                Console.WriteLine("No recipients to display.");
                return;
            }

            // Count own accounts
            int ownAccountCount = recipients.Count(r => r.TryGetValue("IsOwnAccount", out string v) && v == "Yes");

            // Define columns and their max widths (now includes ContactCount, LatestContactDate, and Own indicator)
            var columns = new[] { "Email", "Name", "Type", "Account", "ContactCount", "LatestContactDate", "IsOwnAccount" };
            var headers = new[] { "Email", "Name", "Type", "Account", "Times Contacted", "Latest Contact", "Own" };
            var widths = new int[columns.Length];

            // Calculate column widths
            for (int c = 0; c < columns.Length; c++)
            {
                widths[c] = headers[c].Length;
                foreach (var recipient in recipients)
                {
                    if (recipient.TryGetValue(columns[c], out string val) && !string.IsNullOrEmpty(val))
                    {
                        // For IsOwnAccount, display as "*" for Yes
                        if (columns[c] == "IsOwnAccount" && val == "Yes")
                        {
                            widths[c] = Math.Max(widths[c], 3); // "*" needs at least 1 char
                        }
                        else
                        {
                            widths[c] = Math.Max(widths[c], Math.Min(val.Length, 50)); // Cap at 50 chars
                        }
                    }
                }
            }

            // Print header separator
            Console.WriteLine();
            PrintMatrixSeparator(widths);
            
            // Print header row
            Console.Write("|");
            for (int c = 0; c < headers.Length; c++)
            {
                Console.Write($" {TruncateString(headers[c], widths[c]).PadRight(widths[c])} |");
            }
            Console.WriteLine();
            
            PrintMatrixSeparator(widths);

            // Print data rows (already sorted by ContactCount descending)
            foreach (var recipient in recipients)
            {
                bool isOwnAccount = recipient.TryGetValue("IsOwnAccount", out string ownVal) && ownVal == "Yes";
                
                Console.Write("|");
                for (int c = 0; c < columns.Length; c++)
                {
                    string val = recipient.TryGetValue(columns[c], out string v) ? v : "";
                    
                    // Display "*" for own accounts instead of "Yes"
                    if (columns[c] == "IsOwnAccount")
                    {
                        val = isOwnAccount ? "*" : "";
                    }
                    
                    // Highlight own accounts with color if supported
                    if (isOwnAccount && columns[c] == "Email")
                    {
                        try
                        {
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.Write($" {TruncateString(val, widths[c]).PadRight(widths[c])} ");
                            Console.ResetColor();
                            Console.Write("|");
                        }
                        catch
                        {
                            Console.Write($" {TruncateString(val, widths[c]).PadRight(widths[c])} |");
                        }
                    }
                    else
                    {
                        Console.Write($" {TruncateString(val, widths[c]).PadRight(widths[c])} |");
                    }
                }
                Console.WriteLine();
            }

            PrintMatrixSeparator(widths);
            Console.WriteLine($"\nTotal: {recipients.Count} unique recipients (sorted by most contacted)");
            if (ownAccountCount > 0)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write($"* = Own account");
                Console.ResetColor();
                Console.WriteLine($" ({ownAccountCount} account(s) used for export)");
            }
        }

        private string SerializeToJson(List<Dictionary<string, string>> items)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            
            for (int i = 0; i < items.Count; i++)
            {
                sb.Append("  {");
                var pairs = new List<string>();
                
                foreach (var kvp in items[i])
                {
                    string escapedValue = JsonEscape(kvp.Value);
                    pairs.Add($"\"{kvp.Key}\": \"{escapedValue}\"");
                }
                
                sb.Append(string.Join(", ", pairs));
                sb.Append("}");
                
                if (i < items.Count - 1)
                {
                    sb.AppendLine(",");
                }
                else
                {
                    sb.AppendLine();
                }
            }
            
            sb.Append("]");
            return sb.ToString();
        }

        private string JsonEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
        }

        private string GetSenderEmail(dynamic mailItem)
        {
            try
            {
                // First try the direct sender email address
                string senderAddress = SafeGetProperty(mailItem, "SenderEmailAddress");
                
                // Check if it's an Exchange address (starts with /o= or /O=)
                if (!string.IsNullOrEmpty(senderAddress) && !senderAddress.StartsWith("/"))
                {
                    return senderAddress;
                }
                
                // Try to get SMTP address from sender
                try
                {
                    dynamic sender = mailItem.Sender;
                    if (sender != null)
                    {
                        string smtpAddress = GetSmtpAddressFromAddressEntry(sender);
                        Marshal.ReleaseComObject(sender);
                        if (!string.IsNullOrEmpty(smtpAddress))
                        {
                            return smtpAddress;
                        }
                    }
                }
                catch { }
                
                // Return the original address if we can't resolve it
                return senderAddress;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetSmtpAddress(dynamic recipient)
        {
            try
            {
                dynamic addressEntry = recipient.AddressEntry;
                string smtpAddress = GetSmtpAddressFromAddressEntry(addressEntry);
                Marshal.ReleaseComObject(addressEntry);
                return smtpAddress;
            }
            catch
            {
                return SafeGetProperty(recipient, "Address");
            }
        }

        private string GetSmtpAddressFromAddressEntry(dynamic addressEntry)
        {
            try
            {
                string addressType = SafeGetProperty(addressEntry, "Type");
                
                if (addressType.ToUpper() == "SMTP")
                {
                    return SafeGetProperty(addressEntry, "Address");
                }
                
                // Try to get Exchange user's SMTP address
                try
                {
                    int userType = addressEntry.AddressEntryUserType;
                    
                    // olExchangeUserAddressEntry = 0, olExchangeDistributionListAddressEntry = 1
                    if (userType == 0 || userType == 1)
                    {
                        dynamic exchUser = addressEntry.GetExchangeUser();
                        if (exchUser != null)
                        {
                            string primarySmtp = SafeGetProperty(exchUser, "PrimarySmtpAddress");
                            Marshal.ReleaseComObject(exchUser);
                            if (!string.IsNullOrEmpty(primarySmtp))
                            {
                                return primarySmtp;
                            }
                        }
                    }
                }
                catch { }
                
                // Try PropertyAccessor for PR_SMTP_ADDRESS
                try
                {
                    string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";
                    dynamic propAccessor = addressEntry.PropertyAccessor;
                    string smtpAddr = propAccessor.GetProperty(PR_SMTP_ADDRESS);
                    Marshal.ReleaseComObject(propAccessor);
                    if (!string.IsNullOrEmpty(smtpAddr))
                    {
                        return smtpAddr;
                    }
                }
                catch { }
                
                // Fall back to the Address property
                return SafeGetProperty(addressEntry, "Address");
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetRecipientType(dynamic recipient)
        {
            try
            {
                int type = recipient.Type;
                return type switch
                {
                    olTo => "To (Recipient)",
                    olCC => "CC (Recipient)",
                    olBCC => "BCC (Recipient)",
                    _ => "Unknown"
                };
            }
            catch
            {
                return "Unknown";
            }
        }

        private string SafeGetDateTimeProperty(dynamic obj, string propertyName)
        {
            try
            {
                object value = obj.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    obj,
                    null);

                if (value is DateTime dateValue)
                {
                    return dateValue.ToString("yyyy-MM-dd HH:mm:ss");
                }
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeGetProperty(dynamic obj, string propertyName)
        {
            try
            {
                object value = obj.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    obj,
                    null);
                    
                return value?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeGetDateProperty(dynamic obj, string propertyName)
        {
            try
            {
                object value = obj.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    obj,
                    null);

                if (value is DateTime dateValue)
                {
                    // Check if date is not the default/empty DateTime (Outlook uses 1/1/4501 for empty dates)
                    if (dateValue.Year > 1900 && dateValue.Year < 4500)
                    {
                        return dateValue.ToString("yyyy-MM-dd");
                    }
                }
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string CsvEscape(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            // If value contains comma, quote, or newline, wrap in quotes and escape internal quotes
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }

            return value;
        }

        public void Cleanup()
        {
            try
            {
                if (_namespace != null)
                {
                    Marshal.ReleaseComObject(_namespace);
                    _namespace = null;
                }

                if (_outlook != null)
                {
                    Marshal.ReleaseComObject(_outlook);
                    _outlook = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
}
