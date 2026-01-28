using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ClearFileAssoc
{
    public static class FileAssociationHelper
    {
        [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        private const uint SHCNE_ASSOCCHANGED = 0x08000000;
        private const uint SHCNF_IDLIST = 0x0000;

        public static ClearResult ClearAssociation(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                return new ClearResult(false, "No extension provided");

            // Ensure extension starts with dot
            if (!extension.StartsWith("."))
                extension = "." + extension;

            extension = extension.ToLowerInvariant();

            try
            {
                // The user's file association choice is stored here:
                // HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ext\UserChoice
                string fileExtsPath = $@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{extension}";

                using (var fileExtsKey = Registry.CurrentUser.OpenSubKey(fileExtsPath, writable: true))
                {
                    if (fileExtsKey == null)
                    {
                        return new ClearResult(false, $"No user association found for {extension}");
                    }

                    // Check if UserChoice exists
                    using (var userChoiceKey = fileExtsKey.OpenSubKey("UserChoice"))
                    {
                        if (userChoiceKey == null)
                        {
                            return new ClearResult(false, $"No user association found for {extension}");
                        }
                    }

                    // Delete the UserChoice subkey
                    // This requires taking ownership first on some systems
                    try
                    {
                        fileExtsKey.DeleteSubKeyTree("UserChoice");
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Try alternative approach - use reg.exe with elevated privileges
                        return ClearAssociationViaCmd(extension);
                    }
                }

                // Notify the shell that file associations have changed
                SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);

                return new ClearResult(true, $"Cleared association for {extension}");
            }
            catch (Exception ex)
            {
                return new ClearResult(false, $"Failed to clear association: {ex.Message}");
            }
        }

        private static ClearResult ClearAssociationViaCmd(string extension)
        {
            try
            {
                string keyPath = $@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{extension}\UserChoice";

                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "reg.exe",
                    Arguments = $"delete \"{keyPath}\" /f",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = System.Diagnostics.Process.Start(psi))
                {
                    process.WaitForExit(5000);

                    if (process.ExitCode == 0)
                    {
                        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
                        return new ClearResult(true, $"Cleared association for {extension}");
                    }
                    else
                    {
                        string error = process.StandardError.ReadToEnd();
                        return new ClearResult(false, $"Failed: {error}");
                    }
                }
            }
            catch (Exception ex)
            {
                return new ClearResult(false, $"Failed to clear association: {ex.Message}");
            }
        }

        public static string GetCurrentAssociation(string extension)
        {
            if (!extension.StartsWith("."))
                extension = "." + extension;

            try
            {
                string fileExtsPath = $@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{extension}\UserChoice";

                using (var key = Registry.CurrentUser.OpenSubKey(fileExtsPath))
                {
                    if (key != null)
                    {
                        var progId = key.GetValue("ProgId") as string;
                        if (!string.IsNullOrEmpty(progId))
                        {
                            return progId;
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    public class ClearResult
    {
        public bool Success { get; }
        public string Message { get; }

        public ClearResult(bool success, string message)
        {
            Success = success;
            Message = message;
        }
    }
}
