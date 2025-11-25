using System;
using System.IO;
using System.Runtime.InteropServices;
using Forms = System.Windows.Forms;
using SolidEdgeFileProperties;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        LogBanner("Solid Edge - Closed-file STATUS setter test");

        Log($"WindowsIdentity : {System.Security.Principal.WindowsIdentity.GetCurrent().Name}");
        Log($"Environment user: {Environment.UserDomainName}\\{Environment.UserName}");
        Log($"Machine        : {Environment.MachineName}");
        Log("------------------------------------------------------------");

        // 1) Let you pick a Solid Edge document
        var ofd = new Forms.OpenFileDialog
        {
            Title = "Select Solid Edge document",
            Filter = "Solid Edge (*.par;*.psm;*.asm;*.dft)|*.par;*.psm;*.asm;*.dft",
            Multiselect = false
        };

        if (ofd.ShowDialog() != Forms.DialogResult.OK)
        {
            Log("Canceled by user.");
            return;
        }

        var path = ofd.FileName;
        Log($"Selected file: {path}");

        // 2) Make sure Windows read-only attribute is cleared
        SetFileReadWrite(path);

        // 3) Try to read & write ExtendedSummaryInformation.Status
        SetStatusViaFileProperties(path);

        Log("Done. Press Enter to close.");
        Console.ReadLine();
    }

    // ------------------------------------------------------------
    // Helper: clear NTFS read-only bit (Windows-level read/write)
    // ------------------------------------------------------------
    private static void SetFileReadWrite(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Log($"[FS] Error: file not found at {filePath}");
            return;
        }

        try
        {
            FileInfo fileInfo = new FileInfo(filePath);
            if ((fileInfo.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                fileInfo.Attributes &= ~FileAttributes.ReadOnly;
                Log($"[FS] Cleared Windows Read-only attribute. File is now read/write.");
            }
            else
            {
                Log("[FS] File is already read/write (no Read-only attribute set).");
            }
        }
        catch (Exception ex)
        {
            Log($"[FS] Could not modify file attributes: {ex.GetType().Name}: {ex.Message}");
        }
    }

    // ------------------------------------------------------------
    // Core: write ExtendedSummaryInformation.Status on CLOSED file
    // ------------------------------------------------------------
    private static void SetStatusViaFileProperties(string path)
    {
        Log("");
        Log("---- Closed-file property check (SolidEdgeFileProperties) ----");

        PropertySets sets = null;

        try
        {
            sets = new PropertySets();
            // false = open read/write (not read-only)
            sets.Open(path, false);

            dynamic dSets = sets;

            // Try to get the ExtendedSummaryInformation property set
            dynamic ext = null;
            try
            {
                ext = dSets["ExtendedSummaryInformation"];
            }
            catch (Exception ex)
            {
                Warn($"ExtendedSummaryInformation set not found: {ex.Message}");
                return;
            }

            // Read current Status property (if it exists)
            dynamic statusProp = null;
            int? currentStatus = null;

            try
            {
                statusProp = ext["Status"];
                if (statusProp != null && statusProp.Value != null)
                    currentStatus = Convert.ToInt32(statusProp.Value);
            }
            catch
            {
                // no Status property yet
            }

            Log($"Current ExtendedSummaryInformation.Status = " +
                (currentStatus.HasValue ? currentStatus.ToString() : "<missing>"));

            // Ask user what value to write
            int? desired = AskStatusNumber(currentStatus);
            if (desired == null)
            {
                Log("[FileProps] No new status entered; leaving value unchanged.");
                return;
            }

            try
            {
                if (statusProp != null)
                {
                    Log($"[FileProps] Updating existing Status to {desired.Value}...");
                    statusProp.Value = desired.Value;
                }
                else
                {
                    Log($"[FileProps] Creating Status property with value {desired.Value}...");
                    // Typical COM signature: Add(string Name, object Value, int Id)
                    ext.Add("Status", desired.Value, 0);
                }

                sets.Save();
                Log("[FileProps] Save() completed. Status property written.");
            }
            catch (COMException cex)
            {
                LogCom("[FileProps] COM while writing Status", cex);
            }
            catch (Exception ex)
            {
                Log($"[FileProps] Error while writing Status: {ex.GetType().Name}: {ex.Message}");
            }
        }
        catch (COMException cex)
        {
            LogCom("[FileProps] COM while opening property sets", cex);
        }
        catch (Exception ex)
        {
            Log($"[FileProps] ERROR: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            try { sets?.Close(); } catch { }
            SafeRelease(sets);
            Log("---- End closed-file property check ----");
        }
    }

    // ------------------------------------------------------------
    // Ask user for the numeric Status value (0..3, etc.)
    // ------------------------------------------------------------
    private static int? AskStatusNumber(int? current)
    {
        Console.WriteLine();
        Console.WriteLine("Enter numeric Status value to write via FileProperties:");
        Console.WriteLine("  0 = Available");
        Console.WriteLine("  1 = In Work");
        Console.WriteLine("  2 = In Review");
        Console.WriteLine("  3 = Released");
        Console.WriteLine("  4 = Baselined");
        Console.WriteLine("  5 = Obsolete");
        Console.WriteLine("  (… other values follow whatever your vault uses)");
        Console.WriteLine();

        Console.Write($"Current Status property: {(current.HasValue ? current.ToString() : "<none>")}.  New value (blank = cancel): ");
        var s = (Console.ReadLine() ?? "").Trim();
        if (string.IsNullOrEmpty(s)) return null;

        if (int.TryParse(s, out int n))
            return n;

        Console.WriteLine("Input was not a valid integer. Aborting Status update.");
        return null;
    }

    // ------------------------------------------------------------
    // COM + logging helpers
    // ------------------------------------------------------------
    private static void SafeRelease(object obj)
    {
        try
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.FinalReleaseComObject(obj);
        }
        catch { }
    }

    private static void LogBanner(string title)
        => Console.WriteLine($"\n==================================================\n{title}\n==================================================");

    private static void Log(string msg)
        => Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {msg}");

    private static void Warn(string msg)
        => Console.WriteLine($"*** {msg}");

    private static void LogCom(string ctx, COMException ex)
        => Log($"{ctx}: COM 0x{ex.HResult:X8}: {ex.Message}");
}
