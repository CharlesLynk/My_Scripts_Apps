using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        const string ViewExe = @"C:\Program Files (x86)\ArchestrA\Framework\Bin\ViewAppFramework\View.exe";
        const string ViewAppsDir = @"C:\Program Files (x86)\ArchestrA\Framework\Bin\ViewApps";

        if (!File.Exists(ViewExe))
        {
            MessageBox.Show(
                "View.exe not found:\n" + ViewExe,
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        if (!Directory.Exists(ViewAppsDir))
        {
            MessageBox.Show(
                "ViewApps folder not found:\n" + ViewAppsDir,
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        DirectoryInfo latestViewApp = new DirectoryInfo(ViewAppsDir)
            .GetDirectories()
            .OrderByDescending(d => d.LastWriteTime)
            .FirstOrDefault();

        if (latestViewApp == null)
        {
            MessageBox.Show(
                "No ViewApp folders found under:\n" + ViewAppsDir,
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        string arg = "ViewAppPath:\"" + latestViewApp.FullName + "\"";

        var psi = new ProcessStartInfo
        {
            FileName = ViewExe,
            Arguments = arg,
            WorkingDirectory = Path.GetDirectoryName(ViewExe) ?? string.Empty,
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Maximized
        };

        Process p = Process.Start(psi);
        if (p == null)
        {
            MessageBox.Show(
                "Failed to launch View.exe.",
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        using (p)
        {
            p.WaitForExit();
            return p.ExitCode;
        }
    }
}