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
        string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        if (String.IsNullOrEmpty(programFiles))
            programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);

        string bin = Path.Combine(programFiles, @"ArchestrA\Framework\Bin");
        string viewExe = Environment.GetEnvironmentVariable("AVEVA_OMI_VIEW_EXE");
        if (String.IsNullOrWhiteSpace(viewExe))
            viewExe = Path.Combine(bin, @"ViewAppFramework\View.exe");
        string viewAppsDir = Environment.GetEnvironmentVariable("AVEVA_OMI_VIEWAPPS_DIR");
        if (String.IsNullOrWhiteSpace(viewAppsDir))
            viewAppsDir = Path.Combine(bin, "ViewApps");

        if (!File.Exists(viewExe))
        {
            MessageBox.Show(
                "View.exe not found:\n" + viewExe + "\nFor a custom install, set AVEVA_OMI_VIEW_EXE.",
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        if (!Directory.Exists(viewAppsDir))
        {
            MessageBox.Show(
                "ViewApps folder not found:\n" + viewAppsDir + "\nFor a custom install, set AVEVA_OMI_VIEWAPPS_DIR.",
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        DirectoryInfo latestViewApp = new DirectoryInfo(viewAppsDir)
            .GetDirectories()
            .OrderByDescending(d => d.LastWriteTime)
            .FirstOrDefault();

        if (latestViewApp == null)
        {
            MessageBox.Show(
                "No ViewApp folders found under:\n" + viewAppsDir,
                "OMI Launcher Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        string arg = "ViewAppPath:\"" + latestViewApp.FullName + "\"";

        var psi = new ProcessStartInfo
        {
            FileName = viewExe,
            Arguments = arg,
            WorkingDirectory = Path.GetDirectoryName(viewExe) ?? string.Empty,
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