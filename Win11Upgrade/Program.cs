using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading;
using System.Security.Principal;

namespace Win11Upgrade
{
    internal static class Program
    {
        // --- CONFIG ---
        private const string IsoUrl =
            "https://software-static.download.prss.microsoft.com/dbazure/888969d5-f34g-4e03-ac9d-1f9786c66749/26200.6584.250915-1905.25h2_ge_release_svc_refresh_CLIENT_CONSUMER_x64FRE_en-us.iso";

        private static readonly string IsoPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Win11Upgrade", "Win11.iso");

        private const string SetupArgs = "/product server /eula accept /auto upgrade /migratedrivers all /Compat IgnoreWarning /dynamicupdate disable /showoobe none";
        private const int DriveDetectTimeoutMs = 60000; // 60s
        private const int DriveDetectPollMs = 1000;     // 1s

        private static int Main()
        {
            try
            {
                Console.WriteLine("== Win11Upgrade (Justin Shafer) ==");

                EnsureAdmin();
                EnsureFolders();

                // 1) Download ISO (skip if already fully present and size matches)
                DownloadIso(IsoUrl, IsoPath);

                // 2) Mount ISO (PowerShell)
                Console.WriteLine("Mounting ISO...");
                RunPwsh("Mount-DiskImage -ImagePath '" + EscapePs(IsoPath) + "' -ErrorAction Stop");

                // 3) Detect drive letter
                Console.WriteLine("Detecting mounted drive letter...");
                var driveLetter = DetectMountedDriveLetter(IsoPath, DriveDetectTimeoutMs, DriveDetectPollMs);
                if (string.IsNullOrEmpty(driveLetter))
                    throw new InvalidOperationException("Could not detect mounted ISO drive letter.");

                Console.WriteLine("Mounted at: " + driveLetter + ":\\");
                var setupPath = driveLetter + @":\setup.exe";
                if (!File.Exists(setupPath))
                    throw new FileNotFoundException("setup.exe not found at " + setupPath + ". Is this the correct ISO?");

                // 4) Run setup.exe /product server
                Console.WriteLine("Launching: \"" + setupPath + "\" " + SetupArgs);
                var ec = StartAndWait(setupPath, SetupArgs, true);
                Console.WriteLine("Setup exited with code " + ec + ".");

                // 5) Dismount ISO
                Console.WriteLine("Dismounting ISO...");
                RunPwsh("Dismount-DiskImage -ImagePath '" + EscapePs(IsoPath) + "' -ErrorAction SilentlyContinue");

                Console.WriteLine("All done.");
                return ec;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ERROR: " + ex.Message);
                return 1;
            }
        }

        private static void EnsureFolders()
        {
            var dir = Path.GetDirectoryName(IsoPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        private static void EnsureAdmin()
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    Console.WriteLine("Warning: Not running as Administrator. Mounting may fail.");
                }
            }
            catch { /* ignore */ }
        }

        private static void DownloadIso(string url, string path)
        {
            Console.WriteLine("Downloading ISO to: " + path);

            long remoteLength = -1;
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "HEAD";
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    remoteLength = resp.ContentLength;
                }
            }
            catch
            {
                Console.WriteLine("Could not determine remote file size; proceeding to download.");
            }

            if (File.Exists(path) && remoteLength > 0)
            {
                var localLen = new FileInfo(path).Length;
                if (localLen == remoteLength)
                {
                    Console.WriteLine("ISO already present and size matches remote. Skipping download.");
                    return;
                }
            }

            var tmp = path + ".part";
            if (File.Exists(tmp)) File.Delete(tmp);

            using (var client = new WebClient())
            {
                client.DownloadProgressChanged += (s, e) =>
                {
                    Console.Write("\r{0,3}% of {1}", e.ProgressPercentage, FormatBytes(e.TotalBytesToReceive));
                };
                client.DownloadFileCompleted += (s, e) => Console.WriteLine();
                client.DownloadFile(new Uri(url), tmp);
            }

            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
            Console.WriteLine("Download complete.");
        }

        private static string DetectMountedDriveLetter(string isoPath, int timeoutMs, int pollMs)
        {
            var sw = Stopwatch.StartNew();

            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var dl = RunPwshCapture(
                    "$img = Get-DiskImage -ImagePath '" + EscapePs(isoPath) + "' -ErrorAction SilentlyContinue; " +
                    "if ($img) { $v = $img | Get-Volume -ErrorAction SilentlyContinue; if ($v) { $v.DriveLetter } }"
                ).Trim();

                if (!string.IsNullOrEmpty(dl))
                    return dl.Substring(0, 1).ToUpperInvariant();

                // Fallback: scan CD-ROMs for a root setup.exe
                foreach (var drive in DriveInfo.GetDrives())
                {
                    try
                    {
                        if (drive.DriveType == DriveType.CDRom && drive.IsReady)
                        {
                            var candidate = Path.Combine(drive.Name, "setup.exe");
                            if (File.Exists(candidate))
                                return drive.Name.Substring(0, 1).ToUpperInvariant();
                        }
                    }
                    catch { }
                }

                Thread.Sleep(pollMs);
            }

            return string.Empty;
        }

        private static void RunPwsh(string oneLiner)
        {
            var ec = StartAndWait(PwshExe(),
                "-NoProfile -ExecutionPolicy Bypass -Command \"" + oneLiner + "\"",
                false);
            if (ec != 0)
                throw new InvalidOperationException("PowerShell command failed (exit " + ec + "). Command: " + oneLiner);
        }

        private static string RunPwshCapture(string oneLiner)
        {
            var psi = new ProcessStartInfo
            {
                FileName = PwshExe(),
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"" + oneLiner + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8
            };
            using (var p = Process.Start(psi))
            {
                var output = p.StandardOutput.ReadToEnd();
                var err = p.StandardError.ReadToEnd();
                p.WaitForExit();
                if (p.ExitCode != 0 && !string.IsNullOrEmpty(err))
                    Console.Error.WriteLine("PowerShell error (" + p.ExitCode + "): " + err);
                return output;
            }
        }

        private static string PwshExe()
        {
            var win = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var sys32 = Path.Combine(win, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            return File.Exists(sys32) ? sys32 : "powershell.exe";
        }

        private static int StartAndWait(string exe, string args, bool runAsAdmin)
        {
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = args,
                UseShellExecute = runAsAdmin,
                Verb = runAsAdmin ? "runas" : null,
                CreateNoWindow = false
            };
            using (var p = Process.Start(psi))
            {
                p.WaitForExit();
                return p.ExitCode;
            }
        }

        private static string FormatBytes(long bytes)
        {
            if (bytes < 0) return "unknown";
            string[] units = { "B", "KB", "MB", "GB", "TB" };
            double size = bytes;
            int unit = 0;
            while (size >= 1024 && unit < units.Length - 1)
            {
                size /= 1024;
                unit++;
            }
            return string.Format("{0:0.##} {1}", size, units[unit]);
        }

        private static string EscapePs(string s)
        {
            return s.Replace("'", "''");
        }
    }
}
