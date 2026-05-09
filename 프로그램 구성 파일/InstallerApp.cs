using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Windows.Forms;

namespace WindowBackRecorderInstaller
{
    public static class InstallerApp
    {
        private const string AppFolderName = "백그라운드 영상 녹화 프로그램";
        private const string RecordingsFolderName = "녹화 완료된 동영상";
        private const string PayloadResourceName = "Payload.Zip";

        [STAThread]
        public static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                bool quiet = HasQuietOption(args);
                string installerPath = Assembly.GetExecutingAssembly().Location;
                string installerDir = Path.GetDirectoryName(installerPath);
                string targetDir = Path.Combine(installerDir, AppFolderName);

                if (!quiet && Directory.Exists(targetDir))
                {
                    DialogResult result = MessageBox.Show(
                        "이미 설치 폴더가 있습니다.\n\n새 파일로 덮어쓸까요?\n녹화 완료된 동영상 폴더 안의 파일은 그대로 둡니다.",
                        "백그라운드 영상 녹화 프로그램 설치",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result != DialogResult.Yes)
                    {
                        return 0;
                    }
                }

                Directory.CreateDirectory(targetDir);
                ExtractPayload(targetDir);
                Directory.CreateDirectory(Path.Combine(targetDir, RecordingsFolderName));

                if (quiet)
                {
                    return 0;
                }

                DialogResult open = MessageBox.Show(
                    "설치가 완료되었습니다.\n\n설치된 폴더를 열까요?",
                    "백그라운드 영상 녹화 프로그램 설치",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (open == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(targetDir) { UseShellExecute = true });
                }

                return 0;
            }
            catch (Exception ex)
            {
                if (HasQuietOption(args))
                {
                    Console.Error.WriteLine(ex.Message);
                    return 1;
                }

                MessageBox.Show(
                    "설치 중 문제가 생겼습니다.\n\n" + ex.Message + "\n\n프로그램이 실행 중이라면 닫고 다시 실행해주세요.",
                    "백그라운드 영상 녹화 프로그램 설치",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 1;
            }
        }

        private static bool HasQuietOption(string[] args)
        {
            if (args == null) return false;

            foreach (string arg in args)
            {
                if (string.Equals(arg, "/quiet", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(arg, "/silent", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(arg, "--quiet", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(arg, "--silent", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static void ExtractPayload(string targetDir)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream payload = assembly.GetManifestResourceStream(PayloadResourceName))
            {
                if (payload == null)
                {
                    throw new InvalidOperationException("설치 파일 안의 프로그램 데이터를 찾을 수 없습니다.");
                }

                using (ZipArchive archive = new ZipArchive(payload, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        string normalizedName = entry.FullName.Replace('\\', '/');
                        if (string.IsNullOrWhiteSpace(normalizedName))
                        {
                            continue;
                        }

                        string destinationPath = SafeCombine(targetDir, normalizedName);
                        bool isDirectory = normalizedName.EndsWith("/", StringComparison.Ordinal);
                        bool isRecordingsEntry = normalizedName.StartsWith(RecordingsFolderName + "/", StringComparison.OrdinalIgnoreCase);

                        if (isDirectory)
                        {
                            Directory.CreateDirectory(destinationPath);
                            continue;
                        }

                        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                        if (isRecordingsEntry && File.Exists(destinationPath))
                        {
                            continue;
                        }

                        using (Stream input = entry.Open())
                        using (FileStream output = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            input.CopyTo(output);
                        }
                    }
                }
            }
        }

        private static string SafeCombine(string root, string relativePath)
        {
            string destinationPath = Path.GetFullPath(Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar)));
            string rootPath = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;

            if (!destinationPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("잘못된 설치 경로가 포함되어 있습니다.");
            }

            return destinationPath;
        }
    }
}
