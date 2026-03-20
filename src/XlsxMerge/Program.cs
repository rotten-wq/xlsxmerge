namespace NexonKorea.XlsxMerge
{
    internal static class Program
    {
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "xlsxmerge.log");

        private static void LogException(string context, Exception ex)
        {
            try
            {
                string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}: {ex}\n";
                File.AppendAllText(LogPath, message);
            }
            catch
            {
                // Ignore logging failures
            }
        }

        [STAThread]
        static int Main()
        {
            Application.ThreadException += (sender, e) =>
            {
                LogException("ThreadException", e.Exception);
                MessageBox.Show(
                    $"An error occurred:\n\n{e.Exception.Message}\n\nDetails logged to: {LogPath}",
                    "XlsxMerge Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                if (e.ExceptionObject is Exception ex)
                {
                    LogException("UnhandledException", ex);
                    MessageBox.Show(
                        $"A fatal error occurred:\n\n{ex.Message}\n\nDetails logged to: {LogPath}",
                        "XlsxMerge Fatal Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            };

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            ApplicationConfiguration.Initialize();

            var args = Environment.GetCommandLineArgs();

            MergeArgumentInfo? argumentInfo = null;
            if (args.Length > 1)
            {
                argumentInfo = new MergeArgumentInfo(args);
                if (argumentInfo.ComparisonMode == ComparisonMode.Unknown)
                {
                    argumentInfo = null;
                    MessageBox.Show("명령줄 인수에 잘못되거나 누락된 값이 있습니다.");
                }
            }

            // 폴더 변경은 args 해석 이후에 합니다.
            string? exeFolderPath = Path.GetDirectoryName(path: System.Reflection.Assembly.GetEntryAssembly()?.Location);
            if (String.IsNullOrEmpty(exeFolderPath) == false)
                Directory.SetCurrentDirectory(exeFolderPath);

            if (argumentInfo != null)
            {
                var formMainDiff = new FormMainDiff();
                formMainDiff.MergeArgs = argumentInfo;
                Application.Run(formMainDiff);
                if (formMainDiff.MergeSuccessful)
                    return 0;
            }
            else
            {
                Application.Run(new FormWelcome());
            }

            return 1;
        }
    }
}