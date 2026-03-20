using System.Collections.Concurrent;

namespace NexonKorea.XlsxMerge
{
    /// <summary>
    /// Thread-safe performance logging to %TEMP%\xlsxmerge_perf.log
    /// </summary>
    static class PerfLog
    {
        private static readonly ConcurrentQueue<string> _messages = new();
        private static readonly string _logPath = Path.Combine(Path.GetTempPath(), "xlsxmerge_perf.log");

        public static void Log(string message)
        {
            _messages.Enqueue($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
        }

        public static void Flush()
        {
            var lines = new List<string>();
            while (_messages.TryDequeue(out var msg))
                lines.Add(msg);
            try
            {
                File.WriteAllLines(_logPath, lines);
            }
            catch
            {
                // Ignore logging failures
            }
        }
    }
}
