using System.Collections.Concurrent;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace NexonKorea.XlsxMerge
{
    class XlsxDiff3Core
    {
        public class SheetDiffResult
        {
            public MergeArgumentInfo MergeArgs = null!;
            public string WorksheetName = "";
            public List<DocOrigin> DocsContaining = new List<DocOrigin>();
            public List<DiffHunkInfo> HunkList = new List<DiffHunkInfo>();

            public class DiffHunkInfo
            {
                public Diff3HunkStatus hunkStatus;
                public Dictionary<DocOrigin, RowRange> rowRangeMap = new Dictionary<DocOrigin, RowRange>();
            }
        }

        public MergeArgumentInfo MergeArgs = null!;
        public List<SheetDiffResult> SheetCompareResultList = new List<SheetDiffResult>();
        public Dictionary<DocOrigin, ParsedXlsx> ParsedWorkbookMap = new Dictionary<DocOrigin, ParsedXlsx>();

        public void Run(MergeArgumentInfo mergeArgs)
        {
            var totalSw = Stopwatch.StartNew();
            MergeArgs = mergeArgs;

            // Parallel file loading - each XLWorkbook instance is independent
            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 1단계]", "문서를 읽고 있습니다...");
            var loadSw = Stopwatch.StartNew();

            var loadTasks = new List<Task<(DocOrigin origin, ParsedXlsx parsed)>>
            {
                Task.Run(() => (DocOrigin.Base, new ParsedXlsxGenerator().ParseXlsx(mergeArgs.BasePath!))),
                Task.Run(() => (DocOrigin.Mine, new ParsedXlsxGenerator().ParseXlsx(mergeArgs.MinePath!))),
            };
            if (mergeArgs.ComparisonMode == ComparisonMode.ThreeWay)
                loadTasks.Add(Task.Run(() => (DocOrigin.Theirs, new ParsedXlsxGenerator().ParseXlsx(mergeArgs.TheirsPath!))));

            Task.WhenAll(loadTasks).GetAwaiter().GetResult();
            foreach (var t in loadTasks)
                ParsedWorkbookMap[t.Result.origin] = t.Result.parsed;

            PerfLog.Log($"Total file loading (parallel): {loadSw.ElapsedMilliseconds}ms");

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 3단계]", "엑셀 문서 비교 중..");

            var xlsxList = new List<ParsedXlsx>
            {
                ParsedWorkbookMap[DocOrigin.Base],
                ParsedWorkbookMap[DocOrigin.Mine],
                mergeArgs.ComparisonMode == ComparisonMode.ThreeWay ? ParsedWorkbookMap[DocOrigin.Theirs] : ParsedWorkbookMap[DocOrigin.Base]
            };

            // 비교 대상 워크시트 목록을 추출
            var allSheetNameSet = new HashSet<string>();
            var allSheetNameList = new List<string>();
            foreach (var eachXlsx in xlsxList)
                foreach (string sheetName in eachXlsx.Worksheets.Select(r => r.Name))
                    if (allSheetNameSet.Add(sheetName))
                        allSheetNameList.Add(sheetName);

            // Parallel worksheet diff
            var diffSw = Stopwatch.StartNew();
            var results = new ConcurrentBag<(int index, SheetDiffResult result)>();
            Parallel.ForEach(allSheetNameList.Select((name, idx) => (name, idx)), item =>
            {
                var (worksheetName, idx) = item;
                var sheetSw = Stopwatch.StartNew();

                var newSheetResult = new SheetDiffResult();
                newSheetResult.WorksheetName = worksheetName;
                newSheetResult.MergeArgs = mergeArgs;

                string diff3ResultText;
                {
                    var linesSw = Stopwatch.StartNew();
                    var lines1 = GetWorksheetLines(xlsxList[0], worksheetName);
                    var lines2 = GetWorksheetLines(xlsxList[1], worksheetName);
                    var lines3 = GetWorksheetLines(xlsxList[2], worksheetName);
                    PerfLog.Log($"  GetWorksheetLines '{worksheetName}': {linesSw.ElapsedMilliseconds}ms");

                    if (lines1 != null) newSheetResult.DocsContaining.Add(DocOrigin.Base);
                    if (lines2 != null) newSheetResult.DocsContaining.Add(DocOrigin.Mine);
                    if (lines3 != null) newSheetResult.DocsContaining.Add(DocOrigin.Theirs);

                    var diff3Sw = Stopwatch.StartNew();
                    diff3ResultText = LaunchExternalDiff3Process(lines1, lines2, lines3);
                    PerfLog.Log($"  Diff3 process '{worksheetName}': {diff3Sw.ElapsedMilliseconds}ms");
                }

                newSheetResult.HunkList = ParseDiff3Result(diff3ResultText);
                results.Add((idx, newSheetResult));
                PerfLog.Log($"  Sheet diff total '{worksheetName}': {sheetSw.ElapsedMilliseconds}ms");
            });
            SheetCompareResultList = results.OrderBy(r => r.index).Select(r => r.result).ToList();

            PerfLog.Log($"Total worksheet diff (parallel): {diffSw.ElapsedMilliseconds}ms");
            PerfLog.Log($"TOTAL Run() time: {totalSw.ElapsedMilliseconds}ms");
            PerfLog.Flush();
        }

        public Dictionary<DocOrigin, ParsedXlsx.Worksheet?> GetParsedWorksheetData(string worksheetName)
        {
            var result = new Dictionary<DocOrigin, ParsedXlsx.Worksheet?>();
            foreach (var docOrigin in new[] { DocOrigin.Base, DocOrigin.Mine, DocOrigin.Theirs })
            {
                result[docOrigin] = null;
                if (ParsedWorkbookMap.ContainsKey(docOrigin))
                    result[docOrigin] = ParsedWorkbookMap[docOrigin].Worksheets.FirstOrDefault(r => r.Name == worksheetName);
            }
            return result;
        }

        private static List<string>? GetWorksheetLines(ParsedXlsx xlsxFile, string worksheetName)
        {
            var targetWorksheet = xlsxFile.Worksheets.Find(r => r.Name == worksheetName);
            if (targetWorksheet == null)
                return null;

            var result = new List<string>(targetWorksheet.Cells.Count);
            foreach (var eachRow in targetWorksheet.Cells)
            {
                var columnList = eachRow.Select(r => r.ContentsForDiff3).ToList();
                while (columnList.Count > 0 && columnList[^1] == "")
                    columnList.RemoveAt(columnList.Count - 1);
                // Use SOH delimiter instead of JSON serialization - much faster,
                // still produces unique strings per row since SOH is not in Excel data
                result.Add(string.Join("\u0001", columnList));
            }
            return result;
        }

        private static readonly Regex RegexLineInfo = new("^([123]):([0-9,]+)([ac])$", RegexOptions.Compiled);

        private static readonly Dictionary<string, Diff3HunkStatus> HunkStatusMap = new()
        {
            { "====", Diff3HunkStatus.Conflict },
            { "====1", Diff3HunkStatus.BaseDiffers },
            { "====2", Diff3HunkStatus.MineDiffers },
            { "====3", Diff3HunkStatus.TheirsDiffers }
        };

        private static readonly Dictionary<string, DocOrigin> FileOrderMap = new()
        {
            { "1", DocOrigin.Base },
            { "2", DocOrigin.Mine },
            { "3", DocOrigin.Theirs },
        };

        private static List<SheetDiffResult.DiffHunkInfo> ParseDiff3Result(string diff3ResultText)
        {
            var hunkInfoList = new List<SheetDiffResult.DiffHunkInfo>();
            SheetDiffResult.DiffHunkInfo? curHunk = null;

            using var sr = new StringReader(diff3ResultText);
            string? curLine;
            while ((curLine = sr.ReadLine()) != null)
            {
                if (curLine.StartsWith("===="))
                {
                    curHunk = new SheetDiffResult.DiffHunkInfo();
                    hunkInfoList.Add(curHunk);
                    curHunk.hunkStatus = HunkStatusMap[curLine.Trim()];
                }

                var m = RegexLineInfo.Match(curLine);
                if (!m.Success)
                    continue;

                string fileIndex = m.Groups[1].Value;
                string[] rangeToken = m.Groups[2].Value.Split(',');
                string command = m.Groups[3].Value;

                var rowRangeValue = new RowRange();
                rowRangeValue.RowNumber = int.Parse(rangeToken[0]);
                rowRangeValue.RowCount = 0;

                if (command == "c")
                {
                    rowRangeValue.RowCount = 1;
                    if (rangeToken.Length > 1)
                        rowRangeValue.RowCount = int.Parse(rangeToken[1]) - rowRangeValue.RowNumber + 1;
                }
                if (command == "a")
                {
                    rowRangeValue.RowNumber = int.Parse(rangeToken[0]) + 1;
                    rowRangeValue.RowCount = 0;
                }

                DocOrigin rowRangeDoc = FileOrderMap[fileIndex];
                curHunk!.rowRangeMap[rowRangeDoc] = rowRangeValue;
            }

            return hunkInfoList;
        }

        private static string? FindDiff3Executable()
        {
            // 1. Same folder as XlsxMerge.exe
            string? exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly()?.Location);
            if (!string.IsNullOrEmpty(exeDir))
            {
                string candidate = Path.Combine(exeDir, "diff3.exe");
                if (File.Exists(candidate))
                    return candidate;
            }

            // 2. Current working directory
            string cwdCandidate = Path.GetFullPath("diff3.exe");
            if (File.Exists(cwdCandidate))
                return cwdCandidate;

            // 3. Git for Windows bundled diff3
            try
            {
                var gitPsi = new ProcessStartInfo("where.exe", "git")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using var gitProc = Process.Start(gitPsi);
                if (gitProc != null)
                {
                    string gitPath = gitProc.StandardOutput.ReadLine() ?? "";
                    gitProc.WaitForExit();
                    if (!string.IsNullOrEmpty(gitPath))
                    {
                        // git.exe is typically at <GitRoot>/cmd/git.exe or <GitRoot>/bin/git.exe
                        string? gitRoot = Path.GetDirectoryName(Path.GetDirectoryName(gitPath));
                        if (!string.IsNullOrEmpty(gitRoot))
                        {
                            string gitDiff3 = Path.Combine(gitRoot, "usr", "bin", "diff3.exe");
                            if (File.Exists(gitDiff3))
                                return gitDiff3;
                        }
                    }
                }
            }
            catch
            {
                // Ignore errors in git discovery
            }

            // 4. Common known paths
            string[] knownPaths = new[]
            {
                @"C:\Program Files\Git\usr\bin\diff3.exe",
                @"C:\Program Files (x86)\Git\usr\bin\diff3.exe"
            };
            foreach (string knownPath in knownPaths)
            {
                if (File.Exists(knownPath))
                    return knownPath;
            }

            return null;
        }

        private static string LaunchExternalDiff3Process(List<string>? lines1, List<string>? lines2, List<string>? lines3)
        {
            string tmp1 = Path.GetTempFileName();
            string tmp2 = Path.GetTempFileName();
            string tmp3 = Path.GetTempFileName();

            try
            {
                if (lines1 != null) File.WriteAllLines(tmp1, lines1);
                if (lines2 != null) File.WriteAllLines(tmp2, lines2);
                if (lines3 != null) File.WriteAllLines(tmp3, lines3);

                string? diff3Path = FindDiff3Executable();
                if (diff3Path == null)
                    throw new FileNotFoundException(
                        "diff3.exe not found. Please install Git for Windows or place diff3.exe next to XlsxMerge.exe.");

                var psi = new ProcessStartInfo()
                {
                    FileName = diff3Path,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    Arguments = $"\"{tmp1}\" \"{tmp2}\" \"{tmp3}\""
                };
                psi.WorkingDirectory = Path.GetDirectoryName(psi.FileName);

                using var p = Process.Start(psi);
                if (p == null)
                    throw new InvalidOperationException($"Failed to start diff3 process: {diff3Path}");
                string diff3Result = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                return diff3Result;
            }
            finally
            {
                File.Delete(tmp1);
                File.Delete(tmp2);
                File.Delete(tmp3);
            }
        }
    }
}
