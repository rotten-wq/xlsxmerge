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

            // Workaround: Fork passes working-dir copy as $REMOTE for binary files,
            // so Base and Mine end up identical. Detect this and extract HEAD from git.
            string? gitBaseTemp = null;
            if (mergeArgs.ComparisonMode == ComparisonMode.TwoWay)
                gitBaseTemp = TryResolveGitBase(mergeArgs);

            // Parallel file loading - each XLWorkbook instance is independent
            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 1단계]", "문서를 읽고 있습니다...");
            var loadSw = Stopwatch.StartNew();

            string basePath = gitBaseTemp ?? mergeArgs.BasePath!;
            var loadTasks = new List<Task<(DocOrigin origin, ParsedXlsx parsed)>>
            {
                Task.Run(() => (DocOrigin.Base, new ParsedXlsxGenerator().ParseXlsx(basePath))),
                Task.Run(() => (DocOrigin.Mine, new ParsedXlsxGenerator().ParseXlsx(mergeArgs.MinePath!))),
            };
            if (mergeArgs.ComparisonMode == ComparisonMode.ThreeWay)
                loadTasks.Add(Task.Run(() => (DocOrigin.Theirs, new ParsedXlsxGenerator().ParseXlsx(mergeArgs.TheirsPath!))));

            Task.WhenAll(loadTasks).GetAwaiter().GetResult();
            foreach (var t in loadTasks)
                ParsedWorkbookMap[t.Result.origin] = t.Result.parsed;

            // Clean up temp file
            if (gitBaseTemp != null)
            {
                try { File.Delete(gitBaseTemp); } catch { }
            }

            PerfLog.Log($"Total file loading (parallel): {loadSw.ElapsedMilliseconds}ms");

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 3단계]", "엑셀 문서 비교 중..");

            var xlsxList = new List<ParsedXlsx>
            {
                ParsedWorkbookMap[DocOrigin.Base],
                ParsedWorkbookMap[DocOrigin.Mine],
                mergeArgs.ComparisonMode == ComparisonMode.ThreeWay ? ParsedWorkbookMap[DocOrigin.Theirs] : ParsedWorkbookMap[DocOrigin.Mine]
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
                    PerfLog.Log($"  GetWorksheetLines '{worksheetName}': {linesSw.ElapsedMilliseconds}ms (lines1={lines1?.Count}, lines2={lines2?.Count}, lines3={lines3?.Count})");

                    if (lines1 != null) newSheetResult.DocsContaining.Add(DocOrigin.Base);
                    if (lines2 != null) newSheetResult.DocsContaining.Add(DocOrigin.Mine);
                    if (lines3 != null && mergeArgs.ComparisonMode == ComparisonMode.ThreeWay)
                        newSheetResult.DocsContaining.Add(DocOrigin.Theirs);

                    var diff3Sw = Stopwatch.StartNew();
                    diff3ResultText = LaunchExternalDiff3Process(lines1, lines2, lines3, mergeArgs.ComparisonMode);
                    PerfLog.Log($"  Diff3 process '{worksheetName}': {diff3Sw.ElapsedMilliseconds}ms");
                }

                newSheetResult.HunkList = ParseDiff3Result(diff3ResultText);
                PerfLog.Log($"  Diff3 output '{worksheetName}' ({diff3ResultText.Length} chars): hunks={newSheetResult.HunkList.Count}");
                if (newSheetResult.HunkList.Count > 0)
                    PerfLog.Log($"    First hunk status: {newSheetResult.HunkList[0].hunkStatus}");
                else if (diff3ResultText.Length > 0)
                    PerfLog.Log($"    Raw diff3 (first 200 chars): {diff3ResultText[..Math.Min(200, diff3ResultText.Length)]}");
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

        private static string LaunchExternalDiff3Process(List<string>? lines1, List<string>? lines2, List<string>? lines3, ComparisonMode mode = ComparisonMode.ThreeWay)
        {
            // For 2-way mode use built-in diff to avoid external process failures
            // in environments like Fork git client where diff3.exe returns empty output.
            if (mode == ComparisonMode.TwoWay)
            {
                PerfLog.Log($"  Using built-in 2-way diff (lines1={lines1?.Count}, lines2={lines2?.Count})");
                return BuiltInTwoWayDiff(lines1, lines2);
            }

            // 3-way mode: use external diff3.exe
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
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    Arguments = $"\"{tmp1}\" \"{tmp2}\" \"{tmp3}\""
                };
                psi.WorkingDirectory = Path.GetDirectoryName(psi.FileName);

                using var p = Process.Start(psi);
                if (p == null)
                    throw new InvalidOperationException($"Failed to start diff3 process: {diff3Path}");
                string diff3Result = p.StandardOutput.ReadToEnd();
                string diff3Stderr = p.StandardError.ReadToEnd();
                p.WaitForExit();
                PerfLog.Log($"  diff3 exit={p.ExitCode}, stdout={diff3Result.Length}ch, stderr={diff3Stderr.Length}ch, path={diff3Path}");
                if (diff3Stderr.Length > 0)
                    PerfLog.Log($"  diff3 STDERR: {diff3Stderr[..Math.Min(500, diff3Stderr.Length)]}");
                return diff3Result;
            }
            finally
            {
                File.Delete(tmp1);
                File.Delete(tmp2);
                File.Delete(tmp3);
            }
        }

        /// <summary>
        /// Built-in 2-way diff using Myers algorithm.
        /// Compares lines1 (base) vs lines2 (mine) and generates diff3 hunk format with ====1 markers.
        /// </summary>
        private static string BuiltInTwoWayDiff(List<string>? lines1, List<string>? lines2)
        {
            var a = lines1 ?? new List<string>();
            var b = lines2 ?? new List<string>();

            var edits = MyersDiff(a, b);

            var sb = new StringBuilder();
            int i = 0; // index into edits
            while (i < edits.Count)
            {
                // Skip equal edits
                if (edits[i].op == DiffOp.Equal)
                {
                    i++;
                    continue;
                }

                // Collect a contiguous block of non-equal edits
                int blockStart = i;
                while (i < edits.Count && edits[i].op != DiffOp.Equal)
                    i++;
                int blockEnd = i; // exclusive

                // Determine ranges in file1 (base/lines1) and file2 (mine/lines2)
                int file1Start = -1, file1End = -1;
                int file2Start = -1, file2End = -1;

                foreach (var e in edits.Skip(blockStart).Take(blockEnd - blockStart))
                {
                    if (e.op == DiffOp.Delete || e.op == DiffOp.Equal)
                    {
                        if (file1Start < 0) file1Start = e.aLine;
                        file1End = e.aLine;
                    }
                    if (e.op == DiffOp.Insert || e.op == DiffOp.Equal)
                    {
                        if (file2Start < 0) file2Start = e.bLine;
                        file2End = e.bLine;
                    }
                }

                // Build diff3 hunk: ====1 means file1 (base) differs
                sb.AppendLine("====1");

                // File 1 range (1-based)
                if (file1Start < 0)
                {
                    // Pure insertion: find position in file1 just before this block
                    // The anchor is the aLine of the edit before blockStart (or 0 if none)
                    int anchor1 = blockStart > 0 ? edits[blockStart - 1].aLine : 0;
                    sb.AppendLine($"1:{anchor1}a");
                }
                else
                {
                    int r1s = file1Start + 1; // convert to 1-based
                    int r1e = file1End + 1;
                    sb.AppendLine(r1s == r1e ? $"1:{r1s}c" : $"1:{r1s},{r1e}c");
                }

                // File 2 range (1-based)
                if (file2Start < 0)
                {
                    int anchor2 = blockStart > 0 ? edits[blockStart - 1].bLine : 0;
                    sb.AppendLine($"2:{anchor2}a");
                }
                else
                {
                    int r2s = file2Start + 1;
                    int r2e = file2End + 1;
                    sb.AppendLine(r2s == r2e ? $"2:{r2s}c" : $"2:{r2s},{r2e}c");
                }

                // File 3 = same as file 2 for 2-way mode
                if (file2Start < 0)
                {
                    int anchor3 = blockStart > 0 ? edits[blockStart - 1].bLine : 0;
                    sb.AppendLine($"3:{anchor3}a");
                }
                else
                {
                    int r3s = file2Start + 1;
                    int r3e = file2End + 1;
                    sb.AppendLine(r3s == r3e ? $"3:{r3s}c" : $"3:{r3s},{r3e}c");
                }
            }

            return sb.ToString();
        }

        private enum DiffOp { Equal, Insert, Delete }

        private readonly struct DiffEdit
        {
            public readonly DiffOp op;
            public readonly int aLine; // 0-based index in 'a' (only valid for Delete/Equal)
            public readonly int bLine; // 0-based index in 'b' (only valid for Insert/Equal)
            public DiffEdit(DiffOp op, int aLine, int bLine) { this.op = op; this.aLine = aLine; this.bLine = bLine; }
        }

        /// <summary>
        /// Myers diff algorithm. Returns a list of edit operations describing
        /// how to transform sequence 'a' into sequence 'b'.
        /// Based on the standard Myers 1986 algorithm with forward pass + backtrack.
        /// </summary>
        private static List<DiffEdit> MyersDiff(List<string> a, List<string> b)
        {
            int n = a.Count, m = b.Count;

            if (n == 0 && m == 0)
                return new List<DiffEdit>();

            if (n == 0)
            {
                var inserts = new List<DiffEdit>(m);
                for (int j = 0; j < m; j++)
                    inserts.Add(new DiffEdit(DiffOp.Insert, 0, j));
                return inserts;
            }

            if (m == 0)
            {
                var deletes = new List<DiffEdit>(n);
                for (int ii = 0; ii < n; ii++)
                    deletes.Add(new DiffEdit(DiffOp.Delete, ii, 0));
                return deletes;
            }

            int max = n + m;
            int offset = max; // offset so k can be negative index
            var v = new int[2 * max + 2];
            v[1 + offset] = 0;

            // trace[d] = copy of v after processing step d
            var trace = new List<int[]>(max + 1);

            for (int d = 0; d <= max; d++)
            {
                for (int k = -d; k <= d; k += 2)
                {
                    int kOff = k + offset;
                    int x;
                    if (k == -d || (k != d && v[kOff - 1] < v[kOff + 1]))
                        x = v[kOff + 1];       // move down (insert b[y])
                    else
                        x = v[kOff - 1] + 1;   // move right (delete a[x])

                    int y = x - k;
                    // Follow diagonal (equal lines)
                    while (x < n && y < m && a[x] == b[y])
                    {
                        x++;
                        y++;
                    }
                    v[kOff] = x;
                    if (x >= n && y >= m)
                    {
                        // Save this step's frontier then backtrack
                        var snap = new int[2 * max + 2];
                        Array.Copy(v, snap, v.Length);
                        trace.Add(snap);

                        return Backtrack(a, b, trace, d, offset);
                    }
                }

                // Save frontier after step d
                var snapshot = new int[2 * max + 2];
                Array.Copy(v, snapshot, v.Length);
                trace.Add(snapshot);
            }

            // Should never reach here for valid inputs
            return new List<DiffEdit>();
        }

        private static List<DiffEdit> Backtrack(List<string> a, List<string> b, List<int[]> trace, int dFinal, int offset)
        {
            var edits = new List<DiffEdit>();
            int cx = a.Count, cy = b.Count;

            for (int d = dFinal; d > 0; d--)
            {
                var vPrev = trace[d - 1]; // frontier after step d-1
                int k = cx - cy;
                int kOff = k + offset;

                // Determine which diagonal we came from in step d-1
                int prevK;
                if (k == -d || (k != d && vPrev[kOff - 1] < vPrev[kOff + 1]))
                    prevK = k + 1; // came from k+1: moved down (insert b[y])
                else
                    prevK = k - 1; // came from k-1: moved right (delete a[x])

                int prevX = vPrev[prevK + offset];
                int prevY = prevX - prevK;

                // After the single edit, we're at (endX, endY), then the snake brought us to (cx, cy).
                // endX = prevX + (prevK == k-1 ? 1 : 0)
                // endY = prevY + (prevK == k+1 ? 1 : 0)
                int endX = prevX + (prevK == k - 1 ? 1 : 0);
                int endY = prevY + (prevK == k + 1 ? 1 : 0);

                // Emit equals for the snake (cx,cy) -> (endX,endY), in reverse
                while (cx > endX && cy > endY)
                {
                    cx--;
                    cy--;
                    edits.Add(new DiffEdit(DiffOp.Equal, cx, cy));
                }

                // Emit the single edit
                if (prevK == k - 1)
                {
                    // Moved right: deleted a[prevX] (0-based)
                    cx--;
                    edits.Add(new DiffEdit(DiffOp.Delete, cx, cy));
                }
                else
                {
                    // Moved down: inserted b[prevY] (0-based)
                    cy--;
                    edits.Add(new DiffEdit(DiffOp.Insert, cx, cy));
                }

                cx = prevX;
                cy = prevY;
            }

            // Remaining equal prefix at d=0 diagonal
            while (cx > 0 && cy > 0)
            {
                cx--;
                cy--;
                edits.Add(new DiffEdit(DiffOp.Equal, cx, cy));
            }

            edits.Reverse();
            return edits;
        }
        /// <summary>
        /// Fork passes a working-dir copy as $REMOTE for binary files, making Base == Mine.
        /// Detect this by comparing file hashes. If identical, extract HEAD version from git.
        /// Returns a temp file path with the git HEAD version, or null if not needed.
        /// </summary>
        private static string? TryResolveGitBase(MergeArgumentInfo mergeArgs)
        {
            try
            {
                string basePath = mergeArgs.BasePath!;
                string minePath = mergeArgs.MinePath!;

                // Quick check: if files differ in size, they're genuinely different
                var baseInfo = new FileInfo(basePath);
                var mineInfo = new FileInfo(minePath);
                if (baseInfo.Length != mineInfo.Length)
                    return null;

                // Compare file content via hash
                using var md5 = System.Security.Cryptography.MD5.Create();
                byte[] baseHash, mineHash;
                using (var bs = File.OpenRead(basePath)) baseHash = md5.ComputeHash(bs);
                using (var ms = File.OpenRead(minePath)) mineHash = md5.ComputeHash(ms);
                if (!baseHash.SequenceEqual(mineHash))
                    return null;

                PerfLog.Log("Base and Mine are identical — attempting git HEAD extraction");

                // Find git repo for Mine path
                string? repoFile = FindGitRepoFile(minePath);
                if (repoFile == null)
                {
                    PerfLog.Log("  Not in a git repository");
                    return null;
                }

                // Get relative path within repo
                string repoRoot = Path.GetDirectoryName(repoFile)!;
                string relativePath = Path.GetRelativePath(repoRoot, mineInfo.FullName).Replace('\\', '/');

                // Extract HEAD version via git show
                string tempFile = Path.Combine(Path.GetTempPath(), $"xlsxmerge_gitbase_{Path.GetFileName(minePath)}");
                var psi = new ProcessStartInfo("git", $"show HEAD:\"{relativePath}\"")
                {
                    WorkingDirectory = repoRoot,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = null // binary output
                };

                using var proc = Process.Start(psi);
                if (proc == null)
                    return null;

                // Write binary stdout to temp file
                using (var outStream = File.Create(tempFile))
                {
                    proc.StandardOutput.BaseStream.CopyTo(outStream);
                }
                proc.WaitForExit();

                if (proc.ExitCode != 0)
                {
                    PerfLog.Log($"  git show failed (exit {proc.ExitCode}): {proc.StandardError.ReadToEnd()}");
                    try { File.Delete(tempFile); } catch { }
                    return null;
                }

                // Verify the extracted file differs from working dir
                var extractedInfo = new FileInfo(tempFile);
                if (extractedInfo.Length == mineInfo.Length)
                {
                    using var emd5 = System.Security.Cryptography.MD5.Create();
                    byte[] extractedHash;
                    using (var es = File.OpenRead(tempFile)) extractedHash = emd5.ComputeHash(es);
                    if (extractedHash.SequenceEqual(mineHash))
                    {
                        PerfLog.Log("  HEAD version is same as working dir — no changes to show");
                        try { File.Delete(tempFile); } catch { }
                        return null;
                    }
                }

                PerfLog.Log($"  Using git HEAD as base: {extractedInfo.Length} bytes (working: {mineInfo.Length} bytes)");
                return tempFile;
            }
            catch (Exception ex)
            {
                PerfLog.Log($"  TryResolveGitBase failed: {ex.Message}");
                return null;
            }
        }

        private static string? FindGitRepoFile(string filePath)
        {
            string? dir = Path.GetDirectoryName(Path.GetFullPath(filePath));
            while (dir != null)
            {
                string gitDir = Path.Combine(dir, ".git");
                if (Directory.Exists(gitDir) || File.Exists(gitDir))
                    return gitDir;
                dir = Path.GetDirectoryName(dir);
            }
            return null;
        }
    }
}
