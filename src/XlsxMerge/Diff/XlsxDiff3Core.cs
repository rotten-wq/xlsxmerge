using System.Text;
using System.Text.Json;
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
            MergeArgs = mergeArgs;

            // ClosedXML으로 엑셀 파일 해석 (COM Interop 대비 수십 배 빠름)
            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 1단계]", "문서를 읽고 있습니다...");
            using (var parsedXlsxGenerator = new ParsedXlsxGenerator())
            {
                ParsedWorkbookMap[DocOrigin.Base] = parsedXlsxGenerator.ParseXlsx(mergeArgs.BasePath!);
                ParsedWorkbookMap[DocOrigin.Mine] = parsedXlsxGenerator.ParseXlsx(mergeArgs.MinePath!);
                if (mergeArgs.ComparisonMode == ComparisonMode.ThreeWay)
                    ParsedWorkbookMap[DocOrigin.Theirs] = parsedXlsxGenerator.ParseXlsx(mergeArgs.TheirsPath!);
            }

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

            // 각 워크시트를 비교
            SheetCompareResultList.Clear();
            foreach (var worksheetName in allSheetNameList)
            {
                var newSheetResult = new SheetDiffResult();
                SheetCompareResultList.Add(newSheetResult);
                newSheetResult.WorksheetName = worksheetName;
                newSheetResult.MergeArgs = mergeArgs;

                string diff3ResultText;
                {
                    var lines1 = GetWorksheetLines(xlsxList[0], worksheetName);
                    var lines2 = GetWorksheetLines(xlsxList[1], worksheetName);
                    var lines3 = GetWorksheetLines(xlsxList[2], worksheetName);

                    if (lines1 != null) newSheetResult.DocsContaining.Add(DocOrigin.Base);
                    if (lines2 != null) newSheetResult.DocsContaining.Add(DocOrigin.Mine);
                    if (lines3 != null) newSheetResult.DocsContaining.Add(DocOrigin.Theirs);

                    diff3ResultText = LaunchExternalDiff3Process(lines1, lines2, lines3);
                }

                newSheetResult.HunkList = ParseDiff3Result(diff3ResultText);
            }
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
                result.Add(JsonSerializer.Serialize(columnList));
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

                var psi = new ProcessStartInfo()
                {
                    FileName = Path.GetFullPath(@".\diff3.exe"),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    Arguments = $"\"{tmp1}\" \"{tmp2}\" \"{tmp3}\""
                };
                psi.WorkingDirectory = Path.GetDirectoryName(psi.FileName);

                using var p = Process.Start(psi)!;
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
