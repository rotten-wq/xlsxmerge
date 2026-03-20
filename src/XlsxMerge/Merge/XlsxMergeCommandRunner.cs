using ClosedXML.Excel;

namespace NexonKorea.XlsxMerge
{
    class XlsxMergeCommandRunner : IDisposable
    {
        private Dictionary<DocOrigin, string> xlsxPathMap = new();
        private Dictionary<string, XLWorkbook> workbooksMap = new();

        public void Run(XlsxMergeCommand cmd, string baseFilePath, string mineFilePath, string? theirsFilePath, string mergedFilePath)
        {
            workbooksMap.Clear();
            xlsxPathMap.Clear();

            xlsxPathMap[DocOrigin.Base] = baseFilePath;
            xlsxPathMap[DocOrigin.Mine] = mineFilePath;
            if (!string.IsNullOrEmpty(theirsFilePath))
                xlsxPathMap[DocOrigin.Theirs] = theirsFilePath;

            // 엑셀은 같은 파일명을 가진 워크시트를 동시에 열지 못하므로,
            // 임시 폴더에 엑셀 파일을 이름을 바꾸어 복사합니다.
            string workingFolderPath = Path.Combine(Path.GetTempPath(), "xlsxmerge_" + Path.GetRandomFileName() + DateTime.Now.Ticks);
            Directory.CreateDirectory(workingFolderPath);
            var docOriginList = xlsxPathMap.Keys.ToList();
            foreach (var docOrigin in docOriginList)
            {
                if (string.IsNullOrEmpty(xlsxPathMap[docOrigin]))
                    continue;

                var originalPath = xlsxPathMap[docOrigin];
                var originalExt = Path.GetExtension(originalPath);
                var newPath = Path.Combine(workingFolderPath, $"{docOriginList.IndexOf(docOrigin)}{originalExt}");
                File.Copy(originalPath, newPath);
                xlsxPathMap[docOrigin] = newPath;

                OpenWorkbook(newPath, originalPath);
            }

            int cmdItemIdx = 0;
            foreach (var cmdItem in cmd.CommandList)
            {
                ApplyCommand(cmdItem);
                var progress = (int)(cmdItemIdx * 100 / cmd.CommandList.Count);
                FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 머지 [4단계 중 3단계]", $"머지 내용 적용 중... {progress}%");
                cmdItemIdx++;
            }

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 머지 [4단계 중 4단계]", "결과 저장 중...");
            if (File.Exists(mergedFilePath))
                File.Delete(mergedFilePath);
            GetWorkbook(DocOrigin.Mine)!.SaveAs(Path.GetFullPath(mergedFilePath));

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 머지", "정리 중...");
            Dispose();

            foreach (var tempFile in Directory.GetFiles(workingFolderPath, "*", SearchOption.AllDirectories))
                File.SetAttributes(tempFile, FileAttributes.Normal);
            Directory.Delete(workingFolderPath, true);
            FakeBackgroundWorker.OnUpdateProgress(null);
        }

        private void OpenWorkbook(string xlsxFilePath, string originalPath)
        {
            if (string.IsNullOrEmpty(xlsxFilePath))
                return;

            if (workbooksMap.ContainsKey(xlsxFilePath))
                return;

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 머지 [4단계 중 2단계]", "기반 문서를 준비하고 있습니다.", Path.GetFileName(originalPath), originalPath);
            workbooksMap[xlsxFilePath] = new XLWorkbook(xlsxFilePath);
        }

        private XLWorkbook? GetWorkbook(DocOrigin docOrigin)
        {
            if (!xlsxPathMap.TryGetValue(docOrigin, out var path))
                return null;
            return workbooksMap.GetValueOrDefault(path);
        }

        public void ApplyCommand(XlsxMergeCommandItem cmdItem)
        {
            if (cmdItem.Cmd == "DELETE_SHEET")
            {
                var wb = GetWorkbook(cmdItem.destOrigin!.Value)!;
                wb.Worksheet(cmdItem.param1).Delete();
            }
            else if (cmdItem.Cmd == "COPY_SHEET")
            {
                var wbDest = GetWorkbook(cmdItem.destOrigin!.Value)!;
                var wbSrc = GetWorkbook(cmdItem.sourceOrigin!.Value)!;
                var wsSrc = wbSrc.Worksheet(cmdItem.param1);

                // ClosedXML에서 워크시트 복사
                wsSrc.CopyTo(wbDest, cmdItem.param1, wbDest.Worksheets.Count + 1);
            }
            else if (cmdItem.Cmd == "COPY_ROW")
            {
                var wbDest = GetWorkbook(cmdItem.destOrigin!.Value)!;
                var wbSrc = GetWorkbook(cmdItem.sourceOrigin!.Value)!;
                var wsDest = wbDest.Worksheet(cmdItem.param1);
                var wsSrc = wbSrc.Worksheet(cmdItem.param1);

                int insertAt = cmdItem.intParam1;
                int srcRowStart = cmdItem.intParam2;
                int rowCount = cmdItem.intParam3;

                // 빈 행 삽입
                for (int i = 0; i < rowCount; i++)
                    wsDest.Row(insertAt).InsertRowsAbove(1);

                // 소스에서 대상으로 행 복사
                int srcLastCol = wsSrc.LastColumnUsed()?.ColumnNumber() ?? 1;
                for (int i = 0; i < rowCount; i++)
                {
                    var srcRow = wsSrc.Row(srcRowStart + i);
                    var destRow = wsDest.Row(insertAt + i);

                    for (int col = 1; col <= srcLastCol; col++)
                    {
                        var srcCell = srcRow.Cell(col);
                        var destCell = destRow.Cell(col);

                        if (srcCell.HasFormula)
                        {
                            try
                            {
                                destCell.FormulaR1C1 = srcCell.FormulaR1C1;
                            }
                            catch
                            {
                                // 수식 복사 실패 시 값으로 복사
                                destCell.Value = srcCell.CachedValue;
                            }
                        }
                        else
                        {
                            destCell.Value = srcCell.CachedValue;
                        }

                        // 스타일 복사
                        destCell.Style = srcCell.Style;
                    }
                }
            }
            else if (cmdItem.Cmd == "INSERT_TEXT")
            {
                var wb = GetWorkbook(cmdItem.destOrigin!.Value)!;
                var ws = wb.Worksheet(cmdItem.param1);

                ws.Row(cmdItem.intParam1).InsertRowsAbove(1);
                var cell = ws.Cell(cmdItem.intParam1, 1);
                cell.Value = cmdItem.param2;
                cell.Style.Fill.BackgroundColor = XLColor.Yellow;
            }
            else if (cmdItem.Cmd == "DELETE_ROW")
            {
                var wb = GetWorkbook(cmdItem.destOrigin!.Value)!;
                var ws = wb.Worksheet(cmdItem.param1);

                ws.Rows(cmdItem.intParam1, cmdItem.intParam1 + cmdItem.intParam2 - 1).Delete();
            }
        }

        public void Dispose()
        {
            foreach (var wb in workbooksMap.Values)
                wb.Dispose();
            workbooksMap.Clear();
        }
    }
}
