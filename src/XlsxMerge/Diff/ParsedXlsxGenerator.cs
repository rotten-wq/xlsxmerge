using ClosedXML.Excel;
using System.Diagnostics;

namespace NexonKorea.XlsxMerge
{
    class ParsedXlsxGenerator : IDisposable
    {
        public ParsedXlsx ParseXlsx(string xlsxFilePath)
        {
            var result = new ParsedXlsx();

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 2단계]", "문서를 읽고 있습니다.", Path.GetFileName(xlsxFilePath), xlsxFilePath);

            var sw = Stopwatch.StartNew();
            using var workbook = new XLWorkbook(xlsxFilePath);
            var loadTime = sw.ElapsedMilliseconds;
            PerfLog.Log($"  XLWorkbook load '{Path.GetFileName(xlsxFilePath)}': {loadTime}ms");

            foreach (var worksheet in workbook.Worksheets)
            {
                var sheetSw = Stopwatch.StartNew();
                var parsedWorksheet = new ParsedXlsx.Worksheet();
                result.Worksheets.Add(parsedWorksheet);
                parsedWorksheet.Name = worksheet.Name;
                parsedWorksheet.Cells = new List<List<ParsedXlsx.CellContents>>();

                var lastRowNumber = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                var lastColNumber = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

                if (lastRowNumber != 0 && lastColNumber != 0)
                {
                    // Column widths
                    for (int c = 1; c <= lastColNumber; c++)
                        parsedWorksheet.ColumnWidthList.Add(worksheet.Column(c).Width * 7.5);

                    // Build sparse cell map using CellsUsed() - much faster than rangeUsed.Cell(i,j)
                    var cellMap = new Dictionary<(int row, int col), (string value, string formula)>();
                    foreach (var cell in worksheet.CellsUsed())
                    {
                        cellMap[(cell.Address.RowNumber, cell.Address.ColumnNumber)] =
                            (cell.CachedValue.ToString(), cell.HasFormula ? "=" + cell.FormulaR1C1 : "");
                    }

                    for (int i = 1; i <= lastRowNumber; i++)
                    {
                        var currentRow = new List<ParsedXlsx.CellContents>(lastColNumber);
                        parsedWorksheet.Cells.Add(currentRow);
                        for (int j = 1; j <= lastColNumber; j++)
                        {
                            if (cellMap.TryGetValue((i, j), out var cv))
                                currentRow.Add(new ParsedXlsx.CellContents(cv.value, cv.formula));
                            else
                                currentRow.Add(new ParsedXlsx.CellContents());
                        }
                    }
                }

                PerfLog.Log($"  ParseWorksheet '{worksheet.Name}' ({lastRowNumber}x{lastColNumber}): {sheetSw.ElapsedMilliseconds}ms");
            }

            return result;
        }

        public void Dispose()
        {
            // ClosedXML doesn't need explicit cleanup like COM Interop
        }
    }
}
