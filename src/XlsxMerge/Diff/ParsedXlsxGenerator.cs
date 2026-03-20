using ClosedXML.Excel;

namespace NexonKorea.XlsxMerge
{
    class ParsedXlsxGenerator : IDisposable
    {
        public ParsedXlsx ParseXlsx(string xlsxFilePath)
        {
            var result = new ParsedXlsx();

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교 [3단계 중 2단계]", "문서를 읽고 있습니다.", Path.GetFileName(xlsxFilePath), xlsxFilePath);

            using var workbook = new XLWorkbook(xlsxFilePath);

            foreach (var worksheet in workbook.Worksheets)
            {
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
                        parsedWorksheet.ColumnWidthList.Add(worksheet.Column(c).Width * 7.5); // approximate pixel conversion

                    // Cell data - read all at once for performance
                    var rangeUsed = worksheet.Range(1, 1, lastRowNumber, lastColNumber);

                    for (int i = 1; i <= lastRowNumber; i++)
                    {
                        var currentRow = new List<ParsedXlsx.CellContents>(lastColNumber);
                        parsedWorksheet.Cells.Add(currentRow);
                        for (int j = 1; j <= lastColNumber; j++)
                        {
                            var cell = rangeUsed.Cell(i, j);
                            string value2 = cell.CachedValue.ToString();
                            string formula = cell.HasFormula ? "=" + cell.FormulaR1C1 : "";
                            currentRow.Add(new ParsedXlsx.CellContents(value2, formula));
                        }
                    }
                }
            }

            return result;
        }

        public void Dispose()
        {
            // ClosedXML doesn't need explicit cleanup like COM Interop
        }
    }
}
