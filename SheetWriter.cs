// ✅ SheetWriter.cs - Excel 시트 생성 및 저장
using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class SheetWriter
    {
        public static void WriteIntegrateSheet(string filePath, List<Dictionary<string, string>> rows)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Integrate");

            var headers = rows
                .SelectMany(r => r.Keys)
                .Distinct()
                .OrderBy(h => h)
                .ToList();

            for (int i = 0; i < headers.Count; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (int c = 0; c < headers.Count; c++)
                {
                    if (row.TryGetValue(headers[c], out string val))
                        ws.Cell(r + 2, c + 1).Value = val;
                }
            }

            ws.Columns().AdjustToContents();
            workbook.SaveAs(filePath);
        }
    }
}
