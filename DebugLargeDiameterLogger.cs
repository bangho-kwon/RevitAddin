using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ConnectorSizeExport.Helpers
{
    public static class DebugLargeDiameterLogger
    {
        public static void ExportLargeDiameterDebug(string filePath, List<ConnectorExportRow> rows, List<SettingCondition> conditions)
        {
            // 상위 5개 객체만 확인
            var top5 = rows.Count > 5 ? rows.GetRange(0, 5) : rows;

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("LargeDiameter_Debug");

                // 헤더 작성
                ws.Cell(1, 1).Value = "Index";
                ws.Cell(1, 2).Value = "ElementId";
                ws.Cell(1, 3).Value = "LargeDiameter";
                ws.Cell(1, 4).Value = "Setting_Min";
                ws.Cell(1, 5).Value = "Setting_Max";

                int rowIdx = 2;
                foreach (var row in top5)
                {
                    foreach (var cond in conditions)
                    {
                        ws.Cell(rowIdx, 1).Value = rowIdx - 1;
                        ws.Cell(rowIdx, 2).Value = row.ElementId;
                        ws.Cell(rowIdx, 3).Value = row.LargeDiameter;
                        ws.Cell(rowIdx, 4).Value = cond.LargeDiameterMin;
                        ws.Cell(rowIdx, 5).Value = cond.LargeDiameterMax;
                        rowIdx++;
                    }
                }

                string debugPath = Path.Combine(Path.GetDirectoryName(filePath), "LargeDiameter_Debug.xlsx");
                workbook.SaveAs(debugPath);
            }
        }
    }
}
