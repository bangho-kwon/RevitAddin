// 파일명: ExportAsExcel.cs
using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System;
using System.Collections.Generic;

namespace ConnectorSizeExport.IO
{
    public static class ExportAsExcel
    {
        public static void ExportLargeDiameterCheck(string path, List<ConnectorExportRow> rows, List<SettingCondition> conditions)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("LargeDiameterCheck");

            ws.Cell(1, 1).Value = "Index";
            ws.Cell(1, 2).Value = "ElementId";
            ws.Cell(1, 3).Value = "LargeDiameter";
            ws.Cell(1, 4).Value = "Setting_Min";
            ws.Cell(1, 5).Value = "Setting_Max";

            int rowIdx = 2;
            int index = 1;

            foreach (var row in rows)
            {
                if (!double.TryParse(row.LargeDiameter, out var actual)) continue;

                foreach (var cond in conditions)
                {
                    if (double.TryParse(cond.LargeDiameterMin, out var min) &&
                        double.TryParse(cond.LargeDiameterMax, out var max))
                    {
                        if (actual >= min && actual <= max)
                        {
                            ws.Cell(rowIdx, 1).Value = index++;
                            ws.Cell(rowIdx, 2).Value = row.ElementId;
                            ws.Cell(rowIdx, 3).Value = actual;
                            ws.Cell(rowIdx, 4).Value = min;
                            ws.Cell(rowIdx, 5).Value = max;
                            rowIdx++;
                            break;
                        }
                    }
                }
            }

            wb.SaveAs(path);
        }
    }
}
