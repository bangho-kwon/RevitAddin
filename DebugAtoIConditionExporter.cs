using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System.Collections.Generic;
using System.Linq;

namespace ConnectorSizeExport.IO
{
    public static class DebugAtoIConditionExporter
    {
        public static void Export(string path, List<ConnectorExportRow> rows, List<SettingCondition> conditions)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AtoI_Condition");

            ws.Cell(1, 1).Value = "Index";
            ws.Cell(1, 2).Value = "ElementId";
            ws.Cell(1, 3).Value = "BMArea";
            ws.Cell(1, 4).Value = "BMUnit";
            ws.Cell(1, 5).Value = "BMZone";
            ws.Cell(1, 6).Value = "BMDiscipline";
            ws.Cell(1, 7).Value = "BMSubDiscipline";
            ws.Cell(1, 8).Value = "BMFluid";
            ws.Cell(1, 9).Value = "BMClass";
            ws.Cell(1, 10).Value = "BMScode";

            int rowIdx = 2;
            int index = 1;

            foreach (var row in rows)
            {
                if (!conditions.Any(cond => cond.IsMatch(row)))
                    continue;

                ws.Cell(rowIdx, 1).Value = index++;
                ws.Cell(rowIdx, 2).Value = row.ElementId;
                ws.Cell(rowIdx, 3).Value = row.BMArea;
                ws.Cell(rowIdx, 4).Value = row.BMUnit;
                ws.Cell(rowIdx, 5).Value = row.BMZone;
                ws.Cell(rowIdx, 6).Value = row.BMDiscipline;
                ws.Cell(rowIdx, 7).Value = row.BMSubDiscipline;
                ws.Cell(rowIdx, 8).Value = row.BMFluid;
                ws.Cell(rowIdx, 9).Value = row.BMClass;
                ws.Cell(rowIdx, 10).Value = row.BMScode;
                rowIdx++;
            }

            wb.SaveAs(path); // 🔄 반드시 .xlsx 확장자 사용
        }
    }
}
