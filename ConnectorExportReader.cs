using System.Collections.Generic;
using ClosedXML.Excel;
using ConnectorSizeExport.Models;

namespace ConnectorSizeExport.IO
{
    public static class ConnectorExportReader
    {
        public static List<ConnectorExportRow> ReadExcel(string filePath)
        {
            var rows = new List<ConnectorExportRow>();
            var workbook = new XLWorkbook(filePath);
            var ws = workbook.Worksheet("ConnectorExport");

            var headers = new Dictionary<int, string>();

            // ✅ 실제 헤더는 2행에 존재
            var headerRow = ws.Row(2);
            for (int col = 1; col <= ws.LastColumnUsed().ColumnNumber(); col++)
            {
                string header = headerRow.Cell(col).GetString().Trim();
                if (!string.IsNullOrEmpty(header))
                    headers[col] = header;
            }

            // ✅ 데이터는 3행부터 시작 (1행: 그룹명, 2행: 헤더)
            for (int row = 3; row <= ws.LastRowUsed().RowNumber(); row++)
            {
                var current = new ConnectorExportRow();

                foreach (var kv in headers)
                {
                    string value = ws.Cell(row, kv.Key).GetString().Trim();
                    current.Values[kv.Value] = value;

                    switch (kv.Value)
                    {
                        case "BMArea": current.BMArea = value; break;
                        case "BMUnit": current.BMUnit = value; break;
                        case "BMZone": current.BMZone = value; break;
                        case "BMDiscipline": current.BMDiscipline = value; break;
                        case "BMSubDiscipline": current.BMSubDiscipline = value; break;
                        case "SystemType": current.SystemType = value; break;
                        case "BMFluid": current.BMFluid = value; break;
                        case "BMClass": current.BMClass = value; break;
                        case "BMScode": current.BMScode = value; break;
                        case "LargeDiameter": current.LargeDiameter = value; break;
                        case "SmallDiameter": current.SmallDiameter = value; break;
                        case "ElementId": current.ElementId = value; break;
                    }
                }

                rows.Add(current);
            }

            return rows;
        }
    }
}
