using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ConnectorSizeExport.Helpers
{
    public static class IntegrateExcelExporter
    {
        public static void Export(List<Dictionary<string, string>> filteredRows)
        {
            // 현재 시각 기반 파일 이름 설정
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string filePath = Path.Combine("C:\\Temp", $"ConnectorSizes_Export_Integrate_{timestamp}.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Integrate");

                // 헤더
                string[] headers = new[]
                {
                    "BMArea", "BMUnit", "BMZone", "BMDiscipline", "BMSubDiscipline", "SystemType",
                    "BMFluid", "BMClass", "BMScode", "LargeDiameter", "SmallDiameter",
                    "Size", "Quantity", "CommodityCode", "Description", "Unit"
                };

                for (int i = 0; i < headers.Length; i++)
                    ws.Cell(1, i + 1).Value = headers[i];

                // 내용 입력
                for (int r = 0; r < filteredRows.Count; r++)
                {
                    var row = filteredRows[r];
                    for (int c = 0; c < headers.Length; c++)
                    {
                        ws.Cell(r + 2, c + 1).Value = row.ContainsKey(headers[c]) ? row[headers[c]] : "";
                    }
                }

                workbook.SaveAs(filePath);
            }
        }
    }
}
