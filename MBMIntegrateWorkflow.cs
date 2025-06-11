using ClosedXML.Excel;
using ConnectorSizeExport.Helpers;
using ConnectorSizeExport.IO;
using ConnectorSizeExport.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConnectorSizeExport.Workflows
{
    public static class MBMIntegrateWorkflow
    {
        public static void Run(string connectorExportPath, string settingPath, string outputFolder)
        {
            // 1. 데이터 로드
            var connectorRows = ConnectorExportReader.ReadExcel(connectorExportPath);
            var settingConditions = SettingConditionReader.Read(settingPath);

            // 2. 필터 조건 만족하는 항목 추출
            var filtered = connectorRows
                .Where(row => settingConditions.Any(cond => cond.IsMatch(row)))
                .Select(row => IntegrateRowBuilder.Build(row, settingConditions))
                .ToList();

            // 3. Integrate 파일 이름 변경 (시간 기반)
            string integrateFilePath = Path.Combine(outputFolder, $"ConnectorSizes_Export_Integrate_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            // 4. Integrate 시트 생성 및 저장
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Integrate");

                string[] headers = new[]
                {
                    "BMArea", "BMUnit", "BMZone", "BMDiscipline", "BMSubDiscipline", "SystemType",
                    "BMFluid", "BMClass", "BMScode", "LargeDiameter", "SmallDiameter",
                    "Size", "Quantity", "CommodityCode", "Description", "Unit",
                    "ElementId", "ItemName", "ItemSize"
                };

                for (int i = 0; i < headers.Length; i++)
                    ws.Cell(1, i + 1).Value = headers[i];

                for (int r = 0; r < filtered.Count; r++)
                {
                    var row = filtered[r];
                    for (int c = 0; c < headers.Length; c++)
                    {
                        row.TryGetValue(headers[c], out string val);
                        ws.Cell(r + 2, c + 1).Value = val;
                    }
                }

                workbook.SaveAs(integrateFilePath);
            }

            // 5. Debug 시트 별도 생성 (상위 5개 LargeDiameter 확인용)
            DebugLargeDiameterCheck.Export(integrateFilePath, connectorRows, settingConditions);
        }
    }
}
