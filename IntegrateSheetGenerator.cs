using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class IntegrateSheetGenerator
    {
        public static void CreateIntegrateSheet(string filePath)
        {
            // Excel 파일 열기
            XLWorkbook workbook;
            if (File.Exists(filePath))
            {
                workbook = new XLWorkbook(filePath);
            }
            else
            {
                throw new FileNotFoundException("Excel file not found: " + filePath);
            }

            // 기존 "Integrate" 시트가 존재하면 삭제
            var existingSheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "Integrate");
            if (existingSheet != null)
            {
                workbook.Worksheets.Delete(existingSheet.Name);
            }

            // 새 시트 추가
            var integrateSheet = workbook.Worksheets.Add("Integrate");

            // 헤더 작성
            var headers = new List<string>
            {
                "BMArea", "BMUnit", "BMZone", "BMDiscipline", "BMSubDiscipline", "SystemType",
                "BMFluid", "BMClass", "BMScode", "LargeDiameter Min", "LargeDiameter Max",
                "SmallDiameter Min", "SmallDiameter Max", "Size", "Quantity",
                "CommodityCode", "Description", "Unit"
            };

            for (int i = 0; i < headers.Count; i++)
            {
                integrateSheet.Cell(1, i + 1).Value = headers[i];
                integrateSheet.Cell(1, i + 1).Style.Font.Bold = true;
            }

            // 저장
            workbook.Save();
        }
    }
}
