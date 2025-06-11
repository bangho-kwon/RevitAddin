using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using ConnectorSizeExport.Models;

namespace ConnectorSizeExport.Modules
{
    public static class IntegrateSheetWriter
    {
        public static void Write(string filePath, List<(ConnectorExportRow Row, SettingCondition Cond)> matches)
        {
            // ✅ 파일 존재 여부에 따라 Workbook 생성 방식 분기
            var workbook = File.Exists(filePath)
                ? new XLWorkbook(filePath)
                : new XLWorkbook(); // 없으면 새로 생성

            // 기존 Integrate 시트가 있다면 제거
            if (workbook.Worksheets.Contains("Integrate"))
                workbook.Worksheets.Delete("Integrate");

            var ws = workbook.Worksheets.Add("Integrate");

            // 헤더 작성
            string[] headers = new string[]
            {
                "BMArea", "BMUnit", "BMZone", "BMDiscipline", "BMSubDiscipline", "SystemType",
                "BMFluid", "BMClass", "BMScode",
                "LargeDiameter Min", "LargeDiameter Max",
                "SmallDiameter Min", "SmallDiameter Max",
                "Size", "Quantity", "Commodity Code", "Description", "Unit"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
                ws.Cell(1, i + 1).Style.Font.Bold = true;
            }

            // 데이터 작성
            int row = 2;
            foreach (var (rowData, cond) in matches)
            {
                string Get(string key) =>
                    rowData.Values.ContainsKey(key) ? rowData.Values[key] : "";

                // A~M열 조건 값
                ws.Cell(row, 1).Value = cond.BMArea;
                ws.Cell(row, 2).Value = cond.BMUnit;
                ws.Cell(row, 3).Value = cond.BMZone;
                ws.Cell(row, 4).Value = cond.BMDiscipline;
                ws.Cell(row, 5).Value = cond.BMSubDiscipline;
                ws.Cell(row, 6).Value = Get("SystemType");
                ws.Cell(row, 7).Value = Get("BMFluid");
                ws.Cell(row, 8).Value = Get("BMClass");
                ws.Cell(row, 9).Value = Get("BMScode");
                ws.Cell(row, 10).Value = cond.LargeDiameterMin;
                ws.Cell(row, 11).Value = cond.LargeDiameterMax;
                ws.Cell(row, 12).Value = cond.SmallDiameterMin;
                ws.Cell(row, 13).Value = cond.SmallDiameterMax;

                // N: SizeFormat 적용
                ws.Cell(row, 14).Value = ReplacePlaceholders(cond.SizeFormat, rowData.Values);

                // O: QuantityFormat 계산
                ws.Cell(row, 15).Value = EvaluateQuantity(cond.QuantityFormat, rowData.Values);

                // P~R: 매핑 정보
                ws.Cell(row, 16).Value = cond.CommodityCode;
                ws.Cell(row, 17).Value = cond.Description;
                ws.Cell(row, 18).Value = cond.Unit;

                row++;
            }

            // 저장
            workbook.SaveAs(filePath);
        }

        private static string ReplacePlaceholders(string format, Dictionary<string, string> values)
        {
            if (string.IsNullOrWhiteSpace(format)) return "";

            string result = format;

            // [ExportField]
            var exportMatches = Regex.Matches(result, @"\[(.*?)\]");
            foreach (Match m in exportMatches)
            {
                var key = m.Groups[1].Value;
                if (values.TryGetValue(key, out string val))
                    result = result.Replace(m.Value, val);
            }

            // {ParameterName}
            var paramMatches = Regex.Matches(result, @"\{(.*?)\}");
            foreach (Match m in paramMatches)
            {
                var key = m.Groups[1].Value;
                if (values.TryGetValue(key, out string val))
                    result = result.Replace(m.Value, val);
            }

            return result;
        }

        private static string EvaluateQuantity(string expr, Dictionary<string, string> values)
        {
            if (string.IsNullOrWhiteSpace(expr)) return "";

            try
            {
                string replaced = expr;

                // [ExportField]
                var matches = Regex.Matches(replaced, @"\[(.*?)\]");
                foreach (Match m in matches)
                {
                    var key = m.Groups[1].Value;
                    if (values.TryGetValue(key, out string val) &&
                        double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                    {
                        replaced = replaced.Replace(m.Value, num.ToString(CultureInfo.InvariantCulture));
                    }
                    else return "";
                }

                // 계산 (DataTable 사용)
                var result = new System.Data.DataTable().Compute(replaced, null);
                return Convert.ToDouble(result).ToString("0.##");
            }
            catch
            {
                return "";
            }
        }
    }
}
