using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ConnectorSizeExport.Helpers
{
    public static class IntegrateDebugLogger
    {
        public static void ExportDebugSheet(string filePath, List<ConnectorExportRow> rows, List<SettingCondition> conditions)
        {
            var timestamp = System.DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string debugPath = Path.Combine(Path.GetDirectoryName(filePath), $"ConnectorSizes_Integrate_Debug_{timestamp}.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Debug");

                // 헤더
                var headers = new[]
                {
                    "BMArea", "BMUnit", "BMZone", "BMDiscipline", "BMSubDiscipline",
                    "SystemType", "BMFluid", "BMClass", "BMScode",
                    "LargeDiameter", "SmallDiameter",
                    "Match_BMArea", "Match_BMUnit", "Match_BMZone", "Match_BMDiscipline", "Match_BMSubDiscipline",
                    "Match_SystemType", "Match_Fluid", "Match_CLASS", "Match_SHORTCODE",
                    "Match_LargeDiameter", "Match_SmallDiameter",
                    "Final_Match"
                };

                for (int i = 0; i < headers.Length; i++)
                    ws.Cell(1, i + 1).Value = headers[i];

                int rowIndex = 2;
                foreach (var row in rows)
                {
                    foreach (var cond in conditions)
                    {
                        int col = 1;

                        ws.Cell(rowIndex, col++).Value = row.BMArea;
                        ws.Cell(rowIndex, col++).Value = row.BMUnit;
                        ws.Cell(rowIndex, col++).Value = row.BMZone;
                        ws.Cell(rowIndex, col++).Value = row.BMDiscipline;
                        ws.Cell(rowIndex, col++).Value = row.BMSubDiscipline;
                        ws.Cell(rowIndex, col++).Value = row.SystemType;
                        ws.Cell(rowIndex, col++).Value = row.BMFluid;
                        ws.Cell(rowIndex, col++).Value = row.BMClass;
                        ws.Cell(rowIndex, col++).Value = row.BMScode;
                        ws.Cell(rowIndex, col++).Value = row.LargeDiameter;
                        ws.Cell(rowIndex, col++).Value = row.SmallDiameter;

                        bool Equal(string a, string b) =>
                            string.Equals(a?.Trim(), b?.Trim(), System.StringComparison.OrdinalIgnoreCase);

                        bool MultiMatch(string condition, string value)
                        {
                            var required = (condition ?? "").Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s));
                            var actual = (value ?? "").Split(',').Select(s => s.Trim());
                            return required.All(r => actual.Contains(r, System.StringComparer.OrdinalIgnoreCase));
                        }

                        bool InRange(string minStr, string maxStr, string targetStr)
                        {
                            if (string.IsNullOrWhiteSpace(targetStr)) return false;
                            if (!double.TryParse(targetStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double target)) return false;

                            if (!string.IsNullOrWhiteSpace(minStr) && double.TryParse(minStr, out double min) && target < min)
                                return false;
                            if (!string.IsNullOrWhiteSpace(maxStr) && double.TryParse(maxStr, out double max) && target > max)
                                return false;
                            return true;
                        }

                        // 개별 조건 판단
                        bool[] checks = new bool[]
                        {
                            string.IsNullOrWhiteSpace(cond.BMArea) || Equal(row.BMArea, cond.BMArea),
                            string.IsNullOrWhiteSpace(cond.BMUnit) || Equal(row.BMUnit, cond.BMUnit),
                            string.IsNullOrWhiteSpace(cond.BMZone) || Equal(row.BMZone, cond.BMZone),
                            string.IsNullOrWhiteSpace(cond.BMDiscipline) || Equal(row.BMDiscipline, cond.BMDiscipline),
                            string.IsNullOrWhiteSpace(cond.BMSubDiscipline) || Equal(row.BMSubDiscipline, cond.BMSubDiscipline),
                            MultiMatch(cond.SystemType, row.SystemType),
                            MultiMatch(cond.BMFluid, row.BMFluid),
                            MultiMatch(cond.BMClass, row.BMClass),
                            MultiMatch(cond.BMScode, row.BMScode),
                            InRange(cond.LargeDiameterMin, cond.LargeDiameterMax, row.LargeDiameter),
                            string.IsNullOrWhiteSpace(cond.SmallDiameterMin) && string.IsNullOrWhiteSpace(cond.SmallDiameterMax)
                                || InRange(cond.SmallDiameterMin, cond.SmallDiameterMax, row.SmallDiameter)
                        };

                        foreach (var ok in checks)
                            ws.Cell(rowIndex, col++).Value = ok ? "O" : "X";

                        ws.Cell(rowIndex, col).Value = checks.All(x => x) ? "O" : "X";
                        break; // 조건 여러 개 중 하나만 비교할 경우
                    }
                    rowIndex++;
                }

                workbook.SaveAs(debugPath);
            }
        }
    }
}
