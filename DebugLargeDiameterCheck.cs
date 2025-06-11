using System;
using System.Collections.Generic;
using System.Globalization;
using ClosedXML.Excel;
using ConnectorSizeExport.Models;

namespace ConnectorSizeExport.IO
{
    public static class DebugLargeDiameterCheck
    {
        public static void Export(string path, List<ConnectorExportRow> rows, List<SettingCondition> conditions)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("DiameterCheck");

            // ✅ 헤더 작성
            ws.Cell(1, 1).Value = "Index";
            ws.Cell(1, 2).Value = "ElementId";
            ws.Cell(1, 3).Value = "LargeDiameter";
            ws.Cell(1, 4).Value = "Setting_LargeMin";
            ws.Cell(1, 5).Value = "Setting_LargeMax";
            ws.Cell(1, 6).Value = "SmallDiameter";
            ws.Cell(1, 7).Value = "Setting_SmallMin";
            ws.Cell(1, 8).Value = "Setting_SmallMax";

            ws.Cell(1, 9).Value = "BMArea";
            ws.Cell(1, 10).Value = "BMUnit";
            ws.Cell(1, 11).Value = "BMZone";
            ws.Cell(1, 12).Value = "BMDiscipline";
            ws.Cell(1, 13).Value = "BMSubDiscipline";
            ws.Cell(1, 14).Value = "BMFluid";
            ws.Cell(1, 15).Value = "BMClass";
            ws.Cell(1, 16).Value = "BMScode";

            int index = 1;
            int rowIdx = 2;

            foreach (var row in rows)
            {
                // Diameter 파싱
                bool hasLarge = double.TryParse(row.LargeDiameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var large);
                bool hasSmall = double.TryParse(row.SmallDiameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var small);

                SettingCondition matched = null;

                foreach (var cond in conditions)
                {
                    bool largeMatch = double.TryParse(cond.LargeDiameterMin, out var lmin) &&
                                      double.TryParse(cond.LargeDiameterMax, out var lmax) &&
                                      hasLarge && (large >= lmin && large <= lmax);

                    bool smallMatch = double.TryParse(cond.SmallDiameterMin, out var smin) &&
                                      double.TryParse(cond.SmallDiameterMax, out var smax) &&
                                      hasSmall && (small >= smin && small <= smax);

                    // Large 또는 Small 중 하나라도 조건 만족 시 기록
                    if (largeMatch || smallMatch)
                    {
                        matched = cond;
                        break;
                    }
                }

                if (matched == null) continue;

                // 기록
                ws.Cell(rowIdx, 1).Value = index++;
                ws.Cell(rowIdx, 2).Value = row.ElementId;
                ws.Cell(rowIdx, 3).Value = large;
                ws.Cell(rowIdx, 4).Value = matched.LargeDiameterMin;
                ws.Cell(rowIdx, 5).Value = matched.LargeDiameterMax;
                ws.Cell(rowIdx, 6).Value = small;
                ws.Cell(rowIdx, 7).Value = matched.SmallDiameterMin;
                ws.Cell(rowIdx, 8).Value = matched.SmallDiameterMax;

                ws.Cell(rowIdx, 9).Value = row.BMArea;
                ws.Cell(rowIdx, 10).Value = row.BMUnit;
                ws.Cell(rowIdx, 11).Value = row.BMZone;
                ws.Cell(rowIdx, 12).Value = row.BMDiscipline;
                ws.Cell(rowIdx, 13).Value = row.BMSubDiscipline;
                ws.Cell(rowIdx, 14).Value = row.BMFluid;
                ws.Cell(rowIdx, 15).Value = row.BMClass;
                ws.Cell(rowIdx, 16).Value = row.BMScode;

                rowIdx++;
            }

            wb.SaveAs(path);
        }
    }
}
