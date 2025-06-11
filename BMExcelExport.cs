// BMExcelExport.cs
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class BMExcelExport
    {
        public static void AppendPipeFlexToUnifiedExcel(string filePath, List<UnifiedInfo> data)
        {
            try
            {
                XLWorkbook workbook;
                IXLWorksheet ws;

                if (File.Exists(filePath))
                {
                    workbook = new XLWorkbook(filePath);
                    ws = workbook.Worksheets.First();
                }
                else
                {
                    workbook = new XLWorkbook();
                    ws = workbook.Worksheets.Add("ConnectorSizeExport");
                }

                var headers = new Dictionary<string, int>();
                var headerRow = ws.Row(2);
                int lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;

                for (int col = 1; col <= lastCol; col++)
                {
                    string header = ws.Cell(2, col).GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(header) && !headers.ContainsKey(header))
                        headers[header] = col;
                }

                string[] extractedFields = new[]
                {
                    "ElementId", "TypeName", "BasicSize","SystemType","OutsideDiameter", "InsideDiameter",
                    "Length", "Area", "PipeSegment", "FamilyName",
                    "PartType", "Diameter1","Diameter2","Diameter3","Diameter4",
                    "Width1", "Width2","Width3","Width4","Height1", "Height2",
                    "Height3","Height4", "Count", "ConnectorCount",
                };

                string[] calculatedFields = new[]
                {
                    "ConnectorLength", "LargeDiameter", "SmallDiameter", "LargeWidth", "SmallWidth",
                    "LargeHeight", "SmallHeight", "LargeSize"
                };

                string[] parameterFields = new[]
                {
                    "BMArea", "BMUnit", "BMZone", "BMDiscipline", "BMSubDiscipline",
                    "BMFluid","BMClass","BMScode", "ItemName", "ItemSize"
                };

                int currentCol = lastCol;
                foreach (var field in extractedFields.Concat(calculatedFields).Concat(parameterFields))
                {
                    if (!headers.ContainsKey(field))
                    {
                        currentCol++;
                        ws.Cell(2, currentCol).Value = field;
                        ws.Cell(2, currentCol).Style.Font.Bold = true;
                        headers[field] = currentCol;
                    }
                }

                ws.Row(1).Height = 20;
                ws.Row(2).Height = 18;

                int startExtracted = headers[extractedFields.First()];
                int endExtracted = headers[extractedFields.Last()];
                var extractedRange = ws.Range(1, startExtracted, 1, endExtracted);
                extractedRange.Merge();
                extractedRange.Value = "Extracted Fields";
                extractedRange.Style.Font.FontColor = XLColor.Red;
                extractedRange.Style.Font.Bold = true;
                extractedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                extractedRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                int startCalc = headers[calculatedFields.First()];
                int endCalc = headers[calculatedFields.Last()];
                var calcRange = ws.Range(1, startCalc, 1, endCalc);
                calcRange.Merge();
                calcRange.Value = "Calculated Fields";
                calcRange.Style.Font.FontColor = XLColor.Red;
                calcRange.Style.Font.Bold = true;
                calcRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                calcRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                int startPara = headers[parameterFields.First()];
                int endPara = headers[parameterFields.Last()];
                var paraRange = ws.Range(1, startPara, 1, endPara);
                paraRange.Merge();
                paraRange.Value = "Parameter Fields";
                paraRange.Style.Font.FontColor = XLColor.Red;
                paraRange.Style.Font.Bold = true;
                paraRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                paraRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                int startRow = ws.LastRowUsed()?.RowNumber() + 1 ?? 3;

                foreach (var row in data)
                {
                    var rowMap = new Dictionary<string, string>
                    {
                        { "ElementId", row.ElementId },
                        { "TypeName", row.TypeName },
                        { "OutsideDiameter", row.OutsideDiameter },
                        { "InsideDiameter", row.InsideDiameter },
                        { "BasicSize", row.BasicSize },
                        { "Length", row.Length },
                        { "SystemType", row.SystemType },
                        { "Area", row.Area },
                        { "PipeSegment", row.PipeSegment },
                        { "FamilyName", row.FamilyName },
                        { "PartType", row.PartType },
                        { "Diameter1", row.Diameter1 },
                        { "Diameter2", row.Diameter2 },
                        { "Diameter3", row.Diameter3 },
                        { "Diameter4", row.Diameter4 },
                        { "Width1", row.Width1 },
                        { "Width2", row.Width2 },
                        { "Width3", row.Width3 },
                        { "Width4", row.Width4 },
                        { "Height1", row.Height1 },
                        { "Height2", row.Height2 },
                        { "Height3", row.Height3 },
                        { "Height4", row.Height4 },
                        { "LargeDiameter", row.LargeDiameter },
                        { "SmallDiameter", row.SmallDiameter },
                        { "LargeWidth", row.LargeWidth },
                        { "SmallWidth", row.SmallWidth },
                        { "LargeHeight", row.LargeHeight },
                        { "SmallHeight", row.SmallHeight },
                        { "LargeSize", row.LargeSize },
                        { "BMArea", row.BMArea },
                        { "BMUnit", row.BMUnit },
                        { "BMZone", row.BMZone },
                        { "BMDiscipline", row.BMDiscipline },
                        { "BMSubDiscipline", row.BMSubDiscipline },
                        { "ConnectorLength", row.ConnectorLength },
                        { "BMFluid", row.BMFluid },
                        { "BMClass", row.BMClass },
                        { "BMScode", row.BMScode },
                        { "ItemName", row.ItemName },
                        { "ItemSize", row.ItemSize },
                        { "Count", row.Count },
                        { "ConnectorCount", row.ConnectorCount },
                    };

                    foreach (var kvp in rowMap)
                    {
                        if (headers.TryGetValue(kvp.Key, out int colIdx))
                        {
                            ws.Cell(startRow, colIdx).Value = kvp.Value;
                        }
                    }

                    startRow++;
                }

                ws.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("PipeFlex Export Error", ex.Message);
            }
        }
    }
}
