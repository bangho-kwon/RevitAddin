using System.Collections.Generic;
using ClosedXML.Excel;
using ConnectorSizeExport.Models;
using System;

namespace ConnectorSizeExport.IO
{
    public static class SettingConditionReader
    {
        public static List<SettingCondition> Read(string filePath)
        {
            var list = new List<SettingCondition>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var ws = workbook.Worksheet(1);
                var headers = new Dictionary<int, string>();
                var headerRow = ws.Row(2); // ✅ 실제 헤더는 2행에 존재

                for (int col = 1; col <= ws.LastColumnUsed().ColumnNumber(); col++)
                {
                    string header = headerRow.Cell(col).GetString().Trim();
                    if (!string.IsNullOrEmpty(header))
                        headers[col] = header;
                }

                for (int row = 3; row <= ws.LastRowUsed().RowNumber(); row++) // ✅ 데이터는 3행부터 시작
                {
                    var cond = new SettingCondition();

                    foreach (var kv in headers)
                    {
                        string val = ws.Cell(row, kv.Key).GetString().Trim();
                        string header = kv.Value.Trim().ToLowerInvariant().Replace(" ", ""); // 소문자 + 공백 제거 비교

                        switch (header)
                        {
                            case "bmarea": cond.BMArea = val; break;
                            case "bmunit": cond.BMUnit = val; break;
                            case "bmzone": cond.BMZone = val; break;
                            case "bmdiscipline": cond.BMDiscipline = val; break;
                            case "bmsubdiscipline": cond.BMSubDiscipline = val; break;
                            case "systemtype": cond.SystemType = val; break;
                            case "bmfluid": cond.BMFluid = val; break;
                            case "bmclass": cond.BMClass = val; break;
                            case "bmscode": cond.BMScode = val; break;

                            case "largediameter(from)":
                            case "largediametermin":
                            case "setting_min":
                                cond.LargeDiameterMin = val; break;

                            case "largediameter(to)":
                            case "largediametermax":
                            case "setting_max":
                                cond.LargeDiameterMax = val; break;

                            case "smalldiameter(from)":
                            case "smalldiametermin":
                                cond.SmallDiameterMin = val; break;

                            case "smalldiameter(to)":
                            case "smalldiametermax":
                                cond.SmallDiameterMax = val; break;

                            case "elementid":
                                cond.OutputElementId = val; break;

                            case "itemname":
                                cond.OutputItemName = val; break;

                            case "size":
                            case "sizerule":
                                cond.SizeRule = val; break;

                            case "sizeformat":
                                cond.SizeFormat = val; break;

                            case "quantity":
                            case "quantityrule":
                                cond.QuantityRule = val; break;

                            case "quantityformat":
                                cond.QuantityFormat = val; break;


                            case "commoditycode":
                                cond.CommodityCode = val; break;

                            case "description":
                                cond.Description = val; break;

                            case "unit":
                                cond.Unit = val; break;


                            case "itemsize":
                                cond.ItemSize = val; break;

                            default:
                                break;
                        }
                    }

                    list.Add(cond);
                }
            }

            return list;
        }
    }
}
