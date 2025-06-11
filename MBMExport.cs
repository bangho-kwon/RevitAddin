using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using ConnectorExportUtil;
using ConnectorSizeExport.Helpers;
using ConnectorSizeExport.Modules;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class MBMExport : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var doc = commandData.Application.ActiveUIDocument.Document;

        //// 1. 각 fitting/acc별 BMLength Dictionary 준비
        //var pipeFittingBMLengths = PipeFittingConnectorLength.GetFittingBMLengths(doc);
        //var pipeAccessoryBMLengths = PipeAccessoryConnectorLength.GetAccessoryBMLengths(doc);
        //var ductFittingBMLengths = DuctFittingConnectorLength.GetFittingBMLengths(doc);
        //var ductAccessoryBMLengths = DuctAccessoryConnectorLength.GetAccessoryBMLengths(doc);
        //var cableTrayFittingBMLengths = CableTrayFittingConnectorLength.GetFittingBMLengths(doc);
        ////var conduitFittingBMLengths = ConduitFittingConnectorExtractor.GetFittingBMLengths(doc);

        //// 2. 엑셀 준비
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("ConnectorExport");

        //string[] headers = new[]
        //{
        //    "ElementId","BMScode","BMUnit","FamilyName","TypeName","BasicSize","PartType","TransitionType",
        //    "BMLength","Diameter1","Diameter2","Diameter3","Diameter4",
        //    "WidthHeight1","WidthHeight2","WidthHeight3","WidthHeight4",
        //    "LargeDiameter","SmallDiameter","LargeWidth","SmallWidth",
        //    "LargeHeight","SmallHeight","LargeSize","LargeWidthHeight","SmallWidthHeight"
        //};
        //for (int i = 0; i < headers.Length; i++)
        //    worksheet.Cell(1, i + 1).Value = headers[i];

        //int row = 2;
        //string bmUnit = "M";

        //// 3. fitting/acc 카테고리 정의 및 Dictionary 매핑
        //var categoriesAndDicts = new (BuiltInCategory Bic, string Code, string Kind, Func<int, double?> GetLength)[]
        //{
        //    (BuiltInCategory.OST_PipeFitting,     "FiCP", "Fitting",      id => pipeFittingBMLengths.TryGetValue(id, out var v) ? v : (double?)null),
        //    (BuiltInCategory.OST_PipeAccessory,   "ACCP", "Accessory",    id => pipeAccessoryBMLengths.TryGetValue(id, out var v) ? v : (double?)null),
        //    (BuiltInCategory.OST_DuctFitting,     "FiCD", "Fitting",      id => ductFittingBMLengths.TryGetValue(id, out var v) ? v : (double?)null),
        //    (BuiltInCategory.OST_DuctAccessory,   "ACCD", "Accessory",    id => ductAccessoryBMLengths.TryGetValue(id, out var v) ? v : (double?)null),
        //    (BuiltInCategory.OST_CableTrayFitting,"FiTC", "Fitting",      id => cableTrayFittingBMLengths.TryGetValue(id, out var v) ? v : (double?)null),
        //   // (BuiltInCategory.OST_ConduitFitting,  "FiCN", "Fitting",      id => conduitFittingBMLengths.TryGetValue(id, out var v) ? v : (double?)null),
        //};

        //foreach (var (bic, bmScode, kind, getLength) in categoriesAndDicts)
        //{
        //    var collector = new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType();
        //    foreach (Element e in collector)
        //    {
        //        try
        //        {
        //            string fam = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString() ?? "";
        //            string type = e.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString() ?? "";
        //            string basicSize = e.LookupParameter("Size")?.AsValueString() ?? "";
        //            string partType = PartTypeHelper.GetPartTypeName(e);
        //            string transitionType = (partType == "Transition" || partType.Contains("CableTrayTransition")) ? TransitionTypeHelper.GetTransitionType(e) : "";

        //            // [1] 기본행
        //            WriteRow(worksheet, row++, e.Id.IntegerValue, bmScode, bmUnit, fam, type, basicSize, partType, transitionType, "", GetConnectors(e), doc);

        //            // [2] Add 행
        //            var conns = GetConnectors(e);
        //            foreach (var conn in conns)
        //            {
        //                string sizeStr = GetFormattedConnectorSize(doc, conn);
        //                if (string.IsNullOrEmpty(sizeStr)) continue;
        //                string addBmLength = "0";
        //                var bmLength = getLength(e.Id.IntegerValue);
        //                if (bmLength.HasValue && bmLength.Value != 0)
        //                    addBmLength = bmLength.Value.ToString("0.##", CultureInfo.InvariantCulture);

        //                WriteRow(worksheet, row++, e.Id.IntegerValue, bmScode, bmUnit, fam, type, sizeStr, partType + "-Add", transitionType, addBmLength, new List<Connector> { conn }, doc);
        //            }
        //        }
        //        catch { continue; }
        //    }
        //}

        // 4. 저장
        string filepath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"ConnectorSizes_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        try
        {
            workbook.SaveAs(filepath);
            workbook.Dispose();

            // 5. Pipe/FlexPipe 정보 병합 저장
            var pipeList = PipeFlexExtractor.Extract(doc);       // List<UnifiedInfo>
            var ductList = DuctFlexInfoExtractor.Extract(doc);   // List<UnifiedInfo>
            var conduitList = ConduitInfoExtractor.Extract(doc);   // List<UnifiedInfo>
            var cabletrayList = CabletryinfoExtractor.Extract(doc);   // List<UnifiedInfo>
            var pipefittingaccyList = PipeFittingAccyExtractor.Extract(doc);
            var ductfittingaccyList = DuctFittingAccyExtractor.Extract(doc);
            var cabletrayfittingList = CableTrayFittingExtractor.Extract(doc);
            var conduitfittingList = ConduitFittingExtractor.Extract(doc);
            var pipefittingconlengthList = PipeFittingConnectorExtractor.Extract(doc);
            var pipeaccyconlengthList = PipeAccyConnectorExtractor.Extract(doc);
            var ductfittingconlengthList = DuctFittingConnectorExtractor.Extract(doc);
            var ductaccyconlengthList = DuctAccyConnectorExtractor.Extract(doc);
            var calbeltrayfittingconlengthList = CabletrayFittingConnectorExtractor.Extract(doc);
            var conduitfittingconlengthList = ConduitFittingConnectorExtractor.Extract(doc);


            // 통합
            var allData = new List<UnifiedInfo>();
            allData.AddRange(pipeList);
            allData.AddRange(ductList);
            allData.AddRange(conduitList);
            allData.AddRange(cabletrayList);
            allData.AddRange(pipefittingaccyList);
            allData.AddRange(ductfittingaccyList);
            allData.AddRange(cabletrayfittingList);
            allData.AddRange(conduitfittingList);
            allData.AddRange(pipefittingconlengthList);
            allData.AddRange(pipeaccyconlengthList);
            allData.AddRange(ductfittingconlengthList);
            allData.AddRange(ductaccyconlengthList);
            allData.AddRange(calbeltrayfittingconlengthList);
            allData.AddRange(conduitfittingconlengthList);

            foreach (var row in allData)
            {
                UnifiedInfoSizeHelper.FillSizeStatsFromUnified(row);
            }

            // 엑셀로 저장
            BMExcelExport.AppendPipeFlexToUnifiedExcel(filepath, allData);


            TaskDialog.Show("완료", $"저장됨: {filepath}");
        }
        catch
        {
            TaskDialog.Show("오류", "파일 저장 실패 (열려 있을 수 있음)");
            return Result.Failed;
        }

        return Result.Succeeded;
    }

    //private List<Connector> GetConnectors(Element e)
    //{
    //    if (e is FamilyInstance fi && fi.MEPModel?.ConnectorManager != null)
    //        return fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().ToList();
    //    else if (e is MEPCurve mc && mc.ConnectorManager != null)
    //        return mc.ConnectorManager.Connectors.Cast<Connector>().ToList();
    //    return new List<Connector>();
    //}

    //private string GetFormattedConnectorSize(Document doc, Connector conn)
    //{
    //    var unitTypeId = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId();
    //    if (conn.Shape == ConnectorProfileType.Round)
    //    {
    //        double size = UnitUtils.ConvertFromInternalUnits(conn.Radius * 2, unitTypeId);
    //        return size.ToString("0.##");
    //    }
    //    else if (conn.Shape == ConnectorProfileType.Rectangular || conn.Shape == ConnectorProfileType.Oval)
    //    {
    //        double w = UnitUtils.ConvertFromInternalUnits(conn.Width, unitTypeId);
    //        double h = UnitUtils.ConvertFromInternalUnits(conn.Height, unitTypeId);
    //        return w.ToString("0") + "x" + h.ToString("0");
    //    }
    //    return "";
    //}

    //private void WriteRow(IXLWorksheet ws, int row, int elemId, string bmScode, string bmUnit,
    //string famName, string typeName, string basicSize, string partType, string transitionType,
    //string bmLength, List<Connector> connList, Document doc)
    //{
    //    var stats = ConnectorSizeExport.Helpers.ConnectorSizeCalculator.Calculate(doc, connList);

    //    string[] data = new string[]
    //    {
    //    elemId.ToString(), bmScode, bmUnit, famName, typeName, basicSize, partType, transitionType, bmLength,
    //    stats.Diameters.ElementAtOrDefault(0), stats.Diameters.ElementAtOrDefault(1),
    //    stats.Diameters.ElementAtOrDefault(2), stats.Diameters.ElementAtOrDefault(3),
    //    stats.WidthHeights.ElementAtOrDefault(0), stats.WidthHeights.ElementAtOrDefault(1),
    //    stats.WidthHeights.ElementAtOrDefault(2), stats.WidthHeights.ElementAtOrDefault(3),
    //    stats.LargeDiameter, stats.SmallDiameter,
    //    stats.LargeWidth, stats.SmallWidth,
    //    stats.LargeHeight, stats.SmallHeight,
    //    stats.LargeSize, stats.LargeWidthHeight, stats.SmallWidthHeight
    //    };

    //    for (int i = 0; i < data.Length; i++)
    //        ws.Cell(row, i + 1).Value = data[i];
    //}

}
