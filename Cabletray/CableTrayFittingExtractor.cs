using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ConnectorSizeExport.Modules
{
    public static class CableTrayFittingExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var fittingCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_CableTrayFitting)
                .WhereElementIsNotElementType();

            foreach (var elem in fittingCollector)
            {
                string familyName = elem.Category?.Name ?? "Unknown";
                string partType = familyName;
                string count = "1";
                string connectorCount = "0";

                if (elem is FamilyInstance inst && inst.MEPModel != null)
                {
                    var connectors = inst.MEPModel.ConnectorManager?.Connectors;
                    if (connectors != null)
                        connectorCount = connectors.Size.ToString();
                }

                string typeName = doc.GetElement(elem.GetTypeId())?.Name ?? "";
                string basicSize = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString() ?? "";

                // 매핑된 파라미터 값 추출
                string bmArea = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Area");
                string bmUnit = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Unit");
                string bmZone = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Zone");
                string bmDiscipline = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Discipline");
                string bmSubDiscipline = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM SubDiscipline");
                string ItemName = ParameterMappingHelper.GetMappedValueOrDefault(elem, "Item Name");
                string ItemSize = ParameterMappingHelper.GetMappedValueOrDefault(elem, "Item Size");
                string bmClass = ParameterMappingHelper.GetMappedValueOrDefault(elem, "CLASS");
                string bmScode = ParameterMappingHelper.GetMappedValueOrDefault(elem, "SHORT CODE");

                // Width/Height 추출
                string width1 = "", width2 = "", width3 = "", width4 = "";
                string height1 = "", height2 = "", height3 = "", height4 = "";
                ConnectorSizeExtractor.ExtractWidthsHeights(elem, doc, ref width1, ref width2, ref width3, ref width4, ref height1, ref height2, ref height3, ref height4);

                result.Add(new UnifiedInfo
                {
                    ElementId = elem.Id.IntegerValue.ToString(),
                    TypeName = typeName,
                    BasicSize = basicSize,
                    FamilyName = familyName,
                    PartType = partType,
                    Width1 = width1,
                    Width2 = width2,
                    Width3 = width3,
                    Width4 = width4,
                    Height1 = height1,
                    Height2 = height2,
                    Height3 = height3,
                    Height4 = height4,
                    BMArea = bmArea,
                    BMUnit = bmUnit,
                    BMZone = bmZone,
                    BMDiscipline = bmDiscipline,
                    BMSubDiscipline = bmSubDiscipline,
                    ItemName = ItemName,
                    ItemSize = ItemSize,
                    BMClass = bmClass,
                    BMScode = bmScode,
                    ConnectorCount = connectorCount,
                    Count = count
                });
            }

            return result;
        }

        public static class ConnectorSizeExtractor
        {
            public static void ExtractWidthsHeights(
                Element elem, Document doc,
                ref string width1, ref string width2, ref string width3, ref string width4,
                ref string height1, ref string height2, ref string height3, ref string height4)
            {
                var trayUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.CableTraySize).GetUnitTypeId();

                var widthList = new List<string>();
                var heightList = new List<string>();

                ConnectorManager manager = null;
                if (elem is FamilyInstance fi && fi.MEPModel != null)
                    manager = fi.MEPModel.ConnectorManager;

                if (manager == null) return;

                foreach (Connector conn in manager.Connectors)
                {
                    if ((conn.Shape == ConnectorProfileType.Rectangular || conn.Shape == ConnectorProfileType.Oval) && widthList.Count < 4)
                    {
                        double w = UnitUtils.ConvertFromInternalUnits(conn.Width, trayUnit);
                        double h = UnitUtils.ConvertFromInternalUnits(conn.Height, trayUnit);
                        widthList.Add(w.ToString("0.##", CultureInfo.InvariantCulture));
                        heightList.Add(h.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }

                if (widthList.Count > 0) width1 = widthList[0];
                if (widthList.Count > 1) width2 = widthList[1];
                if (widthList.Count > 2) width3 = widthList[2];
                if (widthList.Count > 3) width4 = widthList[3];

                if (heightList.Count > 0) height1 = heightList[0];
                if (heightList.Count > 1) height2 = heightList[1];
                if (heightList.Count > 2) height3 = heightList[2];
                if (heightList.Count > 3) height4 = heightList[3];
            }
        }
    }
}