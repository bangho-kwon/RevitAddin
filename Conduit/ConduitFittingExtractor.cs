using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ConnectorSizeExport.Modules
{
    public static class ConduitFittingExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var fittingCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ConduitFitting)
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

                string dia1 = "", dia2 = "", dia3 = "", dia4 = "";
                ConnectorSizeExtractor.ExtractDiameters(elem, doc, ref dia1, ref dia2, ref dia3, ref dia4);

                result.Add(new UnifiedInfo
                {
                    ElementId = elem.Id.IntegerValue.ToString(),
                    TypeName = typeName,
                    BasicSize = basicSize,
                    FamilyName = familyName,
                    PartType = partType,
                    Diameter1 = dia1,
                    Diameter2 = dia2,
                    Diameter3 = dia3,
                    Diameter4 = dia4,
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
            public static void ExtractDiameters(Element elem, Document doc,
                ref string diameter1, ref string diameter2, ref string diameter3, ref string diameter4)
            {
                diameter1 = "";
                diameter2 = "";
                diameter3 = "";
                diameter4 = "";

                var diameters = new List<string>();
                var unitType = doc.GetUnits().GetFormatOptions(SpecTypeId.ConduitSize).GetUnitTypeId();

                ConnectorManager manager = null;
                if (elem is FamilyInstance fi && fi.MEPModel != null)
                    manager = fi.MEPModel.ConnectorManager;

                if (manager == null) return;

                foreach (Connector conn in manager.Connectors)
                {
                    if (conn.Shape == ConnectorProfileType.Round && diameters.Count < 4)
                    {
                        double d = conn.Radius * 2;
                        double conv = UnitUtils.ConvertFromInternalUnits(d, unitType);
                        diameters.Add(conv.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }

                if (diameters.Count > 0) diameter1 = diameters[0];
                if (diameters.Count > 1) diameter2 = diameters[1];
                if (diameters.Count > 2) diameter3 = diameters[2];
                if (diameters.Count > 3) diameter4 = diameters[3];
            }
        }
    }
}
