using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class DuctFittingAccyExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var fittingCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .WhereElementIsNotElementType();

            var accessoryCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .WhereElementIsNotElementType();

            foreach (var elem in fittingCollector.Concat(accessoryCollector))
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

                string systemType = "";
                var sysId = elem.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsElementId() ?? ElementId.InvalidElementId;
                if (sysId != ElementId.InvalidElementId)
                    systemType = doc.GetElement(sysId)?.Name ?? "";

                // 매핑된 파라미터 값 추출
                string bmArea = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Area");
                string bmUnit = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Unit");
                string bmZone = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Zone");
                string bmDiscipline = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM Discipline");
                string bmSubDiscipline = ParameterMappingHelper.GetMappedValueOrDefault(elem, "BM SubDiscipline");
                string ItemName = ParameterMappingHelper.GetMappedValueOrDefault(elem, "Item Name");
                string ItemSize = ParameterMappingHelper.GetMappedValueOrDefault(elem, "Item Size");
                string bmFluid = ParameterMappingHelper.GetMappedValueOrDefault(elem, "FLUID");
                string bmClass = ParameterMappingHelper.GetMappedValueOrDefault(elem, "CLASS");
                string bmScode = ParameterMappingHelper.GetMappedValueOrDefault(elem, "SHORT CODE");

                string dia1 = "", dia2 = "", dia3 = "", dia4 = "";
                string width1 = "", width2 = "", width3 = "", width4 = "";
                string height1 = "", height2 = "", height3 = "", height4 = "";

                ConnectorSizeExtractor.ExtractConnectorSizes(doc, elem, ref dia1, ref dia2, ref dia3, ref dia4, ref width1, ref width2, ref width3, ref width4, ref height1, ref height2, ref height3, ref height4);

                result.Add(new UnifiedInfo
                {
                    ElementId = elem.Id.IntegerValue.ToString(),
                    TypeName = typeName,
                    BasicSize = basicSize,
                    SystemType = systemType,
                    FamilyName = familyName,
                    PartType = partType,
                    Diameter1 = dia1,
                    Diameter2 = dia2,
                    Diameter3 = dia3,
                    Diameter4 = dia4,
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
                    BMFluid = bmFluid,
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
            public static void ExtractConnectorSizes(
                Document doc,
                Element elem,
                ref string dia1, ref string dia2, ref string dia3, ref string dia4,
                ref string width1, ref string width2, ref string width3, ref string width4,
                ref string height1, ref string height2, ref string height3, ref string height4)
            {
                var pipeUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.PipeSize).GetUnitTypeId();
                var ductUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.DuctSize).GetUnitTypeId();

                var diameterList = new List<string>();
                var widthList = new List<string>();
                var heightList = new List<string>();

                ConnectorManager manager = null;

                if (elem is FamilyInstance fi && fi.MEPModel != null)
                    manager = fi.MEPModel.ConnectorManager;

                if (manager == null)
                    return;

                foreach (Connector conn in manager.Connectors)
                {
                    if (conn.Shape == ConnectorProfileType.Round && diameterList.Count < 4)
                    {
                        double d = conn.Radius * 2;
                        double conv = UnitUtils.ConvertFromInternalUnits(d, pipeUnit);
                        diameterList.Add(conv.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                    else if ((conn.Shape == ConnectorProfileType.Rectangular || conn.Shape == ConnectorProfileType.Oval) && widthList.Count < 4)
                    {
                        double w = conn.Width;
                        double h = conn.Height;
                        double convW = UnitUtils.ConvertFromInternalUnits(w, ductUnit);
                        double convH = UnitUtils.ConvertFromInternalUnits(h, ductUnit);
                        widthList.Add(convW.ToString("0.##", CultureInfo.InvariantCulture));
                        heightList.Add(convH.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }

                if (diameterList.Count > 0) dia1 = diameterList[0];
                if (diameterList.Count > 1) dia2 = diameterList[1];
                if (diameterList.Count > 2) dia3 = diameterList[2];
                if (diameterList.Count > 3) dia4 = diameterList[3];

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
