using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class PipeFittingConnectorExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();
            var units = doc.GetUnits(); // Project Unit

            var fittings = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeFitting)
                .WhereElementIsNotElementType()
                .OfType<FamilyInstance>();

            foreach (var fi in fittings)
            {
                string elementId = fi.Id.IntegerValue.ToString();
                string familyName = fi.Category?.Name ?? "Unknown";
                string typeName = doc.GetElement(fi.GetTypeId())?.Name ?? "";
                string partType = GetPartTypeName(fi);
                string count = "1";

                string systemType = "";
                var sysId = fi.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsElementId() ?? ElementId.InvalidElementId;
                if (sysId != ElementId.InvalidElementId)
                    systemType = doc.GetElement(sysId)?.Name ?? "";

                string basicSize = GetFormattedBasicSize(fi);

                // 매핑된 파라미터 값 추출
                string bmArea = ParameterMappingHelper.GetMappedValueOrDefault(fi, "BM Area");
                string bmUnit = ParameterMappingHelper.GetMappedValueOrDefault(fi, "BM Unit");
                string bmZone = ParameterMappingHelper.GetMappedValueOrDefault(fi, "BM Zone");
                string bmDiscipline = ParameterMappingHelper.GetMappedValueOrDefault(fi, "BM Discipline");
                string bmSubDiscipline = ParameterMappingHelper.GetMappedValueOrDefault(fi, "BM SubDiscipline");
                string ItemName = ParameterMappingHelper.GetMappedValueOrDefault(fi, "Item Name");
                string ItemSize = ParameterMappingHelper.GetMappedValueOrDefault(fi, "Item Size");
                string bmFluid = ParameterMappingHelper.GetMappedValueOrDefault(fi, "FLUID");
                string bmClass = ParameterMappingHelper.GetMappedValueOrDefault(fi, "CLASS");
                string bmScode = ParameterMappingHelper.GetMappedValueOrDefault(fi, "SHORT CODE");


                if (fi.MEPModel == null || fi.MEPModel.ConnectorManager == null) continue;

                var connectors = fi.MEPModel.ConnectorManager.Connectors;
                var connectorList = connectors.Cast<Connector>().ToList();
                XYZ origin = GetOriginPoint(fi);

                // Cap 처리
                if (connectorList.Count == 1 && IsCap(partType))
                {
                    var conn = connectorList[0];
                    string diameter = GetFormattedDiameter(doc, conn);
                    string connectorLength = "";

                    var bbox = fi.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        var dir = conn.CoordinateSystem.BasisZ.Normalize();
                        var size = bbox.Max - bbox.Min;
                        double capThickness = Math.Abs(
                            size.X * dir.X + size.Y * dir.Y + size.Z * dir.Z
                        );
                        double bmLength = UnitUtils.ConvertFromInternalUnits(capThickness, UnitTypeId.Millimeters);
                        connectorLength = bmLength.ToString("0.##");
                    }


                    result.Add(new UnifiedInfo
                    {
                        ElementId = elementId,
                        TypeName = typeName,
                        BasicSize = basicSize,
                        Diameter1 = diameter,
                        SystemType = systemType,
                        FamilyName = familyName,
                        PartType = partType,
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
                        ConnectorCount = "1",
                        ConnectorLength = connectorLength,
                        Count = count
                    });
                    continue;
                }

                // Transition 처리
                if (connectorList.Count == 2 && IsTransition(partType))
                {
                    var conn1 = connectorList[0];
                    var conn2 = connectorList[1];
                    if (conn1 == null || conn2 == null) continue;

                    double directDist = conn1.Origin.DistanceTo(conn2.Origin);
                    double distMM = UnitUtils.ConvertFromInternalUnits(directDist, UnitTypeId.Millimeters);
                    string avgConnectorLength = (distMM / 2).ToString("0.##");

                    for (int i = 0; i < 2; i++)
                    {
                        var conn = connectorList[i];
                        string diameter = GetFormattedDiameter(doc, conn);


                        result.Add(new UnifiedInfo
                        {
                            ElementId = elementId,
                            TypeName = typeName,
                            BasicSize = basicSize,
                            Diameter1 = diameter,
                            SystemType = systemType,
                            FamilyName = familyName,
                            PartType = partType,
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
                            ConnectorCount = connectorList.Count.ToString(),
                            ConnectorLength = avgConnectorLength, // 두 커넥터에 동일 길이 할당
                            Count = count
                        });
                    }

                    continue;
                }


                // 일반 Connector 처리
                foreach (var conn in connectorList)
                {
                    if (conn == null || conn.Domain != Domain.DomainPiping) continue;

                    string diameter = GetFormattedDiameter(doc, conn);
                    double dist = origin.DistanceTo(conn.Origin);
                    double distMM = UnitUtils.ConvertFromInternalUnits(dist, UnitTypeId.Millimeters);
                    string connectorLength = distMM.ToString("0.##");


                    result.Add(new UnifiedInfo
                    {
                        ElementId = elementId,
                        TypeName = typeName,
                        BasicSize = basicSize,
                        Diameter1 = diameter,
                        SystemType = systemType,
                        FamilyName = familyName,
                        PartType = partType,
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
                        ConnectorCount = "1",
                        ConnectorLength = connectorLength,
                        Count = count
                    });
                }
            }

            return result;
        }

        private static string GetFormattedBasicSize(Element elem)
        {
            return elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString() ?? "";
        }

        private static string GetFormattedDiameter(Document doc, Connector conn)
        {
            if (conn == null || conn.Shape != ConnectorProfileType.Round) return "";

            double diameter = conn.Radius * 2;

            var formatOptions = doc.GetUnits().GetFormatOptions(SpecTypeId.PipeSize);
            var unitTypeId = formatOptions.GetUnitTypeId();

            double converted = UnitUtils.ConvertFromInternalUnits(diameter, unitTypeId);
            return converted.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static XYZ GetOriginPoint(FamilyInstance fi)
        {
            if (fi.Location is LocationPoint lp) return lp.Point;
            var bbox = fi.get_BoundingBox(null);
            return bbox != null ? (bbox.Min + bbox.Max) / 2.0 : XYZ.Zero;
        }

        private static string GetPartTypeName(FamilyInstance fi)
        {
            Parameter partTypeParam = fi.Symbol.Family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE);
            if (partTypeParam == null || !partTypeParam.HasValue)
                return "Unknown";

            int partTypeValue = partTypeParam.AsInteger();
            if (Enum.IsDefined(typeof(PartType), partTypeValue))
                return ((PartType)partTypeValue).ToString();

            return $"Unknown ({partTypeValue})";
        }

        private static bool IsCap(string partType)
        {
            return partType != null && partType.ToLower().Contains("cap");
        }

        private static bool IsTransition(string partType)
        {
            return partType != null && partType.ToLower().Contains("transition");
        }
    }
}
