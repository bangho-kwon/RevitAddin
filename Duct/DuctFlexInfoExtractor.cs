using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class DuctFlexInfoExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var ductCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctCurves)
                .WhereElementIsNotElementType();
            var flexDuctCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_FlexDuctCurves)
                .WhereElementIsNotElementType();

            foreach (var elem in ductCollector.Concat(flexDuctCollector))
            {
                string familyName = elem.Category.Name == "Ducts" ? "Duct" : "Flex Duct";
                string partType = familyName;
                string count = "1";
                string connectorCount = "2";

                string typeName = doc.GetElement(elem.GetTypeId())?.Name ?? "";

                string length = "";
                var lenParam = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (lenParam != null && lenParam.StorageType == StorageType.Double)
                {
                    double val = UnitUtils.ConvertFromInternalUnits(lenParam.AsDouble(), UnitTypeId.Millimeters);
                    length = val.ToString("F1", CultureInfo.InvariantCulture);
                }

                string area = "";
                var areaParam = elem.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA);
                if (areaParam != null && areaParam.StorageType == StorageType.Double)
                {
                    double areaVal = areaParam.AsDouble();
                    double sqm = UnitUtils.ConvertFromInternalUnits(areaVal, UnitTypeId.SquareMeters);
                    area = sqm.ToString("F2", CultureInfo.InvariantCulture);
                }

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

                string diameter1 = "";
                string width1 = "";
                string height1 = "";

                if (elem is MEPCurve mcurve)
                {
                    ConnectorSet connectors = mcurve.ConnectorManager.Connectors;
                    foreach (Connector conn in connectors)
                    {
                        if (conn.Shape == ConnectorProfileType.Round)
                        {
                            var diaParam = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                            if (diaParam != null && diaParam.StorageType == StorageType.Double)
                            {
                                var ductUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.DuctSize).GetUnitTypeId();
                                double val = UnitUtils.ConvertFromInternalUnits(diaParam.AsDouble(), ductUnit);
                                diameter1 = val.ToString("G", CultureInfo.InvariantCulture);
                            }
                            break;
                        }
                        else if (conn.Shape == ConnectorProfileType.Rectangular || conn.Shape == ConnectorProfileType.Oval)
                        {
                            var widthParam = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                            var heightParam = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                            if (widthParam != null && heightParam != null &&
                                widthParam.StorageType == StorageType.Double &&
                                heightParam.StorageType == StorageType.Double)
                            {
                                var ductUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.DuctSize).GetUnitTypeId();
                                double w = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), ductUnit);
                                double h = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), ductUnit);
                                width1 = w.ToString("G", CultureInfo.InvariantCulture);
                                height1 = h.ToString("G", CultureInfo.InvariantCulture);
                            }
                            break;
                        }
                    }
                }

                result.Add(new UnifiedInfo
                {
                    ElementId = elem.Id.IntegerValue.ToString(),
                    TypeName = typeName,
                    BasicSize = basicSize,
                    Length = length,
                    SystemType = systemType,
                    Area = area,
                    FamilyName = familyName,
                    PartType = partType,
                    Diameter1 = diameter1,
                    Width1 = width1,
                    Height1 = height1,
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
                    Count = count,
                    ConnectorCount = connectorCount
                });
            }

            return result;
        }
    }
}