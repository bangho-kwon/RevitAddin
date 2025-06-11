using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class CabletryinfoExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var cabletrayCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_CableTray)
                .WhereElementIsNotElementType();


            foreach (var elem in cabletrayCollector)
            {
                string familyName = "CableTray";
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

                string width1 = "";
                string height1 = "";

                var widthParam = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                var heightParam = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);

                if (widthParam != null && heightParam != null &&
                    widthParam.StorageType == StorageType.Double &&
                    heightParam.StorageType == StorageType.Double)
                {
                    var trayUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.CableTraySize).GetUnitTypeId();
                    double w = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), trayUnit);
                    double h = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), trayUnit);
                    width1 = w.ToString("G", CultureInfo.InvariantCulture);
                    height1 = h.ToString("G", CultureInfo.InvariantCulture);
                }


                result.Add(new UnifiedInfo
                {
                    ElementId = elem.Id.IntegerValue.ToString(),
                    TypeName = typeName,
                    BasicSize = basicSize,
                    Length = length,
                    FamilyName = familyName,
                    PartType = partType,
                    Width1 = width1,
                    Height1 = height1,
                    BMArea = bmArea,
                    BMUnit = bmUnit,
                    BMZone = bmZone,
                    BMDiscipline = bmDiscipline,
                    BMSubDiscipline = bmSubDiscipline,
                    ItemName = ItemName,
                    ItemSize = ItemSize,
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