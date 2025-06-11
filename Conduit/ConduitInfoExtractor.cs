using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Modules
{

    public static class ConduitInfoExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var conduitCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Conduit)
                .WhereElementIsNotElementType();

            foreach (var elem in conduitCollector)
            {
                string familyName = "Conduit";
                string partType = familyName;
                string count = "1";
                string connectorCount = "2"; // 기본값

                string typeName = doc.GetElement(elem.GetTypeId())?.Name ?? "";

                // 실수형 원본 값 + 단위 없이 문자열 변환
                string outDia = "";
                var outDiaParam = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM);
                if (outDiaParam != null && outDiaParam.StorageType == StorageType.Double)
                {
                    double val = UnitUtils.ConvertFromInternalUnits(outDiaParam.AsDouble(), UnitTypeId.Millimeters);
                    outDia = val.ToString("F1", CultureInfo.InvariantCulture);
                }

                string inDia = "";
                var inDiaParam = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_INNER_DIAM_PARAM);
                if (inDiaParam != null && inDiaParam.StorageType == StorageType.Double)
                {
                    double val = UnitUtils.ConvertFromInternalUnits(inDiaParam.AsDouble(), UnitTypeId.Millimeters);
                    inDia = val.ToString("F1", CultureInfo.InvariantCulture);
                }

                string length = "";
                var lenParam = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (lenParam != null && lenParam.StorageType == StorageType.Double)
                {
                    double val = UnitUtils.ConvertFromInternalUnits(lenParam.AsDouble(), UnitTypeId.Millimeters);
                    length = val.ToString("F1", CultureInfo.InvariantCulture);
                }

                                string basicSize = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString() ?? "";

                string diameter1 = "";
                var diaParam = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                if (diaParam != null && diaParam.StorageType == StorageType.Double)
                {
                    // 프로젝트의 Conduit Size 단위 설정을 가져와서 변환
                    var conduitSizeUnit = doc.GetUnits().GetFormatOptions(SpecTypeId.ConduitSize).GetUnitTypeId();
                    double val = UnitUtils.ConvertFromInternalUnits(diaParam.AsDouble(), conduitSizeUnit);
                    diameter1 = val.ToString("G", CultureInfo.InvariantCulture); // 단위 없이 숫자만 저장

                }

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

                result.Add(new UnifiedInfo
                {
                    ElementId = elem.Id.IntegerValue.ToString(),
                    TypeName = typeName,
                    OutsideDiameter = outDia,
                    InsideDiameter = inDia,
                    BasicSize = basicSize,
                    Length = length,
                    FamilyName = familyName,
                    PartType = partType,
                    Diameter1 = diameter1,
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
