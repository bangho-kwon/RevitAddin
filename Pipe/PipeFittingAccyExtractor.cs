using Autodesk.Revit.DB;
using ConnectorSizeExport.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Modules
{

    public static class PipeFittingAccyExtractor
    {
        public static List<UnifiedInfo> Extract(Document doc)
        {
            var result = new List<UnifiedInfo>();

            var pipeFittingCollector = new FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_PipeFitting)
                 .WhereElementIsNotElementType();

            var pipeAccessoryCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeAccessory)
                .WhereElementIsNotElementType();


            foreach (var elem in pipeFittingCollector.Concat(pipeAccessoryCollector))
            {
                // Category 추출
                string familyName = elem.Category != null ? elem.Category.Name : "Unknown";

                // Part Type 추출
                string partType = "Unknown";

                if (elem is FamilyInstance fi)
                {
                    Parameter partTypeParam = fi.Symbol.Family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE);
                    if (partTypeParam != null && partTypeParam.HasValue)
                    {
                        int intVal = partTypeParam.AsInteger();

                        // 열거형 정의 확인
                        if (Enum.IsDefined(typeof(PartType), intVal))
                        {
                            partType = ((PartType)intVal).ToString(); // 예: "Elbow", "Tee", ...
                        }
                        else
                        {
                            partType = "Unknown PartType (" + intVal.ToString() + ")";
                        }
                    }
                    else
                    {
                        partType = elem.Category != null ? elem.Category.Name : "Unknown";
                    }
                }
                else
                {
                    partType = elem.Category != null ? elem.Category.Name : "Unknown";
                }


                string count = "1";

                string typeName = doc.GetElement(elem.GetTypeId())?.Name ?? "";

                // 실수형 원본 값 + 단위 없이 문자열 변환

                string basicSize = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString() ?? "";

                string systemType = "";
                var sysId = elem.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsElementId() ?? ElementId.InvalidElementId;
                if (sysId != ElementId.InvalidElementId)
                    systemType = doc.GetElement(sysId)?.Name ?? "";


                // Pipe Fitting, Accessory Connector 수량을 추출하여 UnifiedInfo에 넣을 때
                string connectorCount = "0";

                if (elem is FamilyInstance inst && inst.MEPModel != null)
                {
                    var connectors = inst.MEPModel.ConnectorManager?.Connectors;
                    if (connectors != null)
                    {
                        connectorCount = connectors.Size.ToString();
                    }
                }

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

                // Pipe Fitting, Accessory Connector Size를 추출(ConnectorDiameterExtractor)하여 UnifiedInfo에 넣을 때
                string dia1, dia2, dia3, dia4;
                ConnectorDiameterExtractor.ExtractDiametersFromConnectors(elem, doc, out dia1, out dia2, out dia3, out dia4);


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
                    Count = count,
                });
            }

            return result;
        }

        //Pipe Fitting, Accessory Connector → Diameter1~4 추출
        public static class ConnectorDiameterExtractor
        {
            public static void ExtractDiametersFromConnectors(Element elem, Document doc, out string diameter1, out string diameter2, out string diameter3, out string diameter4)
            {
                diameter1 = "";
                diameter2 = "";
                diameter3 = "";
                diameter4 = "";

                var diameters = new List<string>();
                var unitType = doc.GetUnits().GetFormatOptions(SpecTypeId.PipeSize).GetUnitTypeId();

                ConnectorManager connectorManager = null;
                if (elem is FamilyInstance fi && fi.MEPModel != null)
                    connectorManager = fi.MEPModel.ConnectorManager;
                else if (elem is MEPCurve mc)
                    connectorManager = mc.ConnectorManager;

                if (connectorManager == null)
                    return;

                int count = 0;
                foreach (Connector conn in connectorManager.Connectors)
                {
                    if (conn.Shape == ConnectorProfileType.Round && count < 4)
                    {
                        double diameter = conn.Radius * 2;
                        double converted = UnitUtils.ConvertFromInternalUnits(diameter, unitType);
                        diameters.Add(converted.ToString("0.##", CultureInfo.InvariantCulture));
                        count++;
                    }
                    if (count >= 4)
                        break;
                }
                // 할당
                if (diameters.Count > 0) diameter1 = diameters[0];
                if (diameters.Count > 1) diameter2 = diameters[1];
                if (diameters.Count > 2) diameter3 = diameters[2];
                if (diameters.Count > 3) diameter4 = diameters[3];
            }
        }
    }

    }

