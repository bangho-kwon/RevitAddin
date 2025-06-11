using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ConnectorSizeExport.Helpers
{
    public static class ParameterMappingHelper
    {
        // 기본 매핑 Dictionary (사용자가 이 값을 수정 가능하도록 확장 가능)
        private static readonly Dictionary<string, string> _parameterMappings = new Dictionary<string, string>
        {
            { "BM Area", "BM Area" },
            { "BM Unit", "BM Unit" },
            { "BM Zone", "BM Zone" },
            { "BM Discipline", "BM Discipline" },
            { "BM SubDiscipline", "BM SubDiscipline" },
            { "FLUID", "SK_UTIL" },
            { "CLASS", "SK_CLASS" },
            { "SHORT CODE", "SK_SCODE" },
        };

        /// <summary>
        /// 매핑 정보를 사용자 정의로 설정할 수 있도록 제공
        /// </summary>
        public static void SetParameterMapping(string logicalName, string actualParameterName)
        {
            if (_parameterMappings.ContainsKey(logicalName))
                _parameterMappings[logicalName] = actualParameterName;
            else
                _parameterMappings.Add(logicalName, actualParameterName);
        }

        /// <summary>
        /// Element에서 매핑된 파라미터 이름으로 값을 가져옴 (Instance → Type 순서로 확인). 없으면 빈 문자열 반환.
        /// </summary>
        public static string GetMappedValueOrDefault(Element elem, string logicalName)
        {
            if (!_parameterMappings.TryGetValue(logicalName, out string actualName))
            {
                actualName = logicalName;
            }

            Parameter param = elem.LookupParameter(actualName);

            // Instance 파라미터 없으면 Type 파라미터 확인
            if (param == null)
            {
                Element typeElem = elem.Document.GetElement(elem.GetTypeId());
                param = typeElem?.LookupParameter(actualName);
            }

            if (param == null || !param.HasValue)
                return "";

            return GetParameterValueAsString(param);
        }

        /// <summary>
        /// 파라미터 값을 StorageType에 맞게 문자열로 반환
        /// </summary>
        private static string GetParameterValueAsString(Parameter param)
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? "";
                case StorageType.Double:
                    return param.AsDouble().ToString();
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.ElementId:
                    return param.AsElementId().IntegerValue.ToString();
                default:
                    return "";
            }
        }

        /// <summary>
        /// 현재 매핑된 파라미터 이름 확인용
        /// </summary>
        public static string GetMappedParameterName(string logicalName)
        {
            return _parameterMappings.TryGetValue(logicalName, out string actualName)
                ? actualName
                : logicalName;
        }

        /// <summary>
        /// 전체 매핑 정보 조회
        /// </summary>
        public static Dictionary<string, string> GetAllMappings()
        {
            return new Dictionary<string, string>(_parameterMappings);
        }
    }
}
