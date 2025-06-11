using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ConnectorSizeExport.Helpers
{
    public static class ParameterMappingHelper
    {
        // �⺻ ���� Dictionary (����ڰ� �� ���� ���� �����ϵ��� Ȯ�� ����)
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
        /// ���� ������ ����� ���Ƿ� ������ �� �ֵ��� ����
        /// </summary>
        public static void SetParameterMapping(string logicalName, string actualParameterName)
        {
            if (_parameterMappings.ContainsKey(logicalName))
                _parameterMappings[logicalName] = actualParameterName;
            else
                _parameterMappings.Add(logicalName, actualParameterName);
        }

        /// <summary>
        /// Element���� ���ε� �Ķ���� �̸����� ���� ������ (Instance �� Type ������ Ȯ��). ������ �� ���ڿ� ��ȯ.
        /// </summary>
        public static string GetMappedValueOrDefault(Element elem, string logicalName)
        {
            if (!_parameterMappings.TryGetValue(logicalName, out string actualName))
            {
                actualName = logicalName;
            }

            Parameter param = elem.LookupParameter(actualName);

            // Instance �Ķ���� ������ Type �Ķ���� Ȯ��
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
        /// �Ķ���� ���� StorageType�� �°� ���ڿ��� ��ȯ
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
        /// ���� ���ε� �Ķ���� �̸� Ȯ�ο�
        /// </summary>
        public static string GetMappedParameterName(string logicalName)
        {
            return _parameterMappings.TryGetValue(logicalName, out string actualName)
                ? actualName
                : logicalName;
        }

        /// <summary>
        /// ��ü ���� ���� ��ȸ
        /// </summary>
        public static Dictionary<string, string> GetAllMappings()
        {
            return new Dictionary<string, string>(_parameterMappings);
        }
    }
}
