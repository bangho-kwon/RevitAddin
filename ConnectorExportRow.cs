using System.Collections.Generic;

namespace ConnectorSizeExport.Models
{
    public class ConnectorExportRow
    {
        public string BMArea { get; set; }
        public string BMUnit { get; set; }
        public string BMZone { get; set; }
        public string BMDiscipline { get; set; }
        public string BMSubDiscipline { get; set; }
        public string SystemType { get; set; }
        public string BMFluid { get; set; }
        public string BMClass { get; set; }
        public string BMScode { get; set; }
        public string LargeDiameter { get; set; }
        public string SmallDiameter { get; set; }
        public string ElementId { get; set; }
        public string FamilyName { get; set; }

        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// {파라미터} 형태를 위한 함수
        /// </summary>
        public string GetParameterValue(string paramName)
        {
            if (string.IsNullOrWhiteSpace(paramName)) return "";
            var key = paramName.Trim();
            if (Values.TryGetValue(key, out var val))
                return val;
            return "";
        }

        /// <summary>
        /// [] 또는 수식 조합을 위한 ExportData 값 접근 함수
        /// </summary>
        public string GetValue(string field)
        {
            if (string.IsNullOrWhiteSpace(field)) return "";

            var key = field.Trim();

            // 먼저 고정 필드 확인 (Export된 주요 열)
            switch (key.ToLower())
            {
                case "bmarea": return BMArea;
                case "bmunit": return BMUnit;
                case "bmzone": return BMZone;
                case "bmdiscipline": return BMDiscipline;
                case "bmsubdiscipline": return BMSubDiscipline;
                case "systemtype": return SystemType;
                case "bmfluid": return BMFluid;
                case "bmclass": return BMClass;
                case "bmscode": return BMScode;
                case "largediameter": return LargeDiameter;
                case "smalldiameter": return SmallDiameter;
                case "elementid": return ElementId;
                case "familyname": return FamilyName;
            }

            // 아니면 Values 딕셔너리에서
            if (Values.TryGetValue(key, out var val))
                return val;

            return "";
        }
    }
}
