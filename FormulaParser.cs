using ConnectorSizeExport.Models;
using System.Text.RegularExpressions;

namespace ConnectorSizeExport.Helpers
{
    public static class FormulaParser
    {
        /// <summary>
        /// 수식 내의 {파라미터} 및 [필드] 값을 실제 row 값으로 치환
        /// </summary>
        public static string Parse(string formula, ConnectorExportRow row)
        {
            if (string.IsNullOrWhiteSpace(formula)) return "";

            string result = formula;

            // {파라미터} → row.GetParameterValue()
            result = Regex.Replace(result, @"\{(.*?)\}", match =>
            {
                string key = match.Groups[1].Value.Trim();
                return row.GetParameterValue(key);
            });

            // [필드] → row.GetValue()
            result = Regex.Replace(result, @"\[(.*?)\]", match =>
            {
                string key = match.Groups[1].Value.Trim();
                return row.GetValue(key);
            });

            return result;
        }
    }
}
