using ConnectorSizeExport.Models;
using System.Text.RegularExpressions;

namespace ConnectorSizeExport.Helpers
{
    public static class SizeBuilder
    {
        /// <summary>
        /// SizeRule 내 [LargeDiameter] 등 []로 감싼 키워드를 실측 값으로 치환
        /// </summary>
        public static string Build(string rule, ConnectorExportRow row, string format)
        {
            if (string.IsNullOrWhiteSpace(rule)) return "";

            // []로 감싸진 모든 매개변수를 찾아 치환
            string result = rule;
            var matches = Regex.Matches(rule, @"\[(.*?)\]");

            foreach (Match match in matches)
            {
                string paramName = match.Groups[1].Value.Trim();
                string value = row.GetValue(paramName);

                // 값이 없으면 공백 처리
                result = result.Replace(match.Value, value ?? "");
            }

            return result;
        }
    }
}
