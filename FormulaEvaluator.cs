using ConnectorSizeExport.Models;
using ConnectorSizeExport.Helpers;
using System;
using System.Data;
using System.Globalization;

namespace ConnectorSizeExport
{
    public static class FormulaEvaluator
    {
        public static string Evaluate(string rule, ConnectorExportRow row, string format = null)
        {
            if (string.IsNullOrWhiteSpace(rule)) return "";

            string parsed = FormulaParser.Parse(rule, row);

            try
            {
                var dt = new DataTable();
                var resultObj = dt.Compute(parsed, null);

                if (resultObj is double d1)
                {
                    return string.IsNullOrEmpty(format)
                        ? d1.ToString(CultureInfo.InvariantCulture)
                        : d1.ToString(format, CultureInfo.InvariantCulture);
                }

                if (double.TryParse(resultObj.ToString(), out double d2))
                {
                    return string.IsNullOrEmpty(format)
                        ? d2.ToString(CultureInfo.InvariantCulture)
                        : d2.ToString(format, CultureInfo.InvariantCulture);
                }

                return resultObj.ToString();
            }
            catch
            {
                // 수식 오류 시 원본 수식 문자열 반환
                return parsed;
            }
        }
    }
}
