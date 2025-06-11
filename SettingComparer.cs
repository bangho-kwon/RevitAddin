using System;
using System.Linq;

namespace ConnectorSizeExport.Helpers
{
    public static class SettingComparer
    {
        /// <summary>
        /// 쉼표로 구분된 다중 조건 중 하나라도 정확히 일치하면 true
        /// </summary>
        public static bool IsFieldMatch(string settingValue, string exportValue)
        {
            if (string.IsNullOrWhiteSpace(settingValue)) return true;
            if (string.IsNullOrWhiteSpace(exportValue)) return false;

            var conditions = settingValue.Split(',')
                .Select(v => v.Trim().ToLowerInvariant());

            var value = exportValue.Trim().ToLowerInvariant();
            return conditions.Any(cond => value == cond);
        }
    }
}
