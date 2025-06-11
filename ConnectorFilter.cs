using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConnectorSizeExport.Models;

namespace ConnectorSizeExport.Modules
{
    public static class ConnectorFilter
    {
        public static List<ConnectorExportRow> FilterRows(List<ConnectorExportRow> allRows, List<SettingCondition> conditions)
        {
            var matched = new List<ConnectorExportRow>();

            foreach (var row in allRows)
            {
                foreach (var cond in conditions)
                {
                    if (IsMatch(row, cond))
                    {
                        matched.Add(row);
                        break; // 하나라도 통과하면 매칭
                    }
                }
            }

            return matched;
        }

        public static bool IsMatch(ConnectorExportRow row, SettingCondition cond)
        {
            string Get(string key)
            {
                if (row.Values.TryGetValue(key, out var value))
                    return value?.Trim();
                return "";
            }

            bool EqualsIgnoreCase(string a, string b)
            {
                return string.Equals(a?.Trim(), b?.Trim(), StringComparison.OrdinalIgnoreCase);
            }

            // 일반 필드 비교 (A~E)
            if (!string.IsNullOrWhiteSpace(cond.BMArea) && !EqualsIgnoreCase(Get("BMArea"), cond.BMArea)) return false;
            if (!string.IsNullOrWhiteSpace(cond.BMUnit) && !EqualsIgnoreCase(Get("BMUnit"), cond.BMUnit)) return false;
            if (!string.IsNullOrWhiteSpace(cond.BMZone) && !EqualsIgnoreCase(Get("BMZone"), cond.BMZone)) return false;
            if (!string.IsNullOrWhiteSpace(cond.BMDiscipline) && !EqualsIgnoreCase(Get("BMDiscipline"), cond.BMDiscipline)) return false;
            if (!string.IsNullOrWhiteSpace(cond.BMSubDiscipline) && !EqualsIgnoreCase(Get("BMSubDiscipline"), cond.BMSubDiscipline)) return false;

            // SystemType, Fluid, SHORT CODE: 다중 조건 (쉼표 구분), 하나라도 맞으면 통과
            var sysTypes = cond.GetSystemTypeConditions();
            if (sysTypes != null && sysTypes.Count > 0)
            {
                var sysType = Get("SystemType")?.ToLowerInvariant();
                if (!sysTypes.Contains(sysType)) return false;
            }

            var fluids = cond.GetFluidConditions();
            if (fluids != null && fluids.Count > 0)
            {
                var fluid = Get("BMFluid")?.ToLowerInvariant();
                if (!fluids.Contains(fluid)) return false;
            }

            var shortCodes = cond.GetShortCodeConditions();
            if (shortCodes != null && shortCodes.Count > 0)
            {
                var code = Get("BMScode")?.ToLowerInvariant();
                if (!shortCodes.Contains(code)) return false;
            }

            // LargeDiameter 비교
            if (!string.IsNullOrWhiteSpace(cond.LargeDiameterMin) || !string.IsNullOrWhiteSpace(cond.LargeDiameterMax))
            {
                if (!double.TryParse(Get("LargeDiameter"), NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                    return false;

                if (!string.IsNullOrWhiteSpace(cond.LargeDiameterMin) &&
                    double.TryParse(cond.LargeDiameterMin, out double min) &&
                    val < min)
                    return false;

                if (!string.IsNullOrWhiteSpace(cond.LargeDiameterMax) &&
                    double.TryParse(cond.LargeDiameterMax, out double max) &&
                    val > max)
                    return false;
            }

            // SmallDiameter 비교
            if (!string.IsNullOrWhiteSpace(cond.SmallDiameterMin) || !string.IsNullOrWhiteSpace(cond.SmallDiameterMax))
            {
                if (double.TryParse(Get("SmallDiameter"), NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                {
                    if (!string.IsNullOrWhiteSpace(cond.SmallDiameterMin) &&
                        double.TryParse(cond.SmallDiameterMin, out double min) &&
                        val < min)
                        return false;

                    if (!string.IsNullOrWhiteSpace(cond.SmallDiameterMax) &&
                        double.TryParse(cond.SmallDiameterMax, out double max) &&
                        val > max)
                        return false;
                }
            }

            return true;
        }
    }
}
