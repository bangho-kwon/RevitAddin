// SettingCondition.cs
using ConnectorSizeExport.Helpers;
using ConnectorSizeExport.Models;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ConnectorSizeExport.Models
{
    public class SettingCondition
    {
        // A~H
        public string BMArea { get; set; }
        public string BMUnit { get; set; }
        public string BMZone { get; set; }
        public string BMDiscipline { get; set; }
        public string BMSubDiscipline { get; set; }
        public string SystemType { get; set; }
        public string BMFluid { get; set; }
        public string BMClass { get; set; }
        public string BMScode { get; set; }

        // J~M
        public string LargeDiameterMin { get; set; }
        public string LargeDiameterMax { get; set; }
        public string SmallDiameterMin { get; set; }
        public string SmallDiameterMax { get; set; }

        // N~O
        public string OutputElementId { get; set; }
        public string OutputItemName { get; set; }

        // P~Q
        public string SizeRule { get; set; }
        public string SizeFormat { get; set; }

        public string QuantityRule { get; set; }
        public string QuantityFormat { get; set; }


        // R~T
        public string CommodityCode { get; set; }
        public string Description { get; set; }
        public string Unit { get; set; }

        // 확장 파라미터 ({} 조합용)
        public string ItemName { get; set; }
        public string ItemSize { get; set; }


        public List<string> GetSystemTypeConditions()
        {
            return ParseMultiValues(SystemType);
        }
        public List<string> GetClassConditions()
        {
            return ParseMultiValues(BMClass);
        }
        public List<string> GetFluidConditions()
        {
            return ParseMultiValues(BMFluid);
        }

        public List<string> GetShortCodeConditions()
        {
            return ParseMultiValues(BMScode);
        }

        private List<string> ParseMultiValues(string input)
        {
            return (input ?? "")
                .Split(',')
                .Select(s => s.Trim().ToLowerInvariant())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
        }

        public int GetPriorityScore()
        {
            if (double.TryParse(LargeDiameterMin, out var min) &&
                double.TryParse(LargeDiameterMax, out var max))
            {
                // 범위가 좁을수록 우선순위 높음 (작은 값을 더 높은 점수로 보기 위해 음수 부여)
                return -(int)((max - min) * 100);
            }

            return 0;
        }
        public bool IsMatch(ConnectorExportRow row)
        {
            // 문자열 동일 비교: 조건값이 없으면 비교 생략 (true)
            bool Match(string cond, string value) =>
                string.IsNullOrWhiteSpace(cond) || string.Equals(cond?.Trim(), value?.Trim(), StringComparison.OrdinalIgnoreCase);

            // 다중 조건 비교 (쉼표로 구분): 조건값이 없으면 비교 생략 (true)
            bool MultiMatch(string condition, string value)
            {
                if (string.IsNullOrWhiteSpace(condition)) return true;

                var required = condition.Split(',')
                                        .Select(s => s.Trim())
                                        .Where(s => !string.IsNullOrWhiteSpace(s));
                var actual = value?.Trim();
                return required.Any(r => string.Equals(actual, r, StringComparison.OrdinalIgnoreCase));
            }

            // 수치 범위 비교: 조건값이 없으면 비교 생략 (true)
            bool InRange(string minStr, string maxStr, string targetStr, string familyName)
            {
                if (string.IsNullOrWhiteSpace(minStr) && string.IsNullOrWhiteSpace(maxStr))
                    return true;
                if (string.IsNullOrWhiteSpace(targetStr)) return false;
                if (!double.TryParse(targetStr, out double target)) return false;

                if (!string.IsNullOrWhiteSpace(minStr) && double.TryParse(minStr, out double min) && target < min)
                    return false;
                if (!string.IsNullOrWhiteSpace(maxStr) && double.TryParse(maxStr, out double max) && target > max)
                    return false;

                return true;
            }

            return
                Match(BMArea, row.BMArea) &&
                Match(BMUnit, row.BMUnit) &&
                Match(BMZone, row.BMZone) &&
                Match(BMDiscipline, row.BMDiscipline) &&
                Match(BMSubDiscipline, row.BMSubDiscipline) &&
                MultiMatch(SystemType, row.SystemType) &&
                MultiMatch(BMFluid, row.BMFluid) &&
                MultiMatch(BMClass, row.BMClass) &&
                MultiMatch(BMScode, row.BMScode) &&
                InRange(LargeDiameterMin, LargeDiameterMax, row.LargeDiameter, row.FamilyName) &&
                InRange(SmallDiameterMin, SmallDiameterMax, row.SmallDiameter, row.FamilyName);
        }


    }
}
