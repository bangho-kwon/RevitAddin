using ConnectorSizeExport.Helpers;
using ConnectorSizeExport.Models;
using System.Collections.Generic;
using System.Linq;

namespace ConnectorSizeExport.IO
{
    public static class IntegrateRowBuilder
    {
        public static Dictionary<string, string> Build(ConnectorExportRow row, List<SettingCondition> conditions)
        {
            // 다중 조건이 있을 경우 가장 우선순위 높은 조건 선택
            var matchedCondition = conditions
                .Where(cond => cond.IsMatch(row))
                .OrderByDescending(cond => cond.GetPriorityScore()) // 범위 좁은 조건 우선
                .FirstOrDefault();

            if (matchedCondition == null)
                return new Dictionary<string, string>(); // 조건이 없을 경우 빈 값

            var result = new Dictionary<string, string>
            {
                // A~H 열 (조건 매칭 필드)
                ["BMArea"] = row.BMArea,
                ["BMUnit"] = row.BMUnit,
                ["BMZone"] = row.BMZone,
                ["BMDiscipline"] = row.BMDiscipline,
                ["BMSubDiscipline"] = row.BMSubDiscipline,
                ["SystemType"] = row.SystemType,
                ["BMFluid"] = row.BMFluid,
                ["BMClass"] = row.BMClass,
                ["BMScode"] = row.BMScode,

                // J~M 열 (지름 조건)
                ["LargeDiameter"] = row.LargeDiameter,
                ["SmallDiameter"] = row.SmallDiameter,

                // P~Q 열 (계산식 적용 필드)
                ["Size"] = SizeBuilder.Build(matchedCondition.SizeRule, row, matchedCondition.SizeFormat),
                ["Quantity"] = FormulaEvaluator.Evaluate(matchedCondition.QuantityRule, row, matchedCondition.QuantityFormat),

                // R~T 열 (최종 출력 매핑 필드)
                ["CommodityCode"] = matchedCondition.CommodityCode,
                ["Description"] = matchedCondition.Description,
                ["Unit"] = matchedCondition.Unit
            };

            // N열: ElementId 출력 여부
            if (!string.IsNullOrWhiteSpace(matchedCondition.OutputElementId))
                result["ElementId"] = row.ElementId;

            // O열: Item Name 출력 필드명에 따라 동적 처리
            if (!string.IsNullOrWhiteSpace(matchedCondition.OutputItemName))
                result["ItemName"] = row.GetValue(matchedCondition.OutputItemName) ?? "";

            // ItemSize 필드도 추가
            if (!string.IsNullOrWhiteSpace(matchedCondition.ItemSize))
                result["ItemSize"] = row.GetValue(matchedCondition.ItemSize) ?? "";

            return result;
        }
    }
}

