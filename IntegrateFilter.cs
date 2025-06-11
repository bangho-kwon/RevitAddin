// ✅ IntegrateFilter.cs - 조건 매칭 및 행 필터링
using ConnectorSizeExport.Models;
using ConnectorSizeExport.IO;
using System.Collections.Generic;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class IntegrateFilter
    {
        public static List<Dictionary<string, string>> FilterMatchingRows(
            List<ConnectorExportRow> connectorRows,
            List<SettingCondition> settingConditions)
        {
            return connectorRows
                .Where(row => settingConditions.Any(cond => cond.IsMatch(row)))
                .Select(row => IntegrateRowBuilder.Build(row, settingConditions))
                .ToList();
        }
    }
}