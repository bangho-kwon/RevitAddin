using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Helpers
{
    public class ConnectorStats
    {
        public string[] Diameters = new string[4];
        public string[] WidthHeights = new string[4];
        public string LargeDiameter = "";
        public string SmallDiameter = "";
        public string LargeWidth = "";
        public string SmallWidth = "";
        public string LargeHeight = "";
        public string SmallHeight = "";
        public string LargeSize = "";
        public string LargeWidthHeight = "";
        public string SmallWidthHeight = "";
    }

    public static class ConnectorSizeCalculator
    {
        public static ConnectorStats Calculate(Document doc, List<Connector> connectors)
        {
            var stats = new ConnectorStats();
            var unitTypeId = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId();

            int diaIndex = 0, whIndex = 0;
            double? maxDia = null, minDia = null, maxW = null, minW = null, maxH = null, minH = null;

            foreach (var conn in connectors)
            {
                if (conn == null) continue;

                if (conn.Shape == ConnectorProfileType.Round)
                {
                    double d = UnitUtils.ConvertFromInternalUnits(conn.Radius * 2, unitTypeId);
                    if (diaIndex < 4)
                        stats.Diameters[diaIndex++] = d.ToString("0.##", CultureInfo.InvariantCulture);
                    maxDia = !maxDia.HasValue ? d : Math.Max(maxDia.Value, d);
                    minDia = !minDia.HasValue ? d : Math.Min(minDia.Value, d);
                }
                else if (conn.Shape == ConnectorProfileType.Rectangular || conn.Shape == ConnectorProfileType.Oval)
                {
                    double w = UnitUtils.ConvertFromInternalUnits(conn.Width, unitTypeId);
                    double h = UnitUtils.ConvertFromInternalUnits(conn.Height, unitTypeId);
                    string whStr = w.ToString("0") + "x" + h.ToString("0");
                    if (whIndex < 4)
                        stats.WidthHeights[whIndex++] = whStr;

                    maxW = !maxW.HasValue ? w : Math.Max(maxW.Value, w);
                    minW = !minW.HasValue ? w : Math.Min(minW.Value, w);
                    maxH = !maxH.HasValue ? h : Math.Max(maxH.Value, h);
                    minH = !minH.HasValue ? h : Math.Min(minH.Value, h);
                }
            }

            stats.LargeDiameter = maxDia?.ToString("0.##") ?? "";
            stats.SmallDiameter = minDia?.ToString("0.##") ?? "";
            stats.LargeWidth = maxW?.ToString("0") ?? "";
            stats.SmallWidth = minW?.ToString("0") ?? "";
            stats.LargeHeight = maxH?.ToString("0") ?? "";
            stats.SmallHeight = minH?.ToString("0") ?? "";
            stats.LargeSize = new[] { maxDia ?? 0, maxW ?? 0, maxH ?? 0 }.Max().ToString("0");

            if (maxW.HasValue && maxH.HasValue)
                stats.LargeWidthHeight = $"{maxW.Value:0}x{maxH.Value:0}";
            if (minW.HasValue && minH.HasValue)
                stats.SmallWidthHeight = $"{minW.Value:0}x{minH.Value:0}";

            return stats;
        }
    }
}
