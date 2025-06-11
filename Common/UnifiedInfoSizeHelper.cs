using ConnectorSizeExport.Modules;
using System;
using System.Globalization;
using System.Linq;

namespace ConnectorSizeExport.Helpers
{
    public static class UnifiedInfoSizeHelper
    {
        private static double? ParseDouble(string s)
        {
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                return val;
            return null;
        }

        public static void FillSizeStatsFromUnified(UnifiedInfo info)
        {
            var diameters = new[] { info.Diameter1, info.Diameter2, info.Diameter3, info.Diameter4 }
                .Select(ParseDouble).Where(v => v.HasValue).Select(v => v.Value).ToList();

            var widths = new[] { info.Width1, info.Width2, info.Width3, info.Width4 }
                .Select(ParseDouble).Where(v => v.HasValue).Select(v => v.Value).ToList();

            var heights = new[] { info.Height1, info.Height2, info.Height3, info.Height4 }
                .Select(ParseDouble).Where(v => v.HasValue).Select(v => v.Value).ToList();

            info.LargeDiameter = diameters.Any() ? diameters.Max().ToString("0.##") : "";
            info.SmallDiameter = diameters.Any() ? diameters.Min().ToString("0.##") : "";

            info.LargeWidth = widths.Any() ? widths.Max().ToString("0") : "";
            info.SmallWidth = widths.Any() ? widths.Min().ToString("0") : "";

            info.LargeHeight = heights.Any() ? heights.Max().ToString("0") : "";
            info.SmallHeight = heights.Any() ? heights.Min().ToString("0") : "";

            // ✅ LargeSize 조건: duct 관련 && shape이 Rectangular 또는 Oval (width/height 존재)
            bool isDuctRelated = !string.IsNullOrEmpty(info.FamilyName) &&
                                 (info.FamilyName.ToLower().Contains("duct"));
            bool hasRectOrOval = widths.Any() && heights.Any();

            if (isDuctRelated && hasRectOrOval)
            {
                double maxSize = Math.Max(widths.Max(), heights.Max());
                info.LargeSize = maxSize.ToString("0");
            }
            else
            {
                info.LargeSize = ""; // ❗ Pipe, Conduit 등은 표시하지 않음
            }
        }
    }
}
