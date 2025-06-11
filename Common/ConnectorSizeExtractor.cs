using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ConnectorExportUtil;

namespace ConnectorSizeExport.Modules
{
    public class ConnectorSizeInfo
    {
        public Element Element;
        public string BMScode;
        public string BMUnit;
        public string FamilyName;
        public string TypeName;
        public string BasicSize;
        public string PartType;
        public string TransitionType;
        public double BMLength;
        public List<Connector> Connectors;

        public double Diameter1, Diameter2, Diameter3, Diameter4;
        public string WidthHeight1, WidthHeight2, WidthHeight3, WidthHeight4;
        public double LargeDiameter, SmallDiameter;
        public double LargeWidth, SmallWidth;
        public double LargeHeight, SmallHeight;
        public string LargeWidthHeight, SmallWidthHeight;
    }

    public static class ConnectorSizeExtractor
    {
        public static ConnectorSizeInfo Extract(Document doc, Element element, ForgeTypeId unitTypeId)
        {
            var connectors = GetConnectors(element);

            ConnectorSizeInfo info = new ConnectorSizeInfo();
            info.Element = element;
            info.BMScode = element.Name;
            info.BMUnit = "M";
            info.FamilyName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString() ?? "";
            info.TypeName = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString() ?? "";
            info.BasicSize = element.LookupParameter("Size")?.AsValueString() ?? "";
            info.PartType = PartTypeHelper.GetPartTypeName(element);
            info.TransitionType = TransitionTypeHelper.GetTransitionType(element);
            info.Connectors = connectors;

            // BMLength 계산 생략 (각 전용 추출기에서 직접 설정해야 함)
            info.BMLength = 0;

            List<double> diameters = new List<double>();
            List<double> widths = new List<double>();
            List<double> heights = new List<double>();
            List<string> widthHeights = new List<string>();

            for (int i = 0; i < connectors.Count; i++)
            {
                var c = connectors[i];
                double d = 0;
                string wh = "";

                if (c.Shape == ConnectorProfileType.Round && c.Radius > 0)
                {
                    d = UnitUtils.ConvertFromInternalUnits(c.Radius * 2, unitTypeId);
                    diameters.Add(d);
                }
                else if ((c.Shape == ConnectorProfileType.Rectangular || c.Shape == ConnectorProfileType.Oval)
                      && c.Width > 0 && c.Height > 0)
                {
                    double w = UnitUtils.ConvertFromInternalUnits(c.Width, unitTypeId);
                    double h = UnitUtils.ConvertFromInternalUnits(c.Height, unitTypeId);
                    widths.Add(w);
                    heights.Add(h);
                    wh = $"{Math.Round(w)}x{Math.Round(h)}";
                    widthHeights.Add(wh);
                }
                else
                {
                    if (!string.IsNullOrEmpty(info.BasicSize))
                    {
                        string[] parts = Regex.Split(info.BasicSize, "[^0-9.]+")
                                              .Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();

                        foreach (string part in parts)
                        {
                            if (double.TryParse(part, out double parsed))
                            {
                                d = parsed;
                                diameters.Add(d);
                                break;
                            }
                        }
                    }
                }

                if (i == 0) { info.Diameter1 = d; info.WidthHeight1 = wh; }
                if (i == 1) { info.Diameter2 = d; info.WidthHeight2 = wh; }
                if (i == 2) { info.Diameter3 = d; info.WidthHeight3 = wh; }
                if (i == 3) { info.Diameter4 = d; info.WidthHeight4 = wh; }
            }

            if (diameters.Count > 0)
            {
                info.LargeDiameter = diameters.Max();
                info.SmallDiameter = diameters.Min();
            }

            if (widthHeights.Count > 0)
            {
                info.LargeWidthHeight = widthHeights.OrderByDescending(s => s).First();
                info.SmallWidthHeight = widthHeights.OrderBy(s => s).First();
            }

            if (widths.Count > 0)
            {
                info.LargeWidth = widths.Max();
                info.SmallWidth = widths.Min();
            }

            if (heights.Count > 0)
            {
                info.LargeHeight = heights.Max();
                info.SmallHeight = heights.Min();
            }

            return info;
        }

        private static List<Connector> GetConnectors(Element e)
        {
            List<Connector> connectors = new List<Connector>();
            if (e is FamilyInstance fi && fi.MEPModel?.ConnectorManager != null)
            {
                foreach (Connector c in fi.MEPModel.ConnectorManager.Connectors)
                    connectors.Add(c);
            }
            else if (e is MEPCurve curve && curve.ConnectorManager != null)
            {
                foreach (Connector c in curve.ConnectorManager.Connectors)
                    connectors.Add(c);
            }
            return connectors;
        }
    }
}
