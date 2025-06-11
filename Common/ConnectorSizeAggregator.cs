// 파일 2: ConnectorSizeAggregator.cs
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace ConnectorSizeExport.Modules
{
    public static class ConnectorSizeAggregator
    {
        public static List<ConnectorSizeInfo> Flatten(List<ConnectorSizeInfo> items)
        {
            var result = new List<ConnectorSizeInfo>();

            foreach (var item in items)
            {
                foreach (var conn in item.Connectors)
                {
                    result.Add(new ConnectorSizeInfo
                    {
                        Element = item.Element,
                        BMScode = item.BMScode,
                        BMUnit = item.BMUnit,
                        FamilyName = item.FamilyName,
                        TypeName = item.TypeName,
                        BasicSize = item.BasicSize,
                        PartType = item.PartType + "-Add",
                        TransitionType = item.TransitionType,
                        BMLength = item.BMLength,
                        Connectors = new List<Connector> { conn }
                    });
                }
            }

            return result;
        }
    }
}


