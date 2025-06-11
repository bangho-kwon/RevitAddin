using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using System;

namespace ConnectorExportUtil
{
    public static class PartTypeHelper
    {
        public static string GetPartTypeName(Element e)
        {
            if (e is FamilyInstance fi && fi.Symbol?.Family != null)
            {
                Parameter partTypeParam = fi.Symbol.Family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE);
                if (partTypeParam == null)
                    return "";

                int partTypeValue = partTypeParam.AsInteger();

                if (Enum.IsDefined(typeof(PartType), partTypeValue))
                {
                    return ((PartType)partTypeValue).ToString();
                }

                return $"Unknown({partTypeValue})";
            }

            return "";
        }
    }
}
