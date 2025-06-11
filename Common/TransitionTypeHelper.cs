using Autodesk.Revit.DB;
using System;
using System.Linq;

namespace ConnectorExportUtil
{
    public static class TransitionTypeHelper
    {
        private const double FeetToMM = 304.8;

        public static string GetTransitionType(Element e)
        {
            FamilyInstance fi = e as FamilyInstance;
            if (fi == null) return "";

            MEPModel mepModel = fi.MEPModel;
            if (mepModel == null || mepModel.ConnectorManager == null) return "";

            var connectors = mepModel.ConnectorManager.Connectors
                .Cast<Connector>()
                .Where(c => c.ConnectorType == ConnectorType.End)
                .OrderBy(c => c.Origin.X)
                .ToList();

            if (connectors.Count != 2)
                return "";

            Connector conn1 = connectors[0];
            Connector conn2 = connectors[1];

            Domain domain1 = conn1.Domain;
            Domain domain2 = conn2.Domain;

            bool isPipe = (domain1 == Domain.DomainPiping && domain2 == Domain.DomainPiping);
            bool isCableTray = (domain1 == Domain.DomainCableTrayConduit && domain2 == Domain.DomainCableTrayConduit);

            XYZ origin1 = conn1.Origin;
            XYZ origin2 = conn2.Origin;

            double deltaY = (origin1.Y - origin2.Y) * FeetToMM;
            double deltaZ = (origin1.Z - origin2.Z) * FeetToMM;

            double tolerance = GetTolerance(conn1, conn2) * FeetToMM;

            if (isPipe)
            {
                bool isEccentric = Math.Abs(deltaY) > tolerance || Math.Abs(deltaZ) > tolerance;
                return isEccentric ? "ECC." : "CON.";
            }
            else if (isCableTray)
            {
                double maxWidth = Math.Max(conn1.Width, conn2.Width) * FeetToMM;
                double widthBasedTolerance = maxWidth * 0.05;
                tolerance = Math.Max(tolerance, widthBasedTolerance);

                if (Math.Abs(deltaY) > tolerance)
                {
                    string eccentricityDirection = deltaY > 0 ? "RIGHT ECC" : "LEFT ECC";
                    return eccentricityDirection;
                }
                else
                {
                    return "CON.";
                }
            }

            return "";
        }

        private static double GetTolerance(Connector conn1, Connector conn2)
        {
            double tolerance = 0.001; // 기본값

            if (conn1.Shape == ConnectorProfileType.Round && conn2.Shape == ConnectorProfileType.Round)
            {
                double maxRadius = Math.Max(conn1.Radius, conn2.Radius);
                tolerance = Math.Max(maxRadius * 0.1, tolerance);
            }
            else if (conn1.Shape == ConnectorProfileType.Rectangular && conn2.Shape == ConnectorProfileType.Rectangular)
            {
                double maxWidth = Math.Max(conn1.Width, conn2.Width);
                tolerance = Math.Max(maxWidth * 0.05, tolerance);
            }

            return tolerance;
        }
    }
}
