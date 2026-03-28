#if NET8_0_OR_GREATER
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace THBIM.MEP.Core;

internal static class ConnectorRoutingService
{
    /// <summary>Get best open (unconnected) connector on element, preferring End type.</summary>
    public static Connector? GetBestOpenConnector(Element element)
        => ConnectorUtils.GetOpenConnectors(element)
            .OrderByDescending(c => c.ConnectorType == ConnectorType.End ? 1 : 0)
            .FirstOrDefault();

    /// <summary>
    /// Get best open connector pair between source and target.
    /// BOTH sides must have open connectors. Prefer closest pair with aligned directions.
    /// </summary>
    public static (Connector Source, Connector Target)? GetBestOpenPair(Element source, Element target)
    {
        var sourceCandidates = ConnectorUtils.GetOpenConnectors(source);
        var targetCandidates = ConnectorUtils.GetOpenConnectors(target);

        if (sourceCandidates.Count == 0 || targetCandidates.Count == 0)
            return null;

        var best = sourceCandidates
            .SelectMany(s => targetCandidates.Select(t =>
            {
                var dist = s.Origin.DistanceTo(t.Origin);
                // Dot product: 1.0 = facing each other (ideal), 0 = perpendicular
                var sDir = s.CoordinateSystem.BasisZ.Normalize();
                var tDir = t.CoordinateSystem.BasisZ.Normalize();
                var dot = Math.Abs(sDir.DotProduct(tDir));
                return new { s, t, dist, dot };
            }))
            .OrderBy(x => x.dist)
            .ThenByDescending(x => x.dot)
            .FirstOrDefault();

        return best is null ? null : (best.s, best.t);
    }

    /// <summary>Direct connect two elements via their best open connector pair.</summary>
    public static bool TryDirectConnect(Element source, Element target)
    {
        var pair = GetBestOpenPair(source, target);
        if (pair is null) return false;

        try
        {
            pair.Value.Source.ConnectTo(pair.Value.Target);
            return true;
        }
        catch { return false; }
    }

    /// <summary>
    /// Create elbow fitting between two elements using Routing Preferences family.
    /// NewElbowFitting automatically uses the elbow family defined in the pipe/duct type's routing preferences.
    /// </summary>
    public static bool TryCreateElbow(Document doc, Element source, Element target)
    {
        var pair = GetBestOpenPair(source, target);
        if (pair is null) return false;

        try
        {
            doc.Create.NewElbowFitting(pair.Value.Source, pair.Value.Target);
            return true;
        }
        catch { return false; }
    }

    /// <summary>
    /// Create tee fitting: branch taps into main.
    /// NewTeeFitting requires: main connector 1, main connector 2, branch connector.
    /// The main connectors MUST be the two End connectors of the main MEPCurve.
    /// </summary>
    public static bool TryCreateTee(Document doc, MEPCurve branch, MEPCurve main)
    {
        // Branch: get best open connector pointing toward main
        var branchConnector = ConnectorUtils.GetOpenConnectors(branch)
            .OrderBy(c => c.Origin.DistanceTo(MepGeometryService.GetElementOrigin(main)))
            .FirstOrDefault();

        if (branchConnector is null)
            return false;

        // Main: MUST use the two END connectors (not just any two closest).
        // NewTeeFitting splits the main pipe at the branch point and creates a tee.
        var mainEndConnectors = ConnectorUtils.ToList(main)
            .Where(c => ConnectorUtils.IsUsable(c) && c.ConnectorType == ConnectorType.End)
            .ToList();

        // Fallback: if no End type connectors found, use any usable connectors
        if (mainEndConnectors.Count < 2)
        {
            mainEndConnectors = ConnectorUtils.ToList(main)
                .Where(ConnectorUtils.IsUsable)
                .ToList();
        }

        if (mainEndConnectors.Count < 2)
            return false;

        // Sort to get the two ends (endpoints of the main curve)
        if (main.Location is LocationCurve locCurve)
        {
            var p0 = locCurve.Curve.GetEndPoint(0);
            var p1 = locCurve.Curve.GetEndPoint(1);

            var c0 = mainEndConnectors.OrderBy(c => c.Origin.DistanceTo(p0)).First();
            var c1 = mainEndConnectors.Where(c => c != c0).OrderBy(c => c.Origin.DistanceTo(p1)).First();

            try
            {
                doc.Create.NewTeeFitting(c0, c1, branchConnector);
                return true;
            }
            catch { return false; }
        }

        // Fallback: just use first two
        try
        {
            doc.Create.NewTeeFitting(mainEndConnectors[0], mainEndConnectors[1], branchConnector);
            return true;
        }
        catch { return false; }
    }
}
#endif
