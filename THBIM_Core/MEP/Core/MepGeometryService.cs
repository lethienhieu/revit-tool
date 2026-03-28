#if NET8_0_OR_GREATER
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace THBIM.MEP.Core;

internal static class MepGeometryService
{
    public static XYZ GetElementOrigin(Element element)
        => element.Location switch
        {
            LocationPoint lp => lp.Point,
            LocationCurve lc => lc.Curve.Evaluate(0.5, true),
            _ => element.get_BoundingBox(null)?.Transform?.Origin ?? XYZ.Zero
        };

    public static XYZ GetElementDirection(Element element)
    {
        if (element.Location is LocationCurve lc)
        {
            var c = lc.Curve;
            return (c.GetEndPoint(1) - c.GetEndPoint(0)).Normalize();
        }

        if (element is FamilyInstance fi && fi.FacingOrientation is not null)
        {
            return fi.FacingOrientation.Normalize();
        }

        return XYZ.BasisX;
    }

    public static void MoveElement(Document document, Element element, XYZ vector)
        => ElementTransformUtils.MoveElement(document, element.Id, vector);

    public static void RotateElement(Document document, Element element, Line axis, double angleRadians)
        => ElementTransformUtils.RotateElement(document, element.Id, axis, angleRadians);

    /// <summary>
    /// Get best rotation axis from connectors.
    /// For fittings: axis through the two end-connectors.
    /// Fallback: vertical axis through element origin.
    /// </summary>
    public static Line GetPreferredRotationAxis(Element element)
    {
        var connectors = ConnectorUtils.ToList(element).Where(ConnectorUtils.IsUsable).ToList();
        if (connectors.Count >= 2)
        {
            var a = connectors[0].Origin;
            var b = connectors[1].Origin;
            if (!a.IsAlmostEqualTo(b))
            {
                return Line.CreateBound(a, b);
            }
        }

        var origin = GetElementOrigin(element);
        return Line.CreateBound(origin, origin + XYZ.BasisZ);
    }

    public static bool TrySetLocationPoint(Element element, XYZ point)
    {
        if (element.Location is LocationPoint lp)
        {
            lp.Point = point;
            return true;
        }
        return false;
    }

    public static bool TryTranslateCurve(Element element, XYZ vector)
    {
        if (element.Location is not LocationCurve lc)
            return false;
        lc.Curve = lc.Curve.CreateTransformed(Transform.CreateTranslation(vector));
        return true;
    }

    /// <summary>
    /// Extend a MEPCurve from its open (unconnected) end.
    /// Supports: Pipe, Duct, Conduit, CableTray (straight Line curves).
    /// FlexPipe/FlexDuct: extends by adding distance to the last point direction.
    /// Returns true if extension succeeded.
    /// </summary>
    public static bool TryExtendOpenEnd(Document document, MEPCurve curve, double distanceFeet)
    {
        if (curve.Location is not LocationCurve locCurve)
            return false;

        var baseCurve = locCurve.Curve;

        // Straight curves (Pipe, Duct, Conduit, CableTray)
        if (baseCurve is Line line)
            return TryExtendLine(locCurve, line, curve, distanceFeet);

        // FlexPipe / FlexDuct: extend the last segment
        // FlexPipe/FlexDuct store a spline — we move the open endpoint outward
        return TryExtendFlexCurve(document, curve, distanceFeet);
    }

    private static bool TryExtendLine(LocationCurve locCurve, Line line, MEPCurve curve, double distanceFeet)
    {
        var p0 = line.GetEndPoint(0);
        var p1 = line.GetEndPoint(1);

        var connectors = ConnectorUtils.ToList(curve).Where(ConnectorUtils.IsUsable).ToList();
        var end0Connected = connectors.Any(c => c.IsConnected && c.Origin.IsAlmostEqualTo(p0, 0.01));
        var end1Connected = connectors.Any(c => c.IsConnected && c.Origin.IsAlmostEqualTo(p1, 0.01));

        if (!end0Connected && !end1Connected)
        {
            var dir = (p1 - p0).Normalize();
            locCurve.Curve = Line.CreateBound(p0, p1 + dir * distanceFeet);
            return true;
        }
        else if (!end1Connected)
        {
            var dir = (p1 - p0).Normalize();
            locCurve.Curve = Line.CreateBound(p0, p1 + dir * distanceFeet);
            return true;
        }
        else if (!end0Connected)
        {
            var dir = (p0 - p1).Normalize();
            locCurve.Curve = Line.CreateBound(p0 + dir * distanceFeet, p1);
            return true;
        }

        return false;
    }

    /// <summary>
    /// Extend a FlexPipe/FlexDuct by moving the open end outward.
    /// Since flex curves are splines, we move the element endpoint using ElementTransformUtils.
    /// </summary>
    private static bool TryExtendFlexCurve(Document document, MEPCurve curve, double distanceFeet)
    {
        // For FlexPipe/FlexDuct, find the open connector and extend in its direction
        var openConnector = ConnectorUtils.GetOpenConnectors(curve)
            .OrderByDescending(c => c.ConnectorType == ConnectorType.End ? 1 : 0)
            .FirstOrDefault();

        if (openConnector is null)
            return false;

        // Extend direction = connector's facing direction
        var dir = openConnector.CoordinateSystem.BasisZ.Normalize();
        var extensionVector = dir * distanceFeet;

        // For flex elements, we use the PointOnFlex approach:
        // Move the entire element endpoint. This is simpler but effective.
        try
        {
            // Get the curve and create a translated version
            if (curve.Location is LocationCurve locCurve)
            {
                var baseCurve = locCurve.Curve;
                var p0 = baseCurve.GetEndPoint(0);
                var p1 = baseCurve.GetEndPoint(1);

                var connectors = ConnectorUtils.ToList(curve).Where(ConnectorUtils.IsUsable).ToList();
                var end0Connected = connectors.Any(c => c.IsConnected && c.Origin.IsAlmostEqualTo(p0, 0.05));

                if (!end0Connected)
                {
                    // Extend from end 0
                    var newP0 = p0 + (p0 - p1).Normalize() * distanceFeet;
                    locCurve.Curve = Line.CreateBound(newP0, p1);
                }
                else
                {
                    // Extend from end 1
                    var newP1 = p1 + (p1 - p0).Normalize() * distanceFeet;
                    locCurve.Curve = Line.CreateBound(p0, newP1);
                }
                return true;
            }
        }
        catch { }

        return false;
    }

    /// <summary>Legacy nudge method. Prefer TryExtendOpenEnd.</summary>
    public static bool TryNudgeOpenEnd(Document document, MEPCurve sourceCurve, XYZ direction, double distanceMm)
    {
        var distanceFeet = UnitUtils.ConvertToInternalUnits(distanceMm, UnitTypeId.Millimeters);
        return TryExtendOpenEnd(document, sourceCurve, distanceFeet);
    }

    /// <summary>Convert mm to internal units (feet).</summary>
    public static double MmToFeet(double mm)
        => UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);

    /// <summary>Convert internal units (feet) to mm.</summary>
    public static double FeetToMm(double feet)
        => UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Millimeters);

    /// <summary>
    /// Create a pipe/duct stub from an open connector on a fitting/accessory.
    /// Automatically determines pipe or duct type from the connector's domain.
    /// </summary>
    public static MEPCurve? TryCreateRunFromConnector(Document doc, Connector connector, XYZ start, XYZ end)
    {
        try
        {
            var domain = connector.Domain;
            var levelId = ElementId.InvalidElementId;

            // Try to get level from the connector's owner
            if (connector.Owner is MEPCurve ownerCurve)
                levelId = ownerCurve.ReferenceLevel.Id;
            else if (connector.Owner is FamilyInstance fi)
                levelId = fi.LevelId;

            if (levelId == ElementId.InvalidElementId)
            {
                // Fallback: get any level
                var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault();
                if (level is not null) levelId = level.Id;
            }

            if (domain == Domain.DomainPiping || domain == Domain.DomainUndefined)
            {
                var sysType = connector.MEPSystem?.GetTypeId()
                    ?? new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType)).ToElementIds().FirstOrDefault();
                var pipeType = new FilteredElementCollector(doc).OfClass(typeof(PipeType)).FirstElementId();
                if (sysType is not null && pipeType != ElementId.InvalidElementId)
                {
                    var pipe = Pipe.Create(doc, sysType, pipeType, levelId, start, end);
                    // Copy diameter from connector
                    if (connector.Shape == ConnectorProfileType.Round)
                    {
                        var d = connector.Radius * 2.0;
                        pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(d);
                    }
                    // Try to connect
                    try { connector.ConnectTo(ConnectorUtils.GetClosestConnector(pipe, start)!); } catch { }
                    return pipe;
                }
            }

            if (domain == Domain.DomainHvac)
            {
                var sysType = connector.MEPSystem?.GetTypeId()
                    ?? new FilteredElementCollector(doc).OfClass(typeof(MechanicalSystemType)).ToElementIds().FirstOrDefault();
                var ductType = new FilteredElementCollector(doc).OfClass(typeof(DuctType)).FirstElementId();
                if (sysType is not null && ductType != ElementId.InvalidElementId)
                {
                    var duct = Duct.Create(doc, sysType, ductType, levelId, start, end);
                    // Copy size
                    if (connector.Shape == ConnectorProfileType.Round)
                        duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(connector.Radius * 2.0);
                    else if (connector.Shape == ConnectorProfileType.Rectangular)
                    {
                        duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.Set(connector.Width);
                        duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.Set(connector.Height);
                    }
                    try { connector.ConnectTo(ConnectorUtils.GetClosestConnector(duct, start)!); } catch { }
                    return duct;
                }
            }
        }
        catch { }
        return null;
    }
}
#endif
