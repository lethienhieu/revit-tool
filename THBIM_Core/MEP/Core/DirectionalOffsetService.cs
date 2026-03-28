#if NET8_0_OR_GREATER
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace THBIM.MEP.Core;

public enum ElbowDirectionKind
{
    Up,
    Down,
    Left,
    Right,
    Down45
}

internal static class DirectionalOffsetService
{
    public static bool TryCreateElbowRun(View view, Element element, XYZ pickPoint, ElbowDirectionKind kind, double distanceMm)
    {
        if (element is not MEPCurve curve)
            return false;

        // Must pick an OPEN (unconnected) connector near the pick point
        var sourceConnector = ConnectorUtils.GetOpenConnectors(curve)
            .OrderBy(c => c.Origin.DistanceTo(pickPoint))
            .FirstOrDefault();

        // Fallback: if no open connector, try any usable connector near pick point
        sourceConnector ??= ConnectorUtils.ToList(curve)
            .Where(ConnectorUtils.IsUsable)
            .OrderBy(c => c.Origin.DistanceTo(pickPoint))
            .FirstOrDefault();

        if (sourceConnector is null)
            return false;

        // Direction based on pipe's local coordinate frame
        var pipeDir = MepGeometryService.GetElementDirection(curve);
        var connDir = sourceConnector.CoordinateSystem?.BasisZ.Normalize();
        var forward = connDir is not null && !connDir.IsZeroLength() ? connDir : pipeDir;
        if (forward.DotProduct(pipeDir) < 0) forward = forward.Negate();

        // Build local frame from pipe direction
        var globalUp = XYZ.BasisZ;
        var upCandidate = Math.Abs(forward.DotProduct(globalUp)) < 0.98 ? globalUp : view.UpDirection.Normalize();
        var right = forward.CrossProduct(upCandidate).Normalize();
        if (right.IsZeroLength()) right = view.RightDirection.Normalize();
        var up = right.CrossProduct(forward).Normalize();
        if (up.IsZeroLength()) up = globalUp;

        var runDirection = kind switch
        {
            ElbowDirectionKind.Up => up,
            ElbowDirectionKind.Down => up.Negate(),
            ElbowDirectionKind.Left => right.Negate(),
            ElbowDirectionKind.Right => right,
            ElbowDirectionKind.Down45 => (up.Negate() + forward).Normalize(),
            _ => right
        };

        var distance = MepGeometryService.MmToFeet(distanceMm);
        var start = sourceConnector.Origin;
        var end = start + runDirection.Normalize() * distance;

        var newRun = CreateRunFromConnector(curve, sourceConnector, start, end);
        if (newRun is null)
            return false;

        // Try to create elbow fitting (uses Routing Preferences family)
        // then fallback to direct connect
        return ConnectorRoutingService.TryCreateElbow(curve.Document, curve, newRun)
               || ConnectorRoutingService.TryDirectConnect(curve, newRun);
    }

    /// <summary>
    /// Create a new MEPCurve run from a source connector, copying size/diameter.
    /// Uses the source connector's properties (Radius/Width/Height) to match the new run.
    /// </summary>
    private static MEPCurve? CreateRunFromConnector(MEPCurve source, Connector sourceConnector, XYZ start, XYZ end)
    {
        var doc = source.Document;
        MEPCurve? newRun = null;

        if (source is Pipe pipe)
        {
            newRun = CreatePipe(doc, pipe, start, end);
            if (newRun is Pipe newPipe)
            {
                // Copy diameter from source connector
                CopyPipeSize(pipe, newPipe, sourceConnector);
            }
        }
        else if (source is Duct duct)
        {
            newRun = CreateDuct(doc, duct, start, end);
            if (newRun is Duct newDuct)
            {
                // Copy width/height/diameter from source
                CopyDuctSize(duct, newDuct, sourceConnector);
            }
        }
        else if (source is FlexPipe flexPipe)
        {
            newRun = CreateFlexPipe(doc, flexPipe, start, end);
            if (newRun is not null)
                CopyMepCurveSize(source, newRun, sourceConnector);
        }
        else if (source is FlexDuct flexDuct)
        {
            newRun = CreateFlexDuct(doc, flexDuct, start, end);
            if (newRun is not null)
                CopyMepCurveSize(source, newRun, sourceConnector);
        }
        else if (source is Conduit conduit)
        {
            newRun = CreateConduit(doc, conduit, start, end);
            if (newRun is not null)
                CopyConduitSize(conduit, newRun);
        }
        else if (source is CableTray cableTray)
        {
            newRun = CreateCableTray(doc, cableTray, start, end);
            if (newRun is not null)
                CopyCableTraySize(cableTray, newRun);
        }

        return newRun;
    }

    #region Pipe creation + size copy

    private static Pipe? CreatePipe(Document doc, Pipe source, XYZ start, XYZ end)
    {
        // Strategy 1: use existing MEPSystem
        var systemTypeId = source.MEPSystem?.GetTypeId();
        if (systemTypeId is not null)
        {
            try { return Pipe.Create(doc, systemTypeId, source.PipeType.Id, source.ReferenceLevel.Id, start, end); }
            catch { }
        }

        // Strategy 2: find any PipingSystemType
        try
        {
            var sysTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .ToElementIds().FirstOrDefault();
            if (sysTypes is not null)
                return Pipe.Create(doc, sysTypes, source.PipeType.Id, source.ReferenceLevel.Id, start, end);
        }
        catch { }

        return null;
    }

    private static void CopyPipeSize(Pipe source, Pipe target, Connector sourceConnector)
    {
        try
        {
            // Use connector Radius (internal units = feet) to set diameter
            if (sourceConnector.Shape == ConnectorProfileType.Round)
            {
                var diameter = sourceConnector.Radius * 2.0; // feet
                target.LookupParameter("Diameter")?.Set(diameter);
                // Also try built-in parameter
                var dParam = target.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                dParam?.Set(diameter);
            }
        }
        catch { }

        // Fallback: copy from source pipe directly
        try
        {
            var srcDiam = source.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            var tgtDiam = target.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (srcDiam is not null && tgtDiam is not null)
                tgtDiam.Set(srcDiam.AsDouble());
        }
        catch { }
    }

    #endregion

    #region Duct creation + size copy

    private static Duct? CreateDuct(Document doc, Duct source, XYZ start, XYZ end)
    {
        var systemTypeId = source.MEPSystem?.GetTypeId();
        if (systemTypeId is not null)
        {
            try { return Duct.Create(doc, systemTypeId, source.DuctType.Id, source.ReferenceLevel.Id, start, end); }
            catch { }
        }

        try
        {
            var sysTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(MechanicalSystemType))
                .ToElementIds().FirstOrDefault();
            if (sysTypes is not null)
                return Duct.Create(doc, sysTypes, source.DuctType.Id, source.ReferenceLevel.Id, start, end);
        }
        catch { }

        return null;
    }

    private static void CopyDuctSize(Duct source, Duct target, Connector sourceConnector)
    {
        try
        {
            if (sourceConnector.Shape == ConnectorProfileType.Round)
            {
                // Round duct: copy diameter
                var srcParam = source.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                var tgtParam = target.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (srcParam is not null && tgtParam is not null)
                    tgtParam.Set(srcParam.AsDouble());
            }
            else if (sourceConnector.Shape == ConnectorProfileType.Rectangular
                     || sourceConnector.Shape == ConnectorProfileType.Oval)
            {
                // Rectangular/Oval duct: copy width + height
                var srcW = source.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var srcH = source.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                var tgtW = target.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var tgtH = target.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (srcW is not null && tgtW is not null) tgtW.Set(srcW.AsDouble());
                if (srcH is not null && tgtH is not null) tgtH.Set(srcH.AsDouble());
            }
        }
        catch { }
    }

    #endregion

    #region FlexPipe / FlexDuct / Conduit / CableTray

    private static FlexPipe? CreateFlexPipe(Document doc, FlexPipe source, XYZ start, XYZ end)
    {
        try
        {
            var systemTypeId = source.MEPSystem?.GetTypeId();
            if (systemTypeId is null)
            {
                var sysType = new FilteredElementCollector(doc)
                    .OfClass(typeof(PipingSystemType)).ToElementIds().FirstOrDefault();
                systemTypeId = sysType;
            }
            if (systemTypeId is null) return null;

            var pts = new List<XYZ> { start, end };
            return FlexPipe.Create(doc, systemTypeId, source.FlexPipeType.Id, source.ReferenceLevel.Id, pts);
        }
        catch { return null; }
    }

    private static FlexDuct? CreateFlexDuct(Document doc, FlexDuct source, XYZ start, XYZ end)
    {
        try
        {
            var systemTypeId = source.MEPSystem?.GetTypeId();
            if (systemTypeId is null)
            {
                var sysType = new FilteredElementCollector(doc)
                    .OfClass(typeof(MechanicalSystemType)).ToElementIds().FirstOrDefault();
                systemTypeId = sysType;
            }
            if (systemTypeId is null) return null;

            var pts = new List<XYZ> { start, end };
            return FlexDuct.Create(doc, systemTypeId, source.FlexDuctType.Id, source.ReferenceLevel.Id, pts);
        }
        catch { return null; }
    }

    private static Conduit? CreateConduit(Document doc, Conduit source, XYZ start, XYZ end)
    {
        try
        {
            var typeId = source.GetTypeId();
            return Conduit.Create(doc, typeId, start, end, source.ReferenceLevel.Id);
        }
        catch { return null; }
    }

    private static CableTray? CreateCableTray(Document doc, CableTray source, XYZ start, XYZ end)
    {
        try
        {
            var typeId = source.GetTypeId();
            return CableTray.Create(doc, typeId, start, end, source.ReferenceLevel.Id);
        }
        catch { return null; }
    }

    #endregion

    #region Size copy for Conduit / CableTray / generic MEPCurve

    /// <summary>Copy size params generically using connector shape (works for FlexPipe/FlexDuct).</summary>
    private static void CopyMepCurveSize(MEPCurve source, MEPCurve target, Connector sourceConnector)
    {
        try
        {
            if (sourceConnector.Shape == ConnectorProfileType.Round)
            {
                var srcD = source.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                var tgtD = target.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (srcD is not null && tgtD is not null) tgtD.Set(srcD.AsDouble());

                // Also try pipe diameter param
                var srcPD = source.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                var tgtPD = target.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (srcPD is not null && tgtPD is not null) tgtPD.Set(srcPD.AsDouble());
            }
            else
            {
                var srcW = source.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var srcH = source.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                var tgtW = target.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var tgtH = target.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (srcW is not null && tgtW is not null) tgtW.Set(srcW.AsDouble());
                if (srcH is not null && tgtH is not null) tgtH.Set(srcH.AsDouble());
            }
        }
        catch { }
    }

    /// <summary>Copy conduit diameter/size from source to target.</summary>
    private static void CopyConduitSize(Conduit source, MEPCurve target)
    {
        try
        {
            // Conduit uses RBS_CONDUIT_DIAMETER_PARAM or RBS_CURVE_DIAMETER_PARAM
            var srcD = source.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                       ?? source.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            var tgtD = target.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                       ?? target.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (srcD is not null && tgtD is not null && !tgtD.IsReadOnly)
                tgtD.Set(srcD.AsDouble());

            // Also copy inner diameter if exists
            var srcInner = source.get_Parameter(BuiltInParameter.RBS_CONDUIT_INNER_DIAM_PARAM);
            var tgtInner = target.get_Parameter(BuiltInParameter.RBS_CONDUIT_INNER_DIAM_PARAM);
            if (srcInner is not null && tgtInner is not null && !tgtInner.IsReadOnly)
                tgtInner.Set(srcInner.AsDouble());
        }
        catch { }
    }

    /// <summary>Copy cable tray width/height from source to target.</summary>
    private static void CopyCableTraySize(CableTray source, MEPCurve target)
    {
        try
        {
            // Cable tray uses RBS_CABLETRAY_WIDTH_PARAM and RBS_CABLETRAY_HEIGHT_PARAM
            var srcW = source.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                       ?? source.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            var srcH = source.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                       ?? source.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            var tgtW = target.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                       ?? target.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            var tgtH = target.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                       ?? target.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            if (srcW is not null && tgtW is not null && !tgtW.IsReadOnly) tgtW.Set(srcW.AsDouble());
            if (srcH is not null && tgtH is not null && !tgtH.IsReadOnly) tgtH.Set(srcH.AsDouble());
        }
        catch { }
    }

    #endregion
}
#endif
