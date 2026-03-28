#if NET8_0_OR_GREATER
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using THBIM.MEP.Core;

namespace THBIM.MEP.Commands;

// ═══════════════════════════════════════════════════════════════
//  LT CREATE
// ═══════════════════════════════════════════════════════════════

#region Bloom

/// <summary>
/// Bloom: extend MEPCurve from open end, OR create short pipe stub
/// from any open connector on fittings/accessories/FamilyInstance.
/// </summary>
[Transaction(TransactionMode.Manual)]
public sealed class BloomCommand : CommandBase
{
    private const double DefaultExtendMm = 500;

    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var distanceFeet = MepGeometryService.MmToFeet(DefaultExtendMm);

        while (true)
        {
            Element picked;
            try { picked = MepSelection.PickMepElement(uidoc, "Pick element to bloom (ESC to finish)."); }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { break; }

            using var tx = new Transaction(doc, "Bloom");
            tx.Start();

            if (picked is MEPCurve curve)
            {
                // Extend existing pipe/duct from open end
                MepGeometryService.TryExtendOpenEnd(doc, curve, distanceFeet);
            }
            else
            {
                // Fitting/Accessory/FamilyInstance: create pipe stubs from ALL open connectors
                var openConnectors = ConnectorUtils.GetOpenConnectors(picked);
                foreach (var conn in openConnectors)
                {
                    var dir = conn.CoordinateSystem.BasisZ.Normalize();
                    var start = conn.Origin;
                    var end = start + dir * distanceFeet;
                    MepGeometryService.TryCreateRunFromConnector(doc, conn, start, end);
                }
            }

            tx.Commit();
        }
        return Result.Succeeded;
    }
}

#endregion

#region Tap

[Transaction(TransactionMode.Manual)]
public sealed class TapCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var branch = MepSelection.PickMepElement(uidoc, "Pick branch MEPCurve.") as MEPCurve;
        var main = MepSelection.PickMepElement(uidoc, "Pick main MEPCurve.") as MEPCurve;
        if (branch is null || main is null) return Result.Cancelled;

        using var tx = new Transaction(doc, "Tap");
        tx.Start();
        if (!ConnectorRoutingService.TryDirectConnect(branch, main))
            if (!ConnectorRoutingService.TryCreateTee(doc, branch, main))
                ConnectorRoutingService.TryCreateElbow(doc, branch, main);
        tx.Commit();
        return Result.Succeeded;
    }
}

#endregion

#region Directional Offset (Elbow Up/Down/Left/Right/Down45)

public abstract class DirectionalOffsetCommandBase : CommandBase
{
    protected abstract string Title { get; }
    protected abstract ElbowDirectionKind DirectionKind { get; }
    protected virtual double DistanceMm => 500;

    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var view = doc.ActiveView;
        var picked = MepSelection.PickMepElementWithPoint(uidoc, $"Pick pipe/duct near end for {Title}.");

        using var tx = new Transaction(doc, Title);
        tx.Start();
        DirectionalOffsetService.TryCreateElbowRun(view, picked.Element, picked.PickPoint, DirectionKind, DistanceMm);
        tx.Commit();
        return Result.Succeeded;
    }
}

[Transaction(TransactionMode.Manual)] public sealed class ElbowUpCommand : DirectionalOffsetCommandBase { protected override string Title => "Elbow Up"; protected override ElbowDirectionKind DirectionKind => ElbowDirectionKind.Up; }
[Transaction(TransactionMode.Manual)] public sealed class ElbowDownCommand : DirectionalOffsetCommandBase { protected override string Title => "Elbow Down"; protected override ElbowDirectionKind DirectionKind => ElbowDirectionKind.Down; }
[Transaction(TransactionMode.Manual)] public sealed class ElbowLeftCommand : DirectionalOffsetCommandBase { protected override string Title => "Elbow Left"; protected override ElbowDirectionKind DirectionKind => ElbowDirectionKind.Left; }
[Transaction(TransactionMode.Manual)] public sealed class ElbowRightCommand : DirectionalOffsetCommandBase { protected override string Title => "Elbow Right"; protected override ElbowDirectionKind DirectionKind => ElbowDirectionKind.Right; }
[Transaction(TransactionMode.Manual)] public sealed class ElbowDown45Command : DirectionalOffsetCommandBase { protected override string Title => "Elbow Down 45"; protected override ElbowDirectionKind DirectionKind => ElbowDirectionKind.Down45; }

#endregion

// ═══════════════════════════════════════════════════════════════
//  LT MODIFY
// ═══════════════════════════════════════════════════════════════

#region Flip Multiple

[Transaction(TransactionMode.Manual)]
public sealed class FlipMultipleCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;

        while (true)
        {
            Element picked;
            try { picked = MepSelection.PickMepElement(uidoc, "Pick instance to flip (ESC to finish)."); }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { break; }
            if (picked is not FamilyInstance fi) continue;

            using var tx = new Transaction(doc, "Flip");
            tx.Start();
            if (fi.CanFlipFacing) fi.flipFacing();
            else if (fi.CanFlipHand) fi.flipHand();
            tx.Commit();
        }
        return Result.Succeeded;
    }
}

#endregion

#region Rotate Fittings (SplitButton: 30/45/60/75/90)

public abstract class RotateFittingCommandBase : CommandBase
{
    protected abstract double AngleDegrees { get; }

    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var angleRad = AngleDegrees * Math.PI / 180.0;

        while (true)
        {
            Element picked;
            try { picked = MepSelection.PickMepElement(uidoc, $"Pick fitting to rotate {AngleDegrees}\u00B0 (ESC to finish)."); }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { break; }
            if (picked is not FamilyInstance fi) continue;

            using var tx = new Transaction(doc, $"Rotate {AngleDegrees}\u00B0");
            tx.Start();

            // Collect connected pairs BEFORE disconnecting
            var connectedPairs = new List<(Connector FittingConn, Connector OtherConn)>();
            foreach (var conn in ConnectorUtils.ToList(fi))
            {
                if (!ConnectorUtils.IsUsable(conn) || !conn.IsConnected) continue;
                foreach (Connector other in conn.AllRefs)
                {
                    if (other.Owner.Id != fi.Id && conn.IsConnectedTo(other))
                        connectedPairs.Add((conn, other));
                }
            }

            // Disconnect all
            ConnectorUtils.TryDisconnectAll(fi);
            doc.Regenerate();

            // Determine rotation axis from the first connected pipe direction
            var pivot = MepGeometryService.GetElementOrigin(fi);
            var axisDir = XYZ.BasisZ;
            if (connectedPairs.Count > 0)
            {
                var firstOther = connectedPairs[0].OtherConn;
                pivot = connectedPairs[0].FittingConn.Origin;
                axisDir = firstOther.CoordinateSystem.BasisZ.Normalize();
                if (axisDir.IsZeroLength()) axisDir = XYZ.BasisZ;
            }

            // Rotate
            var axis = Line.CreateBound(pivot, pivot + axisDir);
            MepGeometryService.RotateElement(doc, fi, axis, angleRad);
            doc.Regenerate();

            // Reconnect: find closest open connector pairs and connect
            foreach (var (_, otherConn) in connectedPairs)
            {
                var bestFitConn = ConnectorUtils.GetClosestOpenConnector(fi, otherConn.Origin);
                if (bestFitConn is not null)
                {
                    try { bestFitConn.ConnectTo(otherConn); } catch { }
                }
            }

            tx.Commit();
        }
        return Result.Succeeded;
    }
}

[Transaction(TransactionMode.Manual)] public sealed class RotateFitting30Command : RotateFittingCommandBase { protected override double AngleDegrees => 30; }
[Transaction(TransactionMode.Manual)] public sealed class RotateFitting45Command : RotateFittingCommandBase { protected override double AngleDegrees => 45; }
[Transaction(TransactionMode.Manual)] public sealed class RotateFitting60Command : RotateFittingCommandBase { protected override double AngleDegrees => 60; }
[Transaction(TransactionMode.Manual)] public sealed class RotateFitting75Command : RotateFittingCommandBase { protected override double AngleDegrees => 75; }
[Transaction(TransactionMode.Manual)] public sealed class RotateFitting90Command : RotateFittingCommandBase { protected override double AngleDegrees => 90; }

#endregion

#region Delete System

/// <summary>
/// Disconnect entire piping/duct network from picked element.
/// Pure BFS walk via physical connectors — robust for Pipe, Duct, Fitting, etc.
/// </summary>
[Transaction(TransactionMode.Manual)]
public sealed class DeleteSystemCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var picked = MepSelection.PickMepElement(uidoc, "Pick MEP element to disconnect its network.");

        // BFS walk via physical connectors
        var visited = new HashSet<ElementId>();
        var toDisconnect = new List<Element>();
        var queue = new Queue<Element>();

        visited.Add(picked.Id);
        queue.Enqueue(picked);
        toDisconnect.Add(picked);

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            foreach (var conn in ConnectorUtils.ToList(current))
            {
                if (!ConnectorUtils.IsUsable(conn) || !conn.IsConnected) continue;
                try
                {
                    foreach (Connector other in conn.AllRefs)
                    {
                        if (other.ConnectorType == ConnectorType.Logical) continue;
                        var owner = other.Owner;
                        if (owner is null || owner.Id == current.Id) continue;
                        if (!visited.Add(owner.Id)) continue;
                        // Only traverse MEP elements
                        if (owner is MEPCurve or FamilyInstance or FabricationPart)
                        {
                            toDisconnect.Add(owner);
                            queue.Enqueue(owner);
                        }
                    }
                }
                catch { /* AllRefs can throw on broken connectors */ }
            }
        }

        using var tx = new Transaction(doc, "Delete System");
        tx.Start();

        // Collect all MEPSystem IDs before disconnecting
        var systemIds = new HashSet<ElementId>();
        foreach (var elem in toDisconnect)
        {
            if (elem is MEPCurve curve && curve.MEPSystem is not null)
                systemIds.Add(curve.MEPSystem.Id);
            foreach (var conn in ConnectorUtils.ToList(elem))
            {
                try { if (conn.MEPSystem is MEPSystem s) systemIds.Add(s.Id); } catch { }
            }
        }

        // Disconnect all connectors
        foreach (var elem in toDisconnect)
            ConnectorUtils.TryDisconnectAll(elem);

        // Delete the MEPSystem elements to fully remove system assignment
        foreach (var sysId in systemIds)
        {
            try { doc.Delete(sysId); } catch { }
        }

        tx.Commit();

        return Result.Succeeded;
    }
}

#endregion

#region Disconnect

[Transaction(TransactionMode.Manual)]
public sealed class DisconnectCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;

        while (true)
        {
            Element picked;
            try { picked = MepSelection.PickMepElement(uidoc, "Pick element to disconnect (ESC to finish)."); }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { break; }

            using var tx = new Transaction(doc, "Disconnect");
            tx.Start();
            ConnectorUtils.TryDisconnectAll(picked);
            tx.Commit();
        }
        return Result.Succeeded;
    }
}

#endregion

#region Move Connect

[Transaction(TransactionMode.Manual)]
public sealed class MoveConnectCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var target = MepSelection.PickMepElement(uidoc, "Pick fixed/reference element.");
        var source = MepSelection.PickMepElement(uidoc, "Pick follower element (will move).");
        var pair = ConnectorRoutingService.GetBestOpenPair(source, target);
        if (pair is null) return Result.Cancelled;

        using var tx = new Transaction(doc, "Move Connect");
        tx.Start();
        MepGeometryService.MoveElement(doc, source, pair.Value.Target.Origin - pair.Value.Source.Origin);
        doc.Regenerate();
        if (!ConnectorRoutingService.TryDirectConnect(source, target))
            ConnectorRoutingService.TryCreateElbow(doc, source, target);
        tx.Commit();
        return Result.Succeeded;
    }
}

#endregion

#region Move Connect Align

[Transaction(TransactionMode.Manual)]
public sealed class MoveConnectAlignCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var target = MepSelection.PickMepElement(uidoc, "Pick fixed/reference element.");
        var source = MepSelection.PickMepElement(uidoc, "Pick follower element (will align + move).");

        using var tx = new Transaction(doc, "Mo-Co Align");
        tx.Start();

        var srcDir = MepGeometryService.GetElementDirection(source);
        var tgtDir = MepGeometryService.GetElementDirection(target);
        var cross = srcDir.CrossProduct(tgtDir);
        var angle = srcDir.AngleTo(tgtDir);
        if (angle > 1e-6 && angle < Math.PI - 1e-6)
        {
            var origin = MepGeometryService.GetElementOrigin(source);
            var axis = !cross.IsZeroLength()
                ? Line.CreateBound(origin, origin + cross.Normalize())
                : Line.CreateBound(origin, origin + XYZ.BasisZ);
            MepGeometryService.RotateElement(doc, source, axis, angle);
        }
        doc.Regenerate();

        var pair = ConnectorRoutingService.GetBestOpenPair(source, target);
        if (pair is null) { tx.RollBack(); return Result.Cancelled; }
        MepGeometryService.MoveElement(doc, source, pair.Value.Target.Origin - pair.Value.Source.Origin);
        doc.Regenerate();
        if (!ConnectorRoutingService.TryDirectConnect(source, target))
            ConnectorRoutingService.TryCreateElbow(doc, source, target);
        tx.Commit();
        return Result.Succeeded;
    }
}

#endregion

// ═══════════════════════════════════════════════════════════════
//  LT ALIGN
// ═══════════════════════════════════════════════════════════════

#region Align in 3D

[Transaction(TransactionMode.Manual)]
public sealed class AlignIn3DCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var target = MepSelection.PickMepElement(uidoc, "Pick fixed/reference element.");
        var source = MepSelection.PickMepElement(uidoc, "Pick follower element (will move).");

        using var tx = new Transaction(doc, "Align In 3D");
        tx.Start();
        MepGeometryService.MoveElement(doc, source,
            MepGeometryService.GetElementOrigin(target) - MepGeometryService.GetElementOrigin(source));
        tx.Commit();
        return Result.Succeeded;
    }
}

#endregion

#region Align Branch

/// <summary>
/// Align Branch: move branch onto main's centerline plane.
/// ONLY moves — does NOT extend or shorten any pipe.
/// </summary>
[Transaction(TransactionMode.Manual)]
public sealed class AlignBranchCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var main = MepSelection.PickMepElement(uidoc, "Pick main/reference element.");
        var branch = MepSelection.PickMepElement(uidoc, "Pick branch element (will move only, no extend).");

        // Project branch onto main's centerline plane
        // Only perpendicular movement — no stretching
        var mainDir = MepGeometryService.GetElementDirection(main);
        var mainOrigin = MepGeometryService.GetElementOrigin(main);
        var branchOrigin = MepGeometryService.GetElementOrigin(branch);

        // Vector from branch to main, projected perpendicular to branch direction
        var branchDir = MepGeometryService.GetElementDirection(branch);
        var delta = mainOrigin - branchOrigin;
        // Remove component along branch direction — only move perpendicular
        var moveVector = delta - branchDir.Multiply(delta.DotProduct(branchDir));

        using var tx = new Transaction(doc, "Align Branch");
        tx.Start();
        MepGeometryService.MoveElement(doc, branch, moveVector);
        tx.Commit();
        return Result.Succeeded;
    }
}

#endregion

#region Align Branch+

/// <summary>
/// Align Branch+: move branch onto main's plane, then try to connect.
/// Does NOT extend or shorten pipes — only moves and attempts fitting creation.
/// </summary>
[Transaction(TransactionMode.Manual)]
public sealed class AlignBranchPlusCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var main = MepSelection.PickMepElement(uidoc, "Pick main/reference element.");
        var branch = MepSelection.PickMepElement(uidoc, "Pick branch element (will move + connect).");

        using var tx = new Transaction(doc, "Align Branch+");
        tx.Start();

        var branchDir = MepGeometryService.GetElementDirection(branch);
        var delta = MepGeometryService.GetElementOrigin(main) - MepGeometryService.GetElementOrigin(branch);
        var moveVector = delta - branchDir.Multiply(delta.DotProduct(branchDir));
        MepGeometryService.MoveElement(doc, branch, moveVector);
        doc.Regenerate();

        // Try connect — no pipe extension
        if (!ConnectorRoutingService.TryDirectConnect(branch, main))
        {
            if (branch is MEPCurve bc && main is MEPCurve mc)
            {
                if (!ConnectorRoutingService.TryCreateTee(doc, bc, mc))
                    ConnectorRoutingService.TryCreateElbow(doc, bc, mc);
            }
        }
        tx.Commit();
        return Result.Succeeded;
    }
}

#endregion

// ═══════════════════════════════════════════════════════════════
//  LT TOOLS
// ═══════════════════════════════════════════════════════════════

#region Sum Parameter

[Transaction(TransactionMode.Manual)]
public sealed class SumParameterCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;

        var refs = uidoc.Selection.PickObjects(ObjectType.Element,
            new ElementPickFilter(MepSelection.IsSupportedElement),
            "Select elements to sum, then Finish.");
        var selected = refs.Select(r => doc.GetElement(r)).Where(e => e is not null).ToList();
        if (selected.Count == 0) return Result.Cancelled;

        var candidates = selected
            .SelectMany(e => e!.Parameters.Cast<Parameter>())
            .Where(p => p.StorageType == StorageType.Double && p.HasValue)
            .GroupBy(p => p.Definition.Name)
            .OrderBy(g => g.Key)
            .ToList();
        if (candidates.Count == 0) return Result.Succeeded;

        var lines = new List<string> { $"Selection: {selected.Count} element(s)", "" };
        foreach (var g in candidates)
        {
            var sum = g.Sum(p => p.AsDouble());
            lines.Add($"{g.Key}: {FormatValue(g.First(), sum, doc)}");
        }

        var dlg = new TaskDialog("Sum Param.") { MainContent = string.Join("\n", lines) };
        dlg.Show();
        return Result.Succeeded;
    }

    private static string FormatValue(Parameter param, double value, Document doc)
    {
        try
        {
            var spec = param.Definition.GetDataType();
            if (spec != null && spec != SpecTypeId.Custom)
                return UnitFormatUtils.Format(doc.GetUnits(), spec, value, false);
        }
        catch { }
        return $"{value:0.###}";
    }
}

#endregion

#region Section

[Transaction(TransactionMode.Manual)]
public sealed class SectionCommand : CommandBase
{
    protected override Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var uidoc = GetUiDoc(commandData);
        var doc = uidoc.Document;
        var element = MepSelection.PickMepElement(uidoc, "Pick element to section.");
        var bbox = element.get_BoundingBox(doc.ActiveView) ?? element.get_BoundingBox(null);
        if (bbox is null) return Result.Cancelled;

        var center = (bbox.Min + bbox.Max) * 0.5;
        var size = bbox.Max - bbox.Min;
        var sectionType = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
            .FirstOrDefault(x => x.ViewFamily == ViewFamily.Section);
        if (sectionType is null) return Result.Failed;

        using var tx = new Transaction(doc, "Create Section");
        tx.Start();
        var dir = MepGeometryService.GetElementDirection(element);
        var bz = dir;
        var by = Math.Abs(bz.DotProduct(XYZ.BasisZ)) > 0.95 ? doc.ActiveView.RightDirection.Normalize() : XYZ.BasisZ;
        var bx = by.CrossProduct(bz).Normalize();
        if (bx.IsZeroLength()) bx = doc.ActiveView.RightDirection.Normalize();
        by = bz.CrossProduct(bx).Normalize();
        var t = Transform.Identity; t.Origin = center; t.BasisX = bx; t.BasisY = by; t.BasisZ = bz;

        var hw = Math.Max(size.X, size.Y) * 0.75 + 1.0;
        var hh = Math.Max(size.Z * 0.75, 1.0);
        var hd = Math.Max(size.GetLength() * 0.5, 1.0);
        var section = ViewSection.CreateSection(doc, sectionType.Id,
            new BoundingBoxXYZ { Transform = t, Min = new XYZ(-hw, -hh, -hd), Max = new XYZ(hw, hh, hd) });
        section.Name = $"MEP Section {element.Id.Value}";
        tx.Commit();

        uidoc.ActiveView = section;
        return Result.Succeeded;
    }
}

#endregion
#endif
