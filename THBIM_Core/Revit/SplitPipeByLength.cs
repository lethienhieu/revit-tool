using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using THBIM.Licensing;
namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallSP : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                return Result.Cancelled;

            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            IList<Reference> references;
            try
            {
                references = uidoc.Selection.PickObjects(ObjectType.Element, new PipeSelectionFilter(), "Select multiple pipes to split");
            }
            catch
            {
                return Result.Cancelled;
            }

            List<Pipe> pipes = references
                .Select(r => doc.GetElement(r))
                .OfType<Pipe>()
                .ToList();

            if (pipes.Count == 0)
            {
                TaskDialog.Show("Error", "No pipes selected.");
                return Result.Cancelled;
            }

            string input = Microsoft.VisualBasic.Interaction.InputBox("Enter split length (mm):", "Split Pipe", "1000");
            if (!double.TryParse(input, out double lengthMM) || lengthMM <= 0)
            {
                TaskDialog.Show("Error", "Invalid value.");
                return Result.Cancelled;
            }

            double splitLengthFT = UnitUtils.ConvertToInternalUnits(lengthMM, UnitTypeId.Millimeters);

            using (Transaction tx = new Transaction(doc, "Split Pipes with Union Fittings"))
            {
                tx.Start();

                foreach (Pipe pipe in pipes)
                {
                    // Split 1 pipe -> returns a list of new segments in order along the route
                    List<Pipe> segments = SplitPipe(doc, pipe, splitLengthFT);

                    // Regenerate before connecting connectors
                    doc.Regenerate();

                    // Create standard Revit unions (couplings) between adjacent segments of THIS pipe
                    AddUnionsBetweenSegments(doc, segments);
                }

                tx.Commit();
            }

            TaskDialog.Show("Completed", "Pipes have been split and Union Fittings (Couplings) created automatically.");
            return Result.Succeeded;
        }

        /// <summary>
        /// Split pipe into equal segments + 1 remainder (if any). Delete original pipe, return list of new Pipes IN ORDER.
        /// (Similar to old structure, still creates new pipes; fittings will be connected later via Union.)
        /// </summary>
        private List<Pipe> SplitPipe(Document doc, Pipe pipe, double splitLength)
        {
            LocationCurve loc = pipe.Location as LocationCurve;
            if (loc == null || !(loc.Curve is Line line)) return new List<Pipe>();

            XYZ start = line.GetEndPoint(0);
            XYZ end = line.GetEndPoint(1);

            // Ensure start -> end order is consistent along the curve direction
            if (start.X > end.X || (start.X == end.X && start.Y > end.Y))
            {
                XYZ temp = start;
                start = end;
                end = temp;
            }

            XYZ direction = (end - start).Normalize();
            double totalLength = line.Length;

            // Number of equal segments
            int numSegments = (int)(totalLength / splitLength);
            double remainder = totalLength - numSegments * splitLength;

            double tolerance = doc.Application.ShortCurveTolerance;
            List<Pipe> newPipes = new List<Pipe>();

            // Create equal segments
            for (int i = 0; i < numSegments; i++)
            {
                XYZ p1 = start + direction * (i * splitLength);
                XYZ p2 = start + direction * ((i + 1) * splitLength);
                Pipe newPipe = CreatePipe(doc, pipe, p1, p2);
                if (newPipe != null) newPipes.Add(newPipe);
            }

            // Final remainder segment (if > tolerance)
            if (remainder > tolerance)
            {
                XYZ p1 = start + direction * (numSegments * splitLength);
                XYZ p2 = end;
                Pipe lastPipe = CreatePipe(doc, pipe, p1, p2);
                if (lastPipe != null) newPipes.Add(lastPipe);
            }

            // Delete original pipe
            doc.Delete(pipe.Id);

            // Note: newPipes is in order along the path (from start -> end)
            return newPipes;
        }

        /// <summary>
        /// Create new pipe copying some basic properties from the original pipe.
        /// </summary>
        private Pipe CreatePipe(Document doc, Pipe original, XYZ p1, XYZ p2)
        {
            // Get SystemTypeId safely (fallback if MEPSystem is null)
            ElementId systemTypeId =
                original.MEPSystem?.GetTypeId()
                ?? original.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsElementId();

            if (systemTypeId == null || systemTypeId == ElementId.InvalidElementId)
                return null;

            ElementId typeId = original.GetTypeId();
            ElementId levelId = original.ReferenceLevel?.Id ?? original.LevelId;

            Pipe pipe = Pipe.Create(doc, systemTypeId, typeId, levelId, p1, p2);
            if (pipe == null) return null;

            // Keep diameter
            double diameter = original.Diameter;
            Parameter param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (param != null && !param.IsReadOnly)
                param.Set(diameter);

            return pipe;
        }

        /// <summary>
        /// Create UNION FITTING (Revit standard Coupling) between adjacent segments (of the same original pipe).
        /// Do not use manual FamilySymbol; Revit auto-selects family based on routing preference/size.
        /// </summary>
        private void AddUnionsBetweenSegments(Document doc, List<Pipe> segmentsInOrder)
        {
            if (segmentsInOrder == null || segmentsInOrder.Count < 2) return;

            // Ensure geometry is updated before connecting
            doc.Regenerate();

            for (int i = 0; i < segmentsInOrder.Count - 1; i++)
            {
                Pipe a = segmentsInOrder[i];
                Pipe b = segmentsInOrder[i + 1];

                // Get END connector of each segment, closest to the other segment
                Connector ca = GetClosestEndConnector(a, b);
                Connector cb = GetClosestEndConnector(b, a);

                if (ca == null || cb == null) continue;
                if (ca.IsConnected || cb.IsConnected) continue;

                try
                {
                    // Create STANDARD UNION FITTING – fitting will auto-align to connector direction (pipe direction)
                    var union = doc.Create.NewUnionFitting(ca, cb);
                    // No manual rotation needed
                }
                catch
                {
                    // Skip if a pair cannot be connected to avoid breaking the transaction
                }
            }
        }

        /// <summary>
        /// Get END type connector of pipeA closest to pipeB (based on distance to pipeB's endpoints).
        /// </summary>
        private Connector GetClosestEndConnector(Pipe pipeA, Pipe pipeB)
        {
            if (pipeA == null || pipeB == null) return null;

            // Get 2 endpoints of pipeB to compare distances
            XYZ bStart, bEnd;
            GetCurveEndpoints(pipeB, out bStart, out bEnd);

            Connector best = null;
            double bestDist = double.MaxValue;

            foreach (Connector c in pipeA.ConnectorManager.Connectors)
            {
                if (c.ConnectorType != ConnectorType.End) continue;

                double d1 = (bStart != null) ? c.Origin.DistanceTo(bStart) : double.MaxValue;
                double d2 = (bEnd != null) ? c.Origin.DistanceTo(bEnd) : double.MaxValue;
                double d = Math.Min(d1, d2);

                if (d < bestDist)
                {
                    bestDist = d;
                    best = c;
                }
            }
            return best;
        }

        /// <summary>
        /// Get start/end points of Pipe (Line-only – keep scope as original code).
        /// </summary>
        private void GetCurveEndpoints(Pipe p, out XYZ start, out XYZ end)
        {
            start = end = null;
            var lc = p.Location as LocationCurve;
            var line = lc?.Curve as Line;
            if (line != null)
            {
                start = line.GetEndPoint(0);
                end = line.GetEndPoint(1);
            }
        }

        private class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element e) => e is Pipe;
            public bool AllowReference(Reference r, XYZ p) => true;
        }
    }
}