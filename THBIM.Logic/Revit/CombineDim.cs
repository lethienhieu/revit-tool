using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CombineDim : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Select Dimensions to combine
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new DimensionSelectionFilter(),
                    "Select dimensions to combine, then press Finish");

                if (refs.Count > 1)
                {
                    List<Dimension> dimList = refs.Select(r => doc.GetElement(r) as Dimension).ToList();

                    // Step 2: Call Logic Processing (Transaction is handled inside this method)
                    CombineDimLogic.Run(doc, dimList);

                    // Step 3: Refresh view after Transaction has committed successfully
                    uidoc.RefreshActiveView();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    public class DimensionSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e) => e is Dimension;
        public bool AllowReference(Reference r, XYZ p) => false;
    }

    public static class CombineDimLogic
    {
        public static void Run(Document doc, List<Dimension> dims)
        {
            // Filter linear dimensions
            var linearDims = dims.Where(d => d.DimensionShape == DimensionShape.Linear).ToList();
            if (linearDims.Count < 2) return;

            // Group collinear dimensions
            List<List<Dimension>> groups = GroupCollinearDimensions(linearDims);

            using (Transaction t = new Transaction(doc, "Combine Dimensions"))
            {
                t.Start();

                foreach (var group in groups)
                {
                    if (group.Count < 2) continue;

                    Line baseLine = group[0].Curve as Line;
                    if (baseLine == null) continue;

                    // Use Dictionary to filter duplicate dimension references
                    Dictionary<string, Reference> uniqueRefs = new Dictionary<string, Reference>();

                    foreach (var d in group)
                    {
                        foreach (Reference r in d.References)
                        {
                            // Get unique stable representation for comparison
                            string refToken = r.ConvertToStableRepresentation(doc);

                            if (!uniqueRefs.ContainsKey(refToken))
                            {
                                uniqueRefs.Add(refToken, r);
                            }
                        }
                    }

                    ReferenceArray combinedRefs = new ReferenceArray();
                    foreach (var r in uniqueRefs.Values)
                    {
                        Element el = doc.GetElement(r.ElementId);
                        if (el is Grid g)
                        {
                            // Re-create reference for Grids to ensure stability
                            combinedRefs.Append(new Reference(g));
                        }
                        else
                        {
                            combinedRefs.Append(r);
                        }
                    }

                    if (combinedRefs.Size < 2) continue;

                    try
                    {
                        // Create new dimension string and delete old ones
                        Dimension newDim = doc.Create.NewDimension(doc.ActiveView, baseLine, combinedRefs);
                        if (newDim != null)
                        {
                            newDim.DimensionType = group[0].DimensionType;
                            List<ElementId> idsToDelete = group.Select(d => d.Id).ToList();
                            doc.Delete(idsToDelete);
                        }
                    }
                    catch { }
                }

                // [IMPORTANT]: Regenerate must be called INSIDE the Transaction
                doc.Regenerate();
                t.Commit();
            }
        }

        private static List<List<Dimension>> GroupCollinearDimensions(List<Dimension> dims)
        {
            List<List<Dimension>> groups = new List<List<Dimension>>();
            List<Dimension> pool = new List<Dimension>(dims);

            while (pool.Count > 0)
            {
                var current = pool[0];
                pool.RemoveAt(0);
                var currentGroup = new List<Dimension> { current };
                Line lineA = current.Curve as Line;

                if (lineA != null)
                {
                    for (int i = pool.Count - 1; i >= 0; i--)
                    {
                        var other = pool[i];
                        Line lineB = other.Curve as Line;
                        if (lineB != null && AreCollinear(lineA, lineB))
                        {
                            currentGroup.Add(other);
                            pool.RemoveAt(i);
                        }
                    }
                }
                groups.Add(currentGroup);
            }
            return groups;
        }

        private static bool AreCollinear(Line a, Line b)
        {
            XYZ dirA = a.Direction.Normalize();
            XYZ dirB = b.Direction.Normalize();

            // Check parallelism
            if (Math.Abs(Math.Abs(dirA.DotProduct(dirB)) - 1.0) > 0.0001) return false;

            // Check if they lie on the same line
            XYZ pA = a.Origin;
            XYZ pB = b.Origin;
            return (pA - pB).CrossProduct(dirB).GetLength() < 0.001;
        }
    }
}