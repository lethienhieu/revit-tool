using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
#nullable disable

namespace THBIM.Supports
{
    public static class USteelFoundationSupport
    {
        public static void Run(ExternalCommandData commandData, Document doc, List<Element> elements, double offsetA, double spacingB)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            TaskDialogResult result = TaskDialog.Show("Mode", "Select support placement mode:\nYes: Single pipe\nNo: Multiple parallel pipes",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                TaskDialogResult.Yes);

            bool isSinglePipeMode = (result == TaskDialogResult.Yes);

            IList<Reference> pickedRefs = uidoc.Selection.PickObjects(ObjectType.Element, new MEPSelectionFilter(), "Select pipes or ducts to place support");
            if (pickedRefs == null || pickedRefs.Count == 0) return;

            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family.Name == "TH_DA_U_Steel_Foundation");

            if (symbol == null)
            {
                TaskDialog.Show("Error", "Family 'TH_DA_U_Steel_Foundation' not found");
                return;
            }

            double parallelSpacing = 0.0;
            XYZ perpOffset = XYZ.Zero;
            Element baseElement = null;

            if (!isSinglePipeMode && pickedRefs.Count >= 2)
            {
                List<Tuple<Element, XYZ, double>> elemData = new List<Tuple<Element, XYZ, double>>();

                foreach (var r in pickedRefs)
                {
                    Element elem = doc.GetElement(r);
                    if (!(elem is Pipe || elem is Duct)) continue;

                    LocationCurve lc = elem.Location as LocationCurve;
                    if (lc == null || !(lc.Curve is Line)) continue;

                    Line line = lc.Curve as Line;
                    XYZ dir = line.Direction.Normalize();

                    if (Math.Abs(dir.DotProduct(XYZ.BasisZ)) > 0.9)
                        continue;

                    XYZ center = (line.GetEndPoint(0) + line.GetEndPoint(1)) / 2.0;
                    double width = GetElementWidth(elem);
                    double insulation = GetInsulationThicknessFromElement(doc, elem);
                    double bottomZ = GetPlacementPoint(elem, center, width, insulation).Z;

                    elemData.Add(Tuple.Create(elem, center, bottomZ));
                }

                if (elemData.Count < 2)
                {
                    TaskDialog.Show("Error", "Not enough valid elements to calculate spacing (minimum 2 required).");
                    return;
                }

                double minZ = elemData.Min(x => x.Item3);
                double maxZ = elemData.Max(x => x.Item3);
                double diffZ = (maxZ - minZ) * 304.8;
                if (diffZ > 50)
                {
                    TaskDialog.Show("Error", $"Bottom elevations differ by {Math.Round(diffZ, 1)}mm. Cannot place support.");
                    return;
                }

                var baseTuple = elemData.First(x => x.Item3 == minZ);
                baseElement = baseTuple.Item1;
                XYZ baseCenter = baseTuple.Item2;

                Line baseLine = (baseElement.Location as LocationCurve)?.Curve as Line;
                XYZ direction = baseLine.Direction;
                XYZ perpFlat = new XYZ(direction.Y, -direction.X, 0).Normalize();

                List<double> projections = elemData.Select(p => p.Item2.DotProduct(perpFlat)).ToList();
                double min = projections.Min();
                double max = projections.Max();
                parallelSpacing = max - min;

                double midProjection = (min + max) / 2.0;
                double baseProjection = baseCenter.DotProduct(perpFlat);
                double signedOffset = midProjection - baseProjection;
                perpOffset = perpFlat.Multiply(signedOffset);
            }

            using (Transaction trans = new Transaction(doc, "Place USteel Foundation Supports"))
            {
                trans.Start();
                if (!symbol.IsActive) symbol.Activate();

                foreach (var r in pickedRefs)
                {
                    Element elem = doc.GetElement(r);
                    if (!(elem is Pipe || elem is Duct)) continue;
                    if (!isSinglePipeMode && elem.Id != baseElement.Id) continue;

                    LocationCurve locCurve = elem.Location as LocationCurve;
                    if (locCurve == null || !(locCurve.Curve is Line)) continue;

                    Line line = locCurve.Curve as Line;
                    XYZ originalStart = line.GetEndPoint(0);
                    XYZ originalEnd = line.GetEndPoint(1);
                    if (originalStart.X > originalEnd.X || (originalStart.X == originalEnd.X && originalStart.Y > originalEnd.Y))
                    {
                        XYZ temp = originalStart;
                        originalStart = originalEnd;
                        originalEnd = temp;
                    }

                    XYZ direction = (originalEnd - originalStart).Normalize();

                    XYZ start = originalStart;
                    double length = originalStart.DistanceTo(originalEnd);
                    double currentDist = offsetA;

                    while (currentDist < length)
                    {
                        XYZ point = start + direction.Multiply(currentDist);
                        if (!isSinglePipeMode)
                            point += perpOffset;

                        double width = GetElementWidth(elem);
                        double insulation = GetInsulationThicknessFromElement(doc, elem);
                        point = GetPlacementPoint(elem, point, width, insulation);

                        FamilyInstance fi = doc.Create.NewFamilyInstance(point, symbol, elem, StructuralType.NonStructural);

                        Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                        double angle = XYZ.BasisX.AngleTo(direction);
                        XYZ cross = XYZ.BasisX.CrossProduct(direction);
                        if (cross.Z < 0) angle = -angle;
                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);

                        // Set USteel_Total Length
                        Parameter totalLengthParam = fi.LookupParameter("USteel_Total Length");
                        if (totalLengthParam != null && !totalLengthParam.IsReadOnly)
                        {
                            double value = 0;
                            if (elem is Pipe)
                            {
                                value = width + 200.0 / 304.8 + 2 * insulation;
                            }
                            else if (elem is Duct duct)
                            {
                                double diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                                if (diameter > 0)
                                    value = diameter + 200.0 / 304.8 + 2 * insulation;
                                else
                                    value = width + 200.0 / 304.8 + 2 * insulation;
                            }
                            if (!isSinglePipeMode)
                                value += parallelSpacing;
                            totalLengthParam.Set(value);
                        }

                        // Set USteel_Height (from pipe bottom to top face of floor, minus 105mm)
                        Parameter heightParam = fi.LookupParameter("USteel_Height");
                        if (heightParam != null && !heightParam.IsReadOnly)
                        {
                            double pipeBottomZ = point.Z;
                            double floorTopZ = FindTopFloorAbove(doc, point);
                            double height = 0;
                            if (floorTopZ < pipeBottomZ)
                                height = (pipeBottomZ - floorTopZ) - (105.0 / 304.8); // Subtract 105mm, in feet
                            heightParam.Set(height > 0 ? height : 0);
                        }

                        currentDist += spacingB;
                        // === NEW: Set parameter "System_Type" = System **Type** name of the host ===
                        try
                        {
                            string systemTypeName = null;

                            // If it is a Pipe
                            if (elem is Pipe pipe)
                            {
                                var mepSystem = pipe.MEPSystem;
                                if (mepSystem != null)
                                {
                                    Element sysTypeElem = doc.GetElement(mepSystem.GetTypeId());
                                    systemTypeName = sysTypeElem?.Name;
                                }
                                if (string.IsNullOrWhiteSpace(systemTypeName))
                                    systemTypeName = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString();
                            }
                            // If it is a Duct
                            else if (elem is Duct duct)
                            {
                                var mepSystem = duct.MEPSystem;
                                if (mepSystem != null)
                                {
                                    Element sysTypeElem = doc.GetElement(mepSystem.GetTypeId());
                                    systemTypeName = sysTypeElem?.Name;
                                }
                                if (string.IsNullOrWhiteSpace(systemTypeName))
                                    systemTypeName = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString();
                            }

                            Parameter pSystemType = fi.LookupParameter("System_Type");
                            if (pSystemType != null && !pSystemType.IsReadOnly && !string.IsNullOrWhiteSpace(systemTypeName))
                            {
                                pSystemType.Set(systemTypeName.Trim());
                            }
                        }
                        catch { /* ignore if error */ }

                    }
                }

                trans.Commit();
            }
        }

        private static double GetElementWidth(Element e)
        {
            if (e is Pipe pipe)
                return pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0;
            else if (e is Duct duct)
                return duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
            return 0;
        }

        private static double GetInsulationThicknessFromElement(Document doc, Element elem)
        {
            var insIds = InsulationLiningBase.GetInsulationIds(doc, elem.Id);
            foreach (var id in insIds)
            {
                Element ins = doc.GetElement(id);
                Parameter p = ins?.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS)
                             ?? ins?.LookupParameter("Insulation Thickness");
                if (p != null && p.HasValue)
                    return p.AsDouble();
            }
            return 0.0;
        }

        private static XYZ GetPlacementPoint(Element elem, XYZ point, double width, double insulation)
        {
            if (elem is Pipe)
            {
                return point - XYZ.BasisZ.Multiply((width / 2.0) + insulation);
            }
            else if (elem is Duct duct)
            {
                double diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                if (diameter > 0)
                    return point - XYZ.BasisZ.Multiply(diameter / 2.0);
                else
                {
                    double height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                    return point - XYZ.BasisZ.Multiply(height / 2.0);
                }
            }
            return point;
        }

        private class MEPSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element e) => e is Pipe || e is Duct;
            public bool AllowReference(Reference r, XYZ p) => true;
        }

        // Find top face of floor (projecting Z down), return elevation of nearest floor top face (in feet)
        private static double FindTopFloorAbove(Document doc, XYZ origin)
        {
            View3D view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);

            if (view3D == null) return origin.Z - (1000.0 / 304.8);

            XYZ direction = -XYZ.BasisZ; // Project downwards
            ReferenceIntersector intersector = new ReferenceIntersector(
                new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                FindReferenceTarget.Face,
                view3D);
            intersector.FindReferencesInRevitLinks = true;

            ReferenceWithContext result = intersector.FindNearest(origin, direction);
            if (result != null && result.GetReference() != null)
            {
                XYZ hitPoint = result.GetReference().GlobalPoint;
                return hitPoint.Z;
            }

            // If no floor found, return 1000mm lower
            return origin.Z - (1000.0 / 304.8);
        }
    }
}