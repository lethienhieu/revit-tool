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
    public static class USteelSupport
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
                .FirstOrDefault(s => s.Family.Name == "TH_DA_U_Steel_Thread Rod");

            if (symbol == null)
            {
                TaskDialog.Show("Error", "Family 'TH_DA_U_Steel_Thread Rod' not found");
                return;
            }

            double parallelSpacing = 0.0;
            XYZ perpOffset = XYZ.Zero;
            Element baseElement = null;

            // === NEW: store elevation difference to set parameter "LAYER 2" ===
            double diffZFeet = 0.0; // feet
            double diffZmm = 0.0;   // mm (reference / display if param is int/string)

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
                    double bottomZ = GetPlacementPoint(elem, center, width, insulation).Z; // pipe bottom (feet)

                    elemData.Add(Tuple.Create(elem, center, bottomZ));
                }

                if (elemData.Count < 2)
                {
                    // Keep safe behavior: stop if not enough valid data
                    TaskDialog.Show("Error", "Not enough valid elements to calculate spacing (minimum 2 required).");
                    return;
                }

                double minZ = elemData.Min(x => x.Item3); // feet
                double maxZ = elemData.Max(x => x.Item3); // feet

                // === NEW: calculate and SAVE elevation difference, do NOT block anymore ===
                diffZFeet = (maxZ - minZ);
                diffZmm = diffZFeet * 304.8;

                var baseTuple = elemData.First(x => x.Item3 == minZ);
                baseElement = baseTuple.Item1;
                XYZ baseCenter = baseTuple.Item2;

                Line baseLine = (baseElement.Location as LocationCurve)?.Curve as Line;
                XYZ direction = baseLine.Direction;
                XYZ perpFlat = new XYZ(direction.Y, -direction.X, 0).Normalize();

                List<double> projections = elemData.Select(p => p.Item2.DotProduct(perpFlat)).ToList();
                double min = projections.Min();
                double max = projections.Max();
                parallelSpacing = max - min; // feet in diagonal coordinate system

                double midProjection = (min + max) / 2.0;
                double baseProjection = baseCenter.DotProduct(perpFlat);
                double signedOffset = midProjection - baseProjection;
                perpOffset = perpFlat.Multiply(signedOffset);
            }

            using (Transaction trans = new Transaction(doc, "Place USteel Supports"))
            {
                trans.Start();
                if (!symbol.IsActive) symbol.Activate();

                foreach (var r in pickedRefs)
                {
                    Element elem = doc.GetElement(r);
                    if (!(elem is Pipe || elem is Duct)) continue;
                    if (!isSinglePipeMode && elem.Id != baseElement?.Id) continue;

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

                        Parameter totalLengthParam = fi.LookupParameter("VSteel_Total Length");
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

                        Parameter floorDistParam = fi.LookupParameter("ElementBottomToFloor");
                        if (floorDistParam != null && !floorDistParam.IsReadOnly)
                        {
                            double elevation = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble();
                            double rodLength = GetRodLengthFromElevation(doc, fi.Location as LocationPoint, elevation);
                            floorDistParam.Set(rodLength);
                        }

                        // === NEW: Write elevation difference to parameter "LAYER 2" ===
                        try
                        {
                            Parameter pLayer2 = fi.LookupParameter("LAYER 2");
                            if (pLayer2 != null && !pLayer2.IsReadOnly)
                            {
                                // Multi: use diffZ; Single: 0
                                double valFeet = (!isSinglePipeMode) ? diffZFeet : 0.0;
                                double valMm = (!isSinglePipeMode) ? diffZmm : 0.0;

                                switch (pLayer2.StorageType)
                                {
                                    case StorageType.Double:
                                        // If type is Length -> feet
                                        pLayer2.Set(valFeet);
                                        break;
                                    case StorageType.Integer:
                                        // If int -> mm as integer
                                        pLayer2.Set((int)Math.Round(valMm));
                                        break;
                                    case StorageType.String:
                                        // If text -> mm as string
                                        pLayer2.Set(Math.Round(valMm, 0).ToString());
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        catch { /* ignore if unable to set */ }

                        // === NEW: Set parameter "Steel_Type" if input exists in UI ===
                        try
                        {
                            string steelType = SupportRuntime.SteelType;
                            if (!string.IsNullOrWhiteSpace(steelType))
                            {
                                Parameter pSteelType = fi.LookupParameter("Steel_Type");
                                if (pSteelType != null && !pSteelType.IsReadOnly)
                                {
                                    pSteelType.Set(steelType.Trim());
                                }
                            }
                        }
                        catch { /* ignore if unable to set */ }

                        currentDist += spacingB;
                        // === NEW: Set parameter "System_Type" = System **Type** name of the host ===
                        try
                        {
                            string systemTypeName = null;

                            // If Pipe
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
                            // If Duct
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

        private static double GetRodLengthFromElevation(Document doc, LocationPoint locationPoint, double elevationFeet)
        {
            if (locationPoint == null) return 1000.0;

            XYZ origin = locationPoint.Point;
            View3D view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);

            if (view3D == null) return 1000.0;

            XYZ direction = XYZ.BasisZ;
            ReferenceIntersector intersector = new ReferenceIntersector(
                new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                FindReferenceTarget.Element,
                view3D);
            intersector.FindReferencesInRevitLinks = true;

            ReferenceWithContext result = intersector.FindNearest(origin, direction);
            if (result != null && result.GetReference() != null)
            {
                XYZ hitPoint = result.GetReference().GlobalPoint;
                double floorBottomElevationFeet = hitPoint.Z;
                if (floorBottomElevationFeet > elevationFeet)
                    return floorBottomElevationFeet - elevationFeet;
            }

            return 1000.0 / 304.8;
        }
    }
}