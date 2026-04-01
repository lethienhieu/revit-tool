using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using THBIM.Supports;
#nullable disable

namespace THBIM.Supports
{
    public static class PipeOmegaSupport
    {
        public static void Run(ExternalCommandData commandData, Document doc, List<Element> elements, double offsetA, double spacingB)
        {
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family.Name == "TH_Pipe Support Omega");

            if (symbol == null)
            {
                TaskDialog.Show("Error", "Family 'TH_Pipe Support Omega' not found");
                return;
            }

            using (Transaction trans = new Transaction(doc, "Place Pipe Supports"))
            {
                trans.Start();
                if (!symbol.IsActive) symbol.Activate();

                foreach (var elem in elements)
                {
                    if (!(elem is Pipe pipe)) continue;

                    LocationCurve locCurve = pipe.Location as LocationCurve;
                    if (locCurve == null || !(locCurve.Curve is Line)) continue;

                    Line line = locCurve.Curve as Line;

                    XYZ start = line.GetEndPoint(0);
                    XYZ end = line.GetEndPoint(1);
                    if (start.X > end.X || (start.X == end.X && start.Y > end.Y))
                    {
                        XYZ temp = start;
                        start = end;
                        end = temp;
                    }


                    XYZ direction = (end - start).Normalize();
                    double length = start.DistanceTo(end);

                    double currentDist = offsetA;

                    while (currentDist < length)
                    {
                        XYZ point = start + direction.Multiply(currentDist);

                        FamilyInstance fi = doc.Create.NewFamilyInstance(point, symbol, pipe, StructuralType.NonStructural);

                        // Rotate to align with pipe direction
                        Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                        double angle = XYZ.BasisX.AngleTo(direction);
                        XYZ cross = XYZ.BasisX.CrossProduct(direction);
                        if (cross.Z < 0) angle = -angle;
                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);

                        // Set TH_DN
                        Parameter dnParam = fi.LookupParameter("TH_DN");
                        if (dnParam != null && !dnParam.IsReadOnly)
                        {
                            dnParam.Set(pipe.Diameter);
                        }

                        // Set threaded rod length
                        Parameter rodLengthParam = fi.LookupParameter("TH_Threaded Rod_Width");
                        if (rodLengthParam != null && !rodLengthParam.IsReadOnly)
                        {
                            double elevation = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble();
                            double rodLength = GetRodLengthFromElevation(doc, fi.Location as LocationPoint, elevation);
                            rodLengthParam.Set(rodLength);
                        }

                        currentDist += spacingB;
                        // NEW: Set "System_Type" parameter = System Type name of the host pipe
                        try
                        {
                            string systemTypeName = null;
                            var mepSys = pipe.MEPSystem;
                            if (mepSys != null)
                            {
                                ElementId typeId = mepSys.GetTypeId();
                                if (typeId != ElementId.InvalidElementId)
                                {
                                    Element sysTypeElem = doc.GetElement(typeId);
                                    systemTypeName = sysTypeElem?.Name;
                                }
                            }
                            if (string.IsNullOrWhiteSpace(systemTypeName))
                            {
                                var prm = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                                if (prm != null) systemTypeName = prm.AsValueString();
                            }

                            if (!string.IsNullOrWhiteSpace(systemTypeName))
                            {
                                Parameter pSysType = fi.LookupParameter("System_Type");
                                if (pSysType != null && !pSysType.IsReadOnly)
                                    pSysType.Set(systemTypeName.Trim());
                            }
                        }
                        catch { /* ignore if unable to set */ }
                    }
                }

                trans.Commit();
            }
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
                {
                    double distanceFeet = floorBottomElevationFeet - elevationFeet;
                    return distanceFeet;
                }
            }

            return 1000.0 / 304.8;
        }

        private class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element e) => e is Pipe;
            public bool AllowReference(Reference r, XYZ p) => true;
        }
    }
}