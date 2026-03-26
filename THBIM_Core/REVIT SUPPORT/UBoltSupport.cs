using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
#nullable disable

namespace THBIM.Supports
{
    public static class UBoltSupport
    {
        public static void Run(ExternalCommandData commandData, Document doc, List<Element> elements, double offsetA, double spacingB)
        {
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family.Name == "TH_U_Bolt");

            if (symbol == null)
            {
                TaskDialog.Show("Error", "Family 'TH_U_Bolt' not found");
                return;
            }

            using (Transaction trans = new Transaction(doc, "Place U-Bolt Supports"))
            {
                trans.Start();
                if (!symbol.IsActive) symbol.Activate();

                foreach (var elem in elements)
                {
                    if (!(elem is Pipe pipe)) continue;

                    LocationCurve locCurve = pipe.Location as LocationCurve;
                    if (locCurve == null || !(locCurve.Curve is Line)) continue;

                    Line line = locCurve.Curve as Line;

                    // Skip vertical pipes (avoid elevation errors due to overlapping Z-axis)
                    XYZ dirChk = line.Direction.Normalize();
                    if (Math.Abs(dirChk.DotProduct(XYZ.BasisZ)) > 0.9) continue;

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

                    // OD and Insulation to calculate pipe BOTTOM elevation (U-bolt placement point)
                    double outsideDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0.0; // feet
                    double insulation = GetInsulationThicknessFromPipe(doc, pipe); // feet

                    double currentDist = offsetA;

                    while (currentDist < length)
                    {
                        XYZ centerPoint = start + direction.Multiply(currentDist);

                        // PLACE AT PIPE BOTTOM: subtract (OD/2 + insulation)
                        XYZ placePoint = centerPoint;

                        FamilyInstance fi = doc.Create.NewFamilyInstance(placePoint, symbol, pipe, StructuralType.NonStructural);

                        // Rotate to align with pipe direction
                        Line axis = Line.CreateBound(placePoint, placePoint + XYZ.BasisZ);
                        double angle = XYZ.BasisX.AngleTo(direction);
                        XYZ cross = XYZ.BasisX.CrossProduct(direction);
                        if (cross.Z < 0) angle = -angle;
                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);

                        // Set DN = pipe outer diameter (feet)
                        Parameter dnParam = fi.LookupParameter("DN");
                        if (dnParam != null && !dnParam.IsReadOnly)
                        {
                            dnParam.Set(outsideDiameter);
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

        // --- Helper: get pipe insulation thickness (feet) ---
        private static double GetInsulationThicknessFromPipe(Document doc, Pipe pipe)
        {
            var insIds = InsulationLiningBase.GetInsulationIds(doc, pipe.Id);
            foreach (var id in insIds)
            {
                Element ins = doc.GetElement(id);
                if (ins == null) continue;

                // Prioritize built-in, fallback to parameter named "Insulation Thickness"
                Parameter p = ins.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS)
                             ?? ins.LookupParameter("Insulation Thickness");
                if (p != null && p.HasValue)
                    return p.AsDouble(); // feet
            }
            return 0.0;
        }

        // Selection filter kept as is (if separate selection command is needed)
        private class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element e) => e is Pipe;
            public bool AllowReference(Reference r, XYZ p) => true;
        }
    }
}