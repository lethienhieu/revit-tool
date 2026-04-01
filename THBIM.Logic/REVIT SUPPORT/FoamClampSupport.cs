using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
#nullable disable

namespace THBIM.Supports
{
    public static class FoamClampSupport
    {
        public static void Run(ExternalCommandData commandData, Document doc, List<Element> elements, double offsetA, double spacingB)
        {
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family.Name == "TH_Foam Pipe Clamp");

            if (symbol == null)
            {
                TaskDialog.Show("Error", "Family 'TH_Foam Pipe Clamp' not found");
                return;
            }

            using (Transaction trans = new Transaction(doc, "Place Foam Clamp Supports"))
            {
                trans.Start();

                if (!symbol.IsActive)
                    symbol.Activate();

                foreach (Element elem in elements)
                {
                    Pipe pipe = elem as Pipe;
                    if (pipe == null)
                        continue;

                    LocationCurve locCurve = pipe.Location as LocationCurve;
                    if (locCurve == null || !(locCurve.Curve is Line))
                        continue;

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
                        XYZ point = start + direction * currentDist;

                        FamilyInstance fi = doc.Create.NewFamilyInstance(point, symbol, pipe, StructuralType.NonStructural);
                        // Rotate to align with pipe direction
                        Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                        double angle = XYZ.BasisX.AngleTo(direction);
                        // If family X-axis needs to be perpendicular to pipe (rotate 90 degrees)
                        // => Add 90 degrees (PI/2 radians)
                        angle += Math.PI / 2;
                        XYZ cross = XYZ.BasisX.CrossProduct(direction);
                        if (cross.Z < 0) angle = -angle;
                        ElementTransformUtils.RotateElement(doc, fi.Id, axis, angle);

                        SetFamilyParameters(fi, pipe, doc);

                        currentDist += spacingB;

                    }
                }

                trans.Commit();
            }
        }


        // Helper function to handle parameter setting for family instance for better maintainability
        private static void SetFamilyParameters(FamilyInstance fi, Pipe pipe, Document doc)
        {
            double outsideDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
            double insulationThickness = GetInsulationThicknessFromPipe(doc, pipe);

            Parameter dnParam = fi.LookupParameter("DN");
            if (dnParam != null && !dnParam.IsReadOnly)
                dnParam.Set(outsideDiameter);

            Parameter insulationParam = fi.LookupParameter("Insulation");
            if (insulationParam != null && !insulationParam.IsReadOnly)
                insulationParam.Set(insulationThickness);
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

        // Get insulation from pipe, fallback is 30mm if no insulation exists
        private static double GetInsulationThicknessFromPipe(Document doc, Pipe pipe)
        {
            ICollection<ElementId> insIds = InsulationLiningBase.GetInsulationIds(doc, pipe.Id);

            if (insIds == null || insIds.Count == 0)
                return 30.0 / 304.8; // fallback: 30mm -> feet

            foreach (ElementId id in insIds)
            {
                PipeInsulation insulation = doc.GetElement(id) as PipeInsulation;
                if (insulation != null)
                {
                    Parameter p = insulation.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS)
                                     ?? insulation.LookupParameter("Insulation Thickness");

                    if (p != null && p.HasValue)
                        return p.AsDouble();
                }
            }

            return 30.0 / 304.8; // fallback when thickness cannot be retrieved
        }
    }
}