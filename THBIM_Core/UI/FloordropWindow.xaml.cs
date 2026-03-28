using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace THBIM.Tools.UI
{
    public partial class FloordropWindow : Window
    {
        private Document _doc;
        private const string FamilyName = "LB_WH_CST_SYB_Floor Drop";

        public bool HasUpdated { get; private set; } = false;
        public FamilySymbol SelectedFamilySymbol { get; private set; }

        public FloordropWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
        }

        // ======================================================
        // AUTO-LOAD FAMILY
        // ======================================================
        private FamilySymbol EnsureFamilyLoaded()
        {
            // 1. Check if already loaded in document
            var existing = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.FamilyName == FamilyName);

            if (existing != null) return existing;

            // 2. Find .rfa file next to DLL
            string asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
            string rfaPath = Path.Combine(asmDir, "rfa", FamilyName + ".rfa");

            if (!File.Exists(rfaPath))
            {
                TaskDialog.Show("THBIM", $"Family file not found:\n{rfaPath}");
                return null;
            }

            // 3. Load family
            using (Transaction t = new Transaction(_doc, "Load Floor Drop Family"))
            {
                t.Start();
                if (_doc.LoadFamily(rfaPath, out Family loadedFamily))
                {
                    var symbolIds = loadedFamily.GetFamilySymbolIds();
                    if (symbolIds.Count > 0)
                    {
                        var symbol = _doc.GetElement(symbolIds.First()) as FamilySymbol;
                        if (symbol != null && !symbol.IsActive)
                            symbol.Activate();
                        t.Commit();
                        return symbol;
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("THBIM", "Failed to load Floor Drop family.");
            return null;
        }

        // ======================================================
        // START PICKING
        // ======================================================
        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            var symbol = EnsureFamilyLoaded();
            if (symbol == null) return;

            SelectedFamilySymbol = symbol;
            this.DialogResult = true;
            this.Close();
        }

        // ======================================================
        // UPDATE ALL INSTANCES IN VIEW
        // ======================================================
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var symbol = EnsureFamilyLoaded();
            if (symbol == null) return;

            long targetTypeId = symbol.Id.GetValue();

            // Find all instances of this family type in active view
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_GenericAnnotation,
                BuiltInCategory.OST_DetailComponents
            };
            var filter = new ElementMulticategoryFilter(categories);

            var elementsInView = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();

            var instances = new List<Element>();
            foreach (var elem in elementsInView)
            {
                if (elem.GetTypeId().GetValue() == targetTypeId)
                    instances.Add(elem);
            }

            if (instances.Count == 0)
            {
                TaskDialog.Show("THBIM", "No Floor Drop instances found in current view.");
                return;
            }

            int successCount = 0;
            int noChangeCount = 0;

            using (Transaction t = new Transaction(_doc, "Update Floor Drops"))
            {
                t.Start();

                foreach (Element inst in instances)
                {
                    LocationPoint loc = inst.Location as LocationPoint;
                    if (loc == null) continue;

                    XYZ point = loc.Point;
                    string newText = GetAdjacentFloorText(_doc, point);

                    if (!string.IsNullOrEmpty(newText))
                    {
                        Parameter param = GetFirstEditableTextParam(inst);
                        if (param != null)
                        {
                            if (param.AsString() != newText)
                            {
                                param.Set(newText);
                                successCount++;
                            }
                            else
                            {
                                noChangeCount++;
                            }
                        }
                    }
                }

                t.Commit();
                HasUpdated = true;
            }

            if (successCount > 0 || noChangeCount > 0)
            {
                TaskDialog.Show("Update Result",
                    $"Update successful!\n" +
                    $"- Values updated: {successCount}\n" +
                    $"- Already correct: {noChangeCount}");
                this.Close();
            }
            else
            {
                TaskDialog.Show("Update Result", "Instances found but could not calculate floor values.");
            }
        }

        // ======================================================
        // HELPER CALCULATION METHODS
        // ======================================================
        private string GetAdjacentFloorText(Document doc, XYZ pickPoint)
        {
            double radius = 1.0;
            XYZ centerBottom = new XYZ(pickPoint.X, pickPoint.Y, -1000.0);

            List<Curve> profile = new List<Curve> {
                Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerBottom), radius, 0, Math.PI),
                Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerBottom), radius, Math.PI, 2 * Math.PI)
            };

            Solid searchSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                new List<CurveLoop> { CurveLoop.Create(profile) }, XYZ.BasisZ, 2000.0);

            var floors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WherePasses(new ElementIntersectsSolidFilter(searchSolid))
                .ToElements();

            if (floors.Count < 2) return null;

            var topFloors = floors.OrderByDescending(x => x.get_BoundingBox(null).Max.Z).Take(2).ToList();

            if (IsFloorSloped(topFloors[0]) || IsFloorSloped(topFloors[1]))
                return "Var.";

            double h1 = topFloors[0].get_BoundingBox(null).Max.Z;
            double h2 = topFloors[1].get_BoundingBox(null).Max.Z;
            double diff = Math.Round(UnitUtils.ConvertFromInternalUnits(Math.Abs(h1 - h2), UnitTypeId.Millimeters), 0);

            if (diff == 0) return null;
            return diff.ToString();
        }

        private bool IsFloorSloped(Element elem)
        {
            Options opt = new Options() { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geo = elem.get_Geometry(opt);
            if (geo == null) return false;

            foreach (GeometryObject obj in geo)
            {
                Solid solid = obj as Solid;
                if (solid != null && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                        if (normal.Z > 0 && Math.Abs(normal.Z) < 0.999) return true;
                    }
                }
            }
            return false;
        }

        private Parameter GetFirstEditableTextParam(Element e)
        {
            foreach (Parameter p in e.Parameters)
                if (!p.IsReadOnly && p.StorageType == StorageType.String && p.Definition.GetDataType() == SpecTypeId.String.Text)
                    return p;
            return null;
        }
    }
}
