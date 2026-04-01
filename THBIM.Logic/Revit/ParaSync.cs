using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM
{
    public class MappingRow
    {
        public string SelectedLinkParam { get; set; }
        public string SelectedHostParam { get; set; }
        public ObservableCollection<string> LinkParams { get; set; }
        public ObservableCollection<string> HostParams { get; set; }
    }

    public class CategorySelectionFilter : ISelectionFilter
    {
        private ElementId _catId;
        public CategorySelectionFilter(ElementId catId) => _catId = catId;
        public bool AllowElement(Element elem) => elem.Category != null && elem.Category.Id == _catId;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    public class ParaSyncProcessor
    {
        private UIDocument _uiDoc;
        private Document _doc;

        public ParaSyncProcessor(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
        }

        public void ExecuteWithSelection(RevitLinkInstance linkInst, ElementId linkCatId, ElementId hostCatId, List<MappingRow> mappings)
        {
            try
            {
                IList<Reference> pickedRefs = _uiDoc.Selection.PickObjects(
                    ObjectType.Element,
                    new CategorySelectionFilter(hostCatId),
                    "Select elements in current file to sync data");

                if (pickedRefs == null || !pickedRefs.Any()) return;

                int selectedCount = pickedRefs.Count;
                int finalSuccessCount = 0;

                Document linkDoc = linkInst.GetLinkDocument();
                Transform tr = linkInst.GetTotalTransform();

                using (Transaction t = new Transaction(_doc, "THBIM - Virtual Solid Sync"))
                {
                    t.Start();
                    foreach (Reference r in pickedRefs)
                    {
                        Element hostElem = _doc.GetElement(r);
                        XYZ basePoint = (hostElem.Location as LocationPoint)?.Point;
                        if (basePoint == null) continue;

                        double radius = GetPileRadius(hostElem);
                        double height = 3000 / 304.8; // 3000mm to Feet

                        Solid virtualSolid = CreateVirtualSolid(basePoint, radius, height);
                        Solid checkSolid = SolidUtils.CreateTransformed(virtualSolid, tr.Inverse);

                        // Collect ALL intersecting elements within the specific Link Category
                        var potentialMatches = new FilteredElementCollector(linkDoc)
                            .OfCategoryId(linkCatId)
                            .WhereElementIsNotElementType()
                            .WherePasses(new ElementIntersectsSolidFilter(checkSolid))
                            .ToElements();

                        Element matchedElem = null;
                        foreach (Element e in potentialMatches)
                        {
                            // Verify existence of the mapping parameter
                            Parameter checkP = e.LookupParameter(mappings[0].SelectedLinkParam);
                            if (checkP != null)
                            {
                                matchedElem = e;
                                break;
                            }
                        }

                        if (matchedElem != null)
                        {
                            bool rowSuccess = false;
                            foreach (var map in mappings)
                            {
                                Parameter sP = matchedElem.LookupParameter(map.SelectedLinkParam);
                                Parameter dP = hostElem.LookupParameter(map.SelectedHostParam);

                                if (sP != null && dP != null && !dP.IsReadOnly)
                                {
                                    string val = sP.AsValueString() ?? sP.AsString();
                                    if (!string.IsNullOrEmpty(val))
                                    {
                                        dP.Set(val);
                                        rowSuccess = true;
                                    }
                                }
                            }
                            if (rowSuccess) finalSuccessCount++;
                        }
                    }
                    t.Commit();
                }

                // Updated Notification showing X of Y success
                TaskDialog.Show("Sync Results",
                    $"Process Completed Successfully!\n\n" +
                    $"Selected Elements: {selectedCount}\n" +
                    $"Successfully Synced: {finalSuccessCount}\n" +
                    $"Success Rate: {(selectedCount > 0 ? (finalSuccessCount * 100 / selectedCount) : 0)}%");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            catch (Exception ex) { TaskDialog.Show("Error", ex.Message); }
        }

        private double GetPileRadius(Element e)
        {
            BoundingBoxXYZ bb = e.get_BoundingBox(null);
            if (bb != null)
            {
                double width = bb.Max.X - bb.Min.X;
                double depth = bb.Max.Y - bb.Min.Y;
                return Math.Min(width, depth) / 2.0;
            }
            return 1.0;
        }

        private Solid CreateVirtualSolid(XYZ basePoint, double radius, double height)
        {
            double safeRadius = radius + 0.16; // +50mm safety buffer
            XYZ startPoint = basePoint - new XYZ(0, 0, 1.6); // Start 500mm below
            double totalHeight = height + 3.2;

            List<Curve> profile = new List<Curve>();
            XYZ p1 = startPoint + new XYZ(-safeRadius, -safeRadius, 0);
            XYZ p2 = startPoint + new XYZ(safeRadius, -safeRadius, 0);
            XYZ p3 = startPoint + new XYZ(safeRadius, safeRadius, 0);
            XYZ p4 = startPoint + new XYZ(-safeRadius, safeRadius, 0);

            profile.Add(Line.CreateBound(p1, p2));
            profile.Add(Line.CreateBound(p2, p3));
            profile.Add(Line.CreateBound(p3, p4));
            profile.Add(Line.CreateBound(p4, p1));

            return GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { CurveLoop.Create(profile) }, XYZ.BasisZ, totalHeight);
        }
    }
}