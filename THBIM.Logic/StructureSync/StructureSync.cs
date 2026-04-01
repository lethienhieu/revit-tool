using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using THBIM.Views;

namespace THBIM
{
    public class MultiCategorySelectionFilter : ISelectionFilter
    {
        private readonly Document _doc;
        private readonly List<BuiltInCategory> _allowedCategories;
        private readonly bool _isLinkMode;

        public MultiCategorySelectionFilter(Document doc, List<BuiltInCategory> allowedCategories, bool isLinkMode)
        {
            _doc = doc; _allowedCategories = allowedCategories; _isLinkMode = isLinkMode;
        }
        public bool AllowElement(Element elem)
        {
            if (_isLinkMode) return elem is RevitLinkInstance;
            if (elem?.Category == null) return false;
            return _allowedCategories.Contains((BuiltInCategory)elem.Category.Id.GetValue());
        }
        public bool AllowReference(Reference reference, XYZ position)
        {
            if (_isLinkMode && reference.LinkedElementId != ElementId.InvalidElementId)
            {
                Element hostElem = _doc.GetElement(reference.ElementId);
                if (hostElem is RevitLinkInstance linkInstance)
                {
                    Document linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        Element linkedElem = linkDoc.GetElement(reference.LinkedElementId);
                        if (linkedElem?.Category != null)
                            return _allowedCategories.Contains((BuiltInCategory)linkedElem.Category.Id.GetValue());
                    }
                }
            }
            return false;
        }
    }

    public class StructureSync
    {
        private UIDocument _uidoc;
        private Document _doc;
        private StrucSyncCore _core;
        private readonly Guid _schemaGuid = new Guid("A1B2C3D4-E5F6-7890-1234-56789ABCDEF0");

        private static StructureReportWindow _activeReportWindow;
        private static UIApplication _cachedUiApp;
        private static EventHandler<Autodesk.Revit.UI.Events.SelectionChangedEventArgs> _selectionHandler;

        public StructureSync(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;
            _core = new StrucSyncCore(_doc);
        }

        public void MatchRelationships(List<Reference> parentRefs, List<ElementId> childIds,
                                       SyncType pType, SyncType cType,
                                       out List<ElementId> validParents, out List<ElementId> validChildren, out List<string> validOffsets)
        {
            _core.MatchRelationships(parentRefs, childIds, pType, cType,
                                     out validParents, out validChildren, out validOffsets, GetRevitLinks());
        }

        public List<RevitLinkInstance> GetRevitLinks() => new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

        public List<Reference> SelectParents(List<BuiltInCategory> categories, bool useLink)
        {
            var selectedRefs = new List<Reference>();
            ISelectionFilter filter = new MultiCategorySelectionFilter(_doc, categories, useLink);
            try
            {
                if (useLink) selectedRefs.AddRange(_uidoc.Selection.PickObjects(ObjectType.LinkedElement, filter, "TAB chọn Parent (Link)"));
                else selectedRefs.AddRange(_uidoc.Selection.PickObjects(ObjectType.Element, filter, "Chọn Parent (Local)"));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            return selectedRefs;
        }

        public List<ElementId> SelectChildren(List<BuiltInCategory> categories)
        {
            ISelectionFilter filter = new MultiCategorySelectionFilter(_doc, categories, false);
            try { return _uidoc.Selection.PickObjects(ObjectType.Element, filter, "Chọn Child").Select(r => r.ElementId).ToList(); }
            catch { return new List<ElementId>(); }
        }

        // --- 1. SYNC TOÀN BỘ (Nút Update) ---
        public void SyncPositions(List<RelationshipItem> activeRelationships, List<RevitLinkInstance> links)
        {
            using (Transaction t = new Transaction(_doc, "Sync Structure"))
            {
                t.Start();

                // Cài đặt màu Xanh
                OverrideGraphicSettings greenSettings = new OverrideGraphicSettings();
                Color greenColor = new Color(46, 204, 113); // Xanh lá
                greenSettings.SetProjectionLineColor(greenColor);
                var solidFill = new FilteredElementCollector(_doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(x => x.GetFillPattern().IsSolidFill);
                if (solidFill != null) { greenSettings.SetSurfaceForegroundPatternId(solidFill.Id); greenSettings.SetSurfaceForegroundPatternColor(greenColor); }

                foreach (var rel in activeRelationships)
                {
                    if (!rel.IsChecked) continue;

                    Enum.TryParse(rel.ParentTypeStr, out SyncType pType);
                    Enum.TryParse(rel.ChildTypeStr, out SyncType cType);

                    // Cho phép dịch theo Z (cao độ) cho đúng cặp PileCap <-> Pile
                    bool allowZ = (pType == SyncType.PileCap && cType == SyncType.Pile) ||
                                  (pType == SyncType.Pile && cType == SyncType.PileCap);

                    var parentCounts = rel.ParentIds.GroupBy(id => id).ToDictionary(g => g.Key, g => g.Count());

                    for (int i = 0; i < rel.ChildIds.Count; i++)
                    {
                        Element child = _doc.GetElement(rel.ChildIds[i]);
                        if (child == null) continue;

                        string offsetStr = (rel.Offsets != null && i < rel.Offsets.Count) ? rel.Offsets[i] : "";
                        int actualChildCount = parentCounts.ContainsKey(rel.ParentIds[i]) ? parentCounts[rel.ParentIds[i]] : 1;

                        if ((pType == SyncType.PileCap && cType == SyncType.Pile) && actualChildCount == 1) offsetStr = "";

                        XYZ targetCenter = _core.CalculateTargetPosition(rel.ParentIds[i], child, pType, cType, links, offsetStr, actualChildCount);
                        if (targetCenter == null) continue;

                        // Di chuyển dạng Point (Cột, Đài, Cọc...)
                        if (child.Location is LocationPoint lp)
                        {
                            XYZ currentPoint = lp.Point;
                            double dz = allowZ ? (targetCenter.Z - currentPoint.Z) : 0;
                            XYZ translation = new XYZ(targetCenter.X - currentPoint.X, targetCenter.Y - currentPoint.Y, dz);
                            if (translation.GetLength() > 1.0e-9) ElementTransformUtils.MoveElement(_doc, child.Id, translation);
                        }
                        // Di chuyển dạng Line (Vách, Dầm...) - Chỉ tịnh tiến, bỏ logic chiều dài
                        else if (child.Location is LocationCurve)
                        {
                            XYZ currentCenter = _core.GetElementCenter(child, Transform.Identity);
                            double dz = allowZ ? (targetCenter.Z - currentCenter.Z) : 0;
                            XYZ translation = new XYZ(targetCenter.X - currentCenter.X, targetCenter.Y - currentCenter.Y, dz);
                            if (translation.GetLength() > 1.0e-9) ElementTransformUtils.MoveElement(_doc, child.Id, translation);
                        }

                        // Tô màu xanh sau khi Sync thành công
                        try { _doc.ActiveView.SetElementOverrides(child.Id, greenSettings); } catch { }
                    }
                    rel.Status = "Synced";
                }
                t.Commit();
            }
        }

        // --- 2. SYNC TỪNG CẶP & HIGHLIGHT XANH ---
        public void SyncSpecificItems(List<ReportItem> itemsToSync, List<RevitLinkInstance> links)
        {
            using (Transaction t = new Transaction(_doc, "Sync Specific & Green Highlight"))
            {
                t.Start();

                OverrideGraphicSettings greenSettings = new OverrideGraphicSettings();
                Color greenColor = new Color(46, 204, 113);
                greenSettings.SetProjectionLineColor(greenColor);
                var solidFill = new FilteredElementCollector(_doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(x => x.GetFillPattern().IsSolidFill);
                if (solidFill != null) { greenSettings.SetSurfaceForegroundPatternId(solidFill.Id); greenSettings.SetSurfaceForegroundPatternColor(greenColor); }

                var parentCounts = itemsToSync.Select(x => x.ParentId).GroupBy(id => id).ToDictionary(g => g.Key, g => g.Count());

                foreach (var rItem in itemsToSync)
                {
                    Enum.TryParse(rItem.SourceItem.ParentTypeStr, out SyncType pType);
                    Enum.TryParse(rItem.SourceItem.ChildTypeStr, out SyncType cType);

                    // Cho phép dịch theo Z (cao độ) cho đúng cặp PileCap <-> Pile
                    bool allowZ = (pType == SyncType.PileCap && cType == SyncType.Pile) ||
                                  (pType == SyncType.Pile && cType == SyncType.PileCap);

                    int actualChildCount = parentCounts.ContainsKey(rItem.ParentId) ? parentCounts[rItem.ParentId] : 1;

                    Element child = _doc.GetElement(rItem.ChildId);
                    if (child == null) continue;

                    string offsetStr = (rItem.SourceItem.Offsets != null && rItem.IndexInSet < rItem.SourceItem.Offsets.Count) ? rItem.SourceItem.Offsets[rItem.IndexInSet] : "";
                    if ((pType == SyncType.PileCap && cType == SyncType.Pile) && actualChildCount == 1) offsetStr = "";

                    XYZ targetCenter = _core.CalculateTargetPosition(rItem.ParentId, child, pType, cType, links, offsetStr, actualChildCount);
                    if (targetCenter == null) continue;

                    // Di chuyển dạng Point
                    if (child.Location is LocationPoint lp)
                    {
                        XYZ currentPoint = lp.Point;
                        double dz = allowZ ? (targetCenter.Z - currentPoint.Z) : 0;
                        XYZ translation = new XYZ(targetCenter.X - currentPoint.X, targetCenter.Y - currentPoint.Y, dz);
                        if (translation.GetLength() > 1.0e-9) ElementTransformUtils.MoveElement(_doc, child.Id, translation);
                    }
                    // Di chuyển dạng Line (Chỉ tịnh tiến)
                    else if (child.Location is LocationCurve)
                    {
                        XYZ currentCenter = _core.GetElementCenter(child, Transform.Identity);
                        double dz = allowZ ? (targetCenter.Z - currentCenter.Z) : 0;
                        XYZ translation = new XYZ(targetCenter.X - currentCenter.X, targetCenter.Y - currentCenter.Y, dz);
                        if (translation.GetLength() > 1.0e-9) ElementTransformUtils.MoveElement(_doc, child.Id, translation);
                    }

                    try { _doc.ActiveView.SetElementOverrides(child.Id, greenSettings); } catch { }

                    rItem.Severity = ReportSeverity.Synced;
                    rItem.StatusDisplay = "Synced";
                    rItem.Diagnosis = "Synced & Highlighted Green.";
                    rItem.TechData = "ΔX=0.000, ΔY=0.000, ΔZ=0 (mm)";
                }
                t.Commit();
            }
        }

        // --- 3. HIGHLIGHT UNSYNCED (Tô Đỏ lỗi) ---
        public void HighlightUnsynced(List<RelationshipItem> activeRelationships, List<RevitLinkInstance> links)
        {
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            ogs.SetProjectionLineColor(new Color(255, 0, 0));
            ogs.SetProjectionLineWeight(5);
            var solidPattern = new FilteredElementCollector(_doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(x => x.GetFillPattern().IsSolidFill);
            if (solidPattern != null) { ogs.SetSurfaceForegroundPatternId(solidPattern.Id); ogs.SetSurfaceForegroundPatternColor(new Color(255, 0, 0)); }

            using (Transaction t = new Transaction(_doc, "Highlight Error"))
            {
                t.Start();
                foreach (var rel in activeRelationships)
                {
                    if (!rel.IsChecked) continue;
                    Enum.TryParse(rel.ParentTypeStr, out SyncType pType);
                    Enum.TryParse(rel.ChildTypeStr, out SyncType cType);

                    var parentCounts = rel.ParentIds.GroupBy(id => id).ToDictionary(g => g.Key, g => g.Count());

                    for (int i = 0; i < rel.ChildIds.Count; i++)
                    {
                        Element child = _doc.GetElement(rel.ChildIds[i]);
                        if (child == null) continue;

                        int actualChildCount = parentCounts.ContainsKey(rel.ParentIds[i]) ? parentCounts[rel.ParentIds[i]] : 1;

                        string offsetStr = (rel.Offsets != null && i < rel.Offsets.Count) ? rel.Offsets[i] : "";
                        if ((pType == SyncType.PileCap && cType == SyncType.Pile) && actualChildCount == 1) offsetStr = "";

                        XYZ targetCenter = _core.CalculateTargetPosition(rel.ParentIds[i], child, pType, cType, links, offsetStr, actualChildCount);

                        if (targetCenter == null) continue;

                        bool isDeviated = false;
                        XYZ currentCenter = _core.GetElementCenter(child, Transform.Identity);
                        double distXY = Math.Sqrt(Math.Pow(targetCenter.X - currentCenter.X, 2) + Math.Pow(targetCenter.Y - currentCenter.Y, 2));
                        double distZ = Math.Abs(targetCenter.Z - currentCenter.Z);

                        if (pType == SyncType.PileCap && cType == SyncType.Pile)
                        {
                            if (distXY > 1.0e-9 || distZ > 1.0e-9) isDeviated = true;
                        }
                        else
                        {
                            if (distXY > 1.0e-9) isDeviated = true;
                        }

                        if (isDeviated)
                        {
                            _doc.ActiveView.SetElementOverrides(child.Id, ogs);
                            Element parent = _doc.GetElement(rel.ParentIds[i]);
                            if (parent != null) try { _doc.ActiveView.SetElementOverrides(parent.Id, ogs); } catch { }
                        }
                    }
                    rel.Status = "Checked";
                }
                t.Commit();
            }
        }

        // --- RESET COLOR ---
        public void ResetColor(List<RelationshipItem> activeRelationships)
        {
            OverrideGraphicSettings clearSettings = new OverrideGraphicSettings();

            using (Transaction t = new Transaction(_doc, "Reset Color"))
            {
                t.Start();
                foreach (var rel in activeRelationships)
                {
                    foreach (var id in rel.ChildIds) try { _doc.ActiveView.SetElementOverrides(id, clearSettings); } catch { }
                    foreach (var id in rel.ParentIds) try { _doc.ActiveView.SetElementOverrides(id, clearSettings); } catch { }
                    rel.Status = "Ready";
                }
                t.Commit();
            }
        }

        // --- SAVE/LOAD ---
        private Schema GetSchema()
        {
            Schema schema = Schema.Lookup(_schemaGuid);
            if (schema != null) return schema;
            SchemaBuilder builder = new SchemaBuilder(_schemaGuid);
            builder.SetReadAccessLevel(AccessLevel.Public);
            builder.SetWriteAccessLevel(AccessLevel.Public);
            builder.SetSchemaName("THBIM_StructureSync_Data");
            builder.AddSimpleField("DataString", typeof(string));
            return builder.Finish();
        }

        // Compute stored offset string in form: "dx,dy,dz"
        // Offset = ChildCenter(host) - ParentCenter(host/linked transformed).
        public string ComputeOffsetStringForPair(ElementId parentId, ElementId childId, List<RevitLinkInstance> links)
        {
            Element child = _doc.GetElement(childId);
            if (child == null) return "";

            var pData = _core.GetParentGeometryData(parentId, links);
            if (pData == null) return "";

            XYZ childCenter = _core.GetElementCenter(child, Transform.Identity);
            XYZ parentCenter = pData.Center;
            XYZ offset = childCenter - parentCenter;
            return $"{offset.X},{offset.Y},{offset.Z}";
        }

        // Regenerate tracking reports from current RelationshipItem sets.
        public List<ReportItem> GenerateReports(List<RelationshipItem> activeRelationships, List<RevitLinkInstance> links)
        {
            return ReportGenerator.Analyze(_core, activeRelationships, links, _doc);
        }

        public void SaveRelationships(List<RelationshipItem> items)
        {
            using (Transaction t = new Transaction(_doc, "Save TH Data"))
            {
                t.Start();
                Element projectInfo = _doc.ProjectInformation;
                Entity entity = new Entity(GetSchema());
                StringBuilder sb = new StringBuilder();
                foreach (var item in items)
                {
                    string pIds = string.Join(",", item.ParentIds.Select(id => id.ToString()));
                    string cIds = string.Join(",", item.ChildIds.Select(id => id.ToString()));
                    string offsets = (item.Offsets != null) ? string.Join("~", item.Offsets) : "";
                    sb.Append($"{item.Name}|{pIds}|{cIds}|{item.ParentTypeStr}|{item.ChildTypeStr}|{(item.ParentIsFromLink ? 1 : 0)}|{offsets};");
                }
                entity.Set("DataString", sb.ToString());
                projectInfo.SetEntity(entity);
                t.Commit();
            }
        }

        public List<RelationshipItem> LoadRelationships()
        {
            List<RelationshipItem> result = new List<RelationshipItem>();
            Entity entity = _doc.ProjectInformation.GetEntity(GetSchema());
            if (entity.IsValid())
            {
                string data = entity.Get<string>("DataString");
                if (!string.IsNullOrEmpty(data))
                {
                    var rows = data.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var row in rows)
                    {
                        var parts = row.Split('|');
                        if (parts.Length >= 3)
                        {
                            var newItem = new RelationshipItem { Name = parts[0], IsChecked = true, Status = "Loaded" };
                            if (!string.IsNullOrEmpty(parts[1])) foreach (var s in parts[1].Split(',')) if (long.TryParse(s, out long id)) newItem.ParentIds.Add(ElementIdCompat.CreateId(id));
                            if (!string.IsNullOrEmpty(parts[2])) foreach (var s in parts[2].Split(',')) if (long.TryParse(s, out long id)) newItem.ChildIds.Add(ElementIdCompat.CreateId(id));

                            bool hasExplicitParentIsFromLink = false;
                            int offsetsIndex = -1;

                            if (parts.Length >= 5) { newItem.ParentTypeStr = parts[3]; newItem.ChildTypeStr = parts[4]; }

                            // New format: |ParentType|ChildType|ParentIsFromLink|Offsets
                            // Old format: |ParentType|ChildType|Offsets
                            if (parts.Length >= 7 && (parts[5] == "0" || parts[5] == "1"))
                            {
                                hasExplicitParentIsFromLink = true;
                                newItem.ParentIsFromLink = parts[5] == "1";
                                offsetsIndex = 6;
                            }
                            else if (parts.Length >= 6)
                            {
                                offsetsIndex = 5;
                            }

                            string offsetsPart = (offsetsIndex >= 0 && offsetsIndex < parts.Length) ? parts[offsetsIndex] : "";
                            if (!string.IsNullOrEmpty(offsetsPart)) newItem.Offsets = offsetsPart.Split('~').ToList();

                            // Backward compatibility: infer ParentIsFromLink if flag is missing.
                            if (!hasExplicitParentIsFromLink && newItem.ParentIds.Count > 0)
                            {
                                try
                                {
                                    ElementId anyParentId = newItem.ParentIds[0];
                                    foreach (var link in GetRevitLinks())
                                    {
                                        var linkDoc = link.GetLinkDocument();
                                        if (linkDoc == null) continue;
                                        if (linkDoc.GetElement(anyParentId) != null)
                                        {
                                            newItem.ParentIsFromLink = true;
                                            break;
                                        }
                                    }
                                }
                                catch { }
                            }

                            while (newItem.Offsets.Count < newItem.ChildIds.Count) newItem.Offsets.Add("");
                            newItem.ChildCount = newItem.ChildIds.Count;
                            result.Add(newItem);
                        }
                    }
                }
            }
            return result;
        }

        // =========================================================================================
        // [SHOW REPORT & QUẢN LÝ WINDOW]
        // =========================================================================================
        public void ShowReportProMax(List<RelationshipItem> activeRelationships, List<RevitLinkInstance> links, ExternalEvent exEvent, StructureSyncRequestHandler handler, UIApplication uiapp)
        {
            HighlightUnsynced(activeRelationships, links);

            var reports = ReportGenerator.Analyze(_core, activeRelationships, links, _doc);
            if (reports.Count == 0)
            {
                TaskDialog.Show("Info", "No active relationships to report.");
                return;
            }

            CleanupOldSession();

            _activeReportWindow = new StructureReportWindow(reports, exEvent, handler, activeRelationships);
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(_activeReportWindow);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            _cachedUiApp = uiapp;
            _selectionHandler = OnRevitSelectionChanged;
            _cachedUiApp.SelectionChanged += _selectionHandler;

            _activeReportWindow.Closed += OnReportWindowClosed;
            _activeReportWindow.Show();

            handler.ReportWindow = _activeReportWindow;
        }

        private void CleanupOldSession()
        {
            try { if (_cachedUiApp != null && _selectionHandler != null) _cachedUiApp.SelectionChanged -= _selectionHandler; }
            catch { }
            finally { _cachedUiApp = null; _selectionHandler = null; }

            if (_activeReportWindow != null)
            {
                try { if (_activeReportWindow.IsLoaded) _activeReportWindow.Close(); }
                catch { }
                finally { _activeReportWindow = null; }
            }
        }

        private static void OnRevitSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e)
        {
            if (_activeReportWindow == null) return;
            try
            {
                if (!_activeReportWindow.IsLoaded) return;
                var selectedIds = e.GetSelectedElements();
                if (selectedIds.Count == 0) return;
                _activeReportWindow.AutoSelectRow(selectedIds);
            }
            catch
            {
                if (_cachedUiApp != null && _selectionHandler != null) { try { _cachedUiApp.SelectionChanged -= _selectionHandler; } catch { } _cachedUiApp = null; }
            }
        }

        private void OnReportWindowClosed(object sender, EventArgs e) => CleanupOldSession();
    }
}
