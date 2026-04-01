using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
// UI Libraries
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Shell;
#nullable disable

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUISplit : IExternalCommand
    {
        private class ColumnProcessData
        {
            public FamilyInstance Column { get; set; }
            public List<SplitPoint> SplitPoints { get; set; }
            public ElementId OriginalTopLevelId { get; set; }
            public double OriginalTopOffset { get; set; }
            public XYZ Center { get; set; }
            public double BaseZ { get; set; }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                {
                    return Result.Cancelled;
                }
                if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                    return Result.Cancelled;

                // 1. SELECT
                IList<Reference> refs = null;
                try
                {
                    refs = uidoc.Selection.PickObjects(ObjectType.Element, new ColumnSelectionFilter(), "Select Columns to split");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                if (refs == null || refs.Count == 0) return Result.Cancelled;

                View3D view3D = new FilteredElementCollector(doc)
                                .OfClass(typeof(View3D))
                                .Cast<View3D>()
                                .FirstOrDefault(v => !v.IsTemplate && !v.IsPerspective);

                if (view3D == null)
                {
                    TaskDialog.Show("Error", "A 3D View is required.");
                    return Result.Failed;
                }

                List<Element> allFloors = new FilteredElementCollector(doc)
                                          .OfCategory(BuiltInCategory.OST_Floors)
                                          .WhereElementIsNotElementType()
                                          .ToElements()
                                          .ToList();

                List<ElementId> columnsToDelete = new List<ElementId>();
                List<ColumnProcessData> processList = new List<ColumnProcessData>();

                // Setup Progress Bar
                using (var pb = new SmoothProgressBarWindow(refs.Count))
                {
                    pb.Show();

                    Stopwatch uiWatch = new Stopwatch();
                    uiWatch.Start();

                    // --- PHASE 1: PRE-CALCULATION ---
                    int count = 0;
                    foreach (Reference r in refs)
                    {
                        count++;
                        // Chỉ cập nhật UI mỗi 50ms để không làm chậm code chính
                        if (uiWatch.ElapsedMilliseconds > 50)
                        {
                            // [UPDATE] Thêm thông điệp "Hãy cười lên"
                            pb.UpdateText($"Happiness is in the waiting... (Analyzing {count}/{refs.Count})");
                            pb.UpdateProgress(0);
                            uiWatch.Restart();
                        }

                        Element elem = doc.GetElement(r);
                        FamilyInstance col = elem as FamilyInstance;
                        if (col == null) continue;

                        if (col.StructuralType != Autodesk.Revit.DB.Structure.StructuralType.Column &&
                            col.StructuralType != Autodesk.Revit.DB.Structure.StructuralType.NonStructural) continue;

                        XYZ colCenter = GetColumnCenter(col);
                        if (colCenter == null) continue;

                        double colBaseZ = GetColumnBaseElevation(doc, col);
                        double colTopZ = GetColumnTopElevation(doc, col);

                        List<SplitPoint> splitPoints = GetValidFloorIntersections(col, colCenter, allFloors, doc, view3D);

                        var validSplits = splitPoints
                            .Where(p => p.GlobalZ > colBaseZ + 0.003 && p.GlobalZ < colTopZ - 0.003)
                            .OrderBy(p => p.GlobalZ)
                            .ToList();

                        processList.Add(new ColumnProcessData
                        {
                            Column = col,
                            SplitPoints = validSplits,
                            OriginalTopLevelId = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId(),
                            OriginalTopOffset = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble(),
                            Center = colCenter,
                            BaseZ = colBaseZ
                        });
                    }

                    if (processList.Count == 0) return Result.Cancelled;

                    // --- PHASE 2: EXECUTION ---
                    pb.SetMax(processList.Count);

                    using (Transaction t = new Transaction(doc, "THBIM: Split Columns"))
                    {
                        t.Start();

                        FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new DeleteWarningSwallower());
                        t.SetFailureHandlingOptions(failOpt);

                        for (int i = 0; i < processList.Count; i++)
                        {
                            // Throttle UI: Chỉ update mỗi 50ms
                            if (uiWatch.ElapsedMilliseconds > 50)
                            {
                                // [UPDATE] Thêm thông điệp "Hãy cười lên"
                                pb.UpdateText($"Happiness is in the waiting... (Processing {i}/{processList.Count})");
                                pb.UpdateProgress(i);
                                uiWatch.Restart();
                            }

                            var data = processList[i];
                            FamilyInstance col = data.Column;
                            var validSplits = data.SplitPoints;

                            if (validSplits.Count > 0)
                            {
                                List<ElementId> allSegmentIds = new List<ElementId>();
                                allSegmentIds.Add(col.Id);

                                if (validSplits.Count > 0)
                                {
                                    for (int k = 0; k < validSplits.Count; k++)
                                    {
                                        var copiedIds = ElementTransformUtils.CopyElement(doc, col.Id, XYZ.Zero);
                                        allSegmentIds.Add(copiedIds.First());
                                    }
                                }

                                UpdateColumnConstraint(doc, allSegmentIds[0], null, 0, validSplits[0].LevelId, validSplits[0].Offset);

                                for (int k = 1; k < allSegmentIds.Count - 1; k++)
                                {
                                    UpdateColumnConstraint(doc, allSegmentIds[k],
                                        validSplits[k - 1].LevelId, validSplits[k - 1].Offset,
                                        validSplits[k].LevelId, validSplits[k].Offset);
                                }

                                UpdateColumnConstraint(doc, allSegmentIds.Last(),
                                    validSplits.Last().LevelId, validSplits.Last().Offset,
                                    data.OriginalTopLevelId, data.OriginalTopOffset);

                                if (!IsSolidFloorAtElevation(doc, view3D, data.Center, data.BaseZ, allFloors))
                                {
                                    columnsToDelete.Add(col.Id);
                                }
                            }
                        }

                        t.Commit();
                    }

                    pb.UpdateProgress(processList.Count, "Done! (Happiness is in the waiting)");
                }

                // --- DELETE PROMPT ---
                if (columnsToDelete.Count > 0)
                {
                    TaskDialog td = new TaskDialog("Clean Up");
                    td.MainInstruction = $"Split complete. Found {columnsToDelete.Count} floating segments.";
                    td.MainContent = "Do you want to delete them?";
                    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                    if (td.Show() == TaskDialogResult.Yes)
                    {
                        using (Transaction tDel = new Transaction(doc, "Delete Columns"))
                        {
                            tDel.Start();
                            doc.Delete(columnsToDelete);
                            tDel.Commit();
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ================= HELPERS =================

        public class DeleteWarningSwallower : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
                foreach (FailureMessageAccessor f in failures)
                {
                    if (f.GetSeverity() == FailureSeverity.Warning)
                    {
                        failuresAccessor.DeleteWarning(f);
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }

        private class SplitPoint { public ElementId LevelId; public double Offset; public double GlobalZ; }

        private bool IsSolidFloorAtElevation(Document doc, View3D view3D, XYZ centerPt, double targetZ, List<Element> allFloors)
        {
            XYZ rayOrigin = new XYZ(centerPt.X, centerPt.Y, targetZ + 1.0);
            XYZ rayDir = new XYZ(0, 0, -1);

            List<ElementId> floorIds = allFloors.Select(f => f.Id).ToList();
            if (floorIds.Count == 0) return false;

            ReferenceIntersector ri = new ReferenceIntersector(floorIds, FindReferenceTarget.Face, view3D);
            ri.FindReferencesInRevitLinks = false;
            ReferenceWithContext rwc = ri.FindNearest(rayOrigin, rayDir);

            if (rwc != null && Math.Abs(rwc.GetReference().GlobalPoint.Z - targetZ) < 0.05) return true;
            return false;
        }

        private List<SplitPoint> GetValidFloorIntersections(Element col, XYZ centerPt, List<Element> allFloors, Document doc, View3D view3D)
        {
            List<SplitPoint> points = new List<SplitPoint>();
            BoundingBoxXYZ colBB = col.get_BoundingBox(null);
            if (colBB == null) return points;

            Outline outline = new Outline(colBB.Min, colBB.Max);
            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);
            var candidateFloors = allFloors.Where(f => bbFilter.PassesFilter(f)).ToList();
            if (candidateFloors.Count == 0) return points;

            List<ElementId> floorIds = candidateFloors.Select(f => f.Id).ToList();
            ReferenceIntersector ri = new ReferenceIntersector(floorIds, FindReferenceTarget.Face, view3D);
            ri.FindReferencesInRevitLinks = false;

            XYZ rayOrigin = new XYZ(centerPt.X, centerPt.Y, colBB.Max.Z + 5.0);
            XYZ rayDir = new XYZ(0, 0, -1);
            IList<ReferenceWithContext> hits = ri.Find(rayOrigin, rayDir);

            foreach (ReferenceWithContext rwc in hits)
            {
                Reference r = rwc.GetReference();
                Element floor = doc.GetElement(r.ElementId);
                if (floor == null) continue;
                Level floorLevel = doc.GetElement(floor.LevelId) as Level;
                if (floorLevel == null) continue;
                double floorOffset = 0;
                Parameter pOff = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                if (pOff != null) floorOffset = pOff.AsDouble();
                points.Add(new SplitPoint { LevelId = floorLevel.Id, Offset = floorOffset, GlobalZ = floorLevel.Elevation + floorOffset });
            }
            return points.GroupBy(p => Math.Round(p.GlobalZ, 4)).Select(g => g.First()).ToList();
        }

        private XYZ GetColumnCenter(Element col) { BoundingBoxXYZ bb = col.get_BoundingBox(null); return bb == null ? null : (bb.Min + bb.Max) / 2.0; }

        private void UpdateColumnConstraint(Document doc, ElementId colId, ElementId baseLvlId, double baseOff, ElementId topLvlId, double topOff)
        {
            Element col = doc.GetElement(colId); if (col == null) return;
            if (baseLvlId != null) { col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).Set(baseLvlId); col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(baseOff); }
            if (topLvlId != null) { col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).Set(topLvlId); col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(topOff); }
        }

        private double GetColumnBaseElevation(Document doc, FamilyInstance col) { ElementId lvlId = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId(); double off = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble(); Level l = doc.GetElement(lvlId) as Level; return (l != null ? l.Elevation : 0) + off; }
        private double GetColumnTopElevation(Document doc, FamilyInstance col) { ElementId lvlId = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId(); double off = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble(); Level l = doc.GetElement(lvlId) as Level; return (l != null ? l.Elevation : 0) + off; }
        public class ColumnSelectionFilter : ISelectionFilter { public bool AllowElement(Element elem) => elem.Category != null && (elem.Category.Id.GetValue() == (int)BuiltInCategory.OST_StructuralColumns || elem.Category.Id.GetValue() == (int)BuiltInCategory.OST_Columns); public bool AllowReference(Reference reference, XYZ position) => false; }

        // ================= SMOOTH FAST UI =================
        public class SmoothProgressBarWindow : Window, IDisposable
        {
            private ProgressBar _pb;
            private TextBlock _txtPercent;
            private TextBlock _txtStatus;
            private double _max;

            public SmoothProgressBarWindow(double max)
            {
                _max = max;
                this.Title = "THBIM Processor";
                this.Width = 500;
                this.Height = 120;
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                this.ResizeMode = ResizeMode.NoResize;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = Brushes.Transparent;
                this.Topmost = true;

                Border mainBorder = new Border
                {
                    Background = Brushes.White,
                    CornerRadius = new CornerRadius(10),
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(20),
                    Effect = new System.Windows.Media.Effects.DropShadowEffect
                    {
                        BlurRadius = 15,
                        Color = System.Windows.Media.Colors.Gray,
                        Opacity = 0.3,
                        ShadowDepth = 5
                    }
                };

                StackPanel panel = new StackPanel();

                _txtStatus = new TextBlock
                {
                    Text = "Initializing...",
                    FontSize = 14,
                    FontWeight = FontWeights.Medium,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(60, 60, 60)),
                    Margin = new Thickness(0, 0, 0, 10)
                };

                System.Windows.Controls.Grid progressGrid = new System.Windows.Controls.Grid();

                _pb = new ProgressBar
                {
                    Minimum = 0,
                    Maximum = max,
                    Height = 8,
                    Value = 0,
                    BorderThickness = new Thickness(0),
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 240, 240)),
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 215))
                };

                _txtPercent = new TextBlock
                {
                    Text = "0%",
                    HorizontalAlignment = HorizontalAlignment.Right,
                    FontSize = 12,
                    Foreground = Brushes.Gray,
                    Margin = new Thickness(0, 15, 0, 0)
                };

                progressGrid.Children.Add(_pb);
                panel.Children.Add(_txtStatus);
                panel.Children.Add(progressGrid);
                panel.Children.Add(_txtPercent);
                mainBorder.Child = panel;
                this.Content = mainBorder;
            }

            public void SetMax(double max)
            {
                _max = max;
                _pb.Maximum = max;
            }

            public void UpdateText(string msg)
            {
                _txtStatus.Text = msg;
                DoEvents();
            }

            public void UpdateProgress(double current, string statusMsg = null)
            {
                if (statusMsg != null) _txtStatus.Text = statusMsg;

                // Animation nhanh 200ms để bắt kịp tốc độ xử lý nhanh
                DoubleAnimation animation = new DoubleAnimation
                {
                    To = current,
                    Duration = TimeSpan.FromMilliseconds(200),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };

                _pb.BeginAnimation(ProgressBar.ValueProperty, animation);

                int percent = (int)((current / _max) * 100);
                if (percent > 100) percent = 100;
                _txtPercent.Text = $"{percent}%";

                DoEvents();
            }

            private void DoEvents()
            {
                if (Application.Current != null)
                {
                    Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new Action(delegate { }));
                }
            }

            public void Dispose()
            {
                this.Close();
            }
        }
    }
}