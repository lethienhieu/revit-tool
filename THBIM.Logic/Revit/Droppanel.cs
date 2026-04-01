using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
#nullable disable

// Thư viện UI (WPF) - Chỉ định rõ namespace để tránh lỗi
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Shell;

namespace THBIM.Revit
{
    public static class DroppanelCore
    {
        public static void Run(Document doc, List<Element> columns, List<Element> inputFloors, FamilySymbol panelSymbol, double userLengthMm, bool isLineMode)
        {
            int successCount = 0;
            List<string> errorLogs = new List<string>();

            // 1. Khởi tạo thanh Progress Bar
            using (var pb = new SmoothProgressBarWindow(columns.Count))
            {
                pb.Show();

                Stopwatch uiWatch = new Stopwatch();
                uiWatch.Start();

                using (Transaction t = new Transaction(doc, "THBIM: DropPanel & PileCap"))
                {
                    t.Start();

                    if (!panelSymbol.IsActive) { panelSymbol.Activate(); doc.Regenerate(); }

                    View3D view3D = Get3DView(doc);
                    if (view3D == null)
                    {
                        TaskDialog.Show("Error", "3D View is required!");
                        t.RollBack();
                        return;
                    }

                    double lengthFeet = userLengthMm / 304.8;
                    var validFloors = inputFloors.Distinct(new ElementIdEqualityComparer()).ToList();

                    bool isFoundation = false;
                    if (panelSymbol.Category != null)
                    {
                        long catId = panelSymbol.Category.Id.GetValue();
                        if (catId == (long)BuiltInCategory.OST_StructuralFoundation) isFoundation = true;
                    }

                    if (isFoundation) isLineMode = false;

                    // --- VÒNG LẶP CHÍNH ---
                    for (int i = 0; i < columns.Count; i++)
                    {
                        // Cập nhật UI mỗi 50ms để tránh lag
                        if (uiWatch.ElapsedMilliseconds > 50)
                        {
                            pb.UpdateText($"Processing Column {i + 1}/{columns.Count}... Happiness is in the waiting");
                            pb.UpdateProgress(i);
                            uiWatch.Restart();
                        }

                        Element col = columns[i];
                        try
                        {
                            BoundingBoxXYZ colBB = col.get_BoundingBox(null);
                            if (colBB == null) continue;
                            XYZ centerBB = (colBB.Min + colBB.Max) / 2.0;

                            // [FIX QUAN TRỌNG] Lấy Max.Z (Đỉnh cột) làm gốc bắn tia
                            // Để đảm bảo tìm được sàn ở các tầng cao (ví dụ tầng 15)
                            XYZ colPoint = new XYZ(centerBB.X, centerBB.Y, colBB.Max.Z);

                            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(new Outline(colBB.Min - new XYZ(0.1, 0.1, 0.1), colBB.Max + new XYZ(0.1, 0.1, 0.1)));
                            var candidateFloors = validFloors.Where(f => bbFilter.PassesFilter(f)).ToList();

                            if (candidateFloors.Count == 0) continue;

                            foreach (Element floor in candidateFloors)
                            {
                                try
                                {
                                    // 2. Bắn tia tìm giao điểm
                                    XYZ hitPoint;
                                    Reference faceRef = GetFloorFaceReference(floor, colPoint, view3D, out hitPoint);
                                    if (faceRef == null) continue;

                                    Level floorLevel = doc.GetElement(floor.LevelId) as Level;

                                    // 3. Kiểm tra giao cắt (Penetration Check)
                                    double floorTopZ = hitPoint.Z;
                                    double floorThickness = 0;
                                    Parameter pThick = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                                    if (pThick != null) floorThickness = pThick.AsDouble();
                                    double floorBottomZ = floorTopZ - floorThickness;

                                    double colBaseZ = 0;
                                    double colTopZ = 0;
                                    GetColumnElevations(doc, col, out colBaseZ, out colTopZ);

                                    double tolerance = 0.001;

                                    // [UPDATE] Logic kiểm tra hợp lệ
                                    bool skipBase = false;
                                    if (isFoundation)
                                    {
                                        // Pile Cap: Chấp nhận Base == FloorTop
                                        // Chỉ bỏ qua nếu Base > FloorTop
                                        if (colBaseZ > floorTopZ + tolerance) skipBase = true;
                                    }
                                    else
                                    {
                                        // Drop Panel: Base phải < FloorTop
                                        if (colBaseZ >= floorTopZ - tolerance) skipBase = true;
                                    }

                                    if (skipBase) continue;

                                    // Kiểm tra đỉnh cột (phải chạm hoặc xuyên qua đáy sàn)
                                    if (colTopZ <= floorBottomZ + tolerance) continue;

                                    // 4. Tạo đối tượng
                                    FamilyInstance instance = null;

                                    if (isLineMode)
                                    {
                                        instance = CreateLineBasedInstance(doc, col, hitPoint, panelSymbol, floorLevel, lengthFeet);
                                    }
                                    else
                                    {
                                        double finalOffsetZ = hitPoint.Z - floorLevel.Elevation;

                                        if (isFoundation)
                                        {
                                            // Pile Cap: Cắt chân cột
                                            Parameter pBaseLvl = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                                            Parameter pBaseOff = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);

                                            if (pBaseLvl != null && pBaseOff != null)
                                            {
                                                pBaseLvl.Set(floorLevel.Id);
                                                pBaseOff.Set(finalOffsetZ);
                                            }
                                            doc.Regenerate();

                                            // Tạo Pile Cap
                                            instance = doc.Create.NewFamilyInstance(hitPoint, panelSymbol, floorLevel, StructuralType.Footing);

                                            // [FIX] Gán Offset ngay lập tức để không bị về 0
                                            if (instance != null) SetOffsetParameters(instance, finalOffsetZ);
                                        }
                                        else
                                        {
                                            // Drop Panel
                                            instance = doc.Create.NewFamilyInstance(hitPoint, panelSymbol, floorLevel, StructuralType.NonStructural);
                                            SetOffsetParameters(instance, finalOffsetZ);
                                        }

                                        // [UPDATE] Xoay thông minh (Cạnh dài theo Cạnh dài)
                                        doc.Regenerate();
                                        AlignRotationByLongestEdge(doc, col, instance, hitPoint);
                                    }

                                    if (instance != null) successCount++;
                                }
                                catch (Exception exInit) { errorLogs.Add($"Col {col.Id}: {exInit.Message}"); }
                            }
                        }
                        catch (Exception ex) { errorLogs.Add($"General Error: {ex.Message}"); }
                    }

                    t.Commit();

                    // Hoàn thành
                    pb.UpdateProgress(columns.Count, "Done! (Happiness)");

                    // Chờ xíu cho đẹp rồi tự tắt
                    System.Threading.Thread.Sleep(300);
                }
            } // Hết khối using -> Cửa sổ Progress Bar tự đóng tại đây

            // --- HIỆN THÔNG BÁO SAU KHI CỬA SỔ ĐÃ ĐÓNG ---
            string resultMsg = successCount > 0
                ? $"Success: Created {successCount} elements!"
                : "No elements created (likely due to no valid geometric intersection).";

            if (errorLogs.Count > 0) resultMsg += $"\nFirst error details: {errorLogs[0]}";

            TaskDialog.Show("THBIM Result", resultMsg);
        }

        // ================= HELPER METHODS =================

        // [MỚI] Hàm xác định vector trục chính (cạnh dài)
        private static XYZ GetMajorAxisVector(Element elem)
        {
            FamilyInstance fi = elem as FamilyInstance;
            if (fi != null)
            {
                Autodesk.Revit.DB.Transform tf = fi.GetTransform();
                double lenX = 0, lenY = 0;

                Options opt = new Options { DetailLevel = ViewDetailLevel.Fine };
                GeometryElement geoElem = elem.get_Geometry(opt);

                if (geoElem != null)
                {
                    foreach (var obj in geoElem)
                    {
                        if (obj is GeometryInstance gi)
                        {
                            BoundingBoxXYZ bb = gi.GetSymbolGeometry().GetBoundingBox();
                            if (bb != null)
                            {
                                lenX = Math.Abs(bb.Max.X - bb.Min.X);
                                lenY = Math.Abs(bb.Max.Y - bb.Min.Y);
                            }
                            break;
                        }
                    }
                }
                return (lenY >= lenX) ? tf.BasisY : tf.BasisX;
            }
            return XYZ.BasisX;
        }

        // [MỚI] Hàm xoay cạnh dài theo cạnh dài
        private static void AlignRotationByLongestEdge(Document doc, Element column, FamilyInstance panel, XYZ rotationCenter)
        {
            try
            {
                XYZ vecCol = GetMajorAxisVector(column);
                XYZ vecPan = GetMajorAxisVector(panel);

                double angle = vecCol.AngleTo(vecPan);
                XYZ cross = vecCol.CrossProduct(vecPan);
                if (cross.Z > 0) angle = -angle;

                if (Math.Abs(angle) > 0.001)
                {
                    Line axis = Line.CreateBound(rotationCenter, rotationCenter + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, panel.Id, axis, angle);
                }
            }
            catch { }
        }

        private static void GetColumnElevations(Document doc, Element col, out double baseZ, out double topZ)
        {
            baseZ = 0;
            topZ = 0;
            try
            {
                Parameter pBaseLvl = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                Parameter pBaseOff = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                if (pBaseLvl != null)
                {
                    Level l = doc.GetElement(pBaseLvl.AsElementId()) as Level;
                    baseZ = (l != null ? l.Elevation : 0) + (pBaseOff != null ? pBaseOff.AsDouble() : 0);
                }

                Parameter pTopLvl = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                Parameter pTopOff = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);
                if (pTopLvl != null)
                {
                    Level l = doc.GetElement(pTopLvl.AsElementId()) as Level;
                    topZ = (l != null ? l.Elevation : 0) + (pTopOff != null ? pTopOff.AsDouble() : 0);
                }
            }
            catch { }
        }

        private static Reference GetFloorFaceReference(Element floor, XYZ colPoint, View3D view3d, out XYZ hitPoint)
        {
            hitPoint = null;
            try
            {
                ReferenceIntersector ri = new ReferenceIntersector(new List<ElementId> { floor.Id }, FindReferenceTarget.Face, view3d);
                ri.FindReferencesInRevitLinks = false;

                // [FIX] Bắn từ trên cao xuống 1 chút (colPoint giờ là Max.Z)
                XYZ rayOrigin = new XYZ(colPoint.X, colPoint.Y, colPoint.Z + 1.0);
                XYZ rayDir = new XYZ(0, 0, -1);

                ReferenceWithContext rwc = ri.FindNearest(rayOrigin, rayDir);
                if (rwc != null)
                {
                    Reference r = rwc.GetReference();
                    hitPoint = r.GlobalPoint;
                    return r;
                }
            }
            catch { }
            return null;
        }

        private static void SetOffsetParameters(FamilyInstance instance, double offsetValue)
        {
            Parameter pGen = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
            if (pGen != null && !pGen.IsReadOnly) { pGen.Set(offsetValue); return; }

            Parameter pFound = instance.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
            if (pFound != null && !pFound.IsReadOnly) { pFound.Set(offsetValue); return; }

            Parameter pS = instance.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);
            Parameter pE = instance.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION);
            if (pS != null && !pS.IsReadOnly) { pS.Set(offsetValue); pE?.Set(offsetValue); }
        }

        public class ElementIdEqualityComparer : IEqualityComparer<Element>
        {
            public bool Equals(Element x, Element y) => x.Id == y.Id;
            public int GetHashCode(Element obj) => obj.Id.GetHashCode();
        }

        private static FamilyInstance CreateLineBasedInstance(Document doc, Element col, XYZ centerPoint, FamilySymbol symbol, Level level, double lengthFeet)
        {
            // Dầm line-based mặc định theo trục dọc nên dùng GetMajorAxisVector là chuẩn
            XYZ colDir = GetMajorAxisVector(col);
            XYZ halfVec = colDir * (lengthFeet / 2.0);
            XYZ startPt = centerPoint - halfVec;
            XYZ endPt = centerPoint + halfVec;
            Curve structuralLine = Line.CreateBound(startPt, endPt);
            return doc.Create.NewFamilyInstance(structuralLine, symbol, level, StructuralType.Beam);
        }

        private static View3D Get3DView(Document doc) => new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate && !v.IsPerspective);

        // ================= UI CLASS (ĐÃ SỬA LỖI CS0104) =================
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
                    // [FIX] Explicit System.Windows.Media.Color
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

                // [FIX] Explicit System.Windows.Controls.Grid
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

            public void UpdateText(string msg)
            {
                _txtStatus.Text = msg;
                DoEvents();
            }

            public void UpdateProgress(double current, string statusMsg = null)
            {
                if (statusMsg != null) _txtStatus.Text = statusMsg;

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