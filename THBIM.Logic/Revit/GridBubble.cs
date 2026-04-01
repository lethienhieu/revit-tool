using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System; // Để dùng Func
using System.Collections.Generic;
using System.Linq;

namespace GridBubble
{
    public class GridBubbleExternalHandler : IExternalEventHandler
    {
        public BubbleMode Mode { get; set; } = BubbleMode.Show;
        public bool SessionRunning { get; private set; } = false;

        private readonly GridBubbleWindow _owner;
        public GridBubbleExternalHandler(GridBubbleWindow owner) => _owner = owner;

        public void Execute(UIApplication app)
        {
            if (SessionRunning) return;
            SessionRunning = true;

            try
            {
                while (true)
                {
                    // FIX: Truyền delegate gọi về _owner để lấy Mode TƯƠI SỐNG từ UI
                    bool esc = GridBubbleCore.RunOnePick(app, () =>
                        _owner.GetModeDirectly() == BubbleMode.Show
                    );

                    if (esc) break;
                }
            }
            finally
            {
                SessionRunning = false;
                try { _owner.EndSessionAndClose(); } catch { }
            }
        }
        public string GetName() => "GridBubble Continuous Session";
    }

    // ======================== CORE: 1 lần Pick ===========================
    public static class GridBubbleCore
    {
        // Nhận vào Func<bool> để lấy trạng thái UI mới nhất
        public static bool RunOnePick(UIApplication uiapp, Func<bool> checkIsShowMode)
        {
            var uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null) return true;

            View view = uidoc.Document.ActiveView;
            if (view == null || view.ViewType == ViewType.ThreeD || view.ViewType == ViewType.DrawingSheet)
                return true;

            // 1. PickBox
            PickedBox picked;
            try
            {
                picked = uidoc.Selection.PickBox(
                    PickBoxStyle.Crossing,
                    "Select the region containing Grid/Level heads (Supports Multi-Segment). Press Esc to exit."
                );
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return true; // ESC
            }
            catch
            {
                return false;
            }
            if (picked == null) return false;

            // 2. Lấy chế độ hiện tại từ UI
            bool setVisible = checkIsShowMode();

            // 3. Chuẩn bị Matrix và Box
            Transform w2v = GetWorldToViewTransform(view);
            XYZ vMin = w2v.OfPoint(picked.Min);
            XYZ vMax = w2v.OfPoint(picked.Max);
            double minX = Math.Min(vMin.X, vMax.X);
            double maxX = Math.Max(vMin.X, vMax.X);
            double minY = Math.Min(vMin.Y, vMax.Y);
            double maxY = Math.Max(vMin.Y, vMax.Y);
            double tol = 1e-6; // Dung sai

            bool isSectionLike = view is ViewSection || view.ViewType == ViewType.Section || view.ViewType == ViewType.Elevation;
            bool isPlanLike = view is ViewPlan || view.ViewType == ViewType.FloorPlan || view.ViewType == ViewType.CeilingPlan || view.ViewType == ViewType.AreaPlan || view.ViewType == ViewType.EngineeringPlan;

            Document doc = uidoc.Document;

            // Lấy tất cả Grid (bao gồm cả các segment con của Multi-Segment Grid)
            var grids = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Grid)).Cast<Grid>().ToList();

            var levels = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Level)).Cast<Level>().ToList();

            using (var tx = new Transaction(doc, (setVisible ? "Show" : "Hide") + " Bubble"))
            {
                try
                {
                    tx.Start();

                    // ===== GRIDS (Xử lý nâng cao cho Multi-Segment) =====
                    foreach (Grid grid in grids)
                    {
                        // Lấy Model Curve (Geometry gốc 3D)
                        Curve modelCurve = grid.Curve;
                        if (modelCurve == null) continue;
                        if (!TryGetEndPoints(modelCurve, out XYZ? pt0, out XYZ? pt1) || pt0 == null || pt1 == null) continue;

                        // HashSet để lưu các đầu cần bật/tắt (tránh trùng lặp nếu cả 2 phương pháp đều tìm thấy)
                        HashSet<DatumEnds> endsToToggle = new HashSet<DatumEnds>();

                        // --- CÁCH 1: KIỂM TRA THEO VIEW CURVES (Visual 2D - Grid thường hay dùng) ---
                        IList<Curve> vCurves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view);
                        if (vCurves == null || vCurves.Count == 0)
                            vCurves = grid.GetCurvesInView(DatumExtentType.Model, view);

                        // --- CÁCH 2: KIỂM TRA THEO MODEL CURVE (Physical 3D - Cứu cánh cho Multi-Segment) ---
                        // Tạo một danh sách phụ chứa ModelCurve đã chiếu xuống View
                        // (Dùng để bắt dính ngay cả khi Revit API ẩn mất đường 2D)
                        List<Curve> fallbackCurves = new List<Curve> { modelCurve };

                        // Gom cả 2 nguồn curves lại để check
                        List<IList<Curve>> curveSources = new List<IList<Curve>>();
                        if (vCurves != null && vCurves.Count > 0) curveSources.Add(vCurves);
                        curveSources.Add(fallbackCurves);

                        foreach (var curvesToCheck in curveSources)
                        {
                            if (isPlanLike || (!isPlanLike && !isSectionLike))
                            {
                                if (CollectEndsInPlan(curvesToCheck, w2v, minX, maxX, minY, maxY, tol, pt0, pt1, out var e))
                                    foreach (var end in e) endsToToggle.Add(end);
                            }
                            else // Section-like
                            {
                                XYZ vPt0 = w2v.OfPoint(pt0);
                                XYZ vPt1 = w2v.OfPoint(pt1);
                                DatumEnds datumTop = (vPt0.Y >= vPt1.Y) ? DatumEnds.End1 : DatumEnds.End0;
                                DatumEnds datumBot = (vPt0.Y >= vPt1.Y) ? DatumEnds.End0 : DatumEnds.End1;

                                if (CollectEndsInSection(curvesToCheck, w2v, minX, maxX, minY, maxY, tol, datumTop, datumBot, out var e))
                                    foreach (var end in e) endsToToggle.Add(end);
                            }
                        }

                        // Thực hiện lệnh trên các đầu đã tìm thấy
                        foreach (var end in endsToToggle)
                        {
                            ApplyDatum(grid, view, setVisible, end);
                        }
                    }

                    // ===== LEVELS (Section/Elevation) - Level ít khi bị lỗi này nên giữ nguyên logic tối ưu =====
                    if (isSectionLike)
                    {
                        foreach (Level level in levels)
                        {
                            if (!(level is DatumPlane dp)) continue;
                            IList<Curve> vCurves = dp.GetCurvesInView(DatumExtentType.ViewSpecific, view);
                            if (vCurves == null || vCurves.Count == 0) vCurves = dp.GetCurvesInView(DatumExtentType.Model, view);
                            if (vCurves == null || vCurves.Count == 0) continue;

                            if (CollectEndsLeftRight(vCurves, w2v, minX, maxX, minY, maxY, tol, out var endsLR))
                                foreach (var end in endsLR) ApplyDatum(dp, view, setVisible, end);
                        }
                    }

                    tx.Commit();
                }
                catch
                {
                    if (tx.GetStatus() == TransactionStatus.Started) tx.RollBack();
                }
            }

            return false; // Chưa ESC -> Tiếp tục
        }

        // ---------- Helpers ----------
        private static Transform GetWorldToViewTransform(View view)
        {
            Transform t = Transform.Identity;
            t.Origin = view.Origin;
            t.BasisX = view.RightDirection;
            t.BasisY = view.UpDirection;
            t.BasisZ = view.ViewDirection;
            return t.Inverse;
        }

        private static bool TryGetEndPoints(Curve c, out XYZ? p0, out XYZ? p1)
        {
            p0 = null; p1 = null;
            try { p0 = c.GetEndPoint(0); p1 = c.GetEndPoint(1); return true; }
            catch { return false; }
        }

        private static bool Inside2D(XYZ p, double minX, double maxX, double minY, double maxY, double tol)
        {
            return p.X >= minX - tol && p.X <= maxX + tol && p.Y >= minY - tol && p.Y <= maxY + tol;
        }

        // Map điểm tìm thấy (có thể là View hoặc Model) về End0/End1 của Model gốc
        private static DatumEnds MapDatumByNearestModelEnd(XYZ p, XYZ pt0, XYZ pt1)
        {
            // So sánh khoảng cách trong không gian 3D để đảm bảo chính xác
            double d0 = p.DistanceTo(pt0);
            double d1 = p.DistanceTo(pt1);
            return (d0 <= d1) ? DatumEnds.End1 : DatumEnds.End0;
            // Lưu ý: Grid.Curve.GetEndPoint(0) thường ứng với End1 (Start) trong Enum, tùy thuộc vào cách vẽ,
            // nhưng logic Distance này sẽ luôn map về đúng cái đầu gần nhất.
        }

        private static bool CollectEndsInPlan(IList<Curve> curves, Transform w2v,
            double minX, double maxX, double minY, double maxY, double tol,
            XYZ pt0, XYZ pt1, out HashSet<DatumEnds> ends)
        {
            ends = new HashSet<DatumEnds>();
            foreach (var c in curves)
            {
                if (!TryGetEndPoints(c, out XYZ? sp, out XYZ? ep) || sp == null || ep == null) continue;

                // Nếu là ModelCurve (check fallback), cần project điểm 3D xuống View 2D
                // Nếu là ViewCurve, nó đã ở vị trí hiển thị, nhưng vẫn cần qua w2v để so với Box chuột

                // Logic w2v.OfPoint xử lý tốt cả 2 trường hợp (chiếu điểm không gian vào mặt phẳng view)
                XYZ vSp = w2v.OfPoint(sp);
                XYZ vEp = w2v.OfPoint(ep);

                if (Inside2D(vSp, minX, maxX, minY, maxY, tol)) ends.Add(MapDatumByNearestModelEnd(sp, pt0, pt1));
                if (Inside2D(vEp, minX, maxX, minY, maxY, tol)) ends.Add(MapDatumByNearestModelEnd(ep, pt0, pt1));
            }
            return ends.Count > 0;
        }

        private static bool CollectEndsInSection(IList<Curve> curves, Transform w2v,
            double minX, double maxX, double minY, double maxY, double tol,
            DatumEnds datumTop, DatumEnds datumBottom, out HashSet<DatumEnds> ends)
        {
            ends = new HashSet<DatumEnds>();
            foreach (var c in curves)
            {
                if (!TryGetEndPoints(c, out XYZ? sp, out XYZ? ep) || sp == null || ep == null) continue;

                XYZ vSp = w2v.OfPoint(sp);
                XYZ vEp = w2v.OfPoint(ep);

                bool spIn = Inside2D(vSp, minX, maxX, minY, maxY, tol);
                bool epIn = Inside2D(vEp, minX, maxX, minY, maxY, tol);

                if (spIn) ends.Add(vSp.Y >= vEp.Y - tol ? datumTop : datumBottom);
                if (epIn) ends.Add(vEp.Y >= vSp.Y - tol ? datumTop : datumBottom);
            }
            return ends.Count > 0;
        }

        private static bool CollectEndsLeftRight(IList<Curve> curves, Transform w2v,
            double minX, double maxX, double minY, double maxY, double tol,
            out HashSet<DatumEnds> ends)
        {
            ends = new HashSet<DatumEnds>();
            foreach (var c in curves)
            {
                if (!TryGetEndPoints(c, out XYZ? sp, out XYZ? ep) || sp == null || ep == null) continue;
                XYZ vSp = w2v.OfPoint(sp);
                XYZ vEp = w2v.OfPoint(ep);

                if (Inside2D(vSp, minX, maxX, minY, maxY, tol)) ends.Add(vSp.X <= vEp.X + tol ? DatumEnds.End0 : DatumEnds.End1);
                if (Inside2D(vEp, minX, maxX, minY, maxY, tol)) ends.Add(vEp.X <= vSp.X + tol ? DatumEnds.End0 : DatumEnds.End1);
            }
            return ends.Count > 0;
        }

        private static void ApplyDatum(Element d, View v, bool setVisible, DatumEnds e)
        {
            try
            {
                // Grid Segment (Multi-segment) đôi khi throw lỗi nếu set bubble ở đoạn giữa
                // Ta bọc try-catch để lờ đi các lỗi không cần thiết
                if (d is Grid g)
                {
                    bool vis = g.IsBubbleVisibleInView(e, v);
                    if (setVisible && !vis) g.ShowBubbleInView(e, v);
                    else if (!setVisible && vis) g.HideBubbleInView(e, v);
                }
                else if (d is DatumPlane dp)
                {
                    bool vis = dp.IsBubbleVisibleInView(e, v);
                    if (setVisible && !vis) dp.ShowBubbleInView(e, v);
                    else if (!setVisible && vis) dp.HideBubbleInView(e, v);
                }
            }
            catch { /* Ignore errors on segments that don't support bubbles */ }
        }
    }
}