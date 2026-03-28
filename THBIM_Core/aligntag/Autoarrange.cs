// ArrangeTagsNoCrossCommand.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class ArrangeTagsNoCrossCommand : IExternalCommand
    {
        private const double TOL_FT = 1.0 / 304.8;     // ~1 mm
        private const double PAD_MM = 1.0;             // đệm giữa các tag
        private static readonly double PAD_FT = PAD_MM / 304.8;
        private const double FALLBACK_SIZE_MM = 1.0;   // nếu không đọc được BB
        private static readonly double FALLBACK_SIZE_FT = FALLBACK_SIZE_MM / 304.8;

        private enum ArrangeAxis { Vertical, Horizontal }

        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }



            var uidoc = cd.Application.ActiveUIDocument;
            var doc = uidoc.Document;
            var view = doc.ActiveView;

            if (!Is2DView(view))
            {
                msg = "Chỉ hỗ trợ trong các view 2D (Plan/Elevation/Section/Drafting).";
                return Result.Failed;
            }

            var tags = GetTagsFromSelectionOrPick(uidoc, doc, view);
            if (tags == null) return Result.Cancelled;
            if (tags.Count < 2) return Result.Cancelled;

            XYZ right = view.RightDirection;
            XYZ up = view.UpDirection;

            // Thu thập info tag
            var infos = new List<TagInfo>(tags.Count);
            foreach (var tag in tags)
            {
                XYZ head = GetTagKeyPoint(tag, view);
                if (head == null) continue;

                // Dùng kích thước ước lượng cố định theo view scale (bỏ qua BB chứa leader)
                double w = 10.0 * view.Scale / 304.8;
                double h = 3.5 * view.Scale / 304.8;

                // Tìm anchor (tâm BB của host element trong view), nếu không có thì fallback vị trí tag
                XYZ anchor = TryGetHostAnchorCenter(tag, view) ?? head;

                infos.Add(new TagInfo(tag,
                    head, Dot(head, right), Dot(head, up),
                    w, h,
                    anchor, Dot(anchor, right), Dot(anchor, up)));
            }

            if (infos.Count < 2)
                return Result.Cancelled;

            // Nhận biết trục giãn (tương tự ArrangeTagsNeat)
            double rangeX = infos.Max(i => i.HeadX) - infos.Min(i => i.HeadX);
            double rangeY = infos.Max(i => i.HeadY) - infos.Min(i => i.HeadY);
            ArrangeAxis axis =
                (rangeX <= 3 * TOL_FT) ? ArrangeAxis.Vertical :
                (rangeY <= 3 * TOL_FT) ? ArrangeAxis.Horizontal :
                ArrangeAxis.Vertical;

            // ===== Thứ tự: tag gần element nhất → dưới, xa dần → lên trên =====
            // Tính khoảng cách từ anchor (host) tới tâm nhóm tag trong view plane
            double avgHeadX = infos.Average(i => i.HeadX);
            double avgHeadY = infos.Average(i => i.HeadY);
            infos = infos.OrderBy(i =>
            {
                double dx = i.AnchorX - avgHeadX;
                double dy = i.AnchorY - avgHeadY;
                return dx * dx + dy * dy;
            }).ToList();

            int moved = 0, skippedPinned = 0, failed = 0;

            using (var t = new Transaction(doc, "THBIM - Arrange Tags (No Cross)"))
            {
                t.Start();

                // === Bi-directional spread từ tâm ===
                // Bước 1: Tính target positions (chưa move) — giãn đều từ tâm
                double[] targets = ComputeSpreadPositions(infos, axis);

                // Bước 2: Move từng tag đến target
                for (int i = 0; i < infos.Count; i++)
                {
                    var cur = infos[i];
                    if (cur.Tag.Pinned) { skippedPinned++; continue; }

                    double current = axis == ArrangeAxis.Vertical ? cur.HeadY : cur.HeadX;
                    double delta1d = targets[i] - current;

                    if (Math.Abs(delta1d) > TOL_FT)
                    {
                        XYZ delta = axis == ArrangeAxis.Vertical
                            ? up.Multiply(delta1d)
                            : right.Multiply(delta1d);

                        if (TryMoveTag(cur.Tag, delta))
                        {
                            moved++;
                            cur.Target = cur.Head + delta;
                        }
                        else
                        {
                            failed++;
                            cur.Target = cur.Head;
                        }
                    }

                    // Luôn apply leader angle cho mọi tag (kể cả không di chuyển)
                    LeaderAngleHelper.ApplyIfNeeded(cur.Tag);
                }

                t.Commit();
            }

            return Result.Succeeded;
        }

        // ===== Helpers =====

        private static bool Is2DView(View v)
        {
            if (v == null) return false;
            return v.ViewType == ViewType.FloorPlan ||
                   v.ViewType == ViewType.EngineeringPlan ||
                   v.ViewType == ViewType.CeilingPlan ||
                   v.ViewType == ViewType.Elevation ||
                   v.ViewType == ViewType.Section ||
                   v.ViewType == ViewType.AreaPlan ||
                   v.ViewType == ViewType.DraftingView;
        }

        private static List<IndependentTag> GetTagsFromSelectionOrPick(UIDocument uidoc, Document doc, View view)
        {
            var picked = new List<IndependentTag>();

            // Dùng selection nếu hợp lệ
            var selIds = uidoc.Selection.GetElementIds();
            if (selIds != null && selIds.Count > 0)
            {
                foreach (var id in selIds)
                {
                    var el = doc.GetElement(id);
                    if (el is IndependentTag it && BelongsToView(it, view))
                        picked.Add(it);
                }
                if (picked.Count >= 2) return picked;
            }

            // Cho quét rectangle
            try
            {
                var rect = uidoc.Selection.PickElementsByRectangle(
                    new TagSelectionFilter(), "Quét chọn các Tag cần sắp xếp (no-cross)");

                foreach (var el in rect)
                {
                    if (el is IndependentTag it && BelongsToView(it, view))
                        picked.Add(it);
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }

            return picked;
        }

        private static bool BelongsToView(IndependentTag tag, View view)
        {
            if (tag.OwnerViewId == view.Id) return true;
            try { return tag.get_BoundingBox(view) != null; } catch { return false; }
        }

        private static BoundingBoxXYZ SafeGetBB(Element e, View v)
        {
            try { return e.get_BoundingBox(v); } catch { return null; }
        }

        private static XYZ GetTagKeyPoint(IndependentTag tag, View view)
        {
            // Ưu tiên đầu tag
            try
            {
                XYZ head = tag.TagHeadPosition;
                if (head != null) return head;
            }
            catch { }

            // Fallback: tâm BB
            var bb = SafeGetBB(tag, view);
            if (bb != null) return (bb.Min + bb.Max) * 0.5;

            // Fallback: LocationPoint
            if (tag.Location is LocationPoint lp && lp.Point != null) return lp.Point;

            return null;
        }

        /// <summary>
        /// Thử lấy tâm BB của phần tử được gắn tag trong view (anchor).
        /// Lưu ý: chỉ xử lý host nội bộ (không xử lý link trong ví dụ này).
        /// </summary>
        // Thay thế hoàn toàn hàm cũ bằng hàm này
        private static XYZ TryGetHostAnchorCenter(IndependentTag tag, View view)
        {
            // 1) Ưu tiên: lấy qua GetTaggedReferences() (ổn định nhiều version)
            try
            {
                // Trả về các Reference mà tag đang bám
                IList<Reference> refs = tag.GetTaggedReferences();
                if (refs != null)
                {
                    foreach (var r in refs)
                    {
                        if (r == null) continue;

                        // Host trong chính document: dùng ElementId
                        ElementId hostId = r.ElementId;
                        if (hostId != null && hostId != ElementId.InvalidElementId)
                        {
                            Element host = view.Document.GetElement(hostId);
                            if (host != null)
                            {
                                BoundingBoxXYZ bb = host.get_BoundingBox(view);
                                if (bb != null) return (bb.Min + bb.Max) * 0.5;
                            }
                        }
                        // Trường hợp host nằm trong LINK:
                        // Revit 2023 không expose sẵn LinkedElementId trên Reference như property public
                        // => nếu cần xử lý link, ta sẽ viết thêm nhánh theo hướng RevitLinkInstance + StableRep (sau).
                    }
                }
            }
            catch
            {
                // fallback bên dưới
            }

            // 2) Fallback: thử dùng reflection cho TaggedElementId (nếu API có ở bản khác)
            try
            {
                var piTaggedElementId = typeof(IndependentTag).GetProperty("TaggedElementId");
                if (piTaggedElementId != null)
                {
                    object linkElemId = piTaggedElementId.GetValue(tag); // kiểu LinkElementId ở một số bản
                    if (linkElemId != null)
                    {
                        var piHostId = linkElemId.GetType().GetProperty("HostElementId");
                        if (piHostId != null)
                        {
                            var hostIdObj = piHostId.GetValue(linkElemId);
                            if (hostIdObj is ElementId hostId && hostId != ElementId.InvalidElementId)
                            {
                                Element host = view.Document.GetElement(hostId);
                                if (host != null)
                                {
                                    BoundingBoxXYZ bb = host.get_BoundingBox(view);
                                    if (bb != null) return (bb.Min + bb.Max) * 0.5;
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            // 3) Bó tay → trả null (lệnh sẽ fallback sắp theo vị trí tag)
            return null;
        }

        /// <summary>
        /// Tính vị trí target cho mỗi tag, spread đều từ tâm (bi-directional).
        /// Tags ở nửa dưới/trái đẩy xuống/trái, nửa trên/phải đẩy lên/phải.
        /// </summary>
        private static double[] ComputeSpreadPositions(List<TagInfo> infos, ArrangeAxis axis)
        {
            int n = infos.Count;
            var result = new double[n];

            // Lấy current position theo trục
            var positions = new double[n];
            for (int i = 0; i < n; i++)
            {
                positions[i] = axis == ArrangeAxis.Vertical ? infos[i].HeadY : infos[i].HeadX;
            }

            // Lấy kích thước trung bình làm khoảng cách đều
            double avgSize = 0;
            for (int i = 0; i < n; i++)
                avgSize += axis == ArrangeAxis.Vertical ? infos[i].H_Up : infos[i].W_Right;
            avgSize /= n;
            double step = avgSize + PAD_FT; // khoảng cách đều giữa các tag

            // Tâm của nhóm tag hiện tại
            double centerCurrent = 0;
            for (int i = 0; i < n; i++) centerCurrent += positions[i];
            centerCurrent /= n;

            // Spread đều từ tâm với step cố định
            double totalSpan = step * (n - 1);
            double start = centerCurrent - totalSpan * 0.5;

            for (int i = 0; i < n; i++)
            {
                result[i] = start + i * step;
            }

            return result;
        }

        private static bool TryMoveTag(IndependentTag tag, XYZ delta)
        {
            if (delta == null || delta.GetLength() <= TOL_FT) return true;
            if (tag.Pinned) return false;

            try
            {
                if (tag.HasLeader)
                {
                    ElementTransformUtils.MoveElement(tag.Document, tag.Id, delta);
                }
                else
                {
                    XYZ cur = tag.TagHeadPosition;
                    if (cur != null)
                    {
                        tag.TagHeadPosition = cur + delta;
                    }
                    else
                    {
                        ElementTransformUtils.MoveElement(tag.Document, tag.Id, delta);
                    }
                }
                return true;
            }
            catch { return false; }
        }

        private static double Dot(XYZ p, XYZ dir) => p.X * dir.X + p.Y * dir.Y + p.Z * dir.Z;

        private static double GetExtentAlong(BoundingBoxXYZ bb, XYZ dir)
        {
            if (bb == null || dir == null) return 0.0;
            var t = bb.Transform ?? Transform.Identity;
            var min = double.PositiveInfinity;
            var max = double.NegativeInfinity;

            foreach (var c in GetCorners(bb))
            {
                var p = t.OfPoint(c);
                double d = Dot(p, dir);
                if (d < min) min = d;
                if (d > max) max = d;
            }
            double extent = max - min;
            return extent < 0 ? 0 : extent;
        }

        private static List<XYZ> GetCorners(BoundingBoxXYZ bb)
        {
            var a = bb.Min; var b = bb.Max;
            return new List<XYZ>
            {
                new XYZ(a.X,a.Y,a.Z), new XYZ(a.X,a.Y,b.Z),
                new XYZ(a.X,b.Y,a.Z), new XYZ(a.X,b.Y,b.Z),
                new XYZ(b.X,a.Y,a.Z), new XYZ(b.X,a.Y,b.Z),
                new XYZ(b.X,b.Y,a.Z), new XYZ(b.X,b.Y,b.Z),
            };
        }

        private class TagInfo
        {
            public IndependentTag Tag { get; }
            public XYZ Head { get; }
            public XYZ Target { get; set; }
            public double HeadX { get; }
            public double HeadY { get; }
            public double W_Right { get; }
            public double H_Up { get; }

            public XYZ Anchor { get; }
            public double AnchorX { get; }
            public double AnchorY { get; }

            public TagInfo(IndependentTag tag, XYZ head, double headX, double headY,
                           double wRight, double hUp, XYZ anchor, double ax, double ay)
            {
                Tag = tag;
                Head = head;
                Target = head;
                HeadX = headX;
                HeadY = headY;
                W_Right = wRight;
                H_Up = hUp;
                Anchor = anchor;
                AnchorX = ax;
                AnchorY = ay;
            }
        }

        private class TagSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is IndependentTag;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}
