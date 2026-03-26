// AlignTagsLeftCommand.cs
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
    public class AlignTagsLeftCommand : IExternalCommand
    {
        private const double TOL_FT = 1.0 / 304.8; // ~1 mm

        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }

            UIDocument uidoc = cd.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            // Chỉ hỗ trợ view 2D
            if (view == null ||
                !(view.ViewType == ViewType.FloorPlan ||
                  view.ViewType == ViewType.EngineeringPlan ||
                  view.ViewType == ViewType.CeilingPlan ||
                  view.ViewType == ViewType.Elevation ||
                  view.ViewType == ViewType.Section ||
                  view.ViewType == ViewType.AreaPlan ||
                  view.ViewType == ViewType.DraftingView))
            {
                msg = "Chỉ hỗ trợ căn tag trong các view 2D (Plan/Elevation/Section/Drafting).";
                return Result.Failed;
            }

            // Lấy danh sách tag người dùng chọn (selection hoặc quét rectangle)
            List<IndependentTag> tags = GetTagsFromSelectionOrPick(uidoc, doc, view);
            if (tags == null) return Result.Cancelled; // user ESC
            if (tags.Count < 2)
            {
                TaskDialog.Show("Align Left", "Cần quét chọn ít nhất 2 Tag.");
                return Result.Cancelled;
            }

            XYZ right = view.RightDirection;

            // Chuẩn bị dữ liệu: key point + toạ độ x theo RightDirection
            var infos = new List<TagInfo>(tags.Count);
            foreach (var tag in tags)
            {
                XYZ p = GetTagKeyPoint(tag, view);
                if (p == null) continue;
                double x = Dot(p, right);
                infos.Add(new TagInfo(tag, p, x));
            }

            if (infos.Count < 2)
            {
                TaskDialog.Show("Align Left", "Không xác định được vị trí Tag.");
                return Result.Cancelled;
            }

            // Anchor = tag có x nhỏ nhất (bên trái ngoài cùng)
            TagInfo anchor = infos.OrderBy(i => i.X).First();
            double xAnchor = anchor.X;

            int moved = 0, skippedPinned = 0, skippedTol = 0, failed = 0;

            using (Transaction t = new Transaction(doc, "THBIM - Align Tags Left"))
            {
                t.Start();

                foreach (var info in infos)
                {
                    if (info.Tag.Id == anchor.Tag.Id) continue;

                    if (info.Tag.Pinned)
                    {
                        skippedPinned++;
                        continue;
                    }

                    double dx = xAnchor - info.X;
                    if (Math.Abs(dx) <= TOL_FT)
                    {
                        skippedTol++;
                        continue;
                    }

                    XYZ delta = right.Multiply(dx);

                    bool done = TryMoveByHeadPosition(info.Tag, delta);
                    if (!done)
                    {
                        try
                        {
                            ElementTransformUtils.MoveElement(doc, info.Tag.Id, delta);
                            done = true;
                        }
                        catch
                        {
                            done = false;
                        }
                    }

                    if (done) moved++; else failed++;
                }

                t.Commit();
            }

            string summary =
                $"Tổng: {infos.Count} tag\n" +
                $"- Anchor (trái ngoài cùng): #{anchor.Tag.Id.GetValue()}\n" +
                $"- Đã căn: {moved}\n" +
                $"- Bỏ qua (pinned): {skippedPinned}\n" +
                $"- Bỏ qua (đã thẳng ~ tolerance): {skippedTol}\n" +
                $"- Lỗi/không di chuyển được: {failed}";
            TaskDialog.Show("Align Left", summary);

            return Result.Succeeded;
        }

        // ===== Helpers =====

        private static List<IndependentTag> GetTagsFromSelectionOrPick(UIDocument uidoc, Document doc, View view)
        {
            var picked = new List<IndependentTag>();

            // 1) Dùng selection hiện tại nếu có
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
                // nếu chưa đủ, cho quét tiếp
            }

            // 2) Quét rectangle để lấy đúng các Tag mày chọn
            try
            {
                IList<Element> rect = uidoc.Selection.PickElementsByRectangle(
                    new TagSelectionFilter(), "Quét chọn các Tag để căn trái");

                foreach (var el in rect)
                {
                    if (el is IndependentTag it && BelongsToView(it, view))
                        picked.Add(it);
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null; // ESC
            }

            return picked;
        }

        private static bool BelongsToView(IndependentTag tag, View view)
        {
            if (tag.OwnerViewId == view.Id) return true;
            try
            {
                var bb = tag.get_BoundingBox(view);
                return bb != null;
            }
            catch { return false; }
        }

        private static XYZ GetTagKeyPoint(IndependentTag tag, View view)
        {
            // Ưu tiên TagHeadPosition
            try
            {
                XYZ head = tag.TagHeadPosition;
                if (head != null) return head;
            }
            catch { }

            // Fallback: tâm BoundingBox
            try
            {
                var bb = tag.get_BoundingBox(view);
                if (bb != null) return (bb.Min + bb.Max) * 0.5;
            }
            catch { }

            // Fallback: LocationPoint
            if (tag.Location is LocationPoint lp && lp.Point != null) return lp.Point;

            return null;
        }

        private static bool TryMoveByHeadPosition(IndependentTag tag, XYZ delta)
        {
            try
            {
                XYZ cur = tag.TagHeadPosition;
                if (cur == null) return false;
                if (delta.GetLength() <= TOL_FT) return true;
                tag.TagHeadPosition = cur + delta;
                return true;
            }
            catch { return false; }
        }

        private static double Dot(XYZ p, XYZ dir) => p.X * dir.X + p.Y * dir.Y + p.Z * dir.Z;

        private class TagInfo
        {
            public IndependentTag Tag { get; }
            public XYZ Point { get; }
            public double X { get; }
            public TagInfo(IndependentTag tag, XYZ point, double x)
            {
                Tag = tag; Point = point; X = x;
            }
        }

        private class TagSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is IndependentTag;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}
