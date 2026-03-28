using System;
using Autodesk.Revit.DB;

namespace THBIM
{
    internal static class LeaderAngleHelper
    {
        /// <summary>
        /// Bẻ leader elbow theo góc đã setting.
        /// Gọi sau khi align tag xong, trong transaction đang mở.
        /// </summary>
        public static void ApplyIfNeeded(IndependentTag tag)
        {
            if (!tag.HasLeader) return;

            LeaderAngleSettings.Load();
            double angleDeg = LeaderAngleSettings.AngleDegrees;
            if (angleDeg <= 0) return;

            double angleRad = angleDeg * Math.PI / 180.0;
            XYZ head = tag.TagHeadPosition;
            if (head == null) return;

            // Chuyển sang Free End nếu đang Attached (cần thiết để set elbow)
            try
            {
                if (tag.LeaderEndCondition != LeaderEndCondition.Free)
                    tag.LeaderEndCondition = LeaderEndCondition.Free;
            }
            catch { }

            var doc = tag.Document;
            var view = doc.GetElement(tag.OwnerViewId) as View;

            try
            {
                foreach (var refObj in tag.GetTaggedReferences())
                {
                    // Luôn set leader end về tâm BB host element
                    XYZ end = null;
                    try
                    {
                        var host = doc.GetElement(refObj.ElementId);
                        if (host != null && view != null)
                        {
                            var bb = host.get_BoundingBox(view);
                            if (bb != null)
                            {
                                end = (bb.Min + bb.Max) * 0.5;
                                tag.SetLeaderEnd(refObj, end);
                            }
                        }
                    }
                    catch { }

                    if (end == null) continue;

                    // Tính elbow theo view direction
                    XYZ up = view != null ? view.UpDirection : XYZ.BasisZ;
                    XYZ right = view != null ? view.RightDirection : XYZ.BasisX;

                    // Bù offset giữa TagHeadPosition và leader attachment point
                    // (Revit nối leader vào giữa text, không phải đỉnh head)
                    double attachOffsetFt = 0.023539;

                    // Project head và end lên view axes
                    double headUp = head.X * up.X + head.Y * up.Y + head.Z * up.Z;
                    double endUp = end.X * up.X + end.Y * up.Y + end.Z * up.Z;
                    double headRight = head.X * right.X + head.Y * right.Y + head.Z * right.Z;
                    double endRight = end.X * right.X + end.Y * right.Y + end.Z * right.Z;

                    // Elbow nằm ở mức head - offset (khớp attachment point)
                    double elbowUp = headUp - attachOffsetFt;
                    double dUp = elbowUp - endUp;
                    double dRight = headRight - endRight;

                    XYZ elbowOffset;
                    if (Math.Abs(angleDeg - 90) < 0.01)
                    {
                        // 90°: element → thẳng lên theo up → ngang vào tag
                        elbowOffset = up * dUp;
                    }
                    else
                    {
                        double offsetRight = Math.Abs(dUp) / Math.Tan(angleRad);
                        double clampedRight = Math.Min(offsetRight, Math.Abs(dRight));
                        elbowOffset = up * dUp + right * (Math.Sign(dRight) * clampedRight);
                    }

                    XYZ elbow = end + elbowOffset;

                    try { tag.SetLeaderElbow(refObj, elbow); }
                    catch { }
                }
            }
            catch { }
        }
    }
}
