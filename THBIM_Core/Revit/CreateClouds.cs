using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#nullable disable

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CreateClouds : IExternalCommand
    {
        // ==== Tuning Parameters (feet) ====
        private const double PAD_FT = 0.5;       // Padding around each sleeve (~150mm)
        private const double MERGE_GAP_FT = 1.0; // Threshold to merge clusters (~300mm)

        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            var uidoc = cd.Application.ActiveUIDocument;
            var doc = uidoc.Document;
            var view = uidoc.ActiveView;

            if (!(view is ViewPlan) && !(view is ViewSection) && !(view is View3D))
            {
                TaskDialog.Show("OPENING", "Please run this command in a Plan, Section, or 3D view.");
                return Result.Cancelled;
            }

            // Retrieve sleeves that need review in the current view
            var sleeves = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(SleeveChangeUpdater.IsSleeveName)
                .Where(SleeveTracking.NeedsReview)
                .ToList();

            if (sleeves.Count == 0)
            {
                TaskDialog.Show("OPENING", "No sleeves found requiring revision clouds in the current view.");
                return Result.Succeeded;
            }

            // View coordinate axes
            XYZ U = view.RightDirection.Normalize();
            XYZ V = view.UpDirection.Normalize();
            XYZ W = view.ViewDirection.Normalize();
            XYZ O = TryGetViewOrigin(view) ?? XYZ.Zero;

            // 1) Get AABB in UV coordinates for each sleeve (expanded by PAD_FT)
            var rects = new List<ViewRect>();
            foreach (var s in sleeves)
            {
                var bb = s.get_BoundingBox(view);
                if (bb == null) continue;

                var corners = GetBoxCorners(bb);
                var uvList = new List<UV>();
                foreach (var p in corners)
                {
                    var d = p - O;
                    double u = Dot(d, U);
                    double v = Dot(d, V);
                    uvList.Add(new UV(u, v));
                }

                double minU = uvList.Min(t => t.U) - PAD_FT;
                double maxU = uvList.Max(t => t.U) + PAD_FT;
                double minV = uvList.Min(t => t.V) - PAD_FT;
                double maxV = uvList.Max(t => t.V) + PAD_FT;

                rects.Add(new ViewRect(minU, minV, maxU, maxV));
            }

            if (rects.Count == 0)
            {
                TaskDialog.Show("OPENING", "Could not determine BoundingBox in the current view.");
                return Result.Succeeded;
            }

            // 2) Cluster: Two rects are in the same cluster if they intersect/overlap after expanding by MERGE_GAP_FT
            var clusters = ClusterRects(rects, MERGE_GAP_FT);

            using (var t = new Transaction(doc, "Create Clouds For Changed Sleeves (Merged)"))
            {
                t.Start();

                // 3) Get or Create Revision
                ElementId revId = GetOrCreateRevisionId(doc, "Openings Update");

                int cloudCount = 0;
                foreach (var cluster in clusters)
                {
                    // Collect all UV vertices of the rects in the cluster
                    var pts = new List<UV>();
                    foreach (var r in cluster)
                    {
                        pts.Add(new UV(r.MinU, r.MinV));
                        pts.Add(new UV(r.MaxU, r.MinV));
                        pts.Add(new UV(r.MaxU, r.MaxV));
                        pts.Add(new UV(r.MinU, r.MaxV));
                    }

                    // Convex hull to get the tightest polygon
                    var hull = ConvexHull(pts);
                    if (hull.Count < 3) continue;

                    // Reverse to ensure cloud scallops point OUTWARD.
                    // Monotone chain usually returns CCW; if CCW, reverse to CW.
                    if (SignedArea(hull) > 0) // CCW
                        hull.Reverse();

                    // Construct 3D loop on the view plane (set w = 0 to align with the view plane)
                    IList<Curve> curves = new List<Curve>();
                    for (int i = 0; i < hull.Count; i++)
                    {
                        var a = Unproject(hull[i], 0, U, V, W, O);
                        var b = Unproject(hull[(i + 1) % hull.Count], 0, U, V, W, O);
                        if (!a.IsAlmostEqualTo(b))
                            curves.Add(Line.CreateBound(a, b));
                    }

                    if (curves.Count >= 3)
                    {
                        RevisionCloud.Create(doc, view, revId, curves);
                        cloudCount++;
                    }
                }

                t.Commit();
                TaskDialog.Show("OPENING",
                    $"Created {cloudCount} merged cloud(s) from {sleeves.Count} changed opening(s).");
            }

            return Result.Succeeded;
        }

        // ========= Helpers =========

        private static XYZ TryGetViewOrigin(View v)
        {
            try { return v.Origin; } catch { return null; }
        }

        private static double Dot(XYZ a, XYZ b) => a.X * b.X + a.Y * b.Y + a.Z * b.Z;

        private static XYZ Unproject(UV uv, double w, XYZ U, XYZ V, XYZ W, XYZ O)
            => O + U.Multiply(uv.U) + V.Multiply(uv.V) + W.Multiply(w);

        private static List<XYZ> GetBoxCorners(BoundingBoxXYZ bb)
        {
            var min = bb.Min;
            var max = bb.Max;

            return new List<XYZ>
            {
                new XYZ(min.X, min.Y, min.Z),
                new XYZ(max.X, min.Y, min.Z),
                new XYZ(max.X, max.Y, min.Z),
                new XYZ(min.X, max.Y, min.Z),

                new XYZ(min.X, min.Y, max.Z),
                new XYZ(max.X, min.Y, max.Z),
                new XYZ(max.X, max.Y, max.Z),
                new XYZ(min.X, max.Y, max.Z)
            };
        }

        private class ViewRect
        {
            public double MinU, MinV, MaxU, MaxV;
            public ViewRect(double minU, double minV, double maxU, double maxV)
            {
                MinU = minU; MinV = minV; MaxU = maxU; MaxV = maxV;
            }
        }

        private static List<List<ViewRect>> ClusterRects(List<ViewRect> rects, double mergeGap)
        {
            int n = rects.Count;
            var visited = new bool[n];
            var clusters = new List<List<ViewRect>>();

            for (int i = 0; i < n; i++)
            {
                if (visited[i]) continue;
                var q = new Queue<int>();
                q.Enqueue(i);
                visited[i] = true;

                var cluster = new List<ViewRect> { rects[i] };

                while (q.Count > 0)
                {
                    int cur = q.Dequeue();
                    for (int j = 0; j < n; j++)
                    {
                        if (visited[j]) continue;
                        if (IsConnected(rects[cur], rects[j], mergeGap))
                        {
                            visited[j] = true;
                            q.Enqueue(j);
                            cluster.Add(rects[j]);
                        }
                    }
                }
                clusters.Add(cluster);
            }
            return clusters;
        }

        private static bool IsConnected(ViewRect a, ViewRect b, double gap)
        {
            bool overlapU = (a.MaxU + gap) >= (b.MinU - gap) && (b.MaxU + gap) >= (a.MinU - gap);
            bool overlapV = (a.MaxV + gap) >= (b.MinV - gap) && (b.MaxV + gap) >= (a.MinV - gap);
            return overlapU && overlapV;
        }

        // Convex hull (Monotone chain algorithm) for UV points
        private static List<UV> ConvexHull(List<UV> pts)
        {
            var unique = pts
                .GroupBy(p => (Math.Round(p.U, 6), Math.Round(p.V, 6)))
                .Select(g => g.First())
                .OrderBy(p => p.U)
                .ThenBy(p => p.V)
                .ToList();

            if (unique.Count <= 1) return unique;

            var lower = new List<UV>();
            foreach (var p in unique)
            {
                while (lower.Count >= 2 && Cross(lower[lower.Count - 2], lower[lower.Count - 1], p) <= 0)
                    lower.RemoveAt(lower.Count - 1);
                lower.Add(p);
            }

            var upper = new List<UV>();
            for (int i = unique.Count - 1; i >= 0; i--)
            {
                var p = unique[i];
                while (upper.Count >= 2 && Cross(upper[upper.Count - 2], upper[upper.Count - 1], p) <= 0)
                    upper.RemoveAt(upper.Count - 1);
                upper.Add(p);
            }

            lower.RemoveAt(lower.Count - 1);
            upper.RemoveAt(upper.Count - 1);
            lower.AddRange(upper);
            return lower; // CCW (Counter-Clockwise)
        }

        private static double Cross(UV a, UV b, UV c)
        {
            // (b - a) x (c - a)
            double x1 = b.U - a.U, y1 = b.V - a.V;
            double x2 = c.U - a.U, y2 = c.V - a.V;
            return x1 * y2 - y1 * x2;
        }

        // Create or reuse a Revision based on the description
        private static ElementId GetOrCreateRevisionId(Document doc, string description)
        {
            var rev = new FilteredElementCollector(doc)
                .OfClass(typeof(Revision))
                .Cast<Revision>()
                .FirstOrDefault(r => string.Equals(r.Description, description, StringComparison.OrdinalIgnoreCase));

            if (rev == null)
            {
                rev = Revision.Create(doc);
                rev.Description = description;
                rev.Issued = false;
            }
            return rev.Id;
        }

        private static double SignedArea(IList<UV> poly)
        {
            double a = 0;
            for (int i = 0; i < poly.Count; i++)
            {
                var p = poly[i];
                var q = poly[(i + 1) % poly.Count];
                a += p.U * q.V - q.U * p.V;
            }
            return 0.5 * a; // >0 => CCW (Counter-Clockwise), <0 => CW (Clockwise)
        }
    }
}