using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using THBIM.Tools;
#nullable disable

namespace ColumnAutoDim.Revit
{
    public static class AutoDimLogic
    {
        private const double TOLERANCE = 1.0e-9;
        private const double ZERO_DIM_TOLERANCE = 0.0033;
        // Khoảng cách grid-to-center nhỏ hơn giá trị này → coi grid đi qua tâm cột
        private const double GRID_AT_CENTER_TOLERANCE = 0.01; // ~3mm

        public static void Run(Document doc, View view, List<Element> columns, List<Grid> allGrids, DimSettings settings)
        {
            using (Transaction t = new Transaction(doc, "THBIM - Auto Dim Columns"))
            {
                t.Start();
                foreach (var el in columns)
                {
                    if (el is FamilyInstance fi) ProcessSingleColumn(doc, view, fi, allGrids, settings);
                }
                t.Commit();
            }
        }

        private static void ProcessSingleColumn(Document doc, View view, FamilyInstance fi, List<Grid> allGrids, DimSettings settings)
        {
            try
            {
                Transform tr = fi.GetTransform();
                XYZ vec1 = tr.BasisX; XYZ vec2 = tr.BasisY;
                GetVisualAxes(vec1, vec2, out XYZ visualRight, out XYZ visualUp);
                XYZ offsetDir_ForVertDim = !settings.IsPlaceLeft ? visualRight : -visualRight;
                XYZ offsetDir_ForHorizDim = settings.IsPlaceTop ? visualUp : -visualUp;

                ProcessDirection(doc, view, fi, offsetDir_ForHorizDim, allGrids, settings);
                ProcessDirection(doc, view, fi, offsetDir_ForVertDim, allGrids, settings);

                if (settings.IsTagEnabled && settings.SelectedTagType != null)
                    CreateTag(doc, view, fi, visualUp, visualRight, settings.SelectedTagType, settings.TagPosition);
            }
            catch { }
        }

        private static void ProcessDirection(Document doc, View view, FamilyInstance fi, XYZ offsetDir, List<Grid> allGrids, DimSettings settings)
        {
            if (!settings.IsDimColumnEnabled) return;
            XYZ safeMeasureDir = new XYZ(-offsetDir.Y, offsetDir.X, 0).Normalize();

            // ========================================================
            // FIX BUG 1: Ưu tiên stable references từ GetReferences()
            // trước, chỉ fallback sang geometry faces nếu không có
            // ========================================================
            Reference ref1 = null; Reference ref2 = null;
            bool isCenterDim = false;

            // Thử lấy stable references trước (luôn hoạt động nếu family có reference planes)
            if (GetStableExtentReferences(fi, safeMeasureDir, out Reference rStableMin, out Reference rStableMax))
            {
                ref1 = rStableMin; ref2 = rStableMax; isCenterDim = false;
            }
            // Fallback: thử geometry faces
            else if (GetExtremeFaces(fi, offsetDir, out Reference rMin, out Reference rMax))
            {
                ref1 = rMin; ref2 = rMax; isCenterDim = false;
            }
            // Fallback cuối: center reference
            else
            {
                Reference rCenter = GetCenterReference(fi, safeMeasureDir);
                if (rCenter != null) { ref1 = rCenter; isCenterDim = true; }
                else return;
            }

            Grid closestGrid = null;
            double gridPos = 0;
            if (settings.IsDimGridEnabled)
                closestGrid = FindClosestParallelGrid(allGrids, safeMeasureDir, fi.GetTransform().Origin, out gridPos);

            double geomMaxProj = GetMaxProjection(fi, offsetDir);
            double centerProj = fi.GetTransform().Origin.DotProduct(offsetDir);
            double centerPosOnMeasure = fi.GetTransform().Origin.DotProduct(safeMeasureDir);
            double baseOffsetLen = (geomMaxProj - centerProj) + (settings.OffsetChainFromColumn / 304.8);
            double extraOffsetStep = settings.OffsetGridStep / 304.8;

            ReferenceArray refArray1 = new ReferenceArray();
            ReferenceArray refArray2 = new ReferenceArray();
            bool isCutting = false;
            double dim1_ExtraPush = 0;

            if (closestGrid == null)
            {
                if (isCenterDim) return;
                refArray1.Append(ref1); refArray1.Append(ref2);
            }
            else
            {
                Reference rGrid = new Reference(closestGrid);

                // ========================================================
                // FIX BUG 2: Check grid có đi qua tâm cột không
                // Nếu có → chỉ tạo 1 dim (grid + 2 faces), KHÔNG tạo 2 lớp
                // ========================================================
                double distGridToCenter = Math.Abs(gridPos - centerPosOnMeasure);
                bool gridAtCenter = distGridToCenter < GRID_AT_CENTER_TOLERANCE;

                if (isCenterDim)
                {
                    if (distGridToCenter > ZERO_DIM_TOLERANCE)
                    {
                        refArray1.Append(rGrid); refArray1.Append(ref1);
                    }
                }
                else if (gridAtCenter)
                {
                    // Grid đi qua tâm cột → chỉ tạo 1 dim chain: grid + 2 faces
                    // Không cần dim riêng cho cột vì grid dim đã bao gồm
                    refArray1.Append(rGrid); refArray1.Append(ref1); refArray1.Append(ref2);
                }
                else
                {
                    // Grid KHÔNG trùng tâm → tạo 2 lớp dim khác nhau
                    refArray1.Append(ref1); refArray1.Append(ref2);
                    isCutting = true;
                    refArray2.Append(rGrid); refArray2.Append(ref1); refArray2.Append(ref2);
                    dim1_ExtraPush = extraOffsetStep;
                }
            }

            CreateDimAtLocation(doc, view, fi.GetTransform().Origin, safeMeasureDir, offsetDir, baseOffsetLen + dim1_ExtraPush, refArray1, settings.SelectedDimType);
            if (isCutting && refArray2.Size > 1 && !isCenterDim)
                CreateDimAtLocation(doc, view, fi.GetTransform().Origin, safeMeasureDir, offsetDir, baseOffsetLen, refArray2, settings.SelectedDimType);
        }

        private static void CreateDimAtLocation(Document doc, View view, XYZ center, XYZ measureDir, XYZ offsetDir, double offsetLen, ReferenceArray refs, DimensionType dimType)
        {
            if (refs.Size < 2) return;
            XYZ dimOrigin = center + offsetDir * offsetLen;
            double viewZ;
            if (view is ViewPlan vp && vp.GenLevel != null) viewZ = vp.GenLevel.Elevation;
            else viewZ = view.Origin.Z;
            XYZ pointOnView = new XYZ(dimOrigin.X, dimOrigin.Y, viewZ);
            XYZ p1 = pointOnView - measureDir * 5.0;
            XYZ p2 = pointOnView + measureDir * 5.0;
            try
            {
                Line dimLine = Line.CreateBound(p1, p2);
                Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
                if (dim != null && dimType != null) dim.DimensionType = dimType;
            }
            catch { }
        }

        // ========================================================
        // FIX BUG 1: Stable references — ưu tiên fi.GetReferences()
        // Đây là cách đáng tin cậy nhất để lấy reference cho dim
        // ========================================================
        private static bool GetStableExtentReferences(FamilyInstance fi, XYZ measureDir, out Reference rMin, out Reference rMax)
        {
            rMin = rMax = null;
            try
            {
                Transform tr = fi.GetTransform();
                double dotX = Math.Abs(measureDir.DotProduct(tr.BasisX.Normalize()));
                double dotY = Math.Abs(measureDir.DotProduct(tr.BasisY.Normalize()));

                IList<Reference> refsA, refsB;
                if (dotX > dotY)
                {
                    refsA = fi.GetReferences(FamilyInstanceReferenceType.Left);
                    refsB = fi.GetReferences(FamilyInstanceReferenceType.Right);
                }
                else
                {
                    refsA = fi.GetReferences(FamilyInstanceReferenceType.Front);
                    refsB = fi.GetReferences(FamilyInstanceReferenceType.Back);
                }

                if (refsA != null && refsA.Count > 0 && refsB != null && refsB.Count > 0)
                {
                    rMin = refsA.First();
                    rMax = refsB.First();
                    return true;
                }

                // Thử thêm Strong/Weak references nếu Left/Right/Front/Back không có
                var strongRefs = fi.GetReferences(FamilyInstanceReferenceType.StrongReference);
                if (strongRefs != null && strongRefs.Count >= 2)
                {
                    // Lọc strong refs song song với measureDir
                    var filtered = new List<Reference>();
                    foreach (var r in strongRefs)
                    {
                        filtered.Add(r);
                    }
                    if (filtered.Count >= 2)
                    {
                        rMin = filtered[0];
                        rMax = filtered[1];
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private static Reference GetCenterReference(FamilyInstance fi, XYZ measureDir)
        {
            try
            {
                Transform tr = fi.GetTransform();
                double dotX = Math.Abs(measureDir.DotProduct(tr.BasisX.Normalize()));
                double dotY = Math.Abs(measureDir.DotProduct(tr.BasisY.Normalize()));
                FamilyInstanceReferenceType refType = (dotX > dotY)
                    ? FamilyInstanceReferenceType.CenterLeftRight
                    : FamilyInstanceReferenceType.CenterFrontBack;
                return fi.GetReferences(refType).FirstOrDefault();
            }
            catch { return null; }
        }

        private static void CreateTag(Document doc, View view, FamilyInstance fi, XYZ upDir, XYZ rightDir, FamilySymbol tagType, string positionMode)
        {
            double offsetMargin = 250.0 / 304.8;
            XYZ locationPt = fi.GetTransform().Origin;
            if (positionMode != "C")
            {
                double maxUp = GetMaxProjection(fi, upDir); double maxDown = GetMaxProjection(fi, -upDir);
                double maxRight = GetMaxProjection(fi, rightDir); double maxLeft = GetMaxProjection(fi, -rightDir);
                XYZ center = fi.GetTransform().Origin;
                double cUp = center.DotProduct(upDir); double cRight = center.DotProduct(rightDir);
                double distUp = maxUp - cUp; double distDown = maxDown - center.DotProduct(-upDir);
                double distRight = maxRight - cRight; double distLeft = maxLeft - center.DotProduct(-rightDir);

                switch (positionMode)
                {
                    case "TR": locationPt = center + (upDir * (distUp + offsetMargin)) + (rightDir * (distRight + offsetMargin - (350.0 / 304.8))); break;
                    case "TL": locationPt = center + (upDir * (distUp + offsetMargin)) - (rightDir * (distLeft + offsetMargin + (550.0 / 304.8))); break;
                    case "BR": locationPt = center - (upDir * (distDown + offsetMargin)) + (rightDir * (distRight + offsetMargin)); break;
                    case "BL": locationPt = center - (upDir * (distDown + offsetMargin)) - (rightDir * (distLeft + offsetMargin + (550.0 / 304.8))); break;
                }
            }
            double z = (view is ViewPlan vp && vp.GenLevel != null) ? vp.GenLevel.Elevation : view.Origin.Z;
            locationPt = new XYZ(locationPt.X, locationPt.Y, z);
            try
            {
                Reference refCol = new Reference(fi);
                IndependentTag.Create(doc, tagType.Id, view.Id, refCol, false, TagOrientation.Horizontal, locationPt);
            }
            catch { }
        }

        private static Grid FindClosestParallelGrid(List<Grid> grids, XYZ measureDir, XYZ center, out double gridPosProj)
        {
            Grid bestGrid = null; double minDist = double.MaxValue; gridPosProj = 0;
            XYZ normDir = measureDir.Normalize();
            foreach (var g in grids)
            {
                if (g.Curve is Line gridLine)
                {
                    if (!IsPerpendicular(gridLine.Direction, measureDir)) continue;
                    double pos = gridLine.Origin.DotProduct(normDir);
                    double dist = Math.Abs(pos - center.DotProduct(normDir));
                    if (dist < minDist) { minDist = dist; bestGrid = g; gridPosProj = pos; }
                }
            }
            return bestGrid;
        }

        // Geometry face extraction — fallback khi GetReferences không có
        private static bool GetExtremeFaces(FamilyInstance fi, XYZ offsetDir, out Reference rMin, out Reference rMax)
        {
            rMin = rMax = null;
            XYZ measureDir = new XYZ(-offsetDir.Y, offsetDir.X, 0).Normalize();
            var faces = CollectPlanarFaces(fi);
            // Chỉ lấy faces có reference hợp lệ
            var candidates = faces.Where(f => f.Reference != null && IsParallel(f.FaceNormal, measureDir)).ToList();
            if (candidates.Count >= 2)
            {
                double minV = double.MaxValue, maxV = double.MinValue;
                PlanarFace fMin = null, fMax = null;
                foreach (var f in candidates)
                {
                    double p = f.Origin.DotProduct(measureDir);
                    if (p < minV) { minV = p; fMin = f; }
                    if (p > maxV) { maxV = p; fMax = f; }
                }
                if (fMin != null && fMax != null && fMin.Reference != null && fMax.Reference != null)
                {
                    rMin = fMin.Reference; rMax = fMax.Reference;
                    return true;
                }
            }
            return false;
        }

        private static double GetMaxProjection(FamilyInstance fi, XYZ dir)
        {
            double maxV = double.MinValue; dir = dir.Normalize();
            foreach (var f in CollectPlanarFaces(fi))
            {
                if (IsParallel(f.FaceNormal, dir))
                {
                    double p = f.Origin.DotProduct(dir);
                    if (p > maxV) maxV = p;
                }
            }
            if (maxV != double.MinValue) return maxV;

            BoundingBoxXYZ bbox = fi.get_BoundingBox(null);
            if (bbox != null)
            {
                XYZ min = bbox.Min; XYZ max = bbox.Max;
                XYZ[] corners = new XYZ[] {
                    new XYZ(min.X, min.Y, min.Z), new XYZ(min.X, min.Y, max.Z),
                    new XYZ(min.X, max.Y, min.Z), new XYZ(min.X, max.Y, max.Z),
                    new XYZ(max.X, min.Y, min.Z), new XYZ(max.X, min.Y, max.Z),
                    new XYZ(max.X, max.Y, min.Z), new XYZ(max.X, max.Y, max.Z)
                };
                foreach (XYZ c in corners)
                {
                    double p = c.DotProduct(dir);
                    if (p > maxV) maxV = p;
                }
                return maxV;
            }
            return fi.GetTransform().Origin.DotProduct(dir) + 1.0;
        }

        private static List<PlanarFace> CollectPlanarFaces(FamilyInstance fi)
        {
            List<PlanarFace> list = new List<PlanarFace>();
            Options opt = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = true };
            GeometryElement ge = fi.get_Geometry(opt);
            if (ge != null)
            {
                foreach (GeometryObject o in ge)
                {
                    if (o is Solid s) AddFaces(s, list);
                    else if (o is GeometryInstance gi)
                        foreach (GeometryObject o2 in gi.GetInstanceGeometry())
                            if (o2 is Solid s2) AddFaces(s2, list);
                }
            }
            return list;
        }

        private static void AddFaces(Solid s, List<PlanarFace> list)
        {
            if (s.Volume < 1e-9) return; // skip empty solids
            foreach (Face f in s.Faces)
                if (f is PlanarFace pf) list.Add(pf);
        }

        private static void GetVisualAxes(XYZ v1, XYZ v2, out XYZ visualRight, out XYZ visualUp)
        {
            if (Math.Abs(v1.X) > Math.Abs(v2.X))
            {
                visualRight = v1.X > 0 ? v1 : -v1;
                visualUp = v2.Y > 0 ? v2 : -v2;
            }
            else
            {
                visualRight = v2.X > 0 ? v2 : -v2;
                visualUp = v1.Y > 0 ? v1 : -v1;
            }
        }

        private static bool IsParallel(XYZ v1, XYZ v2) => Math.Abs(Math.Abs(v1.Normalize().DotProduct(v2.Normalize())) - 1.0) < TOLERANCE;
        private static bool IsPerpendicular(XYZ v1, XYZ v2) => Math.Abs(v1.Normalize().DotProduct(v2.Normalize())) < TOLERANCE;
    }
}
