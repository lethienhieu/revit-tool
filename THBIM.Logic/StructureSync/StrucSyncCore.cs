using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace THBIM
{
    public enum SyncType
    {
        Floor, Column, PileCap, Pile, DropPanel, Wall
    }

    public class RelationshipItem
    {
        public bool IsChecked { get; set; } = true;
        public string Name { get; set; }
        public int ChildCount { get; set; }
        public string Status { get; set; } = "Ready";
        public string ParentTypeStr { get; set; }
        public string ChildTypeStr { get; set; }
        // Whether the selected Parent element(s) for this set were picked from a Revit link.
        // Child is always local in current workflow.
        public bool ParentIsFromLink { get; set; } = false;
        public List<ElementId> ParentIds { get; set; } = new List<ElementId>();
        public List<ElementId> ChildIds { get; set; } = new List<ElementId>();
        public List<string> Offsets { get; set; } = new List<string>();
        public bool IsEccentric { get; set; } = false;
    }


    public class StrucSyncCore
    {
        private readonly Document _doc;

        public StrucSyncCore(Document doc)
        {
            _doc = doc;
        }

        public class ElementGeometryData
        {
            public XYZ Center { get; set; }
            public double MinZ { get; set; }
            public double MaxZ { get; set; }
            public double InsertionZ { get; set; }
            public Transform LocalTransform { get; set; }
        }

        // --- MATCHING LOGIC ---
        public void MatchRelationships(List<Reference> parentRefs, List<ElementId> childIds,
                                       SyncType pType, SyncType cType,
                                       out List<ElementId> validParents, out List<ElementId> validChildren,
                                       out List<string> validOffsets,
                                       List<RevitLinkInstance> links)
        {
            validParents = new List<ElementId>();
            validChildren = new List<ElementId>();
            validOffsets = new List<string>();

            var parentData = new List<Tuple<ElementId, BoundingBoxXYZ, XYZ, Solid>>();

            // 1. CHUẨN BỊ DỮ LIỆU PARENT
            foreach (var r in parentRefs)
            {
                BoundingBoxXYZ box = null; XYZ center = null; Solid solid = null;

                if (r.LinkedElementId != ElementId.InvalidElementId) // Link
                {
                    RevitLinkInstance linkInst = _doc.GetElement(r.ElementId) as RevitLinkInstance;
                    if (linkInst?.GetLinkDocument() is Document linkDoc)
                    {
                        Element p = linkDoc.GetElement(r.LinkedElementId);
                        if (p != null)
                        {
                            box = p.get_BoundingBox(null);
                            if (box != null)
                            {
                                Transform tf = linkInst.GetTotalTransform();
                                var tfBox = GetTransformedBbox(box, tf);
                                center = GetElementCenter(p, tf);
                                solid = GetSolid(p, tf);
                                parentData.Add(new Tuple<ElementId, BoundingBoxXYZ, XYZ, Solid>(r.LinkedElementId, tfBox, center, solid));
                            }
                        }
                    }
                }
                else // Local
                {
                    Element pLocal = _doc.GetElement(r.ElementId);
                    if (pLocal != null)
                    {
                        box = pLocal.get_BoundingBox(null);
                        if (box != null)
                        {
                            center = GetElementCenter(pLocal, Transform.Identity);
                            solid = GetSolid(pLocal, Transform.Identity);
                            parentData.Add(new Tuple<ElementId, BoundingBoxXYZ, XYZ, Solid>(pLocal.Id, box, center, solid));
                        }
                    }
                }
            }

            // 2. QUÉT CHILD VÀ SO KHỚP
            foreach (var cid in childIds)
            {
                Element child = _doc.GetElement(cid);
                if (child == null) continue;

                BoundingBoxXYZ cBox = child.get_BoundingBox(null);
                if (cBox == null) continue;

                var candidates = parentData.Where(p => IsIntersectXY_WithTolerance(p.Item2, cBox, 0.15)).ToList();

                if (candidates.Count == 0) continue;

                Solid cSolid = GetSolid(child, Transform.Identity);

                foreach (var match in candidates)
                {
                    bool isMatched = true;
                    Solid pSolid = match.Item4;

                    if (pSolid != null && cSolid != null)
                    {
                        if ((pType == SyncType.PileCap && cType == SyncType.Pile) || (pType == SyncType.Pile && cType == SyncType.PileCap))
                        {
                            if (!AreSolidsIntersecting3D(pSolid, cSolid)) isMatched = false;
                        }
                        else
                        {
                            if (!AreSolidsOverlappingXY(pSolid, cSolid)) isMatched = false;
                        }
                    }

                    if (isMatched)
                    {
                        validParents.Add(match.Item1);
                        validChildren.Add(cid);

                        XYZ childCenter = GetElementCenter(child, Transform.Identity);
                        XYZ parentCenter = match.Item3;
                        XYZ offset = childCenter - parentCenter;
                        validOffsets.Add($"{offset.X},{offset.Y},{offset.Z}");
                        break;
                    }
                }
            }
        }

        // --- TARGET CALCULATION ---
        public XYZ CalculateTargetPosition(ElementId parentId, Element child, SyncType pType, SyncType cType, List<RevitLinkInstance> links, string storedOffsetStr, int totalChildInGroup)
        {
            ElementGeometryData pData = GetParentGeometryData(parentId, links);
            if (pData == null) return null;

            XYZ storedOffset = XYZ.Zero;
            if (!string.IsNullOrEmpty(storedOffsetStr))
            {
                var parts = storedOffsetStr.Split(',');
                if (parts.Length == 3) { double.TryParse(parts[0], out double x); double.TryParse(parts[1], out double y); double.TryParse(parts[2], out double z); storedOffset = new XYZ(x, y, z); }
            }

            XYZ currentChildCenter = GetElementCenter(child, Transform.Identity);

            if (pType == SyncType.PileCap && cType == SyncType.Pile)
            {
                if (totalChildInGroup > 1)
                {
                    XYZ targetXY = pData.Center + storedOffset;
                    return new XYZ(targetXY.X, targetXY.Y, pData.MinZ);
                }
                else
                {
                    return new XYZ(pData.Center.X, pData.Center.Y, pData.MinZ);
                }
            }

            if (pType == SyncType.Pile && cType == SyncType.PileCap)
            {
                return new XYZ(pData.Center.X, pData.Center.Y, pData.InsertionZ);
            }

            if ((pType == SyncType.Column && cType == SyncType.PileCap) || (pType == SyncType.PileCap && cType == SyncType.Column) ||
                (pType == SyncType.Column && cType == SyncType.DropPanel) || (pType == SyncType.DropPanel && cType == SyncType.Column) ||
                (pType == SyncType.Column && cType == SyncType.Column) ||
                (pType == SyncType.Wall && cType == SyncType.Wall))
            {
                return new XYZ(pData.Center.X, pData.Center.Y, currentChildCenter.Z);
            }

            return new XYZ(pData.Center.X, pData.Center.Y, pData.MinZ);
        }

        // --- HELPER METHODS ---

        // [MỚI] Lấy Curve của vách cha để đồng bộ chiều dài cho Vách con
        public Curve GetParentCurve(ElementId parentId, List<RevitLinkInstance> links)
        {
            if (links != null)
            {
                foreach (var link in links)
                {
                    Document linkDoc = link.GetLinkDocument();
                    if (linkDoc == null) continue;
                    Element p = linkDoc.GetElement(parentId);
                    if (p != null)
                    {
                        if (p.Location is LocationCurve lc)
                        {
                            return lc.Curve.CreateTransformed(link.GetTotalTransform());
                        }
                        break;
                    }
                }
            }
            Element pLocal = _doc.GetElement(parentId);
            if (pLocal != null && pLocal.Location is LocationCurve pLc)
            {
                return pLc.Curve;
            }
            return null;
        }

        private bool IsIntersectXY_WithTolerance(BoundingBoxXYZ a, BoundingBoxXYZ b, double tolerance)
        {
            return (a.Min.X - tolerance <= b.Max.X + tolerance && a.Max.X + tolerance >= b.Min.X - tolerance) &&
                   (a.Min.Y - tolerance <= b.Max.Y + tolerance && a.Max.Y + tolerance >= b.Min.Y - tolerance);
        }

        public ElementGeometryData GetParentGeometryData(ElementId parentId, List<RevitLinkInstance> links)
        {
            BoundingBoxXYZ box = null; Transform tf = Transform.Identity;
            double insertionZ = 0; bool foundInsertion = false; XYZ center = null;

            if (links != null)
            {
                foreach (var link in links)
                {
                    Document linkDoc = link.GetLinkDocument();
                    if (linkDoc == null) continue;
                    Element p = linkDoc.GetElement(parentId);
                    if (p != null)
                    {
                        box = p.get_BoundingBox(null);
                        tf = link.GetTotalTransform();
                        center = GetElementCenter(p, tf);
                        if (p.Location is LocationPoint lp) { insertionZ = tf.OfPoint(lp.Point).Z; foundInsertion = true; }
                        break;
                    }
                }
            }

            if (box == null)
            {
                Element pLocal = _doc.GetElement(parentId);
                if (pLocal != null)
                {
                    box = pLocal.get_BoundingBox(null);
                    center = GetElementCenter(pLocal, Transform.Identity);
                    if (pLocal is FamilyInstance fi) tf = fi.GetTransform();
                    else tf = Transform.Identity;
                    if (pLocal.Location is LocationPoint lp) { insertionZ = lp.Point.Z; foundInsertion = true; }
                }
            }

            if (box == null) return null;
            BoundingBoxXYZ tfBox = GetTransformedBbox(box, tf);
            if (!foundInsertion) insertionZ = tfBox.Min.Z;

            return new ElementGeometryData
            {
                Center = center ?? (tfBox.Min + tfBox.Max) / 2.0,
                MinZ = tfBox.Min.Z,
                MaxZ = tfBox.Max.Z,
                InsertionZ = insertionZ,
                LocalTransform = tf
            };
        }

        private BoundingBoxXYZ GetTransformedBbox(BoundingBoxXYZ box, Transform tf)
        {
            XYZ pMin = tf.OfPoint(box.Min);
            XYZ pMax = tf.OfPoint(box.Max);
            return new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(pMin.X, pMax.X), Math.Min(pMin.Y, pMax.Y), Math.Min(pMin.Z, pMax.Z)),
                Max = new XYZ(Math.Max(pMin.X, pMax.X), Math.Max(pMin.Y, pMax.Y), Math.Max(pMin.Z, pMax.Z))
            };
        }

        public XYZ GetElementCenter(Element elem, Transform tf)
        {
            XYZ localCenter = XYZ.Zero;
            if (elem.Location is LocationPoint lp) localCenter = lp.Point;
            else if (elem.Location is LocationCurve lc)
            {
                XYZ p1 = lc.Curve.GetEndPoint(0); XYZ p2 = lc.Curve.GetEndPoint(1);
                localCenter = (p1 + p2) / 2.0;
            }
            else
            {
                BoundingBoxXYZ box = elem.get_BoundingBox(null);
                if (box != null) localCenter = (box.Min + box.Max) / 2.0;
            }
            return tf.OfPoint(localCenter);
        }

        private bool IsIntersectXY(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            return (a.Min.X <= b.Max.X && a.Max.X >= b.Min.X) && (a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y);
        }

        private Solid GetSolid(Element elem, Transform tf)
        {
            Options opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };
            GeometryElement geoElem = elem.get_Geometry(opt);
            return GetSolidFromGeo(geoElem, tf);
        }

        private Solid GetSolidFromGeo(GeometryElement geoElem, Transform tf)
        {
            if (geoElem == null) return null;
            foreach (GeometryObject obj in geoElem)
            {
                if (obj is Solid s && s.Volume > 0) return SolidUtils.CreateTransformed(s, tf);
                else if (obj is GeometryInstance inst)
                {
                    Transform combinedTf = tf.Multiply(inst.Transform);
                    return GetSolidFromGeo(inst.SymbolGeometry, combinedTf);
                }
            }
            return null;
        }

        private bool AreSolidsOverlappingXY(Solid s1, Solid s2)
        {
            try
            {
                XYZ c1 = s1.ComputeCentroid(); XYZ c2 = s2.ComputeCentroid();
                Transform zMove = Transform.CreateTranslation(new XYZ(0, 0, c1.Z - c2.Z));
                Solid s2Moved = SolidUtils.CreateTransformed(s2, zMove);
                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(s1, s2Moved, BooleanOperationsType.Intersect);
                return intersection != null && intersection.Volume > 0.0001;
            }
            catch { return true; }
        }

        private bool AreSolidsIntersecting3D(Solid s1, Solid s2)
        {
            try
            {
                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(s1, s2, BooleanOperationsType.Intersect);
                return intersection != null && intersection.Volume > 0.0001;
            }
            catch { return false; }
        }

        private void GetSolidZBounds(Solid solid, out double minZ, out double maxZ)
        {
            minZ = double.MaxValue; maxZ = double.MinValue;
            if (solid == null || solid.Edges.Size == 0) { minZ = 0; maxZ = 0; return; }
            foreach (Edge edge in solid.Edges)
            {
                IList<XYZ> points = edge.Tessellate();
                foreach (XYZ pt in points) { if (pt.Z < minZ) minZ = pt.Z; if (pt.Z > maxZ) maxZ = pt.Z; }
            }
        }
    }
}
