using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;    // CableTray
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
#nullable disable

namespace OPENING.MODEL
{
    /// <summary>
    /// TH_RECTANGULAR_SLEEVE (host lên Duct/CableTray – chữ nhật)
    /// Nâng cấp: Cho phép pick Wall/Floor và MEP chữ nhật từ các Revit Link đã tick trong UI.
    /// - Wall/Floor chỉ dùng để tìm vị trí (không host lên đó).
    /// - Duct/CableTray:
    ///     * NGANG → host mặt TRÊN.
    ///     * ĐỨNG → host mặt BÊN theo WIDTH.
    ///     * CableTray + Floor → host mặt BÊN theo HEIGHT (override).
    /// - Sau khi host: gán Width/Height/Insulation/Host Thickness + Clearance/Extrusion/Fill (feet).
    /// </summary>
    public static class TH_RECTANGULAR_SLEEVE
    {
        private const double VERTICAL_THR = 0.75; // |u.Z| > 0.75 xem như đứng
        private const double SIDE_DOT_MIN = 0.707; // ~ cos(45°)

        // Giữ overload cũ để tương thích
        public static void Run(ExternalCommandData commandData, Document doc,
                               double clearance, double extrusion, double fill)
            => Run(commandData, doc, clearance, extrusion, fill, new List<Document>());

        // Overload mới: truyền danh sách Link Docs được tick trong UI
        public static void Run(ExternalCommandData commandData, Document doc,
                               double clearance, double extrusion, double fill,
                               IList<Document> allowedLinkDocs)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            // 0) FamilySymbol "TH_RECTANGULAR_SLEEVE"
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family != null && s.Family.Name == "TH_RECTANGULAR_SLEEVE");

            if (symbol == null)
            {
                TaskDialog.Show("OPENING", "Family 'TH_RECTANGULAR_SLEEVE' not found'.");
                return;
            }

            // Bật/tắt pick link theo link đã tick
            var allowedLinkInstIds = BuildAllowedLinkInstanceIdSet(doc, allowedLinkDocs);
            bool enableLinkPick = allowedLinkInstIds != null && allowedLinkInstIds.Count > 0;

            // ===== 1) Chọn Wall/Floor =====
            var hostPicks = new List<PickedElem>();
            try
            {
                var msgHost = enableLinkPick
                    ? "Select Wall/Floor in the current file (Press Esc to finish host selection)"
                    : "Select Wall/Floor in the current file";
                var hostRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new HostDocFilter(IsWallOrFloor),
                    msgHost);

                foreach (var r in hostRefs)
                {
                    var e = doc.GetElement(r);
                    if (e != null) hostPicks.Add(PickedElem.Host(e));
                }
            }
            catch { /* user Esc */ }

            if (enableLinkPick)
            {
                try
                {
                    var linkRefs = uidoc.Selection.PickObjects(
                        ObjectType.LinkedElement,
                        new LinkedFilter(doc, allowedLinkInstIds, IsRectHost_WallOrFloor),
                        "Select Wall/Floor from the checked LINK files (Press Esc to finish selection)");

                    foreach (var r in linkRefs)
                    {
                        var li = doc.GetElement(r.ElementId) as RevitLinkInstance;
                        var ldoc = li?.GetLinkDocument(); if (ldoc == null) continue;
                        var e = ldoc.GetElement(r.LinkedElementId); if (e == null) continue;
                        hostPicks.Add(PickedElem.Linked(e, li));
                    }
                }
                catch { /* user Esc */ }
            }

            if (hostPicks.Count == 0) return;

            // ===== 2) Chọn MEP chữ nhật (Duct/CableTray) =====
            var mepPicks = new List<PickedElem>();
            try
            {
                var msgHost = enableLinkPick
                    ? "Select Rectangular Duct / Cable Tray in the current file (Press Esc to finish host selection)" 
                    : "Select Rectangular Duct / Cable Tray in the current file";
                var mepRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new HostDocFilter(IsRectMEP),
                    msgHost);

                foreach (var r in mepRefs)
                {
                    var e = doc.GetElement(r);
                    if (e != null) mepPicks.Add(PickedElem.Host(e));
                }
            }
            catch { /* user Esc */ }

            if (enableLinkPick)
            {
                try
                {
                    var mepLinkRefs = uidoc.Selection.PickObjects(
                        ObjectType.LinkedElement,
                        new LinkedFilter(doc, allowedLinkInstIds, IsRectMEP),
                        "Select Rectangular Duct / Cable Tray from the checked LINK files (Press Esc to finish selection)");

                    foreach (var r in mepLinkRefs)
                    {
                        var li = doc.GetElement(r.ElementId) as RevitLinkInstance;
                        var ldoc = li?.GetLinkDocument(); if (ldoc == null) continue;
                        var e = ldoc.GetElement(r.LinkedElementId); if (e == null) continue;
                        mepPicks.Add(PickedElem.Linked(e, li));
                    }
                }
                catch { /* user Esc */ }
            }

            if (mepPicks.Count == 0) return;

            // Lưu view hiện tại để trả lại sau
            ElementId prevViewId = uidoc.ActiveView?.Id ?? ElementId.InvalidElementId;

            try
            {
                foreach (var hostPick in hostPicks)
                {
                    Solid hostSolidL = GetMainSolid(hostPick.Elem);
                    if (hostSolidL == null) continue;

                    // Solid host -> HOST coords
                    Solid hostSolidH = hostPick.ToHost(hostSolidL);

                    double hostThicknessFeet = GetHostThicknessFeet(hostPick.Elem);
                    bool hostIsFloor = hostPick.Elem is Floor;

                    foreach (var mepPick in mepPicks)
                    {
                        if (!(mepPick.Elem?.Location is LocationCurve lc) || lc.Curve == null) continue;

                        // Trục MEP -> HOST coords
                        Curve axisH = mepPick.ToHost(lc.Curve);
                        XYZ uH = GetCurveDirection(axisH);
                        if (IsZeroLength(uH)) continue;

                        // Giao trục với host
                        var sci = hostSolidH.IntersectWithCurve(axisH, new SolidCurveIntersectionOptions());
                        if (sci == null || sci.SegmentCount == 0) continue;

                        // refPoint (HOST coords): mid slab hoặc tâm tường
                        XYZ refPointH;
                        if (hostIsFloor)
                        {
                            refPointH = MidPointOfIntersection(sci);
                        }
                        else
                        {
                            Curve segH = sci.GetCurveSegment(0);
                            XYZ centerOnPlaneH = segH.GetEndPoint(0);
                            if (!TryFindFaceAtPoint(hostSolidH, centerOnPlaneH, out Face hostFaceH, out UV hostUV)) continue;
                            XYZ nHostH = hostFaceH.ComputeNormal(hostUV).Normalize();
                            refPointH = centerOnPlaneH - nHostH * (hostThicknessFeet / 2.0);
                        }

                        // Kích thước RECT (feet)
                        if (!TryGetRectSizesFeet(mepPick.Elem, out double wBare, out double hBare)) continue;

                        // Insulation
                        double insulT = GetRectInsulationThicknessFeet(mepPick.Elem);

                        // Axes RECT (trong MEP doc) -> HOST coords
                        if (!TryGetRectCrossAxes(mepPick.Elem, out XYZ widthAxisL, out XYZ heightAxisL)) continue;
                        XYZ widthAxisH = mepPick.ToHostVec(widthAxisL).Normalize();
                        XYZ heightAxisH = mepPick.ToHostVec(heightAxisL).Normalize();

                        bool isVertical = Math.Abs(uH.Z) > VERTICAL_THR;
                        bool isCableTray = mepPick.Elem is CableTray;

                        // ---- chọn face để host ----
                        PlanarFace hostFace;         // face trong MEP doc (L nếu linked; H nếu host)
                        XYZ faceNormalH;             // normal của face trong HOST coords
                        Reference faceRef;           // reference hợp lệ (host hoặc link)
                        XYZ placePointH;             // điểm đặt (HOST coords)
                        XYZ xDirH;                   // hướng X (HOST coords)

                        if (isCableTray && hostIsFloor)
                        {
                            // CableTray + Floor: mặt BÊN theo HEIGHT
                            if (!TryGetRectSideFaceByAxis(mepPick.Elem, heightAxisL, SIDE_DOT_MIN, out PlanarFace sideFace)) continue;
                            hostFace = sideFace;
                            faceNormalH = mepPick.ToHostVec(hostFace.FaceNormal).Normalize();

                            if (mepPick.IsLinked)
                            {
                                XYZ refPointL = mepPick.ToLinkPoint(refPointH);
                                var proj = hostFace.Project(refPointL);
                                XYZ pL = (proj != null) ? proj.XYZPoint : refPointL;
                                placePointH = mepPick.ToHostPoint(pL);
                            }
                            else
                            {
                                var proj = hostFace.Project(refPointH);
                                placePointH = (proj != null) ? proj.XYZPoint : refPointH;
                            }

                            xDirH = ProjectOnPlane(uH, faceNormalH) ?? ProjectOnPlane(widthAxisH, faceNormalH) ?? AnyPerpTo(faceNormalH);
                            xDirH = xDirH.Normalize();

                            faceRef = mepPick.IsLinked
                                ? hostFace.Reference.CreateLinkReference(mepPick.LinkInst)
                                : hostFace.Reference;
                        }
                        else if (!isVertical)
                        {
                            // NGANG: mặt TRÊN
                            if (!TryGetTopPlanarFace(mepPick.Elem, out PlanarFace topFace)) continue;
                            hostFace = topFace;
                            faceNormalH = mepPick.ToHostVec(hostFace.FaceNormal).Normalize();

                            if (mepPick.IsLinked)
                            {
                                XYZ refPointL = mepPick.ToLinkPoint(refPointH);
                                var proj = hostFace.Project(refPointL);
                                XYZ pL = (proj != null) ? proj.XYZPoint : refPointL;
                                placePointH = mepPick.ToHostPoint(pL);
                            }
                            else
                            {
                                var proj = hostFace.Project(refPointH);
                                placePointH = (proj != null) ? proj.XYZPoint : refPointH;
                            }

                            xDirH = ProjectOnPlane(uH, faceNormalH) ?? AnyPerpTo(faceNormalH);
                            xDirH = xDirH.Normalize();

                            faceRef = mepPick.IsLinked
                                ? hostFace.Reference.CreateLinkReference(mepPick.LinkInst)
                                : hostFace.Reference;
                        }
                        else
                        {
                            // ĐỨNG: mặt BÊN theo WIDTH
                            if (!TryGetRectSideFaceByAxis(mepPick.Elem, widthAxisL, SIDE_DOT_MIN, out PlanarFace sideFace)) continue;
                            hostFace = sideFace;
                            faceNormalH = mepPick.ToHostVec(hostFace.FaceNormal).Normalize();

                            if (mepPick.IsLinked)
                            {
                                XYZ refPointL = mepPick.ToLinkPoint(refPointH);
                                var proj = hostFace.Project(refPointL);
                                XYZ pL = (proj != null) ? proj.XYZPoint : refPointL;
                                placePointH = mepPick.ToHostPoint(pL);
                            }
                            else
                            {
                                var proj = hostFace.Project(refPointH);
                                placePointH = (proj != null) ? proj.XYZPoint : refPointH;
                            }

                            xDirH = ProjectOnPlane(uH, faceNormalH) ?? ProjectOnPlane(heightAxisH, faceNormalH) ?? AnyPerpTo(faceNormalH);
                            xDirH = xDirH.Normalize();

                            faceRef = mepPick.IsLinked
                                ? hostFace.Reference.CreateLinkReference(mepPick.LinkInst)
                                : hostFace.Reference;
                        }

                        // Level & view (trong host doc)
                        ElementId refLvlId = GetMEPReferenceLevelId(doc, mepPick.Elem, placePointH);
                        if (mepPick.IsLinked) refLvlId = GetNearestLevelId(doc, placePointH.Z);
                        ViewPlan plan = FindPlanViewForLevel(doc, refLvlId);
                        if (plan != null) uidoc.ActiveView = plan;

                        using (Transaction t = new Transaction(doc, "Place TH_RECTANGULAR_SLEEVE (link-aware)"))
                        {
                            t.Start();
                            if (!symbol.IsActive) symbol.Activate();

                            FamilyInstance inst = doc.Create.NewFamilyInstance(faceRef, placePointH, xDirH, symbol);

                            // Bảo đảm host đúng MEP
                            bool okHost = true;
                            if (!mepPick.IsLinked)
                                okHost = ((inst.Host as Element)?.Id == mepPick.Elem.Id);
                            else
                                okHost = (inst.Host is RevitLinkInstance rli && rli.Id == mepPick.LinkInst.Id);

                            if (!okHost)
                            {
                                doc.Delete(inst.Id);
                                t.Commit();
                                continue;
                            }

                            doc.Regenerate();

                            // (tuỳ chọn) ghi loại MEP để scoreboard đếm nhanh
                            string mepKind = (mepPick.Elem is Duct) ? "Duct" :
                                             (mepPick.Elem is CableTray) ? "CableTray" :
                                             (mepPick.Elem?.Category?.Name ?? "");
                            SetTextParam(inst, "TH_MEP_KIND", mepKind);

                            // Set tham số
                            SetParam(inst, new[] { "Duct Width", "Width" }, wBare);
                            SetParam(inst, new[] { "Duct Height", "Height" }, hBare);
                            SetParam(inst, new[] { "Insulation" }, insulT > 0 ? insulT : 0.0);
                            SetParam(inst, new[] { "Host Thickness" }, hostThicknessFeet);

                            SetParam(inst, new[] { "Clearance" }, clearance);
                            SetParam(inst, new[] { "Extrusion" }, extrusion);
                            SetParam(inst, new[] { "Fill" }, fill);

                            TrySetScheduleLevel(doc, inst, refLvlId);
                            t.Commit();
                        }

                        // Trả view cũ
                        if (prevViewId != ElementId.InvalidElementId)
                        {
                            View prev = doc.GetElement(prevViewId) as View;
                            if (prev != null) uidoc.ActiveView = prev;
                        }
                    }
                }
            }
            finally
            {
                if (prevViewId != ElementId.InvalidElementId)
                {
                    View prev = doc.GetElement(prevViewId) as View;
                    if (prev != null) uidoc.ActiveView = prev;
                }
            }
        }

        // ===================== Link-aware helper types =====================
        private class PickedElem
        {
            public Element Elem;                // Element trong doc của nó (host hoặc link)
            public RevitLinkInstance LinkInst;  // null nếu host doc
            public bool IsLinked => LinkInst != null;
            public Transform ToHostT;           // Transform elem.Document -> host doc
            public Transform ToLinkT => ToHostT?.Inverse;

            public static PickedElem Host(Element e) => new PickedElem
            {
                Elem = e,
                LinkInst = null,
                ToHostT = Transform.Identity
            };

            public static PickedElem Linked(Element e, RevitLinkInstance li) => new PickedElem
            {
                Elem = e,
                LinkInst = li,
                ToHostT = li.GetTotalTransform()
            };

            public Solid ToHost(Solid s) => IsLinked ? SolidUtils.CreateTransformed(s, ToHostT) : s;
            public Curve ToHost(Curve c) => IsLinked ? c.CreateTransformed(ToHostT) : c;

            public XYZ ToHostPoint(XYZ p) => IsLinked ? ToHostT.OfPoint(p) : p;
            public XYZ ToHostVec(XYZ v) => IsLinked ? ToHostT.OfVector(v) : v;
            public XYZ ToLinkPoint(XYZ pHost) => IsLinked ? ToLinkT.OfPoint(pHost) : pHost;
        }

        private static HashSet<ElementId> BuildAllowedLinkInstanceIdSet(Document hostDoc, IList<Document> allowedLinkDocs)
        {
            var set = new HashSet<ElementId>();
            if (allowedLinkDocs == null || allowedLinkDocs.Count == 0) return set;

            var links = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();

            foreach (var li in links)
            {
                Document ldoc = li.GetLinkDocument();
                if (ldoc == null) continue;
                if (allowedLinkDocs.Contains(ldoc)) set.Add(li.Id);
            }
            return set;
        }

        // ===================== Selection filters (host & link) =====================
        private class HostDocFilter : ISelectionFilter
        {
            private readonly Func<Element, bool> _predicate;
            public HostDocFilter(Func<Element, bool> predicate) { _predicate = predicate; }
            public bool AllowElement(Element e) => e != null && _predicate(e);
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        private class LinkedFilter : ISelectionFilter
        {
            private readonly Document _hostDoc;
            private readonly HashSet<ElementId> _allowedLinkInstIds;
            private readonly Func<Element, bool> _predicate;

            public LinkedFilter(Document hostDoc, HashSet<ElementId> allowedLinkInstIds, Func<Element, bool> predicate)
            {
                _hostDoc = hostDoc;
                _allowedLinkInstIds = allowedLinkInstIds ?? new HashSet<ElementId>();
                _predicate = predicate;
            }

            public bool AllowElement(Element e)
            {
                var li = e as RevitLinkInstance;
                return li != null && _allowedLinkInstIds.Contains(li.Id);
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                if (reference == null) return false;

                // Kiểm tra link instance có trong danh sách cho phép
                ElementId linkInstId = reference.ElementId;
                if (!_allowedLinkInstIds.Contains(linkInstId)) return false;

                var li = _hostDoc.GetElement(linkInstId) as RevitLinkInstance;
                var ldoc = li?.GetLinkDocument(); if (ldoc == null) return false;

                // Lấy element trong link để lọc đúng loại
                Element linkedElem = ldoc.GetElement(reference.LinkedElementId);
                if (linkedElem == null) return false;

                return _predicate(linkedElem);
            }
        }

        // ============================== Predicates ===============================
        private static bool IsWallOrFloor(Element e)
        {
            if (e?.Category == null) return false;
            var bic = (BuiltInCategory)e.Category.Id.GetValue();
            return bic == BuiltInCategory.OST_Walls || bic == BuiltInCategory.OST_Floors;
        }

        private static bool IsRectHost_WallOrFloor(Element e) => IsWallOrFloor(e);

        private static bool IsRectMEP(Element e)
        {
            // Duct (chữ nhật) hoặc CableTray (duct tròn sẽ bị loại ở bước kích thước)
            return (e is Duct) || (e is CableTray);
        }

        // ====================== PARAM & ALIGN HELPERS ======================
        private static void SetParam(FamilyInstance inst, IEnumerable<string> names, double val)
        {
            foreach (var name in names)
            {
                Parameter p = inst.LookupParameter(name);
                if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Double)
                { p.Set(val); return; }
            }
        }

        private static void SetParam(FamilyInstance inst, string name, double val)
        {
            Parameter p = inst.LookupParameter(name);
            if (p != null && !p.IsReadOnly) p.Set(val);
        }

        private static void SetTextParam(Element e, string name, string val)
        {
            var p = e.LookupParameter(name);
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.String) p.Set(val);
        }

        private static XYZ ProjectOnPlane(XYZ v, XYZ normal)
        {
            double dot = v.DotProduct(normal);
            XYZ p = v - dot * normal;
            return (p.GetLength() < 1e-9) ? null : p;
        }

        // ============================== VIEW / LEVEL ===============================
        private static ViewPlan FindPlanViewForLevel(Document doc, ElementId levelId)
        {
            if (levelId == ElementId.InvalidElementId) return null;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.FloorPlan && v.GenLevel != null)
                .FirstOrDefault(v => v.GenLevel.Id == levelId);
        }

        private static ElementId GetMEPReferenceLevelId(Document hostDoc, Element mep, XYZ samplePoint)
        {
            Parameter p = mep.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
            if (p != null && p.StorageType == StorageType.ElementId)
            {
                ElementId id = p.AsElementId();
                if (id != ElementId.InvalidElementId) return id;
            }

            foreach (string name in new[] { "Reference Level", "Referance Level", "Level" })
            {
                p = mep.LookupParameter(name);
                if (p != null && p.StorageType == StorageType.ElementId)
                {
                    ElementId id = p.AsElementId();
                    if (id != ElementId.InvalidElementId) return id;
                }
            }

            p = mep.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (p != null && p.StorageType == StorageType.ElementId)
            {
                ElementId id = p.AsElementId();
                if (id != ElementId.InvalidElementId) return id;
            }

            // Fallback: level gần nhất trong HOST DOC theo cao độ
            return GetNearestLevelId(hostDoc, samplePoint.Z);
        }

        private static void TrySetScheduleLevel(Document doc, FamilyInstance inst, ElementId levelId)
        {
            if (levelId == ElementId.InvalidElementId) return;

            var candidates = new List<BuiltInParameter>
            {
                BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,
                BuiltInParameter.LEVEL_PARAM,
                BuiltInParameter.FAMILY_LEVEL_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM
            };

            foreach (var bip in candidates)
            {
                Parameter p = inst.get_Parameter(bip);
                if (p != null && !p.IsReadOnly)
                {
                    try { p.Set(levelId); return; } catch { }
                }
            }
        }

        private static ElementId GetNearestLevelId(Document doc, double z)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();
            if (levels.Count == 0) return ElementId.InvalidElementId;

            Level best = levels.OrderBy(l => Math.Abs(l.Elevation - z)).FirstOrDefault();
            return best != null ? best.Id : ElementId.InvalidElementId;
        }

        // ============================== GEOMETRY ===============================
        private static Solid GetMainSolid(Element elem)
        {
            var opts = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement ge = elem.get_Geometry(opts);
            if (ge == null) return null;

            Solid best = null; double maxVol = 0.0;
            foreach (GeometryObject go in ge)
            {
                if (go is Solid s && s.Volume > 1e-9)
                { if (s.Volume > maxVol) { best = s; maxVol = s.Volume; } }
                else if (go is GeometryInstance gi)
                {
                    var ige = gi.GetInstanceGeometry();
                    if (ige == null) continue;
                    foreach (GeometryObject igo in ige)
                    {
                        if (igo is Solid si && si.Volume > 1e-9)
                        { if (si.Volume > maxVol) { best = si; maxVol = si.Volume; } }
                    }
                }
            }
            return best;
        }

        // MẶT TRÊN (planar): normal có Z dương, lấy cao nhất
        private static bool TryGetTopPlanarFace(Element mep, out PlanarFace bestFace)
        {
            bestFace = null; double bestZ = double.NegativeInfinity;

            Options opts = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement ge = mep.get_Geometry(opts);
            if (ge == null) return false;

            foreach (var go in ge)
            {
                if (go is Solid s && s.Volume > 1e-9)
                {
                    foreach (Face f in s.Faces)
                    {
                        if (f is PlanarFace pf)
                        {
                            XYZ n = pf.FaceNormal.Normalize();
                            if (n.Z <= 0.25) continue; // hướng lên
                            BoundingBoxUV bb = pf.GetBoundingBox();
                            XYZ p1 = pf.Evaluate(new UV(bb.Min.U, bb.Min.V));
                            XYZ p2 = pf.Evaluate(new UV(bb.Max.U, bb.Max.V));
                            double z = Math.Max(p1.Z, p2.Z);
                            if (z > bestZ) { bestZ = z; bestFace = pf; }
                        }
                    }
                }
            }
            return bestFace != null;
        }

        // Lấy 2 trục ngang của RECT (widthAxis & heightAxis) từ các mặt bên đứng
        private static bool TryGetRectCrossAxes(Element mep, out XYZ widthAxis, out XYZ heightAxis)
        {
            widthAxis = heightAxis = null;

            Options opts = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement ge = mep.get_Geometry(opts);
            if (ge == null) return false;

            var normals = new List<XYZ>();
            foreach (var go in ge)
            {
                if (go is Solid s && s.Volume > 1e-9)
                {
                    foreach (Face f in s.Faces)
                    {
                        if (f is PlanarFace pf)
                        {
                            XYZ n = pf.FaceNormal;
                            if (Math.Abs(n.Z) < 0.25) // mặt đứng → normal ngang
                            {
                                XYZ hn = new XYZ(n.X, n.Y, 0);
                                if (hn.GetLength() > 1e-6) normals.Add(hn.Normalize());
                            }
                        }
                    }
                }
            }
            if (normals.Count == 0) return false;

            XYZ a = normals[0];
            double best = 1.0;
            XYZ bBest = null;
            foreach (var b in normals)
            {
                double dot = Math.Abs(a.DotProduct(b));
                if (dot < best && (b - a).GetLength() > 1e-6 && (b + a).GetLength() > 1e-6)
                { best = dot; bBest = b; }
            }
            if (bBest == null)
            {
                bBest = AnyPerpTo(a);
                bBest = new XYZ(bBest.X, bBest.Y, 0).Normalize();
            }

            widthAxis = a;
            heightAxis = bBest;
            return true;
        }

        // Tìm mặt BÊN có normal ≈ ±axis (axis = widthAxis hoặc heightAxis)
        private static bool TryGetRectSideFaceByAxis(Element mep, XYZ axis, double minDot, out PlanarFace sideFace)
        {
            sideFace = null;
            if (axis == null || axis.GetLength() < 1e-6) return false;

            Options opts = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement ge = mep.get_Geometry(opts);
            if (ge == null) return false;

            double bestDot = -1.0;
            foreach (var go in ge)
            {
                if (go is Solid s && s.Volume > 1e-9)
                {
                    foreach (Face f in s.Faces)
                    {
                        if (f is PlanarFace pf)
                        {
                            XYZ n = pf.FaceNormal.Normalize();
                            if (Math.Abs(n.Z) > 0.25) continue; // chỉ mặt đứng
                            double dot = Math.Abs(n.DotProduct(axis));
                            if (dot > bestDot) { bestDot = dot; sideFace = pf; }
                        }
                    }
                }
            }
            return sideFace != null && bestDot > minDot;
        }

        // Kích thước RECT (feet) – không tính insulation
        private static bool TryGetRectSizesFeet(Element mep, out double width, out double height)
        {
            width = height = 0.0;

            if (mep is Duct d)
            {
                // Nếu là duct tròn → có Diameter → bỏ
                double dia = GetParamDouble(d, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (dia > 1e-9) return false;

                double w = GetParamDouble(d, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                double h = GetParamDouble(d, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (w <= 0 || h <= 0)
                {
                    ElementType dt = d.Document.GetElement(d.GetTypeId()) as ElementType;
                    if (dt != null)
                    {
                        if (w <= 0) w = GetParamDouble(dt, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                        if (h <= 0) h = GetParamDouble(dt, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                    }
                }
                width = w; height = h;
                return (width > 0 && height > 0);
            }

            if (mep is CableTray ct)
            {
                double w = GetParamDouble(ct, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                double h = GetParamDouble(ct, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                if (w <= 0 || h <= 0)
                {
                    ElementType tt = ct.Document.GetElement(ct.GetTypeId()) as ElementType;
                    if (tt != null)
                    {
                        if (w <= 0) w = GetParamDouble(tt, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                        if (h <= 0) h = GetParamDouble(tt, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                    }
                }
                width = w; height = h;
                return (width > 0 && height > 0);
            }

            return false;
        }

        private static double GetParamDouble(Element e, BuiltInParameter bip)
        {
            Parameter p = e.get_Parameter(bip);
            if (p != null && p.StorageType == StorageType.Double) return p.AsDouble();
            return 0.0;
        }

        // Insulation (feet) – duct có thể có; cable tray thường 0
        private static double GetRectInsulationThicknessFeet(Element mep)
        {
            double best = 0.0;
            var deps = mep.GetDependentElements(null);
            if (deps == null || deps.Count == 0) return 0.0;

            Document doc = mep.Document;
            foreach (var id in deps)
            {
                Element e = doc.GetElement(id);
                if (e?.Category == null) continue;
                if (e.Category.Id.GetValue() != (int)BuiltInCategory.OST_DuctInsulations) continue;

                double th = 0.0;
                foreach (string name in new[] { "Thickness", "Insulation Thickness", "Độ dày", "Độ dày cách nhiệt" })
                {
                    Parameter p = e.LookupParameter(name);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        th = p.AsDouble();
                        if (th > 0) break;
                    }
                }

                if (th <= 0.0)
                {
                    foreach (Parameter p in e.Parameters)
                    {
                        if (p.StorageType == StorageType.Double)
                        {
                            double v = p.AsDouble();
                            if (v > th && v < 5.0) th = v;
                        }
                    }
                }
                if (th > best) best = th;
            }
            return best;
        }

        // Midpoint của đoạn giao (cho Floor)
        private static XYZ MidPointOfIntersection(SolidCurveIntersection sci)
        {
            Curve seg = sci.GetCurveSegment(0);
            XYZ p0 = seg.GetEndPoint(0);
            XYZ p1 = seg.GetEndPoint(1);
            return new XYZ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5, (p0.Z + p1.Z) * 0.5);
        }

        private static XYZ GetCurveDirection(Curve c)
        {
            if (c is Line ln) return ln.Direction.Normalize();
            XYZ d = (c.GetEndPoint(1) - c.GetEndPoint(0)).Normalize();
            return IsZeroLength(d) ? XYZ.BasisX : d;
        }

        private static XYZ AnyPerpTo(XYZ u)
        {
            XYZ guess = Math.Abs(u.Z) < 0.9 ? XYZ.BasisZ : XYZ.BasisX;
            XYZ t = guess.CrossProduct(u);
            if (t.GetLength() < 1e-9) t = XYZ.BasisY.CrossProduct(u);
            return t.Normalize();
        }

        private static bool IsZeroLength(XYZ v) => v == null || v.GetLength() < 1e-9;
        // Tìm mặt của solid đi qua đúng điểm p (dùng khi suy ra pháp tuyến tại giao điểm)
        private static bool TryFindFaceAtPoint(Solid solid, XYZ p, out Face face, out UV uv)
        {
            face = null; uv = null;
            if (solid == null) return false;

            foreach (Face f in solid.Faces)
            {
                var proj = f.Project(p);
                if (proj == null) continue;
                if (proj.Distance < 1e-6) // gần như đúng trên mặt
                {
                    face = f;
                    uv = proj.UVPoint;
                    return true;
                }
            }
            return false;
        }


        // ---------------- THICKNESS ----------------
        private static double GetHostThicknessFeet(Element host)
        {
            if (host is Wall w) return w.Width;

            if (host is Floor f)
            {
                Parameter p = f.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                if (p != null && p.StorageType == StorageType.Double) return p.AsDouble();

                ElementType ft = host.Document.GetElement(host.GetTypeId()) as ElementType;
                if (ft != null)
                {
                    Parameter pt = ft.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                    if (pt != null && pt.StorageType == StorageType.Double) return pt.AsDouble();
                }
            }
            return 0.0;
        }
    }
}
