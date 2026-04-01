using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;   // CableTray / Conduit
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
#nullable disable

namespace OPENING.MODEL
{
    public static class TH_RECTANGULAR_MULTI_SLEEVE
    {
        private const double TOL = 1e-6;

        // Giữ overload cũ để tương thích
        public static void Run(ExternalCommandData commandData, Document doc,
                               double clearance, double extrusion, double fill)
            => Run(commandData, doc, clearance, extrusion, fill, new List<Document>());

        // NEW: overload nhận danh sách link được tick
        public static void Run(ExternalCommandData commandData, Document doc,
                               double clearance, double extrusion, double fill,
                               IList<Document> allowedLinkDocs)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            // bật/tắt pick link
            var allowedLinkInstIds = BuildAllowedLinkInstanceIdSet(doc, allowedLinkDocs);
            bool enableLinkPick = allowedLinkInstIds != null && allowedLinkInstIds.Count > 0;

            // 1) Pick host: Wall/Floor (host + link nếu có tick)
            var hostPicks = new List<PickedElem>();
            try
            {
                var msgHost = enableLinkPick
                    ? "Select Wall/Floor in the current file (Press Esc to finish host selection)"
                    : "Select Wall/Floor in the current file";
                var hostRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element, new HostDocFilter(IsWallOrFloor), msgHost);
                foreach (var r in hostRefs)
                {
                    var e = doc.GetElement(r);
                    if (e != null) hostPicks.Add(PickedElem.Host(e));
                }
            }
            catch { /* Esc */ }

            if (enableLinkPick)
            {
                try
                {
                    var linkRefs = uidoc.Selection.PickObjects(
                        ObjectType.LinkedElement,
                        new LinkedFilter(doc, allowedLinkInstIds, IsWallOrFloor),
                        "Select Wall/Floor from the checked LINK files (Press Esc to finish selection)");
                    foreach (var r in linkRefs)
                    {
                        var li = doc.GetElement(r.ElementId) as RevitLinkInstance;
                        var ldoc = li?.GetLinkDocument(); if (ldoc == null) continue;
                        var e = ldoc.GetElement(r.LinkedElementId); if (e == null) continue;
                        hostPicks.Add(PickedElem.Linked(e, li));
                    }
                }
                catch { /* Esc */ }
            }

            if (hostPicks.Count == 0) return;

            // 2) Pick >= 2 MEP (Pipe/Duct/CableTray/Conduit) – host + link nếu có tick
            var mepPicks = new List<PickedElem>();
            try
            {
                var msgHost = enableLinkPick
                    ? "Select 2 or more: Pipe / Duct / Cable Tray / Conduit in the current file (Press Esc to finish host selection)" 
                    : "Select 2 or more: Pipe / Duct / Cable Tray / Conduit in the current file";
                var mepRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element, new HostDocFilter(IsMultiMEP), msgHost);
                foreach (var r in mepRefs)
                {
                    var e = doc.GetElement(r);
                    if (e != null) mepPicks.Add(PickedElem.Host(e));
                }
            }
            catch { /* Esc */ }

            if (enableLinkPick)
            {
                try
                {
                    var mepLinkRefs = uidoc.Selection.PickObjects(
                        ObjectType.LinkedElement,
                        new LinkedFilter(doc, allowedLinkInstIds, IsMultiMEP),
                        "Select additional elements from the checked LINK files (Press Esc to finish selection)");
                    foreach (var r in mepLinkRefs)
                    {
                        var li = doc.GetElement(r.ElementId) as RevitLinkInstance;
                        var ldoc = li?.GetLinkDocument(); if (ldoc == null) continue;
                        var e = ldoc.GetElement(r.LinkedElementId); if (e == null) continue;
                        mepPicks.Add(PickedElem.Linked(e, li));
                    }
                }
                catch { /* Esc */ }
            }

            if (mepPicks.Count < 2)
            {
                TaskDialog.Show("OPENING", "At least 2 MEP elements must be selected.");
                return;
            }

            // 3) Family
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family != null && s.Family.Name == "TH_RECTANGULAR_MULTI_SLEEVE");
            if (symbol == null)
            {
                TaskDialog.Show("OPENING", "Family 'TH_RECTANGULAR_MULTI_SLEEVE' not found'.");
                return;
            }

            // 4) Xử lý cho từng hostPick (có thể chọn nhiều tường/sàn)
            foreach (var hostPick in hostPicks)
            {
                Solid hostSolidL = GetMainSolid(hostPick.Elem);
                if (hostSolidL == null) continue;

                Solid hostSolidH = hostPick.ToHost(hostSolidL); // về host coords
                double hostThk = GetHostThicknessFeet(hostPick.Elem);

                if (hostPick.Elem is Wall)
                {
                    ProcessWall(doc, uidoc, hostPick, hostSolidH, hostThk, mepPicks, symbol,
                                clearance, extrusion, fill);
                }
                else if (hostPick.Elem is Floor)
                {
                    ProcessFloor(doc, uidoc, hostPick, hostSolidH, hostThk, mepPicks, symbol,
                                 clearance, extrusion, fill);
                }
            }
        }

        // ================= WALL (link-aware, host coords) =================
        private static void ProcessWall(
            Document doc, UIDocument uidoc, PickedElem hostPick, Solid wallSolidH, double hostThickness,
            List<PickedElem> mepPicks, FamilySymbol symbol,
            double clearance, double extrusion, double fill)
        {
            // Face để lấy Reference: H cho host, L cho link
            PlanarFace refFaceH = null;
            PlanarFace refFaceL = null;
            Solid wallSolidL = GetMainSolid(hostPick.Elem); // solid gốc (link coords) – để lấy Reference hợp lệ khi host là link

            // Tìm mặt tường tại giao đầu tiên
            foreach (var mp in mepPicks)
            {
                var lc = mp.Elem.Location as LocationCurve;
                if (lc?.Curve == null) continue;

                Curve axisH = mp.ToHost(lc.Curve);
                if (!TryGetHostCurveEndpoints(wallSolidH, axisH, out XYZ a0H, out XYZ a1H)) continue;

                if (!hostPick.IsLinked)
                {
                    if (!TryFindFaceAtPoint_Wall(wallSolidH, a0H, out refFaceH))
                        TryFindFaceAtPoint_Wall(wallSolidH, a1H, out refFaceH);
                    if (refFaceH != null) break;
                }
                else
                {
                    XYZ a0L = hostPick.ToLinkPoint(a0H);
                    XYZ a1L = hostPick.ToLinkPoint(a1H);
                    if (!TryFindFaceAtPoint_Wall(wallSolidL, a0L, out refFaceL))
                        TryFindFaceAtPoint_Wall(wallSolidL, a1L, out refFaceL);
                    if (refFaceL != null) break;
                }
            }
            if (!hostPick.IsLinked && refFaceH == null) return;
            if (hostPick.IsLinked && refFaceL == null) return;

            // Hệ trục trên mặt (HOST coords)
            XYZ nH, oH;
            if (!hostPick.IsLinked) { nH = refFaceH.FaceNormal.Normalize(); oH = refFaceH.Origin; }
            else { nH = hostPick.ToHostVec(refFaceL.FaceNormal).Normalize(); oH = hostPick.ToHostPoint(refFaceL.Origin); }

            XYZ xDirH = XYZ.BasisZ.CrossProduct(nH);
            if (xDirH.GetLength() < 1e-9) xDirH = AnyPerpTo(nH);
            xDirH = xDirH.Normalize();
            XYZ yDirH = nH.CrossProduct(xDirH).Normalize();

            // Gom items → bounding (HOST coords)
            var items = new List<ItemOnFace>();
            foreach (var mp in mepPicks)
            {
                var lc = mp.Elem.Location as LocationCurve;
                if (lc?.Curve == null) continue;

                Curve axisH = mp.ToHost(lc.Curve);
                if (!TryGetHostCurveEndpoints(wallSolidH, axisH, out XYZ p0H, out XYZ p1H)) continue;

                XYZ pOuterH = PickEndpointClosestToFace(new PlanarFaceProxy(nH, oH), p0H, p1H); // dùng proxy để đo khoảng cách theo nH/oH
                XYZ pMidH = pOuterH - nH * (hostThickness / 2.0);

                double ins = GetMEPInsulationThicknessFeet(mp.Elem);
                double hw = GetHalfWidthBare(mp.Elem) + ins;
                double hh = GetHalfHeightBare(mp.Elem) + ins;

                XYZ rel = pMidH - oH;
                items.Add(new ItemOnFace { Elem = mp.Elem, X = rel.DotProduct(xDirH), Y = rel.DotProduct(yDirH), HalfW = hw, HalfH = hh });
            }
            if (items.Count < 2) return;

            ComputeBounding(items, out double xMid, out double yMid, out double width, out double height);
            XYZ placeOnFaceH = oH + xDirH * xMid + yDirH * yMid;

            // Reference hợp lệ
            Reference faceRef = !hostPick.IsLinked
                ? refFaceH.Reference
                : refFaceL.Reference.CreateLinkReference(hostPick.LinkInst);

            PlaceAndParam(doc, uidoc, faceRef, placeOnFaceH, xDirH, symbol,
                          width, height, hostThickness, clearance, extrusion, fill,
                          pushToMid: -hostThickness / 2.0);
        }

        // ================= FLOOR (link-aware, host coords) =================
        private static void ProcessFloor(
            Document doc, UIDocument uidoc, PickedElem hostPick, Solid floorSolidH, double hostThickness,
            List<PickedElem> mepPicks, FamilySymbol symbol,
            double clearance, double extrusion, double fill)
        {
            PlanarFace topFaceH = null;
            PlanarFace topFaceL = null;
            Solid floorSolidL = GetMainSolid(hostPick.Elem);

            XYZ hintDirH = null;

            foreach (var mp in mepPicks)
            {
                var lc = mp.Elem.Location as LocationCurve;
                if (lc?.Curve == null) continue;

                Curve axisH = mp.ToHost(lc.Curve);
                if (!TryGetHostCurveEndpoints(floorSolidH, axisH, out XYZ q0H, out XYZ q1H)) continue;

                if (!hostPick.IsLinked)
                {
                    if (!TryFindTopFaceAtPoint_Floor(floorSolidH, q0H, out topFaceH))
                        TryFindTopFaceAtPoint_Floor(floorSolidH, q1H, out topFaceH);
                    if (topFaceH != null) { hintDirH = GetCurveDirection(axisH); break; }
                }
                else
                {
                    XYZ q0L = hostPick.ToLinkPoint(q0H);
                    XYZ q1L = hostPick.ToLinkPoint(q1H);
                    if (!TryFindTopFaceAtPoint_Floor(floorSolidL, q0L, out topFaceL))
                        TryFindTopFaceAtPoint_Floor(floorSolidL, q1L, out topFaceL);
                    if (topFaceL != null) { hintDirH = GetCurveDirection(axisH); break; }
                }
            }
            if (!hostPick.IsLinked && topFaceH == null) return;
            if (hostPick.IsLinked && topFaceL == null) return;

            // Hệ trục trên mặt (HOST coords)
            XYZ nH, oH;
            if (!hostPick.IsLinked) { nH = topFaceH.FaceNormal.Normalize(); oH = topFaceH.Origin; }
            else { nH = hostPick.ToHostVec(topFaceL.FaceNormal).Normalize(); oH = hostPick.ToHostPoint(topFaceL.Origin); }

            XYZ xDirH = ProjectOnPlane(hintDirH ?? XYZ.BasisX, nH);
            if (xDirH == null) xDirH = AnyPerpTo(nH);
            xDirH = xDirH.Normalize();
            XYZ yDirH = nH.CrossProduct(xDirH).Normalize();

            // Gom items → bounding
            var items = new List<ItemOnFace>();
            foreach (var mp in mepPicks)
            {
                var lc = mp.Elem.Location as LocationCurve;
                if (lc?.Curve == null) continue;

                Curve axisH = mp.ToHost(lc.Curve);
                if (!TryGetHostCurveEndpoints(floorSolidH, axisH, out XYZ p0H, out XYZ p1H)) continue;

                // chọn điểm thuộc mặt TOP (HOST coords)
                XYZ pOnTopH = TryFindTopFaceAtPoint_Floor(floorSolidH, p0H, out _) ? p0H : p1H;
                XYZ pMidH = pOnTopH - nH * (hostThickness / 2.0);

                double ins = GetMEPInsulationThicknessFeet(mp.Elem);
                double hw = GetHalfWidthBare(mp.Elem) + ins;
                double hh = GetHalfHeightBare(mp.Elem) + ins;

                XYZ rel = pMidH - oH;
                items.Add(new ItemOnFace { Elem = mp.Elem, X = rel.DotProduct(xDirH), Y = rel.DotProduct(yDirH), HalfW = hw, HalfH = hh });
            }
            if (items.Count < 2) return;

            ComputeBounding(items, out double xM, out double yM, out double w, out double h);
            XYZ placeOnFaceH = oH + xDirH * xM + yDirH * yM;

            Reference faceRef = !hostPick.IsLinked
                ? topFaceH.Reference
                : topFaceL.Reference.CreateLinkReference(hostPick.LinkInst);

            PlaceAndParam(doc, uidoc, faceRef, placeOnFaceH, xDirH, symbol,
                          w, h, hostThickness, clearance, extrusion, fill,
                          pushToMid: -hostThickness / 2.0);
        }


        // ---------------- PLACE + PARAMS (dùng chung) ----------------
        private static void PlaceAndParam(Document doc, UIDocument uidoc, Reference faceRef, XYZ placeOnFace, XYZ xDir,
                                          FamilySymbol symbol, double width, double height, double hostThickness,
                                          double clearance, double extrusion, double fill, double pushToMid)
        {
            ElementId lvlId = GetNearestLevelId(doc, placeOnFace.Z);
            ViewPlan plan = FindPlanViewForLevel(doc, lvlId);
            ElementId prevViewId = uidoc.ActiveView?.Id ?? ElementId.InvalidElementId;
            if (plan != null) uidoc.ActiveView = plan;

            using (Transaction t = new Transaction(doc, "Place TH_RECTANGULAR_MULTI_SLEEVE"))
            {
                t.Start();
                if (!symbol.IsActive) symbol.Activate();

                var inst = doc.Create.NewFamilyInstance(faceRef, placeOnFace, xDir, symbol);
                doc.Regenerate();

                SetIfExists(inst, "Width", width);
                SetIfExists(inst, "Height", height);
                SetIfExists(inst, "Host Thickness", hostThickness);

                SetIfExists(inst, "Clearance", clearance);
                SetIfExists(inst, "Extrusion", extrusion);
                SetIfExists(inst, "Fill", fill);

                TrySetHostOffset(inst, pushToMid);
                TrySetScheduleLevel(doc, inst, lvlId);
                t.Commit();
            }

            if (plan != null && prevViewId != ElementId.InvalidElementId)
            {
                View prev = doc.GetElement(prevViewId) as View;
                if (prev != null) uidoc.ActiveView = prev;
            }
        }

        // ---------------- BOUNDING (dùng chung) ----------------
        private static void ComputeBounding(List<ItemOnFace> items,
                                            out double xMid, out double yMid,
                                            out double width, out double height)
        {
            var xMinItem = items.OrderBy(it => it.X - it.HalfW).First();
            var xMaxItem = items.OrderByDescending(it => it.X + it.HalfW).First();
            double xMin = xMinItem.X - xMinItem.HalfW;
            double xMax = xMaxItem.X + xMaxItem.HalfW;
            width = xMax - xMin;
            xMid = 0.5 * (xMin + xMax);

            var yMinItem = items.OrderBy(it => it.Y - it.HalfH).First();
            var yMaxItem = items.OrderByDescending(it => it.Y + it.HalfH).First();
            double yMin = yMinItem.Y - yMinItem.HalfH;
            double yMax = yMaxItem.Y + yMaxItem.HalfH;
            height = yMax - yMin;
            yMid = 0.5 * (yMin + yMax);
        }

        private class ItemOnFace
        {
            public Element Elem;
            public double X, Y;
            public double HalfW, HalfH;
        }

        // ---------------- Link-aware types & filters ----------------
        private class PickedElem
        {
            public Element Elem;
            public RevitLinkInstance LinkInst;   // null nếu host
            public bool IsLinked => LinkInst != null;
            public Transform ToHostT;            // elem.Doc -> host Doc
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

        private class HostDocFilter : ISelectionFilter
        {
            private readonly Func<Element, bool> _pred;
            public HostDocFilter(Func<Element, bool> pred) => _pred = pred;
            public bool AllowElement(Element e) => e != null && _pred(e);
            public bool AllowReference(Reference r, XYZ p) => true;
        }

        private class LinkedFilter : ISelectionFilter
        {
            private readonly Document _hostDoc;
            private readonly HashSet<ElementId> _allowedLinkInstIds;
            private readonly Func<Element, bool> _pred;

            public LinkedFilter(Document hostDoc, HashSet<ElementId> allowedLinkInstIds, Func<Element, bool> pred)
            {
                _hostDoc = hostDoc;
                _allowedLinkInstIds = allowedLinkInstIds ?? new HashSet<ElementId>();
                _pred = pred;
            }

            public bool AllowElement(Element e)
            {
                var li = e as RevitLinkInstance;
                return li != null && _allowedLinkInstIds.Contains(li.Id);
            }

            public bool AllowReference(Reference r, XYZ p)
            {
                if (r == null) return false;
                if (!_allowedLinkInstIds.Contains(r.ElementId)) return false;

                var li = _hostDoc.GetElement(r.ElementId) as RevitLinkInstance;
                var ldoc = li?.GetLinkDocument(); if (ldoc == null) return false;
                var linkedElem = ldoc.GetElement(r.LinkedElementId);
                return linkedElem != null && _pred(linkedElem);
            }
        }

        // ---------------- GEOMETRY HELPERS ----------------
        private static bool TryGetHostCurveEndpoints(Solid hostSolid, Curve axis, out XYZ p0, out XYZ p1)
        {
            p0 = p1 = null;
            var sci = hostSolid.IntersectWithCurve(axis, new SolidCurveIntersectionOptions());
            if (sci == null || sci.SegmentCount == 0) return false;
            Curve seg = sci.GetCurveSegment(0);
            p0 = seg.GetEndPoint(0);
            p1 = seg.GetEndPoint(1);
            return true;
        }

        // Chọn mặt VERT cho tường tại điểm p (HOST coords)
        private static bool TryFindFaceAtPoint_Wall(Solid solid, XYZ p, out PlanarFace pf)
        {
            pf = null;
            foreach (Face f in solid.Faces)
            {
                if (!(f is PlanarFace pp)) continue;
                var pr = pp.Project(p);
                if (pr == null || pr.Distance > 1e-6) continue;
                if (Math.Abs(pp.FaceNormal.Z) < 0.25) { pf = pp; return true; }
            }
            return false;
        }

        // Chọn mặt TOP cho sàn tại điểm p (HOST coords)
        private static bool TryFindTopFaceAtPoint_Floor(Solid solid, XYZ p, out PlanarFace top)
        {
            top = null; double bestZ = double.NegativeInfinity;
            foreach (Face f in solid.Faces)
            {
                if (!(f is PlanarFace pf)) continue;
                if (pf.FaceNormal.Z <= 0.25) continue;

                var pr = pf.Project(p);
                if (pr == null || pr.Distance > 1e-6) continue;

                var bb = pf.GetBoundingBox();
                XYZ a = pf.Evaluate(new UV(bb.Min.U, bb.Min.V));
                XYZ b = pf.Evaluate(new UV(bb.Max.U, bb.Max.V));
                double z = Math.Max(a.Z, b.Z);
                if (z > bestZ) { bestZ = z; top = pf; }
            }
            return top != null;
        }

        private static XYZ PickEndpointClosestToFace(PlanarFace face, XYZ a, XYZ b)
        {
            double da = Math.Abs(DistanceToPlane(face, a));
            double db = Math.Abs(DistanceToPlane(face, b));
            return (da <= db) ? a : b;
        }

        private static double DistanceToPlane(PlanarFace face, XYZ p)
        {
            XYZ n = face.FaceNormal.Normalize();
            XYZ o = face.Origin;
            return (p - o).DotProduct(n);
        }

        private static XYZ GetCurveDirection(Curve c)
        {
            if (c is Line ln) return ln.Direction.Normalize();
            XYZ d = (c.GetEndPoint(1) - c.GetEndPoint(0)).Normalize();
            return d.GetLength() < 1e-9 ? XYZ.BasisX : d;
        }

        private static XYZ ProjectOnPlane(XYZ v, XYZ normal)
        {
            if (v == null) return null;
            double dot = v.DotProduct(normal);
            XYZ p = v - dot * normal;
            return (p.GetLength() < 1e-9) ? null : p;
        }

        private static XYZ AnyPerpTo(XYZ n)
        {
            XYZ guess = Math.Abs(n.Z) < 0.9 ? XYZ.BasisZ : XYZ.BasisX;
            XYZ t = guess.CrossProduct(n);
            if (t.GetLength() < 1e-9) t = XYZ.BasisY.CrossProduct(n);
            return t.Normalize();
        }

        private static bool IsWallOrFloor(Element e)
        {
            if (e?.Category == null) return false;
            var bic = (BuiltInCategory)e.Category.Id.GetValue();
            return bic == BuiltInCategory.OST_Walls || bic == BuiltInCategory.OST_Floors;
        }
        private static bool IsMultiMEP(Element e) =>
            (e is Pipe) || (e is Duct) || (e is CableTray) || (e is Conduit);

        // ---------------- SIZE / INSULATION ----------------
        private static double GetHalfWidthBare(Element e)
        {
            if (e is Pipe p)
            {
                double d = GetDouble(p, BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                if (d <= 0) d = GetDouble(p, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                return d > 0 ? d / 2.0 : 0.05;
            }
            if (e is Duct dct)
            {
                double dia = GetDouble(dct, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (dia > TOL) return dia / 2.0; // duct tròn
                double w = GetDouble(dct, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                if (w <= 0)
                {
                    ElementType dt = dct.Document.GetElement(dct.GetTypeId()) as ElementType;
                    if (dt != null) w = GetDouble(dt, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                }
                return w > 0 ? w / 2.0 : 0.05;
            }
            if (e is CableTray ct)
            {
                double w = GetDouble(ct, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                if (w <= 0)
                {
                    ElementType tt = ct.Document.GetElement(ct.GetTypeId()) as ElementType;
                    if (tt != null) w = GetDouble(tt, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                }
                return w > 0 ? w / 2.0 : 0.05;
            }
            if (e is Conduit c)
            {
                double d = GetDouble(c, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                if (d <= 0)
                {
                    ElementType t = c.Document.GetElement(c.GetTypeId()) as ElementType;
                    if (t != null) d = GetDouble(t, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                }
                return d > 0 ? d / 2.0 : 0.05;
            }
            return 0.05;
        }

        // Dùng như một "mặt phẳng" từ normal + origin để đo khoảng cách
        private sealed class PlanarFaceProxy
        {
            public XYZ N { get; }
            public XYZ O { get; }
            public PlanarFaceProxy(XYZ n, XYZ o) { N = n; O = o; }
        }
        private static XYZ PickEndpointClosestToFace(PlanarFaceProxy face, XYZ a, XYZ b)
        {
            double da = Math.Abs((a - face.O).DotProduct(face.N));
            double db = Math.Abs((b - face.O).DotProduct(face.N));
            return (da <= db) ? a : b;
        }

        private static double GetHalfHeightBare(Element e)
        {
            if (e is Pipe p)
            {
                double d = GetDouble(p, BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                if (d <= 0) d = GetDouble(p, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                return d > 0 ? d / 2.0 : 0.05;
            }
            if (e is Duct dct)
            {
                double dia = GetDouble(dct, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (dia > TOL) return dia / 2.0; // duct tròn
                double h = GetDouble(dct, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (h <= 0)
                {
                    ElementType dt = dct.Document.GetElement(dct.GetTypeId()) as ElementType;
                    if (dt != null) h = GetDouble(dt, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                }
                return h > 0 ? h / 2.0 : 0.05;
            }
            if (e is CableTray ct)
            {
                double h = GetDouble(ct, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                if (h <= 0)
                {
                    ElementType tt = ct.Document.GetElement(ct.GetTypeId()) as ElementType;
                    if (tt != null) h = GetDouble(tt, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                }
                return h > 0 ? h / 2.0 : 0.05;
            }
            if (e is Conduit c)
            {
                double d = GetDouble(c, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                if (d <= 0)
                {
                    ElementType t = c.Document.GetElement(c.GetTypeId()) as ElementType;
                    if (t != null) d = GetDouble(t, BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                }
                return d > 0 ? d / 2.0 : 0.05;
            }
            return 0.05;
        }

        private static double GetMEPInsulationThicknessFeet(Element mep)
        {
            if (mep is Conduit) return 0.0; // coi như không có
            double best = 0.0;
            var deps = mep.GetDependentElements(null);
            if (deps == null || deps.Count == 0) return 0.0;

            Document doc = mep.Document;
            foreach (var id in deps)
            {
                Element e = doc.GetElement(id);
                if (e?.Category == null) continue;

                // Revit 2025 dùng long cho ElementId
                long catVal = e.Category.Id.GetValue();

                bool isInsul = (catVal == (long)BuiltInCategory.OST_PipeInsulations) ||
                               (catVal == (long)BuiltInCategory.OST_DuctInsulations);
                if (!isInsul) continue;

                double th = 0.0;
                foreach (string name in new[] { "Thickness", "Insulation Thickness", "Độ dày", "Độ dày cách nhiệt" })
                {
                    Parameter p = e.LookupParameter(name);
                    if (p != null && p.StorageType == StorageType.Double)
                    { th = p.AsDouble(); if (th > 0) break; }
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

        private static double GetDouble(Element e, BuiltInParameter bip)
        {
            var p = e.get_Parameter(bip);
            return (p != null && p.StorageType == StorageType.Double) ? p.AsDouble() : 0.0;
        }

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

        // ---------------- LEVEL / PARAM HELPERS ----------------
        private static void SetIfExists(FamilyInstance inst, string paramName, double val)
        {
            Parameter p = inst.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Double) p.Set(val);
        }

        private static ViewPlan FindPlanViewForLevel(Document doc, ElementId levelId)
        {
            if (levelId == ElementId.InvalidElementId) return null;
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan)).Cast<ViewPlan>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.FloorPlan && v.GenLevel != null)
                .FirstOrDefault(v => v.GenLevel.Id == levelId);
        }

        private static ElementId GetNearestLevelId(Document doc, double z)
        {
            var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            if (levels.Count == 0) return ElementId.InvalidElementId;
            Level best = levels.OrderBy(l => Math.Abs(l.Elevation - z)).FirstOrDefault();
            return best != null ? best.Id : ElementId.InvalidElementId;
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
                { try { p.Set(levelId); return; } catch { } }
            }
        }

        private static void TrySetHostOffset(FamilyInstance inst, double offset)
        {
            foreach (string name in new[] { "Offset", "Host Offset", "Wall Offset", "Depth Offset" })
            {
                Parameter p = inst.LookupParameter(name);
                if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Double)
                { try { p.Set(offset); return; } catch { } }
            }
            foreach (var bip in new[] { BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM })
            {
                Parameter p = inst.get_Parameter(bip);
                if (p != null && !p.IsReadOnly) { try { p.Set(offset); return; } catch { } }
            }
        }
    }
}
