using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;   // Conduit
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
#nullable disable

namespace OPENING.MODEL
{
    /// <summary>
    /// Đặt TH_ROUND_SLEEVE – host trực tiếp lên mặt trụ của Pipe/Duct/Conduit tròn.
    /// NÂNG CẤP: Cho phép pick Wall/Floor và MEP tròn từ các Revit Link đã tick trong UI.
    /// Tự transform hình học từ link về host để tính toán giao, và tạo instance bằng Reference link.
    /// </summary>
    public static class TH_ROUND_SLEEVE
    {
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

            // 0) FamilySymbol "TH_ROUND_SLEEVE"
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family != null && s.Family.Name == "TH_ROUND_SLEEVE");

            if (symbol == null)
            {
                TaskDialog.Show("ERORR", "CAN NOT FOUND Family 'TH_ROUND_SLEEVE'.");
                return;
            }

            // Map danh sách doc link đã tick -> danh sách RevitLinkInstance.Id cho phép
            var allowedLinkInstIds = BuildAllowedLinkInstanceIdSet(doc, allowedLinkDocs);
            bool enableLinkPick = allowedLinkInstIds.Count > 0;

            // ===== 1) Chọn Wall/Floor =====
            var hostPicks = new List<PickedElem>();
            try
            {
                // Host doc luôn cho phép
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

            // Chỉ pick trong LINK khi có link đã tick
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
                        var ldoc = li?.GetLinkDocument();
                        if (ldoc == null) continue;

                        var e = ldoc.GetElement(r.LinkedElementId);
                        if (e == null) continue;

                        hostPicks.Add(PickedElem.Linked(e, li));
                    }
                }
                catch { /* user Esc */ }
            }

            if (hostPicks.Count == 0) return;


            // ===== 2) Chọn Pipe/Duct/Conduit (TRÒN) =====
            var mepPicks = new List<PickedElem>();
            try
            {
                var msgHost = enableLinkPick
                    ? "Select ROUND Pipe/Duct/Conduit in the current file (Press Esc to finish host selection)" 
                    : "Select ROUND Pipe/Duct/Conduit in the current file";
                var mepRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new HostDocFilter(IsRoundMEP),
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
                        new LinkedFilter(doc, allowedLinkInstIds, IsRoundMEP),
                        "Select ROUND Pipe/Duct/Conduit from the checked LINK files (Press Esc to finish selection)");

                    foreach (var r in mepLinkRefs)
                    {
                        var li = doc.GetElement(r.ElementId) as RevitLinkInstance;
                        var ldoc = li?.GetLinkDocument();
                        if (ldoc == null) continue;

                        var e = ldoc.GetElement(r.LinkedElementId);
                        if (e == null) continue;

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
                    Solid hostSolid = GetMainSolid(hostPick.Elem);
                    if (hostSolid == null) continue;

                    // Biến về hệ toạ độ host doc nếu là link
                    Solid hostSolidH = hostPick.ToHost(hostSolid);

                    double hostThicknessFeet = GetHostThicknessFeet(hostPick.Elem);

                    foreach (var mepPick in mepPicks)
                    {
                        if (!(mepPick.Elem?.Location is LocationCurve lc) || lc.Curve == null) continue;

                        // Trục ống trong hệ toạ độ host doc
                        Curve axisH = mepPick.ToHost(lc.Curve);
                        XYZ uH = GetCurveDirection(axisH);
                        if (IsZeroLength(uH)) continue;

                        // 3) Giao giữa trục ống (đã transform về host) và solid host (đã transform)
                        var sci = hostSolidH.IntersectWithCurve(axisH, new SolidCurveIntersectionOptions());
                        if (sci == null || sci.SegmentCount == 0) continue;

                        // Điểm giao phía mặt ngoài host (điểm đầu segment trong host coordinates)
                        Curve insideSegH = sci.GetCurveSegment(0);
                        XYZ centerOnPlaneH = insideSegH.GetEndPoint(0);

                        // Pháp tuyến mặt host tại điểm giao (trong host coordinates)
                        if (!TryFindFaceAtPoint(hostSolidH, centerOnPlaneH, out Face hostFaceH, out UV hostUV))
                            continue;

                        XYZ nHostH = hostFaceH.ComputeNormal(hostUV).Normalize();
                        XYZ centerWallH = centerOnPlaneH - nHostH * (hostThicknessFeet / 2.0);

                        // Lấy cylindrical face của MEP (dò trong doc gốc của MEP)
                        if (!TryGetCylindricalFace(mepPick.Elem, out CylindricalFace cylFace)) continue;

                        // Kích thước ống + cách nhiệt
                        double bareOuterDia = GetRoundOuterDiameterFeet(mepPick.Elem); // feet
                        double bareR = bareOuterDia / 2.0;
                        double insulT = GetInsulationThicknessFeet(mepPick.Elem);

                        // Vector bán kính nằm trên mặt host: rdir = n × u  (đều trong host coords)
                        XYZ rdirH = nHostH.CrossProduct(uH);
                        if (IsZeroLength(rdirH)) rdirH = AnyPerpTo(uH);
                        rdirH = rdirH.Normalize();

                        // Điểm đặt (host coords)
                        XYZ placePointH = centerWallH + rdirH * bareR;

                        // refDir: trục ống (host coords)
                        XYZ refDirH = uH;

                        // Level tham chiếu & chuyển view (trong host doc)
                        ElementId refLvlId = GetMEPReferenceLevelId(doc, mepPick.Elem, placePointH);
                        if (mepPick.IsLinked)  // LevelId từ link không dùng được
                            refLvlId = GetNearestLevelId(doc, placePointH.Z);

                        ViewPlan plan = FindPlanViewForLevel(doc, refLvlId);
                        if (plan != null) uidoc.ActiveView = plan;

                        using (Transaction t = new Transaction(doc, "Place TH_ROUND_SLEEVE (host ON round MEP – link-aware)"))
                        {
                            t.Start();
                            if (!symbol.IsActive) symbol.Activate();

                            // Chuẩn bị Reference hợp lệ trong HOST DOC
                            Reference hostRef = mepPick.IsLinked
                                ? cylFace.Reference.CreateLinkReference(mepPick.LinkInst)
                                : cylFace.Reference;


                            // Tạo instance – HOST LÊN MẶT ỐNG/CONDUIT (kể cả trong link)
                            FamilyInstance inst = doc.Create.NewFamilyInstance(hostRef, placePointH, refDirH, symbol);

                            // Bảo đảm đặt đúng host: 
                            // - Nếu MEP ở host doc: Host Id == mep.Id
                            // - Nếu MEP ở link: Host là RevitLinkInstance có cùng Id
                            bool okHost = true;
                            if (!mepPick.IsLinked)
                            {
                                okHost = ((inst.Host as Element)?.Id == mepPick.Elem.Id);
                            }
                            else
                            {
                                okHost = (inst.Host is RevitLinkInstance rli && rli.Id == mepPick.LinkInst.Id);
                            }

                            if (!okHost)
                            {
                                doc.Delete(inst.Id);
                                t.Commit();
                                continue;
                            }

                            doc.Regenerate();
                            // Xác định loại MEP để đếm về sau
                            string mepKind =
                                (mepPick.Elem is Pipe) ? "Pipe" :
                                (mepPick.Elem is Duct) ? "Duct" :
                                (mepPick.Elem is Conduit) ? "Conduit" :
                                (mepPick.Elem?.Category?.Name ?? "");

                            // Ghi vào tham số instance (text) trong family, ví dụ: TH_MEP_KIND
                            SetTextParam(inst, "TH_MEP_KIND", mepKind);

                            // Xoay căn theo trục (trong host coords)
                            RotateOnCylindricalFace(doc, inst, placePointH, rdirH, refDirH);

                            // --- Gán tham số family ---
                            SetParam(inst, "Inner Radius", bareR);
                            SetParam(inst, "Insulation", insulT > 0 ? insulT : 0.0);
                            SetParam(inst, "Host Thickness", hostThicknessFeet);

                            // 3 tham số UI (feet)
                            SetParam(inst, "Clearance", clearance);
                            SetParam(inst, "Extrusion", extrusion);
                            SetParam(inst, "Fill", fill);

                            // Schedule Level
                            TrySetScheduleLevel(doc, inst, refLvlId);

                            t.Commit();
                        }

                        // Trả lại view ban đầu
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
                // Dự phòng: luôn trả về view cũ
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
            public RevitLinkInstance LinkInst;  // null nếu ở host doc
            public bool IsLinked => LinkInst != null;
            public Transform ToHostT;           // Transform từ elem.Document -> host doc

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
                // Chỉ cho phép pick trên các RevitLinkInstance thuộc danh sách đã tick
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
                var ldoc = li?.GetLinkDocument();
                if (ldoc == null) return false;

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

        private static bool IsRoundMEP(Element e)
        {
            // Cho phép Pipe/Duct/Conduit; nếu Duct không tròn, sẽ bị loại sau khi không tìm thấy CylindricalFace
            return (e is Pipe) || (e is Duct) || (e is Conduit);
        }

        // ====================== PARAM & ALIGN HELPERS (giữ nguyên) ======================
        private static void SetParam(FamilyInstance inst, string name, double val)
        {
            Parameter p = inst.LookupParameter(name);
            if (p != null && !p.IsReadOnly) p.Set(val);
        }

        private static void RotateOnCylindricalFace(Document doc, FamilyInstance inst, XYZ pivot, XYZ radialNormal, XYZ desiredDir)
        {
            XYZ n = radialNormal.Normalize();
            XYZ cur = ProjectOnPlane(inst.HandOrientation, n);
            XYZ des = ProjectOnPlane(desiredDir, n);
            if (cur == null || des == null) return;

            cur = cur.Normalize();
            des = des.Normalize();

            double angle = SignedAngleOnPlane(cur, des, n);
            if (Math.Abs(angle) < 1e-6) return;

            Line axis = Line.CreateBound(pivot, pivot + n);
            try { ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle); }
            catch { /* family bị khoá xoay */ }
        }

        private static XYZ ProjectOnPlane(XYZ v, XYZ normal)
        {
            double dot = v.DotProduct(normal);
            XYZ p = v - dot * normal;
            return (p.GetLength() < 1e-9) ? null : p;
        }

        private static double SignedAngleOnPlane(XYZ a, XYZ b, XYZ normal)
        {
            double ang = a.AngleTo(b);
            double sign = Math.Sign(normal.DotProduct(a.CrossProduct(b)));
            return ang * sign;
        }

        // ============================== VIEW / LEVEL (giữ nguyên) ===============================
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
            // Thử lấy level trực tiếp trên element (kể cả khi mep ở link doc)
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
        private static void SetTextParam(Element e, string name, string val)
        {
            var p = e.LookupParameter(name);
            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.String) p.Set(val);
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

        // ============================== GEOMETRY (giữ nguyên) ===============================
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

        private static bool TryFindFaceAtPoint(Solid solid, XYZ p, out Face face, out UV uv)
        {
            face = null; uv = null;
            foreach (Face f in solid.Faces)
            {
                var proj = f.Project(p);
                if (proj == null) continue;
                if (proj.Distance < 1e-6) { face = f; uv = proj.UVPoint; return true; }
            }
            return false;
        }

        private static bool TryGetCylindricalFace(Element elem, out CylindricalFace cylFace)
        {
            cylFace = null;

            Options opts = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement ge = elem.get_Geometry(opts);
            if (ge == null) return false;

            foreach (GeometryObject go in ge)
            {
                if (go is Solid s && s.Volume > 1e-9)
                {
                    foreach (Face f in s.Faces)
                    {
                        var cf = f as CylindricalFace;
                        if (cf != null) { cylFace = cf; return true; }
                    }
                }
                else if (go is GeometryInstance gi)
                {
                    GeometryElement ige = gi.GetInstanceGeometry();
                    if (ige == null) continue;

                    foreach (GeometryObject igo in ige)
                    {
                        if (igo is Solid si && si.Volume > 1e-9)
                        {
                            foreach (Face f in si.Faces)
                            {
                                var cf = f as CylindricalFace;
                                if (cf != null) { cylFace = cf; return true; }
                            }
                        }
                    }
                }
            }

            return false; // không tìm thấy mặt trụ
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
            if (t.IsZeroLength()) t = XYZ.BasisY.CrossProduct(u);
            return t.Normalize();
        }

        private static double GetRoundOuterDiameterFeet(Element elem)
        {
            if (elem is Pipe p)
            {
                Parameter pi = p.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                if (pi != null && pi.StorageType == StorageType.Double) return pi.AsDouble();

                ElementType pt = p.Document.GetElement(p.GetTypeId()) as ElementType;
                if (pt != null)
                {
                    Parameter ptp = pt.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    if (ptp != null && ptp.StorageType == StorageType.Double) return ptp.AsDouble();
                }

                Parameter pn = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (pn != null && pn.StorageType == StorageType.Double) return pn.AsDouble();

                return 0.1;
            }

            if (elem is Duct d)
            {
                Parameter pd = d.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (pd != null && pd.StorageType == StorageType.Double) return pd.AsDouble();

                ElementType dt = d.Document.GetElement(d.GetTypeId()) as ElementType;
                if (dt != null)
                {
                    Parameter pdt = dt.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                    if (pdt != null && pdt.StorageType == StorageType.Double) return pdt.AsDouble();
                }

                return 0.1;
            }

            if (elem is Conduit cdt)
            {
                Parameter pc = cdt.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                if (pc != null && pc.StorageType == StorageType.Double) return pc.AsDouble();

                ElementType ct = cdt.Document.GetElement(cdt.GetTypeId()) as ElementType;
                if (ct != null)
                {
                    Parameter pct = ct.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                    if (pct != null && pct.StorageType == StorageType.Double) return pct.AsDouble();
                }

                return 0.1;
            }

            return 0.1;
        }

        private static double GetInsulationThicknessFeet(Element mep)
        {
            if (mep is Conduit) return 0.0;

            double best = 0.0;
            ICollection<ElementId> deps = mep.GetDependentElements(null);
            if (deps == null || deps.Count == 0) return 0.0;

            Document d = mep.Document;
            foreach (var id in deps)
            {
                Element e = d.GetElement(id);
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

        // ================================ Utils ================================
        private static bool IsZeroLength(XYZ v) => v == null || v.GetLength() < 1e-9;
    }
}
