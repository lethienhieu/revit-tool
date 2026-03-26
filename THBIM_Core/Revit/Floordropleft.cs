using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace THBIM
{
    public class Floordropleft
    {
        public Result Run(UIDocument uidoc, FamilySymbol symbol, ref string message)
        {
            Document doc = uidoc.Document;

            try
            {
                while (true)
                {
                    // Pick đối tượng
                    Reference pickedRef = uidoc.Selection.PickObject(
                        ObjectType.Edge,
                        new FloorEdgeSelectionFilter(),
                        "Pick the edge between two floors (ESC to exit)");

                    XYZ pickedPoint = pickedRef.GlobalPoint;

                    // Lấy giá trị chuỗi (Số hoặc "Var.")
                    string offsetText = GetAdjacentFloorText(doc, pickedPoint);

                    // Nếu có giá trị hợp lệ thì thực thi
                    if (!string.IsNullOrEmpty(offsetText))
                    {
                        using (Transaction t = new Transaction(doc, "Place & Align & Orient (Left)"))
                        {
                            t.Start();

                            // 1. Đặt Family
                            FamilyInstance tag = doc.Create.NewFamilyInstance(pickedPoint, symbol, doc.ActiveView);
                            doc.Regenerate();

                            // 2. Gán giá trị Text
                            Parameter param = GetFirstEditableTextParam(tag);
                            if (param != null) param.Set(offsetText);

                            // 3. Align song song
                            AlignInstanceToVector(doc, tag, pickedRef, pickedPoint);
                            doc.Regenerate();

                            // 4. Xử lý hướng (Logic ĐẢO NGƯỢC so với bản Right)
                            ProcessOrientationByAngle(doc, tag, pickedRef, pickedPoint);

                            t.Commit();
                        }
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ==========================================================
        // CÁC HÀM XỬ LÝ LOGIC (ĐÃ ĐẢO NGƯỢC ĐIỀU KIỆN MIRROR)
        // ==========================================================

        private void ProcessOrientationByAngle(Document doc, FamilyInstance instance, Reference edgeRef, XYZ pickPoint)
        {
            Element elem = doc.GetElement(edgeRef);
            GeometryObject geomObj = elem.GetGeometryObjectFromReference(edgeRef);
            Curve curve = (geomObj is Edge e) ? e.AsCurve() : (geomObj as Curve);
            if (curve == null) return;

            IntersectionResult res = curve.Project(pickPoint);
            Transform derivatives = curve.ComputeDerivatives(res.Parameter, false);
            XYZ lineDir = derivatives.BasisX.Normalize();

            if (lineDir.X < -0.001 || (Math.Abs(lineDir.X) < 0.001 && lineDir.Y < -0.001))
                lineDir = -lineDir;

            double angleWithY = XYZ.BasisY.AngleTo(lineDir);

            bool isGreaterThan90 = angleWithY > (Math.PI / 2.0 + 0.001);

            if (isGreaterThan90)
            {
                // Logic Góc > 90 độ
                Line axis = Line.CreateBound(pickPoint, pickPoint + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, instance.Id, axis, Math.PI);
                doc.Regenerate();
                CheckAndMirrorStandard(doc, instance, pickPoint, lineDir);
            }
            else
            {
                bool isHorizontal = Math.Abs(lineDir.Y) < 0.001;

                if (isHorizontal)
                {
                    // Logic Trục Hoành
                    Line axis = Line.CreateBound(pickPoint, pickPoint + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, instance.Id, axis, Math.PI);
                    doc.Regenerate();

                    XYZ pointAbove = pickPoint + XYZ.BasisY * 0.5;
                    XYZ pointBelow = pickPoint - XYZ.BasisY * 0.5;
                    double elevAbove = GetFloorElevationAtPoint(doc, pointAbove);
                    double elevBelow = GetFloorElevationAtPoint(doc, pointBelow);

                    // --- ĐẢO NGƯỢC LOGIC: Sàn Trên THẤP HƠN Sàn Dưới -> Mirror ---
                    // (Bản Right là: elevAbove > elevBelow)
                    if (elevAbove < elevBelow && elevAbove != -9999 && elevBelow != -9999)
                    {
                        ExecuteMirror(doc, instance, pickPoint, lineDir);
                    }
                }
                else
                {
                    CheckAndMirrorStandard(doc, instance, pickPoint, lineDir);
                }
            }
        }

        private void CheckAndMirrorStandard(Document doc, FamilyInstance instance, XYZ pickPoint, XYZ lineDir)
        {
            XYZ familyXDir = instance.GetTransform().BasisX.Normalize();
            XYZ pointLeft = pickPoint - familyXDir * 0.5;
            XYZ pointRight = pickPoint + familyXDir * 0.5;

            double elevLeft = GetFloorElevationAtPoint(doc, pointLeft);
            double elevRight = GetFloorElevationAtPoint(doc, pointRight);

            // --- ĐẢO NGƯỢC LOGIC: Trái CAO HƠN Phải -> Mirror ---
            // (Bản Right là: elevLeft < elevRight)
            // Giả định Family này mặc định là Trái Thấp - Phải Cao, nên nếu Trái Cao hơn thì phải lật.
            if (elevLeft > elevRight && elevLeft != -9999 && elevRight != -9999)
            {
                ExecuteMirror(doc, instance, pickPoint, lineDir);
            }
        }

        // --- CÁC HÀM HỖ TRỢ BÊN DƯỚI GIỮ NGUYÊN (VÌ LÀ HÀM TÍNH TOÁN HÌNH HỌC) ---

        private void ExecuteMirror(Document doc, FamilyInstance instance, XYZ pickPoint, XYZ lineDir)
        {
            Plane mirrorPlane = Plane.CreateByThreePoints(pickPoint, pickPoint + lineDir, pickPoint + XYZ.BasisZ);
            ElementTransformUtils.MirrorElement(doc, instance.Id, mirrorPlane);
            doc.Delete(instance.Id);
        }

        private void AlignInstanceToVector(Document doc, FamilyInstance instance, Reference edgeRef, XYZ pickPoint)
        {
            try
            {
                Element elem = doc.GetElement(edgeRef);
                GeometryObject geomObj = elem.GetGeometryObjectFromReference(edgeRef);
                Curve curve = (geomObj is Edge e) ? e.AsCurve() : (geomObj as Curve);
                if (curve == null) return;
                IntersectionResult res = curve.Project(pickPoint);
                Transform derivatives = curve.ComputeDerivatives(res.Parameter, false);
                XYZ targetVector = derivatives.BasisX.Normalize();
                if (targetVector.X < -0.001 || (Math.Abs(targetVector.X) < 0.001 && targetVector.Y < -0.001))
                    targetVector = -targetVector;
                XYZ familyYDir = instance.GetTransform().BasisY.Normalize();
                double angle = familyYDir.AngleTo(targetVector);
                if (familyYDir.CrossProduct(targetVector).Z < 0) angle = -angle;
                if (Math.Abs(angle) > 0.001)
                {
                    Line axis = Line.CreateBound(pickPoint, pickPoint + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle);
                }
            }
            catch { }
        }

        private double GetFloorElevationAtPoint(Document doc, XYZ point)
        {
            double r = 0.3;
            XYZ bot = new XYZ(point.X, point.Y, point.Z - 10);
            List<Curve> profile = new List<Curve> {
                Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, bot), r, 0, Math.PI),
                Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, bot), r, Math.PI, 2 * Math.PI)
            };
            Solid sol = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { CurveLoop.Create(profile) }, XYZ.BasisZ, 20);
            var floors = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Floors).WherePasses(new ElementIntersectsSolidFilter(sol)).ToElements();
            return floors.Count == 0 ? -9999 : floors.Max(f => f.get_BoundingBox(null).Max.Z);
        }

        private string GetAdjacentFloorText(Document doc, XYZ pickPoint)
        {
            double radius = 1.0;
            XYZ centerBottom = new XYZ(pickPoint.X, pickPoint.Y, pickPoint.Z - 25.0);
            List<Curve> profile = new List<Curve> {
                Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerBottom), radius, 0, Math.PI),
                Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerBottom), radius, Math.PI, 2 * Math.PI)
            };
            Solid searchSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { CurveLoop.Create(profile) }, XYZ.BasisZ, 50.0);

            var floors = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WherePasses(new ElementIntersectsSolidFilter(searchSolid))
                .ToElements();

            if (floors.Count < 2) return null;

            var topFloors = floors.OrderByDescending(x => x.get_BoundingBox(null).Max.Z).Take(2).ToList();
            if (IsFloorSloped(topFloors[0]) || IsFloorSloped(topFloors[1]))
            {
                return "Var.";
            }

            double h1 = topFloors[0].get_BoundingBox(null).Max.Z;
            double h2 = topFloors[1].get_BoundingBox(null).Max.Z;
            double diff = Math.Round(UnitUtils.ConvertFromInternalUnits(Math.Abs(h1 - h2), UnitTypeId.Millimeters), 0);

            if (diff == 0) return null;
            return diff.ToString();
        }

        private bool IsFloorSloped(Element elem)
        {
            Options opt = new Options() { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geo = elem.get_Geometry(opt);
            if (geo == null) return false;

            foreach (GeometryObject obj in geo)
            {
                Solid solid = obj as Solid;
                if (solid != null && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                        if (normal.Z > 0)
                        {
                            if (Math.Abs(normal.Z) < 0.999) return true;
                        }
                    }
                }
            }
            return false;
        }

        private Parameter GetFirstEditableTextParam(Element e)
        {
            foreach (Parameter p in e.Parameters)
                if (!p.IsReadOnly && p.StorageType == StorageType.String && p.Definition.GetDataType() == SpecTypeId.String.Text) return p;
            return null;
        }

        public class FloorEdgeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem.Category.Id.GetValue() == (int)BuiltInCategory.OST_Floors || elem.Category.Id.GetValue() == (int)BuiltInCategory.OST_Lines || elem is DetailLine || elem is ModelLine;
            public bool AllowReference(Reference r, XYZ p) => true;
        }
    }
}