using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace colorslapsher.REVIT
{
    // Class chứa dữ liệu màu
    public class ColorItem
    {
        public string ValueName { get; set; } = "";
        public Color RevitColor { get; set; } = new Color(0, 0, 0);
        public byte R { get; set; }
        public byte G { get; set; }
        public byte B { get; set; }
    }

    public class ColorSplasher
    {
        private Document _doc;
        private Random _rand = new Random();

        public ColorSplasher(Document doc)
        {
            _doc = doc;
        }

        // 1. LẤY CATEGORIES
        public List<Category> GetCategoriesInActiveView()
        {
            var elements = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                           .WhereElementIsNotElementType()
                           .ToElements();

            return elements
                   .Select(e => e.Category)
                   .Where(c => c != null && c.HasMaterialQuantities)
                   .GroupBy(c => c.Id)
                   .Select(g => g.First())
                   .OrderBy(c => c.Name)
                   .ToList();
        }

        // 2. LẤY PARAMETERS
        public List<Parameter> GetParametersOfCategory(Category cat)
        {
            if (cat == null) return new List<Parameter>();

            Element firstElem = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                                .OfCategoryId(cat.Id)
                                .WhereElementIsNotElementType()
                                .FirstElement();

            if (firstElem == null) return new List<Parameter>();

            return firstElem.Parameters.Cast<Parameter>()
                   .OrderBy(p => p.Definition.Name)
                   .ToList();
        }

        // 3. PHÂN TÍCH GIÁ TRỊ VÀ TẠO MÀU
        public List<ColorItem> AnalyzeValuesAndGenerateColors(Category cat, string paramName)
        {
            var result = new List<ColorItem>();

            // HashSet để kiểm soát các giá trị Parameter đã xử lý (tránh lặp item)
            var processedValues = new HashSet<string>();

            // HashSet để kiểm soát các MÀU đã sử dụng (tránh trùng màu)
            // Lưu dưới dạng int: (R << 16) | (G << 8) | B
            var usedColorCodes = new HashSet<int>();

            // Lấy Elements
            var elements = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                           .OfCategoryId(cat.Id)
                           .WhereElementIsNotElementType()
                           .ToElements();

            foreach (var elem in elements)
            {
                // 1. Lấy giá trị Parameter
                Parameter p = elem.LookupParameter(paramName);
                if (p == null) continue;

                string val = p.AsValueString();
                if (string.IsNullOrEmpty(val) && p.StorageType == StorageType.String)
                    val = p.AsString();
                if (string.IsNullOrEmpty(val)) val = "<Empty/Null>";

                // Nếu giá trị này chưa có trong danh sách result
                if (!processedValues.Contains(val))
                {
                    processedValues.Add(val);

                    Color finalColor = null;

                    // 2. KIỂM TRA MÀU HIỆN CÓ (LOAD NGƯỢC)
                    OverrideGraphicSettings ogs = _doc.ActiveView.GetElementOverrides(elem.Id);

                    // Kiểm tra xem element có bị tô màu Foreground không
                    if (ogs.SurfaceForegroundPatternColor.IsValid)
                    {
                        // Lấy màu hiện tại của element
                        finalColor = ogs.SurfaceForegroundPatternColor;
                    }

                    // 3. NẾU CHƯA CÓ MÀU -> RANDOM MỚI (KHÔNG TRÙNG)
                    if (finalColor == null)
                    {
                        finalColor = GetUniqueRandomColor(usedColorCodes);
                    }
                    else
                    {
                        // Nếu lấy màu cũ, nhớ đánh dấu là màu này đã dùng để thằng sau không random trúng
                        int rgbCode = (finalColor.Red << 16) | (finalColor.Green << 8) | finalColor.Blue;
                        if (!usedColorCodes.Contains(rgbCode))
                        {
                            usedColorCodes.Add(rgbCode);
                        }
                    }

                    // 4. Thêm vào kết quả
                    result.Add(new ColorItem
                    {
                        ValueName = val,
                        RevitColor = finalColor,
                        R = finalColor.Red,
                        G = finalColor.Green,
                        B = finalColor.Blue
                    });
                }
            }
            return result.OrderBy(x => x.ValueName).ToList();
        }

        // Hàm phụ trợ: Random màu không trùng lặp
        private Color GetUniqueRandomColor(HashSet<int> usedCodes)
        {
            int safetyCounter = 0; // Tránh vòng lặp vô tận nếu hết màu
            while (safetyCounter < 1000)
            {
                byte r = (byte)_rand.Next(30, 225); // Tránh màu quá đen hoặc quá trắng
                byte g = (byte)_rand.Next(30, 225);
                byte b = (byte)_rand.Next(30, 225);

                int code = (r << 16) | (g << 8) | b;

                // Nếu chưa dùng màu này -> Chốt đơn
                if (!usedCodes.Contains(code))
                {
                    usedCodes.Add(code);
                    return new Color(r, g, b);
                }
                safetyCounter++;
            }
            // Nếu đen đủi quá random mãi vẫn trùng (hiếm) thì lấy đại
            return new Color((byte)_rand.Next(0, 255), (byte)_rand.Next(0, 255), (byte)_rand.Next(0, 255));
        }

        // 4. RESET MÀU
        public void ResetGraphics(Category cat)
        {
            using (Transaction t = new Transaction(_doc, "Reset Colors"))
            {
                t.Start();
                var elements = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                               .OfCategoryId(cat.Id)
                               .WhereElementIsNotElementType()
                               .ToElements();

                OverrideGraphicSettings clearSettings = new OverrideGraphicSettings();

                foreach (var e in elements)
                {
                    _doc.ActiveView.SetElementOverrides(e.Id, clearSettings);
                }
                t.Commit();
            }
        }

        // 5. APPLY MÀU (SET COLORS)
        public void ApplyColorSplash(Category cat, string paramName, List<ColorItem> colorMap)
        {
            using (Transaction t = new Transaction(_doc, "Apply Color Splash"))
            {
                t.Start();

                FillPatternElement solidFill = GetSolidFillPattern();

                var elements = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                               .OfCategoryId(cat.Id)
                               .WhereElementIsNotElementType()
                               .ToElements();

                foreach (var elem in elements)
                {
                    Parameter p = elem.LookupParameter(paramName);
                    string val = p?.AsValueString() ?? p?.AsString() ?? "<Empty/Null>";

                    var colorItem = colorMap.FirstOrDefault(x => x.ValueName == val);

                    if (colorItem != null && solidFill != null)
                    {
                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                        ogs.SetSurfaceForegroundPatternColor(colorItem.RevitColor);
                        ogs.SetCutForegroundPatternId(solidFill.Id);
                        ogs.SetCutForegroundPatternColor(colorItem.RevitColor);
                        ogs.SetProjectionLineColor(colorItem.RevitColor);

                        _doc.ActiveView.SetElementOverrides(elem.Id, ogs);
                    }
                }
                t.Commit();
            }
        }

        // 6. TẠO FILTER (Hàm bạn đang thiếu)
        // Thay thế hàm cũ bằng hàm này
        // Thay thế toàn bộ hàm cũ bằng hàm này
        // Thay thế hàm CreateFiltersForValues cũ bằng hàm này
        public void CreateFiltersForValues(Category cat, string paramName, List<ColorItem> colorItems, string prefix)
        {
            using (Transaction t = new Transaction(_doc, "Create View Filters"))
            {
                t.Start();
                View activeView = _doc.ActiveView;
                FillPatternElement solidFill = GetSolidFillPattern();

                // --- CẬP NHẬT: TÌM PARAMETER ID CHUẨN ---
                // Thay vì lấy từ firstElem.LookupParameter, ta dùng hàm Helper vừa viết
                ElementId paramId = GetValidFilterableParameterId(_doc, cat.Id, paramName);

                // Nếu không tìm thấy ID chuẩn, thử fallback về cách cũ (phòng hờ) hoặc dừng lại
                if (paramId == ElementId.InvalidElementId)
                {
                    // Fallback: Thử lấy theo cách cũ nếu hàm chuẩn không tìm ra (hiếm khi)
                    Element firstElem = new FilteredElementCollector(_doc, activeView.Id)
                                        .OfCategoryId(cat.Id)
                                        .WhereElementIsNotElementType()
                                        .FirstElement();
                    if (firstElem != null)
                    {
                        Parameter p = firstElem.LookupParameter(paramName);
                        if (p != null) paramId = p.Id;
                    }
                }

                // Nếu vẫn không có ID thì chịu thua -> Rollback
                if (paramId == ElementId.InvalidElementId)
                {
                    t.RollBack();
                    // Có thể throw exception hoặc return tuỳ bạn
                    return;
                }

                List<ElementId> catIds = new List<ElementId> { cat.Id };

                // Cần lấy lại collector để tìm giá trị mẫu (cho logic StorageType)
                var collector = new FilteredElementCollector(_doc, activeView.Id)
                                .OfCategoryId(cat.Id)
                                .WhereElementIsNotElementType()
                                .ToElements();

                foreach (var item in colorItems)
                {
                    string filterName = $"{prefix}_{item.ValueName}";

                    // --- ĐOẠN DƯỚI NÀY GIỮ NGUYÊN LOGIC CŨ ---
                    // (Chỉ lưu ý: paramId bây giờ đã là ID chuẩn xịn)

                    // Tìm element mẫu để xác định kiểu dữ liệu (Double/Int/String)
                    Element sampleElem = collector.FirstOrDefault(e =>
                    {
                        Parameter p = e.LookupParameter(paramName); // Lookup để lấy giá trị thì vẫn OK
                        if (p == null) return false;

                        string valStr = p.AsValueString();
                        if (string.IsNullOrEmpty(valStr) && p.StorageType == StorageType.String)
                            valStr = p.AsString();

                        return (valStr ?? "").Equals(item.ValueName);
                    });

                    // ... (Tiếp tục logic Switch case StorageType và Tạo Filter như cũ) ...

                    // NOTE: Copy lại đoạn Switch case StorageType cũ của bạn vào đây
                    FilterRule rule = null;
                    if (sampleElem != null)
                    {
                        Parameter targetParam = sampleElem.LookupParameter(paramName);
                        switch (targetParam.StorageType)
                        {
                            case StorageType.Integer:
                                rule = ParameterFilterRuleFactory.CreateEqualsRule(paramId, targetParam.AsInteger());
                                break;
                            case StorageType.Double:
                                rule = ParameterFilterRuleFactory.CreateEqualsRule(paramId, targetParam.AsDouble(), 0.001);
                                break;
                            case StorageType.ElementId:
                                rule = ParameterFilterRuleFactory.CreateEqualsRule(paramId, targetParam.AsElementId());
                                break;
                            default:
                                rule = ParameterFilterRuleFactory.CreateEqualsRule(paramId, item.ValueName);
                                break;
                        }
                    }
                    else
                    {
                        // Nếu không tìm thấy mẫu, mặc định String
                        rule = ParameterFilterRuleFactory.CreateEqualsRule(paramId, item.ValueName);
                    }

                    if (rule == null) continue;

                    // Tạo Filter (đoạn này giữ nguyên như code trước)
                    ParameterFilterElement filterElement = new FilteredElementCollector(_doc)
                        .OfClass(typeof(ParameterFilterElement))
                        .Cast<ParameterFilterElement>()
                        .FirstOrDefault(f => f.Name.Equals(filterName));

                    if (filterElement == null)
                    {
                        ElementParameterFilter paramFilter = new ElementParameterFilter(rule);
                        LogicalOrFilter orFilter = new LogicalOrFilter(new List<ElementFilter> { paramFilter });
                        filterElement = ParameterFilterElement.Create(_doc, filterName, catIds, orFilter);
                    }

                    // Add view & set màu (giữ nguyên)
                    if (!activeView.GetFilters().Contains(filterElement.Id)) activeView.AddFilter(filterElement.Id);

                    if (solidFill != null)
                    {
                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                        ogs.SetSurfaceForegroundPatternColor(item.RevitColor);
                        ogs.SetCutForegroundPatternId(solidFill.Id);
                        ogs.SetCutForegroundPatternColor(item.RevitColor);
                        ogs.SetProjectionLineColor(item.RevitColor);
                        activeView.SetFilterOverrides(filterElement.Id, ogs);
                        activeView.SetFilterVisibility(filterElement.Id, true);
                    }
                }
                t.Commit();
            }
        }
        // Hàm tìm ID chuẩn cho Filter (Thay thế cách lấy ID thủ công dễ gây lỗi)
        private ElementId GetValidFilterableParameterId(Document doc, ElementId catId, string paramName)
        {
            // 1. Lấy danh sách tất cả Parameter ID được phép dùng cho Filter của Category này
            ICollection<ElementId> filterableParamIds = ParameterFilterUtilities.GetFilterableParametersInCommon(doc, new List<ElementId> { catId });

            foreach (ElementId id in filterableParamIds)
            {
                string name = string.Empty;

                // 2. Kiểm tra tên Parameter để so sánh
                if (id.GetValue() < 0) // Là BuiltInParameter (ID âm)
                {
                    try
                    {
                        // Lấy tên hiển thị (Localized Name)
                        BuiltInParameter bip = (BuiltInParameter)id.GetValue();
                        name = LabelUtils.GetLabelFor(bip);
                    }
                    catch { continue; }
                }
                else // Là Shared/Project Parameter (ID dương)
                {
                    try
                    {
                        Element paramElem = doc.GetElement(id);
                        if (paramElem != null) name = paramElem.Name;
                    }
                    catch { continue; }
                }

                // 3. So sánh tên (Không phân biệt hoa thường)
                if (name.Equals(paramName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return id;
                }
            }
            return ElementId.InvalidElementId;
        }

        // Helper tìm Solid Fill an toàn
        private FillPatternElement GetSolidFillPattern()
        {
            return new FilteredElementCollector(_doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .FirstOrDefault(fp => fp.Name.Contains("Solid") || fp.Name.Contains("solid"));
        }


        public void CreateLegendView(Category cat, string paramName, List<ColorItem> colorItems)
        {
            using (Transaction t = new Transaction(_doc, "Create Legend"))
            {
                t.Start();

                // 1. Đặt tên View
                string viewName = $"Legend_{cat.Name}_{paramName}";

                // 2. Xóa view cũ nếu đã tồn tại để tạo mới
                View existingView = new FilteredElementCollector(_doc)
                                    .OfClass(typeof(View))
                                    .Cast<View>()
                                    .FirstOrDefault(v => v.Name.Equals(viewName));

                if (existingView != null)
                {
                    _doc.Delete(existingView.Id);
                }

                // 3. Tìm các Type cần thiết (Drafting View, Filled Region, Text)
                ViewFamilyType viewFamilyType = new FilteredElementCollector(_doc)
                                                .OfClass(typeof(ViewFamilyType))
                                                .Cast<ViewFamilyType>()
                                                .FirstOrDefault(x => x.ViewFamily == ViewFamily.Drafting);

                FilledRegionType regionType = new FilteredElementCollector(_doc)
                                              .OfClass(typeof(FilledRegionType))
                                              .Cast<FilledRegionType>()
                                              .FirstOrDefault(); // Lấy loại Region mặc định

                TextNoteType textType = new FilteredElementCollector(_doc)
                                        .OfClass(typeof(TextNoteType))
                                        .Cast<TextNoteType>()
                                        .FirstOrDefault(); // Lấy loại Text mặc định

                if (viewFamilyType != null && regionType != null && textType != null)
                {
                    // 4. Tạo View Drafting
                    ViewDrafting legendView = ViewDrafting.Create(_doc, viewFamilyType.Id);
                    legendView.Name = viewName;
                    legendView.Scale = 50;

                    // Lấy mẫu tô Solid
                    FillPatternElement solidFill = GetSolidFillPattern();

                    // 5. Vẽ từng dòng màu
                    double xPos = 0;
                    double yPos = 0;
                    double rowHeight = 0.5; // Khoảng cách dòng (feet)
                    double boxSize = 0.3;   // Kích thước ô màu (feet)

                    foreach (var item in colorItems)
                    {
                        // A. Vẽ hình vuông
                        List<CurveLoop> profileLoops = new List<CurveLoop>();
                        CurveLoop loop = new CurveLoop();
                        XYZ p1 = new XYZ(xPos, yPos, 0);
                        XYZ p2 = new XYZ(xPos + boxSize, yPos, 0);
                        XYZ p3 = new XYZ(xPos + boxSize, yPos - boxSize, 0);
                        XYZ p4 = new XYZ(xPos, yPos - boxSize, 0);
                        loop.Append(Line.CreateBound(p1, p2));
                        loop.Append(Line.CreateBound(p2, p3));
                        loop.Append(Line.CreateBound(p3, p4));
                        loop.Append(Line.CreateBound(p4, p1));
                        profileLoops.Add(loop);

                        FilledRegion region = FilledRegion.Create(_doc, legendView.Id, regionType.Id, profileLoops);

                        // B. Override màu cho hình vuông
                        if (solidFill != null)
                        {
                            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                            ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                            ogs.SetSurfaceForegroundPatternColor(item.RevitColor);
                            legendView.SetElementOverrides(region.Id, ogs);
                        }

                        // C. Viết Text chú thích
                        XYZ textPos = new XYZ(xPos + boxSize + 0.2, yPos - (boxSize / 2), 0);
                        TextNote.Create(_doc, legendView.Id, textPos, item.ValueName, textType.Id);

                        // D. Xuống dòng
                        yPos -= rowHeight;
                    }
                }

                t.Commit();
            }
        }
    }
}