using System.Collections.Generic;
using System.Linq;
using System.Text; 
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    public static class CombineParam
    {
        public static void Execute(Document doc, List<ElementId> categoryIds, bool isAllCategories, List<string> sourceParamNames, string targetParamName, string separator)
        {
            // 1. THU THẬP DỮ LIỆU 
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            if (isAllCategories)
            {
                collector.WhereElementIsNotElementType();
            }
            else
            {
                if (categoryIds == null || categoryIds.Count == 0)
                {
                    TaskDialog.Show("THBIM", "Please select at least one Category.");
                    return;
                }
                ElementMulticategoryFilter multiCatFilter = new ElementMulticategoryFilter(categoryIds);
                collector.WherePasses(multiCatFilter).WhereElementIsNotElementType();
            }

            ICollection<ElementId> elementIds = collector.ToElementIds();

            if (elementIds.Count == 0)
            {
                TaskDialog.Show("THBIM", "No elements found.");
                return;
            }

            // 2. CHẠY TRANSACTION & THỐNG KÊ
            using (Transaction t = new Transaction(doc, "Combine Project Params"))
            {
                t.Start();

                // Dictionary để thống kê: Key = Tên Category, Value = [Số Instance, Số Type]
                Dictionary<string, int[]> stats = new Dictionary<string, int[]>();

                HashSet<ElementId> processedTypeIds = new HashSet<ElementId>();
                int totalUpdates = 0;

                foreach (ElementId id in elementIds)
                {
                    Element ele = doc.GetElement(id);
                    if (ele == null) continue;

                    // --- CHECK LOGIC ---
                    Parameter targetParam = GetParamOnInstanceOrType(ele, targetParamName);

                    if (targetParam == null || targetParam.IsReadOnly || targetParam.StorageType != StorageType.String)
                        continue;

                    // --- TYPE OPTIMIZATION ---
                    bool isTypeParam = false;
                    ElementId paramOwnerId = targetParam.Element.Id;
                    if (paramOwnerId != ele.Id)
                    {
                        isTypeParam = true;
                        if (processedTypeIds.Contains(paramOwnerId)) continue;
                    }

                    // --- SET VALUE & COUNT ---
                    string combinedValue = GenerateCombinedString(ele, sourceParamNames, separator);

                    if (targetParam.AsString() != combinedValue)
                    {
                        targetParam.Set(combinedValue);
                        totalUpdates++;

                        // Lấy tên Category để thống kê
                        string catName = ele.Category != null ? ele.Category.Name : "Other";

                        // Khởi tạo nếu chưa có trong danh sách
                        if (!stats.ContainsKey(catName))
                            stats[catName] = new int[] { 0, 0 }; // [0]: Instance, [1]: Type

                        // Cộng dồn
                        if (isTypeParam)
                        {
                            processedTypeIds.Add(paramOwnerId);
                            stats[catName][1]++;
                        }
                        else
                        {
                            stats[catName][0]++;
                        }
                    }
                }

                t.Commit();

                // 3. XUẤT BÁO CÁO CHI TIẾT
                ShowDetailedReport(stats, totalUpdates, elementIds.Count);
            }
        }

        // Hàm hiển thị thông báo đẹp mắt
        private static void ShowDetailedReport(Dictionary<string, int[]> stats, int totalUpdates, int totalScanned)
        {
            StringBuilder sb = new StringBuilder();

            if (totalUpdates == 0)
            {
                sb.AppendLine("Scan completed but NO parameters needed update.");
                sb.AppendLine($"Checked {totalScanned} elements.");
            }
            else
            {
                sb.AppendLine("Done! Update Report:");
                sb.AppendLine("------------------------------------------------");

                // Sắp xếp theo tên Category cho dễ nhìn
                foreach (var item in stats.OrderBy(x => x.Key))
                {
                    string catName = item.Key;
                    int instCount = item.Value[0];
                    int typeCount = item.Value[1];

                    // Ví dụ: - Doors: 50 Inst, 2 Types
                    string detail = $"- {catName}: ";
                    if (instCount > 0) detail += $"{instCount} Inst";
                    if (instCount > 0 && typeCount > 0) detail += ", ";
                    if (typeCount > 0) detail += $"{typeCount} Types";

                    sb.AppendLine(detail);
                }

                sb.AppendLine("------------------------------------------------");
                sb.AppendLine($"Total operations: {totalUpdates}");
            }

            TaskDialog.Show("THBIM Result", sb.ToString());
        }

        // --- CÁC HÀM HELPER 
        public static string GenerateCombinedString(Element ele, List<string> paramNames, string separator)
        {
            List<string> values = new List<string>();
            foreach (string name in paramNames)
            {
                if (string.IsNullOrEmpty(name)) continue;
                Parameter param = GetParamOnInstanceOrType(ele, name);
                values.Add(GetParameterValueAsString(param));
            }
            return string.Join(separator, values);
        }

        private static Parameter GetParamOnInstanceOrType(Element ele, string paramName)
        {
            Parameter p = ele.LookupParameter(paramName);
            if (p != null) return p;
            ElementId typeId = ele.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                Element typeEle = ele.Document.GetElement(typeId);
                return typeEle?.LookupParameter(paramName);
            }
            return null;
        }

        private static string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue) return "";
            if (param.StorageType == StorageType.String) return param.AsString() ?? "";
            return param.AsValueString() ?? "";
        }
    }
}