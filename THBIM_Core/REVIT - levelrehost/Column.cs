using System;
using Autodesk.Revit.DB;

namespace LevelRehost.REVIT
{
    // LỖI DO DÒNG NÀY: Bạn đang để là "public static class WallProcessor"
    // HÃY SỬA THÀNH:
    public static class ColumnProcessor
    {
        // VÀ SỬA TÊN HÀM NÀY TỪ RehostWall -> RehostColumn
        public static bool RehostColumn(Document doc, FamilyInstance column, Level newLevel)
        {
            try
            {
                // ... (Phần logic giữ nguyên như code tôi gửi ở bước trước)

                // 1. Base Level
                Parameter baseLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                if (baseLevelParam == null || baseLevelParam.IsReadOnly) return false;

                ElementId oldLevelId = baseLevelParam.AsElementId();
                Level oldLevel = doc.GetElement(oldLevelId) as Level;

                if (oldLevel == null || oldLevel.Id == newLevel.Id) return true;

                // 2. Base Offset
                Parameter baseOffsetParam = null;
                if (column.IsSlantedColumn)
                    baseOffsetParam = column.get_Parameter(BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM);
                else
                    baseOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);

                if (baseOffsetParam == null) return false;

                // 3. Tính toán
                double oldOffset = baseOffsetParam.AsDouble();
                double oldLevelElev = oldLevel.ProjectElevation;
                double absElevation = oldLevelElev + oldOffset;

                double newLevelElev = newLevel.ProjectElevation;
                double newOffset = absElevation - newLevelElev;

                // 4. Apply
                baseLevelParam.Set(newLevel.Id);
                baseOffsetParam.Set(newOffset);

                // 5. Xử lý Top (Nếu cần)
                Parameter topLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                if (topLevelParam != null && topLevelParam.AsElementId() == oldLevelId)
                {
                    Parameter topOffsetParam = column.IsSlantedColumn
                       ? column.get_Parameter(BuiltInParameter.SCHEDULE_TOP_LEVEL_OFFSET_PARAM)
                       : column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);

                    if (topOffsetParam != null)
                    {
                        double oldTopOffset = topOffsetParam.AsDouble();
                        double absTopElev = oldLevelElev + oldTopOffset;
                        double newTopOffset = absTopElev - newLevelElev;

                        topLevelParam.Set(newLevel.Id);
                        topOffsetParam.Set(newTopOffset);
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}