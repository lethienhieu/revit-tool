using System;
using Autodesk.Revit.DB;

namespace LevelRehost.REVIT
{
    public static class DoorProcessor
    {
        public static bool RehostDoor(Document doc, FamilyInstance door, Level newLevel)
        {
            try
            {
                // 1. Lấy Level hiện tại
                // Cửa thường dùng FAMILY_LEVEL_PARAM
                Parameter levelParam = door.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                if (levelParam == null || levelParam.IsReadOnly) return false;

                ElementId oldLevelId = levelParam.AsElementId();
                Level oldLevel = doc.GetElement(oldLevelId) as Level;

                if (oldLevel == null || oldLevel.Id == newLevel.Id) return true;

                // 2. Lấy Sill Height hiện tại (Offset)
                // Tham số: INSTANCE_SILL_HEIGHT_PARAM
                Parameter sillHeightParam = door.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
                if (sillHeightParam == null) return false;

                double oldSillHeight = sillHeightParam.AsDouble();

                // 3. Tính toán cao độ tuyệt đối
                double oldLevelElev = oldLevel.ProjectElevation;
                double absElevation = oldLevelElev + oldSillHeight;

                // 4. Tính Sill Height mới
                double newLevelElev = newLevel.ProjectElevation;
                double newSillHeight = absElevation - newLevelElev;

                // 5. Apply
                // Lưu ý: Với cửa, đôi khi đổi Level sẽ làm Sill Height tự nhảy về 0 hoặc giá trị mặc định,
                // nên việc set lại Sill Height ngay sau đó là rất quan trọng.
                levelParam.Set(newLevel.Id);
                sillHeightParam.Set(newSillHeight);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}