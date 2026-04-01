using System;
using Autodesk.Revit.DB;

namespace LevelRehost.REVIT
{
    public static class FloorProcessor
    {
        public static bool RehostFloor(Document doc, Floor floor, Level newLevel)
        {
            try
            {
                // 1. Lấy Level hiện tại
                Parameter levelParam = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (levelParam == null || levelParam.IsReadOnly) return false;

                ElementId oldLevelId = levelParam.AsElementId();
                Level oldLevel = doc.GetElement(oldLevelId) as Level;

                if (oldLevel == null || oldLevel.Id == newLevel.Id) return true;

                // 2. Lấy Offset hiện tại (Height Offset From Level)
                Parameter offsetParam = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                if (offsetParam == null) return false;

                double oldOffset = offsetParam.AsDouble();

                // 3. Tính toán cao độ tuyệt đối
                double oldLevelElev = oldLevel.ProjectElevation;
                double absElevation = oldLevelElev + oldOffset;

                // 4. Tính Offset mới
                double newLevelElev = newLevel.ProjectElevation;
                double newOffset = absElevation - newLevelElev;

                // 5. Apply
                levelParam.Set(newLevel.Id);
                offsetParam.Set(newOffset);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}