using System;
using Autodesk.Revit.DB;

namespace LevelRehost.REVIT
{
    public static class WallProcessor
    {
        /// <summary>
        /// Chuyển Wall sang Level mới và tính toán lại Offset để giữ nguyên vị trí 3D.
        /// </summary>
        /// <param name="doc">Revit Document</param>
        /// <param name="wall">Đối tượng tường cần xử lý</param>
        /// <param name="newLevel">Level đích</param>
        /// <returns>True nếu thành công, False nếu lỗi</returns>
        public static bool RehostWall(Document doc, Wall wall, Level newLevel)
        {
            try
            {
                // 1. Lấy thông tin Base Level hiện tại
                Parameter baseLevelParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
                if (baseLevelParam == null || baseLevelParam.IsReadOnly) return false;

                ElementId oldLevelId = baseLevelParam.AsElementId();
                if (oldLevelId == ElementId.InvalidElementId) return false;

                Level oldLevel = doc.GetElement(oldLevelId) as Level;
                if (oldLevel == null) return false;

                // Nếu Level mới trùng Level cũ thì bỏ qua
                if (oldLevel.Id == newLevel.Id) return true;

                // 2. Lấy Base Offset hiện tại
                Parameter baseOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
                double oldOffset = baseOffsetParam.AsDouble(); // Đơn vị: Feet

                // 3. Tính toán Cao độ tuyệt đối (Z_Absolute)
                // ProjectElevation: Cao độ của Level so với gốc 0.0 của dự án (Internal Origin)
                double oldLevelElevation = oldLevel.ProjectElevation;
                double currentAbsoluteBaseElevation = oldLevelElevation + oldOffset;

                // 4. Tính toán Offset mới (Compensation)
                double newLevelElevation = newLevel.ProjectElevation;
                double newOffset = currentAbsoluteBaseElevation - newLevelElevation;

                // 5. Áp dụng thay đổi
                // Cần đổi Level trước, sau đó đổi Offset ngay lập tức
                // Lưu ý: Trong một số trường hợp, Revit có thể cảnh báo "Joined Elements",
                // nhưng API thường tự xử lý hoặc ta chấp nhận warning.

                baseLevelParam.Set(newLevel.Id);
                baseOffsetParam.Set(newOffset);

                // --- XỬ LÝ PHỤ (OPTIONAL): TOP CONSTRAINT ---
                // Kiểm tra nếu Top Constraint đang trỏ vào chính Level cũ
                // Thì nên chuyển Top Constraint sang Level mới luôn để tường không bị "lộn ngược" hoặc lỗi 0 height
                Parameter topLevelParam = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                if (topLevelParam != null && topLevelParam.HasValue)
                {
                    ElementId topLevelId = topLevelParam.AsElementId();
                    if (topLevelId == oldLevelId) // Nếu đỉnh và đáy cùng 1 level cũ
                    {
                        Parameter topOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
                        double oldTopOffset = topOffsetParam.AsDouble();

                        // Tính lại Top Offset tương tự Base
                        double currentAbsoluteTop = oldLevelElevation + oldTopOffset;
                        double newTopOffset = currentAbsoluteTop - newLevelElevation;

                        topLevelParam.Set(newLevel.Id);
                        topOffsetParam.Set(newTopOffset);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                // Ghi log lỗi nếu cần thiết (ví dụ System.Diagnostics.Debug.WriteLine)
                return false;
            }
        }
    }
}