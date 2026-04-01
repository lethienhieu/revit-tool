using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace THBIM.Helpers
{
    public class QTOSelectionFilter : ISelectionFilter
    {
        private string _mode;

        public QTOSelectionFilter(string mode)
        {
            _mode = mode;
        }

        public bool AllowElement(Element elem)
        {
            // Nếu mode là "All", cho phép cả Pipe và Duct (và các phụ kiện)
            if (_mode == "All")
            {
                return IsPipeCategory(elem) || IsDuctCategory(elem);
            }

            if (_mode == "Piping") return IsPipeCategory(elem);
            if (_mode == "Ducting") return IsDuctCategory(elem);

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }

        // Hàm phụ trợ kiểm tra Piping
        private bool IsPipeCategory(Element elem)
        {
            if (elem.Category == null) return false;
            long catId = elem.Category.Id.GetValue(); // Hoặc .Value với Revit 2024+
            return catId == (long)BuiltInCategory.OST_PipeCurves ||
                   catId == (long)BuiltInCategory.OST_PipeFitting ||
                   catId == (long)BuiltInCategory.OST_PipeAccessory ||
                   catId == (long)BuiltInCategory.OST_FlexPipeCurves ||
                   catId == (long)BuiltInCategory.OST_PipeInsulations;
        }

        // Hàm phụ trợ kiểm tra Ducting
        private bool IsDuctCategory(Element elem)
        {
            if (elem.Category == null) return false;
            long catId = elem.Category.Id.GetValue();
            return catId == (long)BuiltInCategory.OST_DuctCurves ||
                   catId == (long)BuiltInCategory.OST_DuctFitting ||
                   catId == (long)BuiltInCategory.OST_DuctAccessory ||
                   catId == (long)BuiltInCategory.OST_FlexDuctCurves ||
                   catId == (long)BuiltInCategory.OST_DuctInsulations; // Revit 2014+ là DuctInsulations
        }
    }
}