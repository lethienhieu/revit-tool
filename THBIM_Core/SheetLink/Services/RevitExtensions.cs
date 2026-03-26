using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM.Services
{
    public static class RevitDocumentCache
    {
        public static Document Current { get; set; }
        public static UIDocument CurrentUi { get; set; }
    }

    public static class RevitExtensions
    {
        public static Document GetDocument(this RevitDataService svc)
            => RevitDocumentCache.Current;
    }
}
