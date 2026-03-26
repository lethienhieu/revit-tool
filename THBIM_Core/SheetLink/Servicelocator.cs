using THBIM.Services;

namespace THBIM
{
    using THBIM.Services;

    public static class ServiceLocator
    {
        public static RevitDataService RevitData { get; set; }
        public static ExcelService Excel { get; set; }
        public static RevitViewService RevitView { get; set; }

        public static bool IsRevitMode => RevitData != null;

        public static void Reset()
        {
            RevitData = null;
            Excel = null;
            RevitView = null;
        }
    }
}