using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class LeaderAngleSettingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {
            LeaderAngleSettings.Load();
            var win = new LeaderAngleSettingWindow();
            win.ShowDialog();
            return Result.Succeeded;
        }
    }
}
