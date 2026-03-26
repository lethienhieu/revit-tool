using System;
using System.IO;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Services;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SheetLinkCommand : IExternalCommand
    {
        private static SheetLinkWindow _window;

        public Result Execute(ExternalCommandData commandData,
                              ref string message, ElementSet elements)
        {
            try
            {
                if (commandData?.Application == null)
                {
                    message = "Revit command context is not available.";
                    return Result.Failed;
                }

                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                if (uiDoc == null)
                {
                    message = "Please open a Revit model before running SheetLink.";
                    return Result.Failed;
                }

                var doc = uiDoc.Document;
                RevitDocumentCache.Current = doc;
                RevitDocumentCache.CurrentUi = uiDoc;

                var revitSvc = new RevitDataService(doc);
                var excelSvc = new ExcelService(revitSvc);
                var viewSvc = new RevitViewService(uiDoc);

                ServiceLocator.RevitData = revitSvc;
                ServiceLocator.Excel = excelSvc;
                ServiceLocator.RevitView = viewSvc;

                RevitEventHandler.Initialize();

                if (_window != null && _window.IsVisible)
                {
                    _window.Activate();
                    return Result.Succeeded;
                }

                var win = new SheetLinkWindow();
                var helper = new WindowInteropHelper(win)
                {
                    Owner = uiApp.MainWindowHandle
                };
                win.Closed += (_, _) =>
                {
                    _window = null;
                    ServiceLocator.Reset();
                    RevitDocumentCache.Current = null;
                    RevitDocumentCache.CurrentUi = null;
                };

                _window = win;
                win.Show();
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                var logPath = WriteErrorLog(ex);
                var root = ex;
                while (root.InnerException != null)
                    root = root.InnerException;

                message = $"{root.GetType().Name}: {root.Message}\nLog: {logPath}";
                return Result.Failed;
            }
        }

        private static string WriteErrorLog(Exception ex)
        {
            try
            {
                var folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "THBIM",
                    "SheetLink");
                Directory.CreateDirectory(folder);

                var path = Path.Combine(folder, "last-error.log");
                var text = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] SheetLink error{Environment.NewLine}{ex}";
                File.WriteAllText(path, text);
                return path;
            }
            catch
            {
                return "Unable to write error log.";
            }
        }
    }
}
