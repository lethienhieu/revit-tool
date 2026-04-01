using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows;

namespace THBIM
{
    public enum StructureSyncRequestId
    {
        None = 0,
        Save = 1,
        Update = 2,
        Highlight = 3,
        ResetColor = 4,
        SelectElements = 5,
        Report = 6,
        SyncSpecific = 7
    }

    public class StructureSyncRequestHandler : IExternalEventHandler
    {
        public StructureSyncRequestId Request { get; set; } = StructureSyncRequestId.None;
        public StructureSync Logic { get; set; }
        public List<RelationshipItem> Relationships { get; set; }
        public List<RevitLinkInstance> Links { get; set; }
        public List<ElementId> ElementsToSelect { get; set; }

        // <--- 2. MỚI: Cần biến này để truyền cho Window báo cáo (để nó gọi lại Select)
        public ExternalEvent AppExEvent { get; set; }
        public List<ReportItem> ItemsToSync { get; set; }
        public Window ReportWindow { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (Logic == null) return;
                UIDocument uidoc = app.ActiveUIDocument;

                switch (Request)
                {
                    case StructureSyncRequestId.Save:
                        if (Relationships != null) Logic.SaveRelationships(Relationships);
                        break;

                    case StructureSyncRequestId.Update:
                        if (Relationships != null) Logic.SyncPositions(Relationships, Links);
                        break;

                    case StructureSyncRequestId.Highlight:
                        if (Relationships != null) Logic.HighlightUnsynced(Relationships, Links);
                        break;

                    case StructureSyncRequestId.ResetColor:
                        if (Relationships != null) Logic.ResetColor(Relationships);
                        break;

                    case StructureSyncRequestId.SelectElements:
                        if (ElementsToSelect != null && ElementsToSelect.Count > 0)
                        {
                            uidoc.Selection.SetElementIds(ElementsToSelect);
                        }
                        break;

                    case StructureSyncRequestId.Report:
                        if (Relationships != null)
                        {
                            // [CẬP NHẬT] Truyền thêm biến 'app' (UIApplication) vào cuối hàm
                            Logic.ShowReportProMax(Relationships, Links, AppExEvent, this, app);
                        }
                        break;

                    case StructureSyncRequestId.SyncSpecific: // <--- MỚI: Xử lý Sync 1 cặp
                        if (ItemsToSync != null && ItemsToSync.Count > 0)
                        {
                            // Gọi hàm Sync đặc biệt từ Logic
                            Logic.SyncSpecificItems(ItemsToSync, Links);

                            // Cập nhật lại giao diện bảng ngay lập tức
                            if (ReportWindow is Views.StructureReportWindow win)
                            {
                                win.Dispatcher.Invoke(() => win.RefreshDataGrid());
                            }
                        }
                        break;

                    default: break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "External Event Failed: " + ex.Message);
            }
            finally
            {
                Request = StructureSyncRequestId.None;
            }
        }

        public string GetName() => "THBIM Structure Sync Handler";
    }
}
