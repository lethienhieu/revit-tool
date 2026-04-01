using System;
using System.Collections.Generic;
using System.Linq; // [QUAN TRỌNG] Cần để dùng GroupBy
using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace THBIM
{
    public enum ReportSeverity
    {
        Synced,     // Xanh
        Warning,    // Vàng
        Critical,   // Đỏ
        Info        // Xám
    }

    public class ReportItem : INotifyPropertyChanged
    {
        private ReportSeverity _severity;
        private string _statusDisplay;
        private string _diagnosis;
        private string _techData;

        public int Id { get; set; }
        public string SetName { get; set; }

        public ReportSeverity Severity
        {
            get { return _severity; }
            set { _severity = value; OnPropertyChanged(); OnPropertyChanged("SeverityColor"); }
        }
        public string StatusDisplay
        {
            get { return _statusDisplay; }
            set { _statusDisplay = value; OnPropertyChanged(); }
        }
        public string Diagnosis
        {
            get { return _diagnosis; }
            set { _diagnosis = value; OnPropertyChanged(); }
        }
        public string TechData
        {
            get { return _techData; }
            set { _techData = value; OnPropertyChanged(); }
        }

        public RelationshipItem SourceItem { get; set; }
        public int IndexInSet { get; set; }
        public ElementId ParentId { get; set; }
        public ElementId ChildId { get; set; }

        public string SeverityColor
        {
            get
            {
                switch (Severity)
                {
                    case ReportSeverity.Synced: return "#2ECC71";
                    case ReportSeverity.Warning: return "#F1C40F";
                    case ReportSeverity.Critical: return "#E74C3C";
                    default: return "#95A5A6";
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public static class ReportGenerator
    {
        public static List<ReportItem> Analyze(StrucSyncCore core, List<RelationshipItem> items, List<RevitLinkInstance> links, Document doc)
        {
            var reports = new List<ReportItem>();
            int counter = 1;
            double tolerance = 1.0e-9; // Dung sai siêu nhỏ (0.0000003 mm)

            foreach (var rel in items)
            {
                if (!rel.IsChecked) continue;

                Enum.TryParse(rel.ParentTypeStr, out SyncType pType);
                Enum.TryParse(rel.ChildTypeStr, out SyncType cType);

                // [FIX LỖI GROUP MODE]
                // Thay vì lấy tổng số lượng (rel.ChildIds.Count), ta đếm xem mỗi Parent cụ thể có bao nhiêu con.
                // Điều này giúp phân biệt: 100 cái đài đơn (Count=1) KHÁC VỚI 1 cái đài 100 cọc (Count=100).
                var parentCounts = rel.ParentIds
                                      .GroupBy(id => id)
                                      .ToDictionary(g => g.Key, g => g.Count());

                for (int i = 0; i < rel.ChildIds.Count; i++)
                {
                    var rItem = new ReportItem
                    {
                        Id = counter++,
                        SetName = rel.Name,
                        SourceItem = rel,
                        IndexInSet = i,
                        ParentId = rel.ParentIds[i],
                        ChildId = rel.ChildIds[i]
                    };

                    Element child = doc.GetElement(rel.ChildIds[i]);
                    if (child == null)
                    {
                        rItem.Severity = ReportSeverity.Critical;
                        rItem.StatusDisplay = "Missing Child";
                        reports.Add(rItem); continue;
                    }

                    // Lấy số lượng con thực tế của cha này
                    int actualChildCount = parentCounts.ContainsKey(rel.ParentIds[i]) ? parentCounts[rel.ParentIds[i]] : 1;

                    string offsetStr = (rel.Offsets != null && i < rel.Offsets.Count) ? rel.Offsets[i] : "";

                    // [LOGIC CỐT LÕI CHO PILE]
                    // Nếu là Cọc Đơn (1 Đài - 1 Cọc), ép Offset về rỗng -> Bắt buộc về tâm đài.
                    // Nếu là Group (1 Đài - Nhiều Cọc), giữ nguyên Offset -> Giữ vị trí tương đối.
                    if ((pType == SyncType.PileCap && cType == SyncType.Pile) && actualChildCount == 1)
                    {
                        offsetStr = "";
                    }

                    // Truyền actualChildCount vào hàm tính toán (Thay vì totalChild cũ)
                    XYZ targetPos = core.CalculateTargetPosition(rel.ParentIds[i], child, pType, cType, links, offsetStr, actualChildCount);

                    if (targetPos == null)
                    {
                        rItem.Severity = ReportSeverity.Critical;
                        rItem.StatusDisplay = "Missing Parent";
                        reports.Add(rItem); continue;
                    }

                    // --- TÍNH TOÁN ĐỘ LỆCH (DEVIATION) ---
                    XYZ currentPos = core.GetElementCenter(child, Transform.Identity);
                    if (child.Location is LocationPoint lp) currentPos = lp.Point;

                    double devX = currentPos.X - targetPos.X;
                    double devY = currentPos.Y - targetPos.Y;
                    double devZ = currentPos.Z - targetPos.Z;

                    // Convert sang mm
                    double devX_mm = devX * 304.8;
                    double devY_mm = devY * 304.8;
                    double devZ_mm = devZ * 304.8;
                    double distXY = Math.Sqrt(devX * devX + devY * devY);

                    // Hiển thị 3 số lẻ
                    rItem.TechData = $"ΔX={Math.Round(devX_mm, 3)}, ΔY={Math.Round(devY_mm, 3)}, ΔZ={Math.Round(devZ_mm, 3)} (mm)";

                    // --- PHÂN LOẠI DIAGNOSIS (KHÔNG ẢNH HƯỞNG CỘT/VÁCH) ---

                    // CASE A: CỌC & ĐÀI (Yêu cầu cả XY và Z phải đúng)
                    if ((pType == SyncType.PileCap && cType == SyncType.Pile) || (pType == SyncType.Pile && cType == SyncType.PileCap))
                    {
                        bool isXYOk = distXY <= tolerance;
                        bool isZOk = Math.Abs(devZ) <= tolerance;

                        string suffix = (actualChildCount > 1) ? " (Group)" : ""; // Chỉ hiện chữ Group nếu thực sự là Group

                        if (isXYOk && isZOk)
                        {
                            rItem.Severity = ReportSeverity.Synced;
                            rItem.StatusDisplay = "Synced";
                            rItem.Diagnosis = $"Perfect Match{suffix}.";
                        }
                        else
                        {
                            rItem.Severity = ReportSeverity.Critical;
                            rItem.StatusDisplay = "Deviation";

                            List<string> errs = new List<string>();
                            if (!isXYOk) errs.Add($"XY Shift: {Math.Round(distXY * 304.8, 2)}mm");
                            if (!isZOk) errs.Add($"Z Diff: {Math.Round(devZ_mm, 2)}mm");

                            rItem.Diagnosis = $"{string.Join(", ", errs)}{suffix}";
                        }
                    }
                    // CASE B: CỘT, VÁCH (Logic cũ: Chỉ check XY, bỏ qua Z)
                    else
                    {
                        // Logic bỏ qua Z cho Cột/Vách
                        // (Ở đây ta kiểm tra lại điều kiện Z ignore cho chắc chắn)
                        bool isZIgnored = true;

                        if (distXY <= tolerance)
                        {
                            rItem.Severity = ReportSeverity.Synced;
                            rItem.StatusDisplay = "Synced";
                            rItem.Diagnosis = $"Aligned XY. Z diff ({Math.Round(devZ_mm, 1)}mm) ignored.";
                        }
                        else
                        {
                            rItem.Severity = ReportSeverity.Critical;
                            rItem.StatusDisplay = "Deviation";
                            rItem.Diagnosis = $"Shifted XY: {Math.Round(distXY * 304.8, 3)}mm";
                        }
                    }

                    reports.Add(rItem);
                }
            }
            return reports;
        }
    }
}
