using Autodesk.Revit.DB;
using Autodesk.Revit.UI; // TaskDialog, ExternalCommandData
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Shapes = System.Windows.Shapes;
using OPENING.MODEL;
using MediaColor = System.Windows.Media.Color;
using RvtColor = Autodesk.Revit.DB.Color;
#nullable disable

// ====== ADDED SECTIONS ======
using System.Collections.Generic;
using System.Windows.Media;
// ====== /ADDED SECTIONS ======

// Alias to avoid conflict with Autodesk.Revit.UI.ComboBox
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfComboBoxItem = System.Windows.Controls.ComboBoxItem;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace OPENING
{
    public partial class OPENING : Window, INotifyPropertyChanged
    {
        private ExternalCommandData _commandData;
        private Document _doc;
        private bool _uiReady = false;

        // ===== COUNTERS (data bound to the UI scoreboard) =====
        private int _ductsCount;
        public int DuctsCount { get => _ductsCount; set { _ductsCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalCount)); } }

        private int _pipesCount;
        public int PipesCount { get => _pipesCount; set { _pipesCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalCount)); } }

        private int _cableTraysCount;
        public int CableTraysCount { get => _cableTraysCount; set { _cableTraysCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalCount)); } }

        private int _conduitsCount;
        public int ConduitsCount { get => _conduitsCount; set { _conduitsCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalCount)); } }

        private int _multiCount;
        public int MultiCount { get => _multiCount; set { _multiCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(TotalCount)); } }

        public int TotalCount => DuctsCount + PipesCount + CableTraysCount + ConduitsCount + MultiCount;

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public OPENING()
        {
            InitializeComponent();
            DataContext = this;

            this.Loaded += (s, e) =>
            {
                _uiReady = true;

                if (cbOpeningType != null && cbOpeningType.SelectedIndex < 0)
                    cbOpeningType.SelectedIndex = 0;

                // Draw canvas preview
                DrawCanvas();

                if (_doc != null)
                {
                    try { RecountOpenings(); } catch { }

                    // ====== ADDED: Populate the list of Revit Links once _doc is available ======
                    try { PopulateRevitLinks(); } catch { }
                    // ====== /ADDED ======
                }
            };
        }

        public OPENING(ExternalCommandData commandData, Document doc) : this()
        {
            _commandData = commandData;
            _doc = doc;

            if (_doc != null)
            {
                try { RecountOpenings(); } catch { }
            }
        }

        // ===== Helpers =====
        private static double MmToFt(double mm) => mm / 304.8;
        private static double ParseOrZero(string s) => double.TryParse(s, out var v) ? v : 0;

        // --- [FIXED] Safe method to get directory path (compatible with .NET 8) ---
        private static string GetAssemblyDir()
        {
            try
            {
                return PluginPaths.BaseDir;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string FindExistingPath(params string[] paths)
        {
            foreach (var p in paths)
            {
                // Check for null/empty string before checking file existence to avoid errors
                if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;
            }
            return null;
        }

        // ===== Canvas Preview (dynamic drawing) =====
        private void UpdatePreview(object sender, TextChangedEventArgs e)
        {
            if (this.IsLoaded) DrawCanvas();
        }

        private void cbOpeningType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_uiReady) return;
            DrawCanvas();
        }

        private void DrawCanvas()
        {
            if (cvsPreview == null) return;
            cvsPreview.Children.Clear();

            double clearance = 10, extrusion = 20, fill = 25;
            double.TryParse(txtClearance?.Text, out clearance);
            double.TryParse(txtExtrusion?.Text, out extrusion);
            double.TryParse(txtFill?.Text, out fill);

            double baseMEP = 60;
            double safeFill = Math.Min(Math.Max(0, fill), 40);
            double safeClearance = Math.Min(Math.Max(0, clearance), 30);
            double safeExtrusion = Math.Min(Math.Max(0, extrusion), 60);

            double innerWhiteSize = baseMEP + (safeFill * 2);
            double outerGraySize = innerWhiteSize + (safeClearance * 2);

            // ================= FRONT VIEW =================
            double cxFront = 140, cyFront = 160;

            AddRect(cxFront, cyFront, 240, 240, "#B85550");
            AddRect(cxFront, cyFront, outerGraySize, outerGraySize, "#6E6E6E");
            AddRect(cxFront, cyFront, innerWhiteSize, innerWhiteSize, "#FFFFFF");

            // Cross lines (fill corners)
            AddLine(cxFront - outerGraySize / 2, cyFront - outerGraySize / 2, cxFront - innerWhiteSize / 2, cyFront - innerWhiteSize / 2, "#999999");
            AddLine(cxFront + outerGraySize / 2, cyFront - outerGraySize / 2, cxFront + innerWhiteSize / 2, cyFront - innerWhiteSize / 2, "#999999");
            AddLine(cxFront - outerGraySize / 2, cyFront + outerGraySize / 2, cxFront - innerWhiteSize / 2, cyFront + innerWhiteSize / 2, "#999999");
            AddLine(cxFront + outerGraySize / 2, cyFront + outerGraySize / 2, cxFront + innerWhiteSize / 2, cyFront + innerWhiteSize / 2, "#999999");

            // MEP core
            AddRect(cxFront, cyFront, baseMEP, baseMEP, "#0071CE");

            // Dimension annotations
            DrawDimension(cxFront + baseMEP / 2, cyFront + 10, cxFront + innerWhiteSize / 2, cyFront + 10, fill.ToString(), 0, -5);
            DrawDimension(cxFront + innerWhiteSize / 2, cyFront - 20, cxFront + outerGraySize / 2, cyFront - 20, clearance.ToString(), 0, -5);

            // ================= SIDE VIEW =================
            double cxSide = 380, cySide = 160;
            double wallThickness = 70;
            double openingSideWidth = wallThickness + (safeExtrusion * 2);

            AddRect(cxSide, cySide, wallThickness, 300, "#B85550");
            AddRect(cxSide, cySide, openingSideWidth, outerGraySize, "#6E6E6E");

            double dimY = cySide - outerGraySize / 2 - 15;

            // Left extrusion
            DrawDimension(cxSide - openingSideWidth / 2, dimY, cxSide - wallThickness / 2, dimY, extrusion.ToString(), 0, -5);
            AddLine(cxSide - openingSideWidth / 2, cySide - outerGraySize / 2, cxSide - openingSideWidth / 2, dimY, "#333333", true);
            AddLine(cxSide - wallThickness / 2, cySide - outerGraySize / 2, cxSide - wallThickness / 2, dimY, "#333333", true);

            // Right extrusion
            DrawDimension(cxSide + wallThickness / 2, dimY, cxSide + openingSideWidth / 2, dimY, extrusion.ToString(), 0, -5);
            AddLine(cxSide + wallThickness / 2, cySide - outerGraySize / 2, cxSide + wallThickness / 2, dimY, "#333333", true);
            AddLine(cxSide + openingSideWidth / 2, cySide - outerGraySize / 2, cxSide + openingSideWidth / 2, dimY, "#333333", true);
        }

        // ===== Canvas drawing helpers =====
        private void AddRect(double cx, double cy, double w, double h, string hexFill)
        {
            var rect = new Shapes.Rectangle { Width = w, Height = h };
            rect.Fill = (SolidColorBrush)new BrushConverter().ConvertFrom(hexFill);
            Canvas.SetLeft(rect, cx - w / 2);
            Canvas.SetTop(rect, cy - h / 2);
            cvsPreview.Children.Add(rect);
        }

        private void AddLine(double x1, double y1, double x2, double y2, string hexStroke, bool isDashed = false)
        {
            var line = new Shapes.Line { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2 };
            line.Stroke = (SolidColorBrush)new BrushConverter().ConvertFrom(hexStroke);
            line.StrokeThickness = 1;
            if (isDashed) line.StrokeDashArray = new DoubleCollection { 2, 2 };
            cvsPreview.Children.Add(line);
        }

        private void DrawDimension(double x1, double y1, double x2, double y2, string text, double offsetX, double offsetY)
        {
            AddLine(x1, y1, x2, y2, "#000000");

            var dot1 = new Shapes.Ellipse { Width = 4, Height = 4, Fill = Brushes.Black };
            Canvas.SetLeft(dot1, x1 - 2); Canvas.SetTop(dot1, y1 - 2);
            cvsPreview.Children.Add(dot1);

            var dot2 = new Shapes.Ellipse { Width = 4, Height = 4, Fill = Brushes.Black };
            Canvas.SetLeft(dot2, x2 - 2); Canvas.SetTop(dot2, y2 - 2);
            cvsPreview.Children.Add(dot2);

            if (Math.Abs(x2 - x1) > 5 || text.Length > 0)
            {
                var tb = new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.Black, FontWeight = FontWeights.SemiBold };
                tb.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                Canvas.SetLeft(tb, (x1 + x2) / 2 - tb.DesiredSize.Width / 2 + offsetX);
                Canvas.SetTop(tb, y1 + offsetY - 14);
                cvsPreview.Children.Add(tb);
            }
        }

        // ===== Load Family =====
        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            string familyName = (cbOpeningType.SelectedItem as WpfComboBoxItem)?.Content?.ToString();
            if (string.IsNullOrWhiteSpace(familyName))
            {
                TaskDialog.Show("OPENING", "No Family selected.");
                return;
            }

            string dir = GetAssemblyDir();
            // --- [FIXED] Validate directory path ---
            if (string.IsNullOrEmpty(dir))
            {
                TaskDialog.Show("OPENING", "Cannot determine Plugin installation folder.");
                return;
            }

            string p1 = Path.Combine(dir, "Resources", familyName + ".rfa");
            string p2 = Path.Combine(dir, "Resource", familyName + ".rfa");
            string p3 = Path.Combine(dir, "rfa", familyName + ".rfa");
            string rfa = FindExistingPath(p1, p2, p3);

            if (rfa == null)
            {
                TaskDialog.Show("OPENING", $"RFA file not found for '{familyName}'.\nSearched in:\n- {p1}\n- {p2}\n- {p3}");
                return;
            }

            using (var t = new Transaction(_doc, "Load Family"))
            {
                t.Start();
                _doc.LoadFamily(rfa);
                t.Commit();
            }
            TaskDialog.Show("OPENING", $"Loaded '{familyName}'.");
        }

        // ===== Create Opening (Hide UI then allow pick) =====
        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            string familyName = (cbOpeningType.SelectedItem as WpfComboBoxItem)?.Content?.ToString();
            if (string.IsNullOrWhiteSpace(familyName))
            {
                TaskDialog.Show("OPENING", "No Family selected.");
                return;
            }

            double clearanceFt = MmToFt(ParseOrZero(txtClearance.Text));
            double extrusionFt = MmToFt(ParseOrZero(txtExtrusion.Text));
            double fillFt = MmToFt(ParseOrZero(txtFill.Text));

            var cmd = _commandData;
            var d = _doc;

            // List of checked (selected) links
            var selectedLinks = GetSelectedLinkDocs();

            this.Close(); // Close UI to allow user picking

            try
            {
                switch (familyName)
                {
                    case "TH_ROUND_SLEEVE":
                        // ⟵ PASS selectedLinks into the function
                        TH_ROUND_SLEEVE.Run(cmd, d, clearanceFt, extrusionFt, fillFt, selectedLinks);
                        break;

                    case "TH_RECTANGULAR_SLEEVE":
                        TH_RECTANGULAR_SLEEVE.Run(cmd, d, clearanceFt, extrusionFt, fillFt, selectedLinks);
                        break;

                    case "TH_RECTANGULAR_MULTI_SLEEVE":
                        TH_RECTANGULAR_MULTI_SLEEVE.Run(cmd, d, clearanceFt, extrusionFt, fillFt, selectedLinks);
                        break;

                    default:
                        TaskDialog.Show("OPENING", "Invalid Family.");
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("THBIM Tools - OPENING", $"{ex.GetType().Name}: {ex.Message}\n\nStack:\n{ex.StackTrace}");
            }
        }


        // ===== Count openings for the scoreboard =====
        private void RecountOpenings()
        {
            int ducts = 0, pipes = 0, trays = 0, conduits = 0, multi = 0;

            var fis = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    var fam = fi.Symbol?.Family?.Name;
                    return fam == "TH_ROUND_SLEEVE" ||
                           fam == "TH_RECTANGULAR_SLEEVE" ||
                           fam == "TH_RECTANGULAR_MULTI_SLEEVE";
                });

            foreach (var fi in fis)
            {
                string fam = fi.Symbol.Family.Name;

                if (fam == "TH_RECTANGULAR_MULTI_SLEEVE")
                {
                    multi++;
                    continue;
                }

                // 1) Case: Host is an MEP element in the CURRENT DOC: count by host category
                var host = fi.Host as Element;
                if (host?.Category != null)
                {
                    // Used .Value (required for Revit 2025 API compatibility)
                    var bic = (BuiltInCategory)host.Category.Id.GetValue();
                    if (bic == BuiltInCategory.OST_PipeCurves) { pipes++; continue; }
                    if (bic == BuiltInCategory.OST_DuctCurves) { ducts++; continue; }
                    if (bic == BuiltInCategory.OST_CableTray) { trays++; continue; }
                    if (bic == BuiltInCategory.OST_Conduit) { conduits++; continue; }
                }

                // 2) Case: Host is in a LINK: retrieve the MEP element from the link via HostFace
                if (TryGetLinkedHostedElement(fi, out Element linked))
                {
                    if (linked is Pipe) { pipes++; continue; }
                    if (linked is Duct) { ducts++; continue; }
                    if (linked is Conduit) { conduits++; continue; }
                    // Cable Tray in link (future proofing if hosting is supported) - Used .Value (for 2025)
                    if (linked.Category?.Id.GetValue() == (long)BuiltInCategory.OST_CableTray) { trays++; continue; }
                }

                // 3) Fallback (if detection fails): try reading a text parameter storing MEP type (if you added one)
                var kind = fi.LookupParameter("TH_MEP_KIND")?.AsString();
                if (string.Equals(kind, "Pipe", StringComparison.OrdinalIgnoreCase)) { pipes++; continue; }
                if (string.Equals(kind, "Duct", StringComparison.OrdinalIgnoreCase)) { ducts++; continue; }
                if (string.Equals(kind, "Conduit", StringComparison.OrdinalIgnoreCase)) { conduits++; continue; }
                if (string.Equals(kind, "CableTray", StringComparison.OrdinalIgnoreCase)) { trays++; continue; }
            }

            DuctsCount = ducts;
            PipesCount = pipes;
            CableTraysCount = trays;
            ConduitsCount = conduits;
            MultiCount = multi;
        }
        private static bool TryGetLinkedHostedElement(FamilyInstance fi, out Element linkedElem)
        {
            linkedElem = null;

            // Only proceed if host is a RevitLinkInstance
            var rli = fi.Host as RevitLinkInstance;
            if (rli == null) return false;

            // For face-hosted families, HostFace holds the reference to the hosted face
            Reference rf = fi.HostFace;
            if (rf == null || rf.LinkedElementId == ElementId.InvalidElementId) return false;

            Document ldoc = rli.GetLinkDocument();
            if (ldoc == null) return false;

            linkedElem = ldoc.GetElement(rf.LinkedElementId);
            return linkedElem != null;
        }


        // ====== ADDED: Revit Links helper ======
        private void PopulateRevitLinks()
        {
            // Requirement: XAML must contain a StackPanel named "spLinks"
            if (_doc == null || spLinks == null) return;

            spLinks.Children.Clear();

            var links = new FilteredElementCollector(_doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .OrderBy(l => l.Name)
                .ToList();

            if (links.Count == 0)
            {
                spLinks.Children.Add(new TextBlock
                {
                    Text = "No Revit links found.",
                    Foreground = new SolidColorBrush(MediaColor.FromRgb(120, 120, 120)),
                    Margin = new Thickness(0, 2, 0, 0)
                });
                return;
            }

            foreach (var rli in links)
            {
                bool isUsable = rli.GetLinkDocument() != null; // Loaded & accessible
                var cb = new CheckBox
                {
                    Content = rli.Name,
                    Margin = new Thickness(0, 2, 0, 2),
                    IsChecked = false,            // Automatically check (tick) if the link is usable (defaulted to false here)
                    Tag = rli.Id,                 // Store ElementId to retrieve the element later
                    ToolTip = isUsable ? "Ready to use" : "Link not loaded"
                };
                spLinks.Children.Add(cb);
            }
        }

        private List<Document> GetSelectedLinkDocs()
        {
            var result = new List<Document>();
            if (_doc == null || spLinks == null) return result;

            foreach (var child in spLinks.Children.OfType<CheckBox>())
            {
                if (child.IsChecked != true) continue;

                if (child.Tag is ElementId id)
                {
                    var inst = _doc.GetElement(id) as RevitLinkInstance;
                    var linkDoc = inst?.GetLinkDocument();
                    if (linkDoc != null) result.Add(linkDoc);
                }
            }
            return result;
        }
        // ====== /ADDED ======
    }
}