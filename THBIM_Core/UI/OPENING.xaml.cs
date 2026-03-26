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

                // Load static preview image from Resources/OPENINGPRV.png
                LoadPreview();

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
                string asm = Assembly.GetExecutingAssembly().Location;
                // Check for null/empty to prevent Path.GetDirectoryName crash
                if (string.IsNullOrEmpty(asm)) return string.Empty;
                return Path.GetDirectoryName(asm);
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

        // ===== Preview: Always use OPENINGPRV.png =====
        // --- [FIXED] Safe Image Loading Method ---
        private void LoadPreview()
        {
            if (imgPreview == null) return;

            try
            {
                string dir = GetAssemblyDir();

                // Abort immediately if directory is missing to prevent errors
                if (string.IsNullOrEmpty(dir)) return;

                string p1 = Path.Combine(dir, "Resources", "OPENINGPRV.png");

                if (File.Exists(p1))
                {
                    // Safe image loading for WPF: prevents file locking and format errors
                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad; // Release the file handle immediately after loading
                    bi.UriSource = new Uri(p1);
                    bi.EndInit();
                    imgPreview.Source = bi;
                }
                else
                {
                    imgPreview.Source = null;
                }
            }
            catch
            {
                // If image fails, leave empty; do not throw exception to avoid crashing the app
                imgPreview.Source = null;
            }
        }

        private void cbOpeningType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_uiReady) return;
            // Preview is static, just reload it (if needed)
            LoadPreview();
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
                TaskDialog.Show("OPENING", ex.Message);
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