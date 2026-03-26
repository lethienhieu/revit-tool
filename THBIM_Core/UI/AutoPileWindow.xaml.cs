using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;

namespace THBIM
{
    public partial class AutoPileWindow : Window
    {
        private Document _doc;

        // --- OUTPUT PROPERTIES ---
        public FamilySymbol SelectedPileSymbol { get; private set; }

        // Create Params (Override Length)
        public string OverrideParameterName { get; private set; }
        public double? OverrideValue { get; private set; }
        public bool IsLinkMode { get; private set; }

        public AutoPileWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            this.Loaded += AutoPileWindow_Loaded;
            LoadFamilies();
        }

        private void AutoPileWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (SessionSettings.HasRun) RestoreSettings();
        }

        private void RestoreSettings()
        {
            try
            {
                chkIsLink.IsChecked = SessionSettings.LastLinkMode;

                if (!string.IsNullOrEmpty(SessionSettings.LastFamilyName))
                {
                    PileTypeWrapper target = null;
                    foreach (PileTypeWrapper item in cmbFamilies.Items)
                    {
                        if (item.DisplayName == SessionSettings.LastFamilyName) { target = item; break; }
                    }

                    if (target != null)
                    {
                        cmbFamilies.SelectedItem = target;
                        LoadLengthParameters(target.Symbol);
                    }
                }

                if (!string.IsNullOrEmpty(SessionSettings.LastParamName) && cmbParameters.HasItems)
                {
                    foreach (PileParameter item in cmbParameters.Items)
                    {
                        if (item.Name == SessionSettings.LastParamName) { cmbParameters.SelectedItem = item; break; }
                    }
                }
                txtParamValue.Text = SessionSettings.LastParamValue;
            }
            catch { }
        }

        private void LoadFamilies()
        {
            var collector = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>().ToList();

            List<PileTypeWrapper> displayList = new List<PileTypeWrapper>();
            foreach (var symbol in collector)
            {
                string fName = symbol.FamilyName;
                bool isCap = fName.IndexOf("pilecap", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             fName.IndexOf("cap", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isCap) continue;
                displayList.Add(new PileTypeWrapper(symbol));
            }
            cmbFamilies.ItemsSource = displayList.OrderBy(x => x.DisplayName);

            // Mặc định chọn cái đầu tiên nếu chưa có Session
            if (displayList.Count > 0 && !SessionSettings.HasRun) cmbFamilies.SelectedIndex = 0;
        }

        private void CmbFamilies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!this.IsLoaded) return;
            var selectedWrapper = cmbFamilies.SelectedItem as PileTypeWrapper;
            if (selectedWrapper == null) return;

            LoadLengthParameters(selectedWrapper.Symbol);
        }

        private void LoadLengthParameters(FamilySymbol symbol)
        {
            cmbParameters.ItemsSource = null;
            cmbParameters.IsEnabled = false;
            HashSet<string> validNames = new HashSet<string>();

            using (Transaction t = new Transaction(_doc, "Temp Read Len Params"))
            {
                t.Start();
                try
                {
                    if (!symbol.IsActive) symbol.Activate();
                    Level dummyLevel = new FilteredElementCollector(_doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
                    if (dummyLevel != null)
                    {
                        FamilyInstance dummyPile = _doc.Create.NewFamilyInstance(XYZ.Zero, symbol, dummyLevel, Autodesk.Revit.DB.Structure.StructuralType.Footing);
                        foreach (Parameter p in dummyPile.Parameters) { if (IsValidLengthParam(p)) validNames.Add(p.Definition.Name); }
                    }
                    foreach (Parameter p in symbol.Parameters) { if (IsValidLengthParam(p)) validNames.Add(p.Definition.Name); }
                }
                catch { }
                t.RollBack();
            }
            if (validNames.Count > 0)
            {
                cmbParameters.ItemsSource = validNames.OrderBy(n => n).Select(n => new PileParameter { Name = n }).ToList();
                cmbParameters.DisplayMemberPath = "Name";
                cmbParameters.IsEnabled = true;
            }
        }

        private bool IsValidLengthParam(Parameter p)
        {
            if (p.IsReadOnly) return false;
            if (p.StorageType == StorageType.Double)
            {
                try { if (p.Definition.GetDataType() == SpecTypeId.Length) return true; } catch { }
            }
            return false;
        }

        // --- BUTTON EVENTS ---

        // NÚT APPLY (CREATE PILE)
        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            var selectedWrapper = cmbFamilies.SelectedItem as PileTypeWrapper;
            if (selectedWrapper == null)
            {
                MessageBox.Show("Please select a pile type!");
                return;
            }
            SelectedPileSymbol = selectedWrapper.Symbol;

            if (cmbParameters.SelectedIndex != -1 && !string.IsNullOrWhiteSpace(txtParamValue.Text))
            {
                var p = cmbParameters.SelectedItem as PileParameter;
                OverrideParameterName = p?.Name;
                if (double.TryParse(txtParamValue.Text, out double valMm)) OverrideValue = valMm / 304.8;
                else { MessageBox.Show("Invalid number format!"); return; }
            }
            else
            {
                OverrideParameterName = null;
                OverrideValue = null;
            }

            IsLinkMode = chkIsLink.IsChecked == true;

            SaveSession();
            this.DialogResult = true;
            this.Close();
        }

        private void SaveSession()
        {
            SessionSettings.HasRun = true;
            SessionSettings.LastLinkMode = IsLinkMode;
            SessionSettings.LastFamilyName = (cmbFamilies.SelectedItem as PileTypeWrapper)?.DisplayName;
            SessionSettings.LastParamName = OverrideParameterName;
            SessionSettings.LastParamValue = txtParamValue.Text;
        }
    }

    public class PileTypeWrapper
    {
        public FamilySymbol Symbol { get; }
        public string DisplayName { get; }
        public PileTypeWrapper(FamilySymbol symbol)
        {
            Symbol = symbol;
            DisplayName = $"{symbol.FamilyName} : {symbol.Name}";
        }
    }

    public class PileParameter { public string Name { get; set; } }

    public static class SessionSettings
    {
        public static string LastFamilyName { get; set; } = null;
        public static bool LastLinkMode { get; set; } = false;
        public static bool HasRun { get; set; } = false;
        public static string LastParamName { get; set; } = null;
        public static string LastParamValue { get; set; } = "";
    }
}