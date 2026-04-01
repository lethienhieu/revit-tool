using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;

namespace THBIM
{
    public partial class ZoneWindow : Window
    {
        private Document _doc;
        public string SelectedZoneParam { get; private set; }
        public string SelectedZoneValue { get; private set; }
        public string SelectedNumberParam { get; private set; }
        public string Prefix { get; private set; }
        public int StartNumber { get; private set; } = 1;
        public int Digits { get; private set; } = 3;
        public BuiltInCategory SelectedCategory { get; private set; }

        public ZoneWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            LoadCategories();
            this.Loaded += (s, e) => { if (ZoneSession.HasRun) RestoreSettings(); else cmbCategories.SelectedIndex = 0; };
        }

        private void LoadCategories()
        {
            cmbCategories.ItemsSource = new List<CategoryWrapper>
            {
                new CategoryWrapper("Structural Foundations", BuiltInCategory.OST_StructuralFoundation),
                new CategoryWrapper("Structural Framing", BuiltInCategory.OST_StructuralFraming),
                new CategoryWrapper("Structural Columns", BuiltInCategory.OST_StructuralColumns),
                new CategoryWrapper("Floors", BuiltInCategory.OST_Floors),
                new CategoryWrapper("Roofs", BuiltInCategory.OST_Roofs)
            };
        }

        private void RestoreSettings()
        {
            var target = cmbCategories.Items.Cast<CategoryWrapper>().FirstOrDefault(x => x.Name == ZoneSession.LastCategoryName);
            if (target != null) cmbCategories.SelectedItem = target;
            txtZoneValue.Text = ZoneSession.LastValue;
            txtPrefix.Text = ZoneSession.LastPrefix;
        }

        private void CmbCategories_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (cmbCategories.SelectedItem is CategoryWrapper selected) LoadParameters(selected.UICategory);
        }

        private void LoadParameters(BuiltInCategory bic)
        {
            var sample = new FilteredElementCollector(_doc).OfCategory(bic).WhereElementIsNotElementType().FirstOrDefault();
            HashSet<string> names = new HashSet<string>();

            if (sample != null)
            {
                foreach (Parameter p in sample.Parameters) if (!p.IsReadOnly && p.StorageType == StorageType.String) names.Add(p.Definition.Name);
                Element type = _doc.GetElement(sample.GetTypeId());
                if (type != null) foreach (Parameter p in type.Parameters) if (!p.IsReadOnly && p.StorageType == StorageType.String) names.Add(p.Definition.Name);
            }

            var source = names.OrderBy(n => n).Select(n => new ParameterModel { Name = n }).ToList();
            cmbZoneParams.ItemsSource = source;
            cmbNumberParams.ItemsSource = source;
            cmbZoneParams.IsEnabled = cmbNumberParams.IsEnabled = source.Count > 0;
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            // LOGIC MỚI: Chỉ cần 1 trong 2 được chọn
            bool isZoneSelected = cmbZoneParams.SelectedItem != null;
            bool isNumberSelected = cmbNumberParams.SelectedItem != null;

            if (!isZoneSelected && !isNumberSelected)
            {
                MessageBox.Show("Please select at least one parameter (Zone or Numbering) to proceed.");
                return;
            }

            SelectedCategory = (cmbCategories.SelectedItem as CategoryWrapper).UICategory;

            if (isZoneSelected)
            {
                SelectedZoneParam = (cmbZoneParams.SelectedItem as ParameterModel).Name;
                SelectedZoneValue = txtZoneValue.Text;
            }

            if (isNumberSelected)
            {
                SelectedNumberParam = (cmbNumberParams.SelectedItem as ParameterModel).Name;
                Prefix = txtPrefix.Text;
                int.TryParse(txtStart.Text, out int s); StartNumber = s;
                int.TryParse(txtDigits.Text, out int d); Digits = d;
            }

            ZoneSession.HasRun = true;
            ZoneSession.LastCategoryName = (cmbCategories.SelectedItem as CategoryWrapper).Name;
            ZoneSession.LastValue = txtZoneValue.Text;
            ZoneSession.LastPrefix = txtPrefix.Text;

            this.DialogResult = true;
            this.Close();
        }
    }

    public class CategoryWrapper { public string Name { get; } public BuiltInCategory UICategory { get; } public CategoryWrapper(string n, BuiltInCategory c) { Name = n; UICategory = c; } }
    public class ParameterModel { public string Name { get; set; } }
    public static class ZoneSession { public static bool HasRun = false; public static string LastCategoryName, LastValue, LastPrefix; }
}