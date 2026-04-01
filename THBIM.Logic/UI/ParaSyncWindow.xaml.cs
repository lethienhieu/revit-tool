using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    public partial class ParaSyncWindow : Window
    {
        private UIDocument _uiDoc;
        private Document _doc;
        private ParaSyncProcessor _processor;
        public ObservableCollection<MappingRow> MappingRows { get; set; }

        public ParaSyncWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
            _processor = new ParaSyncProcessor(uiDoc);

            MappingRows = new ObservableCollection<MappingRow>();
            icParamMapping.ItemsSource = MappingRows;

            LoadLinkInstances();
            LoadSpecificCategories();
        }

        private void LoadLinkInstances()
        {
            var links = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
            cbLinkInstance.ItemsSource = links;
            if (links.Count > 0) cbLinkInstance.SelectedIndex = 0;
        }

        private void LoadSpecificCategories()
        {
            List<BuiltInCategory> filterCats = new List<BuiltInCategory> {
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralFoundation
            };
            var categories = _doc.Settings.Categories.Cast<Category>()
                .Where(c => filterCats.Contains((BuiltInCategory)c.Id.GetHashCode())).OrderBy(c => c.Name).ToList();

            cbLinkCategory.ItemsSource = categories;
            cbHostCategory.ItemsSource = categories;
            cbLinkCategory.DisplayMemberPath = "Name";
            cbHostCategory.DisplayMemberPath = "Name";
        }

        private void BtnAddRow_Click(object sender, RoutedEventArgs e)
        {
            var linkInst = cbLinkInstance.SelectedItem as RevitLinkInstance;
            var linkCat = cbLinkCategory.SelectedItem as Category;
            var hostCat = cbHostCategory.SelectedItem as Category;
            if (linkInst == null || linkCat == null || hostCat == null) return;

            MappingRows.Add(new MappingRow
            {
                LinkParams = GetParamsFromDoc(linkInst.GetLinkDocument(), linkCat),
                HostParams = GetParamsFromDoc(_doc, hostCat)
            });
        }

        private ObservableCollection<string> GetParamsFromDoc(Document doc, Category cat)
        {
            if (doc == null || cat == null) return new ObservableCollection<string>();
            Element example = new FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType().FirstOrDefault();
            if (example == null) return new ObservableCollection<string>();
            return new ObservableCollection<string>(example.Parameters.Cast<Parameter>().Select(p => p.Definition.Name).OrderBy(n => n));
        }

        private void BtnRemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is MappingRow row) MappingRows.Remove(row);
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            var linkInst = cbLinkInstance.SelectedItem as RevitLinkInstance;
            var linkCat = cbLinkCategory.SelectedItem as Category;
            var hostCat = cbHostCategory.SelectedItem as Category;

            if (linkInst == null || linkCat == null || hostCat == null || !MappingRows.Any()) return;

            this.Hide();
            // Pass ElementId linkCat.Id to fix CS1503
            _processor.ExecuteWithSelection(linkInst, linkCat.Id, hostCat.Id, MappingRows.ToList());
            this.Show();
        }
    }
}