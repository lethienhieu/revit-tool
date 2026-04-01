using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Autodesk.Revit.DB;

namespace THBIM
{
    public partial class CombineParamWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;

        // Models
        public class CategoryModel { public string Name { get; set; } public ElementId Id { get; set; } }
        public class ParameterModel { public string Name { get; set; } }

        // --- BINDING DATA ---

        // Fix lỗi CS0103: Khai báo biến private đầy đủ
        private bool _isAllCategories;
        public bool IsAllCategories
        {
            get => _isAllCategories;
            set
            {
                _isAllCategories = value;
                OnPropertyChanged("IsAllCategories");
                RefreshParameterLists();
                // Logic phụ trợ: Disable list category nếu chọn All
                if (icCategories != null) icCategories.IsEnabled = !value;
            }
        }

        public List<CategoryModel> AllCategoriesList { get; set; }
        public ObservableCollection<CatWrapper> SelectedCategoryRows { get; set; }

        public List<ParameterModel> SourceParamList { get; set; }
        public ObservableCollection<ParamWrapper> SourceParameters { get; set; }

        public List<ParameterModel> TargetParamList { get; set; }
        public ParameterModel TargetParameter { get; set; }

        private string _separator = "_";
        public string Separator
        {
            get => _separator;
            set { _separator = value; OnPropertyChanged("Separator"); UpdatePreview(); }
        }

        // Fix lỗi CS0103: Khai báo biến private đầy đủ
        private string _previewResult;
        public string PreviewResult
        {
            get => _previewResult;
            set { _previewResult = value; OnPropertyChanged("PreviewResult"); }
        }


        // --- CONSTRUCTOR ---
        public CombineParamWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            this.DataContext = this;

            SelectedCategoryRows = new ObservableCollection<CatWrapper>();
            SourceParameters = new ObservableCollection<ParamWrapper>();

            LoadAllCategories();

            AddCategoryRow();
            AddParameterRow();

            RefreshParameterLists();
        }

        // --- LOGIC LOAD DỮ LIỆU ---
        private void LoadAllCategories()
        {
            AllCategoriesList = new List<CategoryModel>();
            foreach (Category cat in _doc.Settings.Categories)
            {
                if (cat.CategoryType == CategoryType.Model && cat.AllowsBoundParameters)
                {
                    AllCategoriesList.Add(new CategoryModel { Name = cat.Name, Id = cat.Id });
                }
            }
            AllCategoriesList = AllCategoriesList.OrderBy(x => x.Name).ToList();
        }

        private void RefreshParameterLists()
        {
            HashSet<string> uniqueParams = new HashSet<string>();

            if (IsAllCategories)
            {
                var it = _doc.ParameterBindings.ForwardIterator();
                while (it.MoveNext()) if (it.Key != null) uniqueParams.Add(it.Key.Name);

                string[] common = { "Mark", "Comments", "Type Mark", "Type Name", "Level", "Family", "Family and Type" };
                foreach (var c in common) uniqueParams.Add(c);
            }
            else
            {
                foreach (var wrapper in SelectedCategoryRows)
                {
                    if (wrapper.SelectedCategory == null) continue;
                    Element elem = new FilteredElementCollector(_doc)
                                   .OfCategoryId(wrapper.SelectedCategory.Id)
                                   .WhereElementIsNotElementType()
                                   .FirstOrDefault();
                    if (elem != null)
                    {
                        foreach (Parameter p in elem.Parameters) uniqueParams.Add(p.Definition.Name);
                        Element typeElem = _doc.GetElement(elem.GetTypeId());
                        if (typeElem != null)
                            foreach (Parameter p in typeElem.Parameters) uniqueParams.Add(p.Definition.Name);
                    }
                }
            }

            var sorted = uniqueParams.Where(n => !string.IsNullOrEmpty(n)).OrderBy(n => n)
                                     .Select(n => new ParameterModel { Name = n }).ToList();

            SourceParamList = sorted;
            TargetParamList = sorted;

            OnPropertyChanged("SourceParamList");
            OnPropertyChanged("TargetParamList");

            UpdatePreview();
        }


        // --- EVENTS THÊM/XÓA DÒNG ---
        private void BtnAddCat_Click(object sender, RoutedEventArgs e) => AddCategoryRow();
        private void BtnRemoveCat_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is CatWrapper item)
            {
                item.PropertyChanged -= OnCatChanged;
                SelectedCategoryRows.Remove(item);
                RefreshParameterLists();
            }
        }
        private void AddCategoryRow()
        {
            var item = new CatWrapper();
            item.PropertyChanged += OnCatChanged;
            SelectedCategoryRows.Add(item);
        }
        private void OnCatChanged(object sender, PropertyChangedEventArgs e) => RefreshParameterLists();


        private void BtnAddParam_Click(object sender, RoutedEventArgs e) => AddParameterRow();
        private void BtnRemoveParam_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ParamWrapper item)
            {
                item.PropertyChanged -= OnParamChanged;
                SourceParameters.Remove(item);
                UpdatePreview();
            }
        }
        private void AddParameterRow()
        {
            var item = new ParamWrapper();
            item.PropertyChanged += OnParamChanged;
            SourceParameters.Add(item);
        }
        private void OnParamChanged(object sender, PropertyChangedEventArgs e) => UpdatePreview();


        // --- EXECUTE & PREVIEW ---
        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (TargetParameter == null) { MessageBox.Show("Select Target Param!"); return; }

            List<ElementId> catIds = new List<ElementId>();
            if (!IsAllCategories)
            {
                catIds = SelectedCategoryRows.Where(x => x.SelectedCategory != null)
                                             .Select(x => x.SelectedCategory.Id).ToList();
            }

            List<string> srcNames = SourceParameters.Select(x => x.SelectedParam?.Name).ToList();

            CombineParam.Execute(_doc, catIds, IsAllCategories, srcNames, TargetParameter.Name, Separator);
            this.Close();
        }

        private void UpdatePreview()
        {
            Element elem = null;
            FilteredElementCollector col = new FilteredElementCollector(_doc).WhereElementIsNotElementType();

            if (!IsAllCategories && SelectedCategoryRows.Count > 0 && SelectedCategoryRows[0].SelectedCategory != null)
            {
                col.OfCategoryId(SelectedCategoryRows[0].SelectedCategory.Id);
            }

            elem = col.FirstOrDefault();

            if (elem == null) { PreviewResult = "No element to preview."; return; }

            List<string> srcNames = SourceParameters.Where(x => x.SelectedParam != null).Select(x => x.SelectedParam.Name).ToList();
            PreviewResult = CombineParam.GenerateCombinedString(elem, srcNames, Separator);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // --- WRAPPERS & CONVERTER ---

    public class CatWrapper : INotifyPropertyChanged
    {
        private CombineParamWindow.CategoryModel _selectedCategory;
        public CombineParamWindow.CategoryModel SelectedCategory
        {
            get => _selectedCategory;
            set { _selectedCategory = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("SelectedCategory")); }
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class ParamWrapper : INotifyPropertyChanged
    {
        private CombineParamWindow.ParameterModel _selectedParam;
        public CombineParamWindow.ParameterModel SelectedParam
        {
            get => _selectedParam;
            set { _selectedParam = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("SelectedParam")); }
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    // Class 
    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is bool b) return !b;
            return value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => null;
    }
}