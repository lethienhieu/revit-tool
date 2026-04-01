using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace THBIM
{
    // ==========================================
    // 2. VIEW MODEL CHO BẢNG VÀ CỘT
    // ==========================================
    public class CategoryTableViewModel : ViewModelBase
    {
        public MultiCategoryExportViewModel ParentVM { get; private set; }
        public ObservableCollection<string> AvailableCategories => ParentVM.AvailableCategories;
        public ObservableCollection<string> AvailableLevels => ParentVM.AvailableLevels;

        public ObservableCollection<string> TableParameters { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> MainParameterValues { get; } = new ObservableCollection<string>();

        public ObservableCollection<string> AvailableLevelParameters { get; } = new ObservableCollection<string>();

        private string _selectedLevelParameter;
        public string SelectedLevelParameter
        {
            get => _selectedLevelParameter;
            set { if (_selectedLevelParameter != value) { _selectedLevelParameter = value; OnPropertyChanged(); LoadMainParameterValues(); foreach (var r in Rows) foreach (var c in r.Cells) c.UpdateValue(); } }
        }

        private string _selectedLevelValue = "All";
        public string SelectedLevelValue
        {
            get => _selectedLevelValue;
            set { if (_selectedLevelValue != value) { _selectedLevelValue = value; OnPropertyChanged(); LoadMainParameterValues(); foreach (var r in Rows) foreach (var c in r.Cells) c.UpdateValue(); } }
        }

        private string _mainParameterName;
        public string MainParameterName
        {
            get => _mainParameterName;
            set { _mainParameterName = value; MainParameterHeading = value; OnPropertyChanged(); LoadMainParameterValues(); }
        }

        private string _mainParameterHeading;
        public string MainParameterHeading
        {
            get => _mainParameterHeading;
            set { _mainParameterHeading = value; OnPropertyChanged(); }
        }

        public ObservableCollection<TableColumnViewModel> Columns { get; } = new ObservableCollection<TableColumnViewModel>();
        public ObservableCollection<RowViewModel> Rows { get; } = new ObservableCollection<RowViewModel>();
        public ObservableCollection<TotalCellViewModel> TotalCells { get; } = new ObservableCollection<TotalCellViewModel>();

        private string _selectedCategory;
        public string SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                if (_selectedCategory != value)
                {
                    _selectedCategory = value;
                    OnPropertyChanged();
                    LoadDataForCategory();
                    // Để trống Parameter ban đầu đúng như bạn yêu cầu
                    MainParameterName = "";
                    foreach (var c in Columns) c.ParameterName = "";
                }
            }
        }

        public ICommand AddRowCommand { get; set; }
        public ICommand AddAllRowsCommand { get; set; }
        public ICommand RemoveRowCommand { get; set; }
        public ICommand EditFieldsCommand { get; set; }

        public CategoryTableViewModel(MultiCategoryExportViewModel parentVM)
        {
            ParentVM = parentVM;

            AddRowCommand = new RelayCommand<object>(o => Rows.Add(new RowViewModel(this)));

            AddAllRowsCommand = new RelayCommand<object>(o =>
            {
                if (MainParameterValues == null || !MainParameterValues.Any()) return;
                var emptyRows = Rows.Where(r => string.IsNullOrEmpty(r.MainParameterValue)).ToList();
                foreach (var er in emptyRows) Rows.Remove(er);
                foreach (var val in MainParameterValues)
                {
                    if (!Rows.Any(r => r.MainParameterValue == val))
                    {
                        var newRow = new RowViewModel(this);
                        newRow.MainParameterValue = val;
                        Rows.Add(newRow);
                    }
                }
                RecalculateTotals();
            });

            RemoveRowCommand = new RelayCommand<RowViewModel>(row => { if (row != null) { Rows.Remove(row); RecalculateTotals(); } });

            EditFieldsCommand = new RelayCommand<object>(o =>
            {
                var currentSelected = new List<string>();
                if (!string.IsNullOrEmpty(MainParameterName)) currentSelected.Add(MainParameterName);
                currentSelected.AddRange(Columns.Select(c => c.ParameterName).Where(p => !string.IsNullOrEmpty(p)));

                var vm = new ParameterSelectionViewModel(TableParameters, currentSelected);
                var window = new ParameterSelectionWindow(vm);

                if (window.ShowDialog() == true)
                {
                    var newSelected = vm.ScheduledFields.ToList();
                    if (newSelected.Count > 0)
                    {
                        if (MainParameterName != newSelected[0]) MainParameterName = newSelected[0];
                        var colFields = newSelected.Skip(1).ToList();
                        while (ParentVM.GlobalColumnIds.Count < colFields.Count) ParentVM.AddGlobalColumnCommand.Execute(null);

                        for (int i = 0; i < Columns.Count; i++)
                        {
                            if (i < colFields.Count) { if (Columns[i].ParameterName != colFields[i]) Columns[i].ParameterName = colFields[i]; }
                            else { Columns[i].ParameterName = ""; }
                        }
                    }
                    else { MainParameterName = ""; foreach (var c in Columns) c.ParameterName = ""; }
                }
            });
        }

        public void AddColumn(string columnId)
        {
            var newCol = new TableColumnViewModel(this, columnId);
            Columns.Add(newCol);
            TotalCells.Add(new TotalCellViewModel(this, newCol));
            foreach (var row in Rows) row.Cells.Add(new CellViewModel(row, newCol));
        }

        public void RemoveColumn(string columnId)
        {
            var colToRemove = Columns.FirstOrDefault(c => c.ColumnId == columnId);
            if (colToRemove != null)
            {
                Columns.Remove(colToRemove);
                var totalToRemove = TotalCells.FirstOrDefault(t => t.Column == colToRemove);
                if (totalToRemove != null) TotalCells.Remove(totalToRemove);
                foreach (var row in Rows)
                {
                    var cellToRemove = row.Cells.FirstOrDefault(c => c.Column == colToRemove);
                    if (cellToRemove != null) row.Cells.Remove(cellToRemove);
                }
                RecalculateTotals();
            }
        }

        private void LoadDataForCategory()
        {
            if (string.IsNullOrEmpty(SelectedCategory)) return;
            TableParameters.Clear();
            AvailableLevelParameters.Clear();

            Category revitCat = ParentVM.Doc.Settings.Categories.Cast<Category>().FirstOrDefault(c => c.Name == SelectedCategory);
            if (revitCat == null) return;

            var sampleElement = new FilteredElementCollector(ParentVM.Doc).OfCategoryId(revitCat.Id).WhereElementIsNotElementType().FirstOrDefault();
            if (sampleElement != null)
            {
                var paramNames = new HashSet<string>();
                var levelParams = new HashSet<string>();

                foreach (Parameter p in sampleElement.Parameters)
                {
                    paramNames.Add(p.Definition.Name);

                    if (p.StorageType == StorageType.ElementId)
                    {
                        bool isLevel = false;
                        string n = p.Definition.Name.ToLower();
                        if (n.Contains("level") || n.Contains("constraint") || n.Contains("tầng") || n.Contains("base") || n.Contains("top")) isLevel = true;
                        if (isLevel) levelParams.Add(p.Definition.Name);
                    }
                }

                Element typeElem = ParentVM.Doc.GetElement(sampleElement.GetTypeId());
                if (typeElem != null) foreach (Parameter p in typeElem.Parameters) paramNames.Add(p.Definition.Name);

                // ĐÃ KHÔI PHỤC LẠI ĐẦY ĐỦ CÁC BIẾN ẢO
                paramNames.Add("Type Name");
                paramNames.Add("Family Name");
                paramNames.Add("Family and Type");

                foreach (var p in paramNames.OrderBy(n => n)) TableParameters.Add(p);
                foreach (var lp in levelParams.OrderBy(n => n)) AvailableLevelParameters.Add(lp);

                if (AvailableLevelParameters.Count > 0) SelectedLevelParameter = AvailableLevelParameters[0];
            }
        }

        public string GetParamValue(Document doc, Element elem, string paramName)
        {
            if (string.IsNullOrEmpty(paramName)) return "";

            // ĐÃ KHÔI PHỤC LẠI LOGIC LẤY GIÁ TRỊ FAMILY AND TYPE
            if (paramName == "Type Name") return elem.Name;
            if (paramName == "Family Name") { Element tElem = doc.GetElement(elem.GetTypeId()); if (tElem is ElementType et) return et.FamilyName; return ""; }
            if (paramName == "Family and Type") { Element tElem = doc.GetElement(elem.GetTypeId()); if (tElem is ElementType et) return $"{et.FamilyName}: {elem.Name}"; return ""; }

            Parameter p = elem.LookupParameter(paramName);
            if (p == null) { Element tElem = doc.GetElement(elem.GetTypeId()); if (tElem != null) p = tElem.LookupParameter(paramName); }
            if (p == null) return "";

            if (p.StorageType == StorageType.String) return p.AsString();
            if (p.StorageType == StorageType.Double) return p.AsValueString() ?? p.AsDouble().ToString();
            if (p.StorageType == StorageType.Integer)
            {
                try { if (p.Definition.GetDataType() == SpecTypeId.Boolean.YesNo) return p.AsInteger() == 1 ? "Yes" : "No"; } catch { }
                return p.AsValueString() ?? p.AsInteger().ToString();
            }
            if (p.StorageType == StorageType.ElementId)
            {
                ElementId id = p.AsElementId();
                if (id == ElementId.InvalidElementId) return "";
                Element target = doc.GetElement(id);
                return target != null ? target.Name : id.ToString();
            }
            return "";
        }

        public void LoadMainParameterValues()
        {
            MainParameterValues.Clear();
            if (string.IsNullOrEmpty(SelectedCategory) || string.IsNullOrEmpty(MainParameterName)) return;

            Category revitCat = ParentVM.Doc.Settings.Categories.Cast<Category>().FirstOrDefault(c => c.Name == SelectedCategory);
            if (revitCat == null) return;

            var elements = new FilteredElementCollector(ParentVM.Doc)
                .OfCategoryId(revitCat.Id)
                .WhereElementIsNotElementType()
                .ToList();

            if (!string.IsNullOrEmpty(SelectedLevelParameter) && !string.IsNullOrEmpty(SelectedLevelValue) && SelectedLevelValue != "All")
            {
                elements = elements.Where(e => GetParamValue(ParentVM.Doc, e, SelectedLevelParameter) == SelectedLevelValue).ToList();
            }

            var values = new HashSet<string>();
            foreach (var elem in elements)
            {
                string val = GetParamValue(ParentVM.Doc, elem, MainParameterName);
                if (!string.IsNullOrEmpty(val)) values.Add(val);
            }

            foreach (var v in values.OrderBy(n => n)) MainParameterValues.Add(v);
        }

        public void RecalculateTotals()
        {
            foreach (var totalCell in TotalCells)
            {
                var cellsInCol = Rows.SelectMany(r => r.Cells).Where(c => c.Column == totalCell.Column).ToList();
                double sum = cellsInCol.Sum(c => c.NumericValue);
                var validCell = cellsInCol.FirstOrDefault(c => c.ParamDataType != null);

                if (sum > 0 && validCell != null)
                {
                    try { totalCell.DisplayValue = UnitFormatUtils.Format(ParentVM.Doc.GetUnits(), validCell.ParamDataType, sum, false); }
                    catch { totalCell.DisplayValue = Math.Round(sum, 2).ToString(); }
                }
                else { totalCell.DisplayValue = sum > 0 ? Math.Round(sum, 2).ToString() : ""; }
            }
            ParentVM.RecalculateGrandTotals();
        }
    }

    public class TableColumnViewModel : ViewModelBase
    {
        public CategoryTableViewModel ParentTable { get; }
        public string ColumnId { get; }

        private string _parameterName;
        public string ParameterName
        {
            get => _parameterName;
            set { if (_parameterName != value) { _parameterName = value; Heading = value; OnPropertyChanged(); ParameterChanged?.Invoke(this, EventArgs.Empty); } }
        }

        private string _heading;
        public string Heading { get => _heading; set { _heading = value; OnPropertyChanged(); } }

        public event EventHandler ParameterChanged;

        public TableColumnViewModel(CategoryTableViewModel parentTable, string columnId)
        {
            ParentTable = parentTable;
            ColumnId = columnId;
        }
    }

    public class TotalCellViewModel : ViewModelBase
    {
        public CategoryTableViewModel ParentTable { get; }
        public TableColumnViewModel Column { get; }
        private string _displayValue;
        public string DisplayValue { get => _displayValue; set { _displayValue = value; OnPropertyChanged(); } }
        public TotalCellViewModel(CategoryTableViewModel parentTable, TableColumnViewModel column) { ParentTable = parentTable; Column = column; }
    }

    public class GrandTotalCellViewModel : ViewModelBase
    {
        public string ColumnId { get; }
        private string _displayValue;
        public string DisplayValue { get => _displayValue; set { _displayValue = value; OnPropertyChanged(); } }
        public GrandTotalCellViewModel(string columnId) { ColumnId = columnId; }
    }

    // ==========================================
    // 3. VIEW MODEL CHO HÀNG & Ô (CELL)
    // ==========================================
    public class RowViewModel : ViewModelBase
    {
        public CategoryTableViewModel ParentTable { get; private set; }
        public ObservableCollection<CellViewModel> Cells { get; } = new ObservableCollection<CellViewModel>();

        private string _note;
        public string Note { get => _note; set { _note = value; OnPropertyChanged(); } }

        private string _mainParameterValue;
        public string MainParameterValue
        {
            get => _mainParameterValue;
            set { if (_mainParameterValue != value) { _mainParameterValue = value; OnPropertyChanged(); foreach (var cell in Cells) cell.UpdateValue(); } }
        }

        public RowViewModel(CategoryTableViewModel parentTable)
        {
            ParentTable = parentTable;
            foreach (var col in ParentTable.Columns) Cells.Add(new CellViewModel(this, col));
        }
    }

    public class CellViewModel : ViewModelBase
    {
        public RowViewModel ParentRow { get; }
        public TableColumnViewModel Column { get; }

        private string _cellValue;
        public string CellValue { get => _cellValue; set { _cellValue = value; OnPropertyChanged(); } }

        public double NumericValue { get; private set; }
        public ForgeTypeId ParamDataType { get; private set; }
        public bool IsNumeric { get; private set; }

        public CellViewModel(RowViewModel parentRow, TableColumnViewModel column)
        {
            ParentRow = parentRow;
            Column = column;
            Column.ParameterChanged += (s, e) => UpdateValue();
        }

        public void UpdateValue()
        {
            NumericValue = 0;
            ParamDataType = null;
            IsNumeric = false;

            if (string.IsNullOrEmpty(ParentRow.MainParameterValue) || string.IsNullOrEmpty(Column.ParameterName))
            {
                CellValue = "";
                ParentRow.ParentTable.RecalculateTotals();
                return;
            }

            Document doc = ParentRow.ParentTable.ParentVM.Doc;
            string catName = ParentRow.ParentTable.SelectedCategory;
            string mainParamName = ParentRow.ParentTable.MainParameterName;
            string mainParamValue = ParentRow.MainParameterValue;
            string targetParamName = Column.ParameterName;

            Category revitCat = doc.Settings.Categories.Cast<Category>().FirstOrDefault(c => c.Name == catName);
            if (revitCat == null) return;

            var allInstances = new FilteredElementCollector(doc).OfCategoryId(revitCat.Id).WhereElementIsNotElementType().ToList();

            string levelParam = ParentRow.ParentTable.SelectedLevelParameter;
            string levelVal = ParentRow.ParentTable.SelectedLevelValue;
            if (!string.IsNullOrEmpty(levelParam) && !string.IsNullOrEmpty(levelVal) && levelVal != "All")
            {
                allInstances = allInstances.Where(e => ParentRow.ParentTable.GetParamValue(doc, e, levelParam) == levelVal).ToList();
            }

            var matchingInstances = new List<Element>();
            foreach (var elem in allInstances)
            {
                string val = ParentRow.ParentTable.GetParamValue(doc, elem, mainParamName);
                if (val == mainParamValue) matchingInstances.Add(elem);
            }

            if (!matchingInstances.Any()) { CellValue = "0"; ParentRow.ParentTable.RecalculateTotals(); return; }

            var firstInst = matchingInstances.First();
            bool isTypeParam = false;

            // ĐÃ KHÔI PHỤC KIỂM TRA FAMILY AND TYPE TẠI ĐÂY
            Parameter targetParam = null;
            if (targetParamName != "Type Name" && targetParamName != "Family Name" && targetParamName != "Family and Type")
            {
                targetParam = firstInst.LookupParameter(targetParamName);
                if (targetParam == null) { Element typeElem = doc.GetElement(firstInst.GetTypeId()); if (typeElem != null) { targetParam = typeElem.LookupParameter(targetParamName); isTypeParam = true; } }
            }

            // ĐÃ KHÔI PHỤC XUẤT CHUỖI FAMILY AND TYPE TẠI ĐÂY
            if (targetParamName == "Type Name") { CellValue = firstInst.Name; }
            else if (targetParamName == "Family Name") { Element tElem = doc.GetElement(firstInst.GetTypeId()); if (tElem is ElementType et) CellValue = et.FamilyName; }
            else if (targetParamName == "Family and Type") { Element tElem = doc.GetElement(firstInst.GetTypeId()); if (tElem is ElementType et) CellValue = $"{et.FamilyName}: {firstInst.Name}"; }
            else if (targetParam != null)
            {
                if (targetParam.StorageType == StorageType.Double)
                {
                    IsNumeric = true;
                    ParamDataType = targetParam.Definition.GetDataType();

                    double sum = 0;
                    if (isTypeParam) { foreach (var inst in matchingInstances) { Element tElem = doc.GetElement(inst.GetTypeId()); if (tElem != null) { Parameter tp = tElem.LookupParameter(targetParamName); if (tp != null) sum += tp.AsDouble(); } } }
                    else { sum = matchingInstances.Sum(e => e.LookupParameter(targetParamName)?.AsDouble() ?? 0); }

                    NumericValue = sum;
                    CellValue = UnitFormatUtils.Format(doc.GetUnits(), ParamDataType, NumericValue, false);
                }
                else if (targetParam.StorageType == StorageType.Integer && targetParam.Definition.GetDataType() != SpecTypeId.Boolean.YesNo)
                {
                    IsNumeric = true;
                    int sum = 0;
                    if (isTypeParam) { foreach (var inst in matchingInstances) { Element tElem = doc.GetElement(inst.GetTypeId()); if (tElem != null) { Parameter tp = tElem.LookupParameter(targetParamName); if (tp != null) sum += tp.AsInteger(); } } }
                    else { sum = matchingInstances.Sum(e => e.LookupParameter(targetParamName)?.AsInteger() ?? 0); }

                    NumericValue = sum;
                    CellValue = NumericValue.ToString();
                }
                else { CellValue = targetParam.AsValueString() ?? targetParam.AsString(); }
            }
            else { CellValue = "N/A"; }

            ParentRow.ParentTable.RecalculateTotals();
        }
    }

    // ==========================================
    // 4. VIEW MODEL CHỌN PARAMETER CÓ SEARCH
    // ==========================================
    public class ParameterSelectionViewModel : ViewModelBase
    {
        private List<string> _allAvailableFields = new List<string>();
        public ObservableCollection<string> AvailableFields { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> ScheduledFields { get; } = new ObservableCollection<string>();

        private string _selectedAvailableField;
        public string SelectedAvailableField
        {
            get => _selectedAvailableField;
            set { _selectedAvailableField = value; OnPropertyChanged(); }
        }

        private string _selectedScheduledField;
        public string SelectedScheduledField
        {
            get => _selectedScheduledField;
            set { _selectedScheduledField = value; OnPropertyChanged(); }
        }

        private string _searchText = "";
        public string SearchText
        {
            get => _searchText;
            set { if (_searchText != value) { _searchText = value; OnPropertyChanged(); FilterAvailableFields(); } }
        }

        public ICommand AddCommand { get; }
        public ICommand RemoveCommand { get; }
        public ICommand MoveUpCommand { get; }
        public ICommand MoveDownCommand { get; }

        public ParameterSelectionViewModel(IEnumerable<string> available, IEnumerable<string> scheduled)
        {
            foreach (var s in scheduled) ScheduledFields.Add(s);
            foreach (var a in available.OrderBy(x => x)) if (!ScheduledFields.Contains(a)) _allAvailableFields.Add(a);

            FilterAvailableFields();

            AddCommand = new RelayCommand<object>(o => {
                if (!string.IsNullOrEmpty(SelectedAvailableField))
                {
                    string item = SelectedAvailableField;
                    ScheduledFields.Add(item);
                    _allAvailableFields.Remove(item);
                    FilterAvailableFields();
                }
            });

            RemoveCommand = new RelayCommand<object>(o => {
                if (!string.IsNullOrEmpty(SelectedScheduledField))
                {
                    string item = SelectedScheduledField;
                    ScheduledFields.Remove(item);
                    _allAvailableFields.Add(item);
                    _allAvailableFields.Sort();
                    FilterAvailableFields();
                }
            });

            MoveUpCommand = new RelayCommand<object>(o => {
                int index = ScheduledFields.IndexOf(SelectedScheduledField);
                if (index > 0) ScheduledFields.Move(index, index - 1);
            });

            MoveDownCommand = new RelayCommand<object>(o => {
                int index = ScheduledFields.IndexOf(SelectedScheduledField);
                if (index >= 0 && index < ScheduledFields.Count - 1) ScheduledFields.Move(index, index + 1);
            });
        }

        private void FilterAvailableFields()
        {
            AvailableFields.Clear();
            var filtered = string.IsNullOrWhiteSpace(SearchText)
                ? _allAvailableFields
                : _allAvailableFields.Where(x => x.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var item in filtered) AvailableFields.Add(item);
        }
    }
}