using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using THBIM.Supports;
using System.Linq;
#nullable disable

namespace HangerGenerator
{
    public partial class SupportGeneratorWindow : Window
    {
        private ExternalCommandData commandData;
        private Document doc;

        private List<Element> selectedElements = new List<Element>();

        // Delegate containing the appropriate handler based on support type
        private Action<ExternalCommandData, Document, List<Element>, double, double> currentHandler;

        public SupportGeneratorWindow(ExternalCommandData commandData, Document doc)
        {
            InitializeComponent();
            this.commandData = commandData;
            this.doc = doc;
            this.PreviewKeyDown += new KeyEventHandler(SuppressEscape);


            // NEW: Default lock Steel Type (only enable when U/V is selected)
            try
            {
                if (txtSteelType != null)
                {
                    txtSteelType.IsEnabled = false;
                    txtSteelType.Text = string.Empty;
                }
            }
            catch { }
            // NEW: Default lock LAYER 2
            try
            {
                if (chkLayer2 != null)
                {
                    chkLayer2.IsEnabled = false;
                    chkLayer2.IsChecked = false;
                }
            }
            catch { }
        }


        private void SuppressEscape(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
            }
        }

        private void cbSupportType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selected = (cbSupportType.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (string.IsNullOrEmpty(selected))
            {
                imgPreview.Source = null;
                currentHandler = null;

                // NEW: If nothing is selected, lock Steel Type
                if (txtSteelType != null)
                {
                    txtSteelType.IsEnabled = false;
                    txtSteelType.Text = string.Empty;
                }
                return;
            }
            // NEW: Enable/disable LAYER 2 CheckBox based on U/V family
            try
            {
                if (chkLayer2 != null)
                {
                    bool enableLayer2 = IsUVSelected(selected);
                    chkLayer2.IsEnabled = enableLayer2;

                    // When not U/V, uncheck to avoid keeping old state
                    if (!enableLayer2)
                        chkLayer2.IsChecked = false;
                }
            }
            catch { }


            switch (selected)
            {
                case "TH_Pipe Support Omega":
                    currentHandler = PipeOmegaSupport.Run;
                    break;
                case "TH_Foam Pipe Clamp":
                    currentHandler = FoamClampSupport.Run;
                    break;
                case "TH_U_Bolt":
                    currentHandler = UBoltSupport.Run;
                    break;
                case "TH_DA_U_Steel_Thread Rod":
                    currentHandler = USteelSupport.Run;
                    break;
                case "TH_DA_V_Steel_Thread Rod":
                    currentHandler = VSteelSupport.Run;
                    break;
                default:
                    currentHandler = null;
                    break;
                case "TH_DA_U_Steel_Foundation":
                    currentHandler = USteelFoundationSupport.Run;
                    break;
            }

            // NEW: Enable/disable Steel Type based on U/V family
            try
            {
                if (txtSteelType != null)
                {
                    bool enable = IsUVSelected(selected);
                    txtSteelType.IsEnabled = enable;
                    if (!enable) txtSteelType.Text = string.Empty;
                }
            }
            catch { }

            // Update preview image
            try
            {
                if (imgPreview == null) return;
                string baseDir = PluginPaths.BaseDir;
                string imagePath = Path.Combine(baseDir, "Resources", selected + ".png");

                if (File.Exists(imagePath))
                {
                    imgPreview.Source = new BitmapImage(new Uri(imagePath, UriKind.Absolute));
                }
                else
                {
                    imgPreview.Source = null;
                }
            }
            catch
            {
                imgPreview.Source = null;
            }
        }

        private void btnPickElements_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();

            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                var pickedRefs = uidoc.Selection.PickObjects(ObjectType.Element, new PipeDuctSelectionFilter(), "Select pipes or ducts");

                selectedElements.Clear();
                foreach (var r in pickedRefs)
                    selectedElements.Add(doc.GetElement(r));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
            }
            finally
            {
                this.Show();
                this.Activate();
            }
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string baseDir = PluginPaths.BaseDir;
                string resourceDir = Path.Combine(baseDir, "rfa");

                string[] familyFiles = Directory.GetFiles(resourceDir, "*.rfa", SearchOption.TopDirectoryOnly);

                int loadedCount = 0;
                using (Transaction t = new Transaction(doc, "Load Families"))
                {
                    t.Start();

                    foreach (var famPath in familyFiles)
                    {
                        string familyName = Path.GetFileNameWithoutExtension(famPath);

                        bool alreadyLoaded = new FilteredElementCollector(doc)
                            .OfClass(typeof(Family))
                            .Cast<Family>()
                            .Any(fam => fam.Name.Equals(familyName, StringComparison.InvariantCultureIgnoreCase));

                        if (!alreadyLoaded)
                        {
                            doc.LoadFamily(famPath);
                            loadedCount++;
                        }
                    }
                    t.Commit();
                }

                if (loadedCount == 0)
                    TaskDialog.Show("Notification", "All families already exist in the project.");
                else
                    TaskDialog.Show("Success", $"Loaded {loadedCount} families into the model.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading family: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool IsUVSelected(string familyName)
        {
            return familyName == "TH_DA_U_Steel_Thread Rod" || familyName == "TH_DA_V_Steel_Thread Rod";
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            if (currentHandler == null)
            {
                MessageBox.Show("Support type not selected.");
                return;
            }

            if (!double.TryParse(txtOffsetA.Text, out double aMM))
            {
                MessageBox.Show("Invalid spacing A.");
                return;
            }

            if (!double.TryParse(txtSpacingB.Text, out double bMM))
            {
                MessageBox.Show("Invalid spacing B.");
                return;
            }

            double offsetA = aMM / 304.8;
            double spacingB = bMM / 304.8;

            string selected = (cbSupportType.SelectedItem as ComboBoxItem)?.Content?.ToString();

            // NEW: get Steel Type from UI and pass to runtime for U/V to read
            string steelTypeValue = string.Empty;
            if (txtSteelType != null && IsUVSelected(selected))
                steelTypeValue = (txtSteelType.Text ?? string.Empty).Trim();
            THBIM.Supports.SupportRuntime.SteelType = steelTypeValue; // leave empty if not entered -> keep family default

            if (selected == "TH_DA_U_Steel_Thread Rod" || selected == "TH_DA_V_Steel_Thread Rod" || selected == "TH_DA_U_Steel_Foundation")
            {
                try
                {
                    currentHandler.Invoke(commandData, doc, null, offsetA, spacingB);

                    // Layer 2 only applies to U/V
                    if (chkLayer2 != null && chkLayer2.IsChecked == true && !string.IsNullOrEmpty(selected) && IsUVSelected(selected))
                        ApplyLayer2FlagForFamily(selected);

                    MessageBox.Show("Support created successfully.");

                    // Reset UI and Steel Type
                    cbSupportType.SelectedIndex = -1;
                    txtOffsetA.Text = "200";
                    txtSpacingB.Text = "3000";
                    if (txtSteelType != null)
                    {
                        txtSteelType.Text = string.Empty;
                        txtSteelType.IsEnabled = false;
                    }
                    THBIM.Supports.SupportRuntime.SteelType = string.Empty; // reset runtime
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    this.Show();
                    this.Activate();
                }
            }
            else
            {
                this.Hide();

                try
                {
                    UIDocument uidoc = commandData.Application.ActiveUIDocument;
                    var pickedRefs = uidoc.Selection.PickObjects(ObjectType.Element, new PipeDuctSelectionFilter(), "Select pipes or ducts");

                    List<Element> elements = new List<Element>();
                    foreach (var r in pickedRefs)
                        elements.Add(doc.GetElement(r));

                    currentHandler.Invoke(commandData, doc, elements, offsetA, spacingB);

                    // Layer 2 only applies to U/V
                    if (chkLayer2 != null && chkLayer2.IsChecked == true && !string.IsNullOrEmpty(selected) && IsUVSelected(selected))
                        ApplyLayer2FlagForFamily(selected, elements);

                    MessageBox.Show("Support created successfully.");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    this.Show();
                    this.Activate();
                }
            }

            // Final reset
            cbSupportType.SelectedIndex = -1;
            txtOffsetA.Text = "200";
            txtSpacingB.Text = "3000";
            if (txtSteelType != null)
            {
                txtSteelType.Text = string.Empty;
                txtSteelType.IsEnabled = false;
            }
            THBIM.Supports.SupportRuntime.SteelType = string.Empty; // reset runtime
        }

        private void ApplyLayer2FlagForFamily(string familyName, List<Element> hosts = null)
        {
            try
            {
                using (Transaction tr = new Transaction(doc, "Enable LAYER_2 flag"))
                {
                    tr.Start();

                    var instances = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol != null && fi.Symbol.Family != null && fi.Symbol.Family.Name == familyName);

                    if (hosts != null && hosts.Count > 0)
                    {
                        var hostIds = new HashSet<ElementId>(hosts.Select(h => h.Id));
                        instances = instances.Where(fi => fi.Host != null && hostIds.Contains(fi.Host.Id));
                    }

                    foreach (var fi in instances)
                    {
                        Parameter p = fi.LookupParameter("LAYER_2") ?? fi.LookupParameter(" LAYER_2");
                        if (p == null || p.IsReadOnly) continue;

                        switch (p.StorageType)
                        {
                            case StorageType.Integer:
                                p.Set(1);
                                break;
                            case StorageType.Double:
                                p.Set(1.0);
                                break;
                            case StorageType.String:
                                p.Set("1");
                                break;
                            default:
                                break;
                        }
                    }

                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot enable LAYER_2: " + ex.Message);
            }
        }

        private class PipeDuctSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Pipe || elem is Duct;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }
    }
}

// === Runtime holder for SteelType to avoid changing handler signature ===
namespace THBIM.Supports
{
    public static class SupportRuntime
    {
        public static string SteelType = string.Empty;
    }
}