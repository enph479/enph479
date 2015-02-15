using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using Excel = NetOffice.ExcelApi;
using TextBox = System.Windows.Controls.TextBox;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for SheetSelectionDialog.xaml
    /// </summary>
    public partial class SheetSelectionDialog
    {
        public bool HasValidExcelFile { get; set; }
        public string ValidName { get; set; }

        public Document Document { get; set; }

        public string ScheduleName { get; set; }

        public string WorkbookPath { get; set; }

        public SheetSelectionDialog(Document doc)
        {
            this.DataContext = this;

            InitializeComponent();

            Document = doc;

            HasValidExcelFile = false;
            SheetComboBox.IsEnabled = false;

            Closing += SheetSelectionDialog_Closing;
            
            var rule = (ScheduleNameValidator)ScheduleNameTextBox.GetBindingExpression(TextBox.TextProperty).ParentBinding.ValidationRules.First();
            rule.RevitDocument = Document;

        }

        void SheetSelectionDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var scheduleName = ScheduleNameTextBox.Text;

            if (DialogResult == true && ValidName != null && scheduleName != ValidName)
            {
                var existingSchedule =
                    new FilteredElementCollector(Document).OfClass(typeof(PanelScheduleView))
                        .Cast<PanelScheduleView>().Any(s => s.Name == scheduleName);

                if (existingSchedule)
                {
                    var dlg = new TaskDialog("Cannot rename schedule");
                    dlg.MainInstruction = String.Format("There is already a panel schedule with name '{0}.'",
                        scheduleName);
                    dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                    dlg.Show();
                    e.Cancel = true;
                    Focus();
                }
            }
        }

        void FilePathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!FilePathTextBox.GetBindingExpression(TextBox.TextProperty).HasError)
            {
                HasValidExcelFile = true;
                SheetComboBox.IsEnabled = true;
                OpenWorkbookButton.IsEnabled = true;

                try
                {
                    using (var workbook = ExcelSingleton.OpenWorkbook(FilePathTextBox.Text))
                    {
                        var worksheets = workbook.Worksheets.Cast<Excel.Worksheet>().Select(s => s.Name).ToList();
                        SheetComboBox.ItemsSource = worksheets;

                        if (worksheets.Any())
                            SheetComboBox.SelectedIndex = 0;
                    }
                }
                catch (Exception)
                {

                }
            }
            else
            {
                SheetComboBox.ItemsSource = Enumerable.Empty<string>();
                HasValidExcelFile = false;
                SheetComboBox.IsEnabled = false;
                OpenWorkbookButton.IsEnabled = false;
            }
        }

        private void FileDialogButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Multiselect = false,
                CheckFileExists = true
            };

            var dialogResult = fileDialog.ShowDialog(owner: this);

            if (dialogResult == true)
            {
                FilePathTextBox.Text = fileDialog.FileName;
            }
        }

        private void CanExecuteHandler(object sender, CanExecuteRoutedEventArgs e)
        {
            bool hasSelectedSheet = SheetComboBox.SelectedIndex > -1;

            bool validationErrors = ScheduleNameTextBox.GetBindingExpression(TextBox.TextProperty).HasError
                                    || FilePathTextBox.GetBindingExpression(TextBox.TextProperty).HasError;

            e.CanExecute = !validationErrors && hasSelectedSheet;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void OpenWorkbookButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(FilePathTextBox.Text);
        }
    }
}
