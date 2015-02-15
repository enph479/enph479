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
            
            var rule = (ScheduleNameValidator)ScheduleNameTextBox.GetBindingExpression(TextBox.TextProperty).ParentBinding.ValidationRules.First();
            rule.RevitDocument = Document;

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
                        else
                            SheetComboBox.SelectedIndex = -1;

                        workbook.Close();
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
            try
            {
                Process.Start(FilePathTextBox.Text);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Unable to start Excel", ex.Message);
            }
        }
    }
}
