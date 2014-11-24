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

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for LinkConfigurationDialog.xaml
    /// </summary>
    public partial class LinkConfigurationDialog
    {
        public bool HasValidExcelFile { get; set; }
        private readonly Document _document;

        public LinkConfigurationDialog(LinkConfigurationViewModel viewModel)
        {
//            _excelApplication = excelApplication;
//            _document = doc;
//            _schedule = schedule;

//            HasValidExcelFile = false;
            DataContext = viewModel;
            InitializeComponent();
//            SheetComboBox.IsEnabled = false;

            Closing += SheetSelectionDialog_Closing;
        }

        void SheetSelectionDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (DialogResult == true)
                ((LinkConfigurationViewModel) DataContext).SheetSelectionDialog_Closing(sender, e);
        }

        void FilePathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ((LinkConfigurationViewModel) DataContext).Link.WorkbookPath = FilePathTextBox.Text;

//            if (System.IO.File.Exists(FilePathTextBox.Text))
//            {
//                HasValidExcelFile = true;
//                SheetComboBox.IsEnabled = true;
//                OpenWorkbookButton.IsEnabled = true;
//
//                using (var workbook = _excelApplication.Workbooks.Open(FilePathTextBox.Text, false, true))
//                {
//                    var worksheets = workbook.Worksheets.Cast<Excel.Worksheet>().Select(s => s.Name).ToList();
//                    SheetComboBox.ItemsSource = worksheets;
//
//                    if (worksheets.Any())
//                        SheetComboBox.SelectedIndex = 0;
//
//                    workbook.Close();
//                }
//            }
//            else
//            {
//                SheetComboBox.ItemsSource = Enumerable.Empty<string>();
//                HasValidExcelFile = false;
//                SheetComboBox.IsEnabled = false;
//                OpenWorkbookButton.IsEnabled = false;
//            }
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

            var link = ((LinkConfigurationViewModel) DataContext).Link;

            e.CanExecute = hasSelectedSheet && link.WorkbookExists && link.Error == "";
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void SheetComboBox_SourceUpdated(object sender, System.Windows.Data.DataTransferEventArgs e)
        {
            var wsName = ((LinkConfigurationViewModel)DataContext).Link.WorksheetName;

            if (SheetComboBox.Items.Cast<string>().Contains(wsName))
            {
                SheetComboBox.SelectedIndex = SheetComboBox.Items.Cast<String>().ToList().IndexOf(wsName);
            }
            else
            {
                SheetComboBox.SelectedIndex = SheetComboBox.Items.Count > 0 ? 0 : -1;
            }
        }

        private void SheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SheetComboBox.SelectedIndex >= 0)
                ((LinkConfigurationViewModel) DataContext).Link.WorksheetName = (string) SheetComboBox.SelectedItem;
        }

        private void ScheduleNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
//            ((LinkConfigurationViewModel) DataContext).Link.ScheduleName = ScheduleNameTextBox.Text;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Grid1.Focus();
        }
    }
}
