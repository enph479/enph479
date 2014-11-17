using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for SheetSelectionDialog.xaml
    /// </summary>
    public partial class SheetSelectionDialog
    {
        public bool HasValidExcelFile { get; set; }
        private readonly Excel.Application _excelApplication;

        public SheetSelectionDialog(Excel.Application excelApplication)
        {
            _excelApplication = excelApplication;
            HasValidExcelFile = false;
            InitializeComponent();
            SheetComboBox.IsEnabled = false;
        }

        void FilePathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (System.IO.File.Exists(FilePathTextBox.Text))
            {
                HasValidExcelFile = true;
                SheetComboBox.IsEnabled = true;
                OpenWorkbookButton.IsEnabled = true;

                // TODO This is probably way too slow to put here
                using (var workbook = _excelApplication.Workbooks.Open(FilePathTextBox.Text, false, true))
                {
                    var worksheets = workbook.Worksheets.Cast<Excel.Worksheet>().Select(s => s.Name).ToList();
                    SheetComboBox.ItemsSource = worksheets;

                    if (worksheets.Any())
                        SheetComboBox.SelectedIndex = 0;

                    workbook.Close();
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

            e.CanExecute = HasValidExcelFile && hasSelectedSheet;
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
