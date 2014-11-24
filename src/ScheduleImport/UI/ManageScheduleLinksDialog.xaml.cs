using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for ManageScheduleLinksDialog.xaml
    /// </summary>
    public partial class ManageScheduleLinksDialog
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;

        public bool PressedCreate { get; private set; }

        public ManageScheduleLinksDialog(Document document, UIDocument uiDocument)
        {
            _document = document;
            _uiDocument = uiDocument;
            InitializeComponent();
        }

        private void ManagedScheduleLinksDataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            bool hasSelection = ManagedScheduleLinksDataGrid.SelectedIndex != -1;

            if (hasSelection)
            {
                ReloadButton.IsEnabled = true;
                ReloadFromButton.IsEnabled = true;
                RemoveButton.IsEnabled = true;
            }
            else
            {
                ReloadButton.IsEnabled = false;
                ReloadFromButton.IsEnabled = false;
                RemoveButton.IsEnabled = false;
            }
        }

        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedLink = (ManagedScheduleLink) ManagedScheduleLinksDataGrid.SelectedItem;

            var schedule = selectedLink.GetSchedule();
            var workbookPath = selectedLink.WorkbookPath;
            var worksheetName = selectedLink.WorksheetName;

            Cursor = Cursors.Wait;
            ExternalCommand.ImportSchedule(schedule, workbookPath, worksheetName);
            Cursor = Cursors.Arrow;
        }

        private void ReloadFromButton_Click(object sender, RoutedEventArgs e)
        {
            string workbookPath;
            string worksheetName;

            using (var excelApplication = new NetOffice.ExcelApi.Application { DisplayAlerts = false })
            {
                var wnd = new SheetSelectionDialog(excelApplication);

                if (wnd.ShowDialog() != true)
                    return;

                workbookPath = wnd.FilePathTextBox.Text;
                worksheetName = (string)wnd.SheetComboBox.SelectedItem;

                excelApplication.Quit();
            }

            var selectedLink = (ManagedScheduleLink)ManagedScheduleLinksDataGrid.SelectedItem;
            var schedule = selectedLink.GetSchedule();
            var schema = ExternalCommand.GetSchema();
            Debug.Assert(schedule.DeleteEntity(schema));

            Cursor = Cursors.Wait;
            ExternalCommand.ImportSchedule(schedule, workbookPath, worksheetName);
            ExternalCommand.StoreImportInformation(schedule, workbookPath, worksheetName);
            Cursor = Cursors.Arrow;

            ManagedScheduleLinksDataGrid.ItemsSource = ExternalCommand.GetManagedSchedules(_document);
            ManagedScheduleLinksDataGrid.Items.Refresh();
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedLink = (ManagedScheduleLink)ManagedScheduleLinksDataGrid.SelectedItem;
            var schedule = selectedLink.GetSchedule();

            var schema = ExternalCommand.GetSchema();
            Debug.Assert(schedule.DeleteEntity(schema));

            Cursor = Cursors.Wait;
            ManagedScheduleLinksDataGrid.ItemsSource = ExternalCommand.GetManagedSchedules(_document);
            Cursor = Cursors.Arrow;

            ManagedScheduleLinksDataGrid.Items.Refresh();
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            PressedCreate = true;
            DialogResult = true;

//            Visibility = System.Windows.Visibility.Collapsed;
//
//            var selectedPanel = new ElementSelector(_uiDocument).SelectSingle() as FamilyInstance;
//
//            Debug.Assert(selectedPanel != null);
//
//            Visibility = System.Windows.Visibility.Visible;
        }
    }
}
