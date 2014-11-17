using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElectricalToolSuite.ScheduleImport.Annotations;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for ManageScheduleLinksDialog.xaml
    /// </summary>
    public partial class ManageScheduleLinksDialog : Window
    {
        private Document _document;

        public ManageScheduleLinksDialog(Document document)
        {
            _document = document;
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

            var prevCurs = Cursor;
            Cursor = Cursors.Wait;

            ExternalCommand.ImportSchedule(schedule, workbookPath, worksheetName);

            Cursor = prevCurs;
        }

        private void ReloadFromButton_Click(object sender, RoutedEventArgs e)
        {
            string workbookPath;
            string worksheetName;

            using (var excelApplication = new NetOffice.ExcelApi.Application { DisplayAlerts = false })
            {
                var wnd = new UI.SheetSelectionDialog(excelApplication);

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

            var prevCurs = Cursor;
            Cursor = Cursors.Wait;
            ExternalCommand.ImportSchedule(schedule, workbookPath, worksheetName);
            ExternalCommand.StoreImportInformation(schedule, workbookPath, worksheetName);
            Cursor = prevCurs;

            ManagedScheduleLinksDataGrid.ItemsSource = ExternalCommand.GetManagedSchedules(_document);
            ManagedScheduleLinksDataGrid.Items.Refresh();
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedLink = (ManagedScheduleLink)ManagedScheduleLinksDataGrid.SelectedItem;
            var schedule = selectedLink.GetSchedule();

            var schema = ExternalCommand.GetSchema();
            Debug.Assert(schedule.DeleteEntity(schema));

            var prevCurs = Cursor;
            Cursor = Cursors.Wait;
            ManagedScheduleLinksDataGrid.ItemsSource = ExternalCommand.GetManagedSchedules(_document);
            Cursor = prevCurs;

            ManagedScheduleLinksDataGrid.Items.Refresh();
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }
    }
}
