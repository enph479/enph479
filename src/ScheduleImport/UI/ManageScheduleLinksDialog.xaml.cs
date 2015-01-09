﻿using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    /// <summary>
    /// Interaction logic for ManageScheduleLinksDialog.xaml
    /// </summary>
    public partial class ManageScheduleLinksDialog
    {
        private readonly Document _document;

        public bool PressedCreate { get; private set; }

        public ManageScheduleLinksDialog(Document document)
        {
            _document = document;
            InitializeComponent();
        }

        public void UpdateButtons()
        {
            bool hasSelection = ManagedScheduleLinksDataGrid.SelectedIndex != -1;

            if (hasSelection)
            {
                EditButton.IsEnabled = true;
                RemoveButton.IsEnabled = true;
                ReloadButton.IsEnabled =
                    ((ManagedScheduleLink) ManagedScheduleLinksDataGrid.SelectedItem).WorkbookExists;
            }
            else
            {
                ReloadButton.IsEnabled = false;
                EditButton.IsEnabled = false;
                RemoveButton.IsEnabled = false;
            }

            var anyToReload = ManagedScheduleLinksDataGrid.Items
                .Cast<ManagedScheduleLink>()
                .Any(link => link.WorkbookExists);

            ReloadAllButton.IsEnabled = anyToReload;

            var any = ManagedScheduleLinksDataGrid.Items.Cast<ManagedScheduleLink>().Any();
            RefreshButton.IsEnabled = any;
        }

        private void ManagedScheduleLinksDataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            UpdateButtons();
        }

        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedLink = (ManagedScheduleLink) ManagedScheduleLinksDataGrid.SelectedItem;

            selectedLink.Reload();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            string workbookPath;
            string worksheetName;
            string scheduleName;
            string scheduleType;

            var selectedLink = (ManagedScheduleLink) ManagedScheduleLinksDataGrid.SelectedItem;

            using (var excelApplication = new NetOffice.ExcelApi.Application {DisplayAlerts = false})
            {
                var wnd = new SheetSelectionDialog(excelApplication, _document);

                wnd.ScheduleNameTextBox.Text = selectedLink.ScheduleName;
                wnd.ScheduleTypeTextBox.Text = selectedLink.ScheduleType;

                wnd.FilePathTextBox.Text = selectedLink.WorkbookPath;

                wnd.OkButton.Content = "Save";

                wnd.ValidName = selectedLink.ScheduleName;

                wnd.Title = "Edit Excel Schedule Link";

                if (wnd.ShowDialog() != true)
                    return;

                workbookPath = wnd.FilePathTextBox.Text;
                worksheetName = (string) wnd.SheetComboBox.SelectedItem;
                scheduleName = wnd.ScheduleNameTextBox.Text;
                scheduleType = wnd.ScheduleTypeTextBox.Text;


                excelApplication.Quit();
            }

            selectedLink.WorkbookPath = workbookPath;
            selectedLink.WorksheetName = worksheetName;
            selectedLink.ScheduleName = scheduleName;
            selectedLink.ScheduleType = scheduleType;
            selectedLink.Reload();
            ManagedScheduleLinksDataGrid.ItemsSource = ExternalCommand.GetManagedSchedules(_document);
            ManagedScheduleLinksDataGrid.Items.Refresh();

            UpdateButtons();
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedLink = (ManagedScheduleLink) ManagedScheduleLinksDataGrid.SelectedItem;
            var schedule = selectedLink.GetSchedule();

            LinkGateway.DeleteLink(schedule);
            _document.Delete(schedule.Id);

            ManagedScheduleLinksDataGrid.ItemsSource = ExternalCommand.GetManagedSchedules(_document);

            ManagedScheduleLinksDataGrid.Items.Refresh();
            UpdateButtons();
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
            UpdateButtons();
        }

        private void ReloadAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (ManagedScheduleLink link in ManagedScheduleLinksDataGrid.Items)
            {
                link.Reload();
            }

            UpdateButtons();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            ManagedScheduleLinksDataGrid.Items.Refresh();
        }
    }
}