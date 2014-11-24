using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using ElectricalToolSuite.ScheduleImport.Annotations;
using Microsoft.Practices.Prism.Commands;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    public class ManageLinksViewModel : INotifyPropertyChanged
    {
        public DelegateCommand<ManagedScheduleLink> EditCommand { get; private set; }
        public DelegateCommand<ManagedScheduleLink> RemoveCommand { get; private set; }
        public DelegateCommand RefreshCommand { get; private set; }
        public DelegateCommand<ManagedScheduleLink> ReloadCommand { get; private set; }
        public DelegateCommand ReloadAllCommand { get; private set; }


        public Document Document { get; set; }

        public ObservableCollection<ManagedScheduleLink> Links { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public ManageLinksViewModel(List<ManagedScheduleLink> links)
        {
            Links = new ObservableCollection<ManagedScheduleLink>(links);
            EditCommand = new DelegateCommand<ManagedScheduleLink>(Edit, IsValidSelection);
            RemoveCommand = new DelegateCommand<ManagedScheduleLink>(Remove, IsValidSelection);
            RefreshCommand = new DelegateCommand(Refresh, () => HasAnyLinks);
            ReloadCommand = new DelegateCommand<ManagedScheduleLink>(Reload, IsValidSelection);
            ReloadAllCommand = new DelegateCommand(ReloadAll, () => HasAnyLinks);
        }

        [NotifyPropertyChangedInvocator]
        public virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public bool HasAnyLinks
        {
            get { return Links.Count > 0; }
        }

        private bool IsValidSelection(ManagedScheduleLink link)
        {
            return link != null;
        }

        private void Edit(ManagedScheduleLink link)
        {
            string workbookPath;
            string worksheetName;
            string scheduleName;
            string scheduleType;
            
            using (var excelApplication = new NetOffice.ExcelApi.Application { DisplayAlerts = false })
            {
                var wnd = new SheetSelectionDialog(excelApplication, Document, link.GetSchedule());

                wnd.ScheduleNameTextBox.Text = link.ScheduleName;
                wnd.ScheduleTypeTextBox.Text = link.ScheduleType;
                wnd.FilePathTextBox.Text = link.WorkbookPath;
                            
                wnd.OkButton.Content = "Save";
                wnd.Title = "Edit Excel Schedule Link";
            
                if (wnd.ShowDialog() != true)
                    return;
            
                workbookPath = wnd.FilePathTextBox.Text;
                worksheetName = (string) wnd.SheetComboBox.SelectedItem;
                scheduleName = wnd.ScheduleNameTextBox.Text;
                scheduleType = wnd.ScheduleTypeTextBox.Text;
            
                excelApplication.Quit();
            }

            link.WorkbookPath = workbookPath;
            link.WorksheetName = worksheetName;
            link.ScheduleName = scheduleName;
            link.ScheduleType = scheduleType;
            link.Reload();
        }

        private void Remove(ManagedScheduleLink link)
        {
            Links.Remove(link);
            LinkGateway.DeleteLink(link.GetSchedule());
            Document.Delete(link.GetSchedule().Id);
        }

        private void Refresh()
        {
            foreach (var link in Links)
                link.OnPropertyChanged("StatusText");
            OnPropertyChanged("Links");
        }

        private void Reload(ManagedScheduleLink link)
        {
            link.Reload();
        }

        private void ReloadAll()
        {
            foreach (var link in Links)
                link.Reload();
        }
    }
}
