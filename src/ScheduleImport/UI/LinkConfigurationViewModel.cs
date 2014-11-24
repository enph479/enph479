using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using ElectricalToolSuite.ScheduleImport.Annotations;
using Microsoft.Practices.Prism.Commands;
using Microsoft.Win32;
using Excel = NetOffice.ExcelApi;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    public class LinkConfigurationViewModel : INotifyPropertyChanged
    {
        private Excel.Application _excelApp;

        public DelegateCommand OpenWorkbookCommand { get; set; }

//        public DelegateCommand FileDialogCommand { get; set; }

        public ManagedScheduleLink Link { get; set; }

        public IList<string> Worksheets
        {
            get
            {
                if (!Link.WorkbookExists)
                    return Enumerable.Empty<string>().ToList();

                using (var workbook = _excelApp.Workbooks.Open(Link.WorkbookPath, false, true))
                {
                    var sheets = workbook.Worksheets.Cast<Excel.Worksheet>().Select(s => s.Name).ToList();
                    workbook.Close();
                    return sheets;
                }
            }
        }

        public LinkConfigurationViewModel(Document document, Excel.Application excelApplication, ManagedScheduleLink link)
        {
            _excelApp = excelApplication;
            Link = link;
            OpenWorkbookCommand = new DelegateCommand(OpenWorkbook, () => Link.WorkbookExists);

            Link.PropertyChanged += Link_PropertyChanged;
//            FileDialogCommand = new DelegateCommand(OpenFileDialog, () => true);
        }

        void Link_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "WorkbookPath")
            {
                OnPropertyChanged("Worksheets");
                OpenWorkbookCommand.RaiseCanExecuteChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        private void OpenWorkbook()
        {
            Process.Start(Link.WorkbookPath);
        }

        public void SheetSelectionDialog_Closing(object sender, CancelEventArgs e)
        {
//            if (_schedule != null && scheduleName == _schedule.Name)
//                return;
//
//            var existingSchedule =
//                new FilteredElementCollector(_document).OfClass(typeof(PanelScheduleView))
//                    .Cast<PanelScheduleView>()
//                    .Any(s => s.Name == scheduleName);
//
//            if (!existingSchedule)
//                return;
//
//            TaskDialog.Show("Cannot rename schedule",
//                String.Format("There is already a panel schedule with name '{0}.'", scheduleName));
//
//            e.Cancel = true;
//            Focus();
        }
    }
}
