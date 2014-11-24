using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB.Electrical;
using System;
using ElectricalToolSuite.ScheduleImport.Annotations;
using Revit = Autodesk.Revit.Exceptions;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ManagedScheduleLink : INotifyPropertyChanged, IDataErrorInfo
    {
        private readonly PanelScheduleView _schedule;
        private string _temporaryName;

        public ManagedScheduleLink(PanelScheduleView schedule)
        {
            if (!LinkGateway.IsLinked(schedule))
                LinkGateway.CreateLink(schedule, "", "", "");

            _schedule = schedule;
        }

        public string ScheduleName
        {
            get
            {
                if (Error != "")
                    return _temporaryName;
                return _schedule.ViewName; 
            }
            set
            {
                try
                {
                    _schedule.ViewName = value;
                    OnPropertyChanged("ScheduleName");
                }
                catch (Revit.ArgumentException)
                {
                    Error = "There is already a schedule with this name.";
                    _temporaryName = value;
                    return;
                }
                Error = "";
            }
        }

        public string ScheduleType
        {
            get { return LinkGateway.GetScheduleType(_schedule); }
            set
            {
                LinkGateway.SetScheduleType(_schedule, value);
                OnPropertyChanged("ScheduleType");
            }
        }

        public string WorkbookPath
        {
            get { return LinkGateway.GetWorkbookPath(_schedule); }
            set
            {
                LinkGateway.SetWorkbookPath(_schedule, value);
                OnPropertyChanged("WorkbookPath");
                OnPropertyChanged("WorkbookExists");
                OnPropertyChanged("StatusText");
            }
        }

        public string WorksheetName
        {
            get { return LinkGateway.GetWorksheetName(_schedule); }
            set 
            { 
                LinkGateway.SetWorksheetName(_schedule, value);
                OnPropertyChanged("WorksheetName");
            }
        }

        public PanelScheduleView GetSchedule()
        {
            return _schedule;
        }

        public void Reload()
        {
            if (WorkbookExists)
                ExternalCommand.ImportSchedule(_schedule, WorkbookPath, WorksheetName);            
        }

        public string StatusText
        {
            get
            {
                if (WorkbookExists)
                    return "Loaded";
                return "Not found";
            }
        }

        public bool WorkbookExists
        {
            get { return File.Exists(WorkbookPath); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        public virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public string this[string columnName]
        {
            get { return Error; }
        }

        public string Error { get; private set; }
    }
}
