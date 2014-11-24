using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB.Electrical;
using System;
using ElectricalToolSuite.ScheduleImport.Annotations;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ManagedScheduleLink : INotifyPropertyChanged
    {
        private PanelScheduleView _schedule;

        public ManagedScheduleLink(PanelScheduleView schedule)
        {
            if (!LinkGateway.IsLinked(schedule))
                throw new ArgumentException("Schedule is not linked to an Excel sheet", "schedule");

            _schedule = schedule;
        }

        public string ScheduleName
        {
            get
            {
                return _schedule.ViewName; 
            }
            set
            {
                _schedule.ViewName = value;
                OnPropertyChanged("ScheduleName");
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
    }
}
