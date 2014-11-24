using System.IO;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ManagedScheduleLink
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
            }
        }

        public string ScheduleType
        {
            get { return LinkGateway.GetScheduleType(_schedule); }
            set { LinkGateway.SetScheduleType(_schedule, value); }
        }

        public string WorkbookPath
        {
            get { return LinkGateway.GetWorkbookPath(_schedule); }
            set { LinkGateway.SetWorkbookPath(_schedule, value); }
        }

        public string WorksheetName
        {
            get { return LinkGateway.GetWorksheetName(_schedule); }
            set { LinkGateway.SetWorksheetName(_schedule, value); }
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
    }
}
