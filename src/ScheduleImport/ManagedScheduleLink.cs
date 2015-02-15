using System;
using System.IO;
using Autodesk.Revit.DB.Electrical;

namespace ElectricalToolSuite.ScheduleImport
{
    public class ManagedScheduleLink
    {
        private readonly PanelScheduleView _schedule;

        public static ManagedScheduleLink CreateNew(PanelScheduleView schedule, string workbookPath,
            string worksheetName, string scheduleType)
        {
            var importer = new ScheduleImporter(schedule);
            importer.ImportFromFile(workbookPath, worksheetName);
            LinkGateway.CreateLink(schedule, workbookPath, worksheetName, scheduleType);
            return new ManagedScheduleLink(schedule);
        }

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
            {
                var importer = new ScheduleImporter(_schedule);
                importer.ImportFromFile(WorkbookPath, WorksheetName);
            }
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
