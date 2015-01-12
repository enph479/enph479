using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElectricalToolSuite.ScheduleImport.UI
{
    public class ManagedScheduleLink
    {
        private PanelScheduleView _schedule;

        public ManagedScheduleLink(PanelScheduleView schedule, string workbookPath, string worksheetName)
        {
            _schedule = schedule;
            WorkbookPath = workbookPath;
            WorksheetName = worksheetName;
        }

        public string ScheduleName { get { return _schedule.Name; }}

        public string WorkbookPath { get; private set; }

        public string WorksheetName { get; private set; }

        public PanelScheduleView GetSchedule()
        {
            return _schedule;
        }
    }
}
