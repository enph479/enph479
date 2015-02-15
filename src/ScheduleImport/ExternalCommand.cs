using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Data;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using ElectricalToolSuite.ScheduleImport.UI;
using Excel = NetOffice.ExcelApi;
using TextBox = System.Windows.Controls.TextBox;

namespace ElectricalToolSuite.ScheduleImport
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            while (true)
            {
                var mgWnd = new ManageScheduleLinksDialog(doc, uiDoc);
                var managedSchedules = GetManagedSchedules(doc);
                mgWnd.ManagedScheduleLinksDataGrid.ItemsSource = managedSchedules;
                mgWnd.UpdateButtons();
                mgWnd.ShowDialog();

                if (mgWnd.PressedCreate)
                {
                    CreateLink(uiDoc, doc);
                }
                else
                {
                    break;
                }
            }

            return Result.Succeeded;
        }

        private static void CreateLink(UIDocument uiDoc, Document doc)
        {
            var selectedPanel = new ElementSelector(uiDoc).SelectSingle() as FamilyInstance;

            if (selectedPanel == null)
            {
                TaskDialog.Show("Invalid selection", "Please select a panel.");
                {
                    return;
                }
            }

            var wnd = new SheetSelectionDialog(doc);
            wnd.ScheduleNameTextBox.Text = selectedPanel.Name;
//            wnd.ScheduleName = selectedPanel.Name;
            wnd.ScheduleNameTextBox.GetBindingExpression(TextBox.TextProperty).UpdateSource();

            if (wnd.ShowDialog() != true)
            {
                return;
            }

            var workbookPath = wnd.WorkbookPath;
            var worksheetName = (string) wnd.SheetComboBox.SelectedItem;
            var scheduleType = wnd.ScheduleTypeTextBox.Text;
            var scheduleName = wnd.ScheduleName;

            var schedule = PanelScheduleView.CreateInstanceView(doc, selectedPanel.Id);
            schedule.ViewName = scheduleName;

            ManagedScheduleLink.CreateNew(schedule, workbookPath, worksheetName, scheduleType);
        }

        public static List<ManagedScheduleLink> GetManagedSchedules(Document doc)
        {
            var linkedSchedules = new FilteredElementCollector(doc).OfClass(typeof (PanelScheduleView))
                .Cast<PanelScheduleView>()
                .Where(LinkGateway.IsLinked);

            var managedSchedules =
                (from sched in linkedSchedules
                    select new ManagedScheduleLink(sched))
                    .OrderBy(l => l.ScheduleName)
                    .ToList();

            return managedSchedules;
        }
    }
}
