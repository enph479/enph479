using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using ElectricalToolSuite.ScheduleImport.UI;
using Excel = NetOffice.ExcelApi;

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

            string workbookPath;
            string worksheetName;
            string scheduleName;
            string scheduleType;

            using (var excelApplication = new Excel.Application {DisplayAlerts = false})
            {
                var wnd = new SheetSelectionDialog(excelApplication, doc);

                wnd.ScheduleNameTextBox.Text = selectedPanel.Name;

                if (wnd.ShowDialog() != true)
                {
                    return;
                }

                workbookPath = wnd.FilePathTextBox.Text;
                worksheetName = (string) wnd.SheetComboBox.SelectedItem;
                scheduleType = wnd.ScheduleTypeTextBox.Text;
                scheduleName = wnd.ScheduleNameTextBox.Text;

                excelApplication.Quit();
            }

            PanelScheduleView schedule = PanelScheduleView.CreateInstanceView(doc, selectedPanel.Id);
            schedule.ViewName = scheduleName;

            var importer = new ScheduleImporter(schedule);
            importer.ImportFromFile(workbookPath, worksheetName);

            StoreImportInformation(schedule, workbookPath, worksheetName, scheduleType);
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

        public static void StoreImportInformation(PanelScheduleView schedule, string workbookPath, string worksheetName, string scheduleType)
        {
            LinkGateway.CreateLink(schedule, workbookPath, worksheetName, scheduleType);
        }
    }
}
