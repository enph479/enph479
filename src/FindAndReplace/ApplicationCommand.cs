using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Redbolts.DockableUITest.UI;

namespace ElectricalToolSuite.FindAndReplace
{
    public class ApplicationCommand:IExternalApplication
    {
        private FindResultsWindow _findResultsWindow;
        public Result OnStartup(UIControlledApplication application)
        {
            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("NewRibbonPanel");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("FindCmd",
            "Find Tool", thisAssemblyPath, "ElectricalToolSuite.FindAndReplace.ExternalCommand");

            ribbonPanel.AddItem(buttonData);

            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            application.Idling += new EventHandler<IdlingEventArgs>(ForceShowDockablePane);

            var dPid = new DockablePaneId(DockConstants.Id);
            if (!DockablePane.PaneIsRegistered(dPid))
            {
                var state = new DockablePaneState();
                _findResultsWindow = new FindResultsWindow();
                DockExtensions.RegisterDockablePane2(application, DockConstants.Id, DockConstants.Name, _findResultsWindow, state);
            }
            return Result.Succeeded;
        }

        private void ForceShowDockablePane(object sender, IdlingEventArgs e)
        {
            var app = sender as UIApplication;
            var dPid = new DockablePaneId(DockConstants.Id);
            try
            {
                var pane = app.GetDockablePane(dPid);
                pane.Show();
                app.Idling -= ForceShowDockablePane;
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception.Message);
            }
        }

        void OnIdling(object sender, IdlingEventArgs e)
        {
            var app = sender as UIApplication;

            if (app != null && Globals.MatchingElementSet != null)
            {
                //UIDocument uidoc = app.ActiveUIDocument;
                //uidoc.Selection.SetElementIds(new Collection<ElementId> {Globals.SelectedElement});
                _findResultsWindow.UpdateElements(Globals.MatchingElementSet);
                Globals.MatchingElementSet = null;
            }
        }

        

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
