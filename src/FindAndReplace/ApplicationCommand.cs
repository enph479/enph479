using System;
using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using ElectricalToolSuite.FindAndReplace.Composition;
using Redbolts.DockableUITest.UI;

namespace ElectricalToolSuite.FindAndReplace
{
    public class ApplicationCommand:IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {

            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            
            var container = CommandContainer.Instance();
            if (container.Valid)
            {
                container.BuildRibbon(application);
            }

            var dPid = new DockablePaneId(DockConstants.Id);
            if (!DockablePane.PaneIsRegistered(dPid))
            {
                var state = new DockablePaneState {DockPosition = DockPosition.Right};
                var element = new FindResultsWindow();
                application.RegisterDockablePane2(DockConstants.Id, DockConstants.Name, element, state);
            }
            return Result.Succeeded;
        }

        void OnIdling(object sender, IdlingEventArgs e)
        {
            //should be able to call revit API stuff here?
            // access active document from sender:

            var app = sender as UIApplication;

            if (app != null && Globals.SelectedElement != null)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                var doc = uidoc.Document;
                uidoc.Selection.SetElementIds(new Collection<ElementId> {Globals.SelectedElement});
                Globals.SelectedElement = null;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
