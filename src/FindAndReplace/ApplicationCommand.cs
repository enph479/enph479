using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace ElectricalToolSuite.FindAndReplace
{
    public class ApplicationCommand:IExternalApplication
    {
        private FindResultsWindow _resultsWindow;
        public Result OnStartup(UIControlledApplication application)
        {
            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Find Tool");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("FindCmd",
            "Find Tool", thisAssemblyPath, "ElectricalToolSuite.FindAndReplace.ExternalCommand");

            ribbonPanel.AddItem(buttonData);

            application.Idling += OnIdling;
            application.Idling += ForceHideDockablePane;

            var dPid = new DockablePaneId(DockConstants.Id);
            if (!DockablePane.PaneIsRegistered(dPid))
            {
                var state = new DockablePaneState();
                _resultsWindow = new FindResultsWindow();
                DockExtensions.RegisterDockablePane2(application, DockConstants.Id, DockConstants.Name, _resultsWindow, state);
            }
            return Result.Succeeded;
        }

        private void ForceHideDockablePane(object sender, IdlingEventArgs e)
        {
            var app = sender as UIApplication;
            var dPid = new DockablePaneId(DockConstants.Id);
            if (app != null)
            {
                var pane = app.GetDockablePane(dPid);
                pane.Hide();
                app.Idling -= ForceHideDockablePane;
            }
        }

        void OnIdling(object sender, IdlingEventArgs e)
        {
            var app = sender as UIApplication;

            if (app != null && Globals.MatchingElementSet != null)
            {
                _resultsWindow.UpdateElements(Globals.MatchingElementSet);
                Globals.MatchingElementSet = null;
            }
            if (app != null && Globals.SelectedElement != null)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                uidoc.Selection.SetElementIds(new Collection<ElementId> {Globals.SelectedElement});
                var document = uidoc.Document;
                var selectedElement = document.GetElement(Globals.SelectedElement);
                selectedElement.get_BoundingBox(uidoc.ActiveView);

                //Changing the views and stuff
                View currentView = uidoc.ActiveView;
                UIView uiview = null;
                IList<UIView> uiviews = uidoc.GetOpenUIViews(); //this is dumb but is the way thebuildingcoder does it

                foreach (UIView uv in uiviews)
                {
                    if (!uv.ViewId.Equals(currentView.Id)) continue;
                    uiview = uv;
                    break;
                }
                if (Globals.SelectedElement != null && uiview != null)
                {
                    var elem = document.GetElement(Globals.SelectedElement);
                    var boundingbox = elem.get_BoundingBox(uidoc.ActiveView);
                    uiview.ZoomAndCenterRectangle(boundingbox.get_Bounds(0), boundingbox.get_Bounds(1));
                }
                Globals.SelectedElement = null;
            }
        }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
