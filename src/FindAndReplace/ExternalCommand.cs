using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

namespace ElectricalToolSuite.FindAndReplace
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                ref string message, ElementSet elements)
        {
            var uiDocument = commandData.Application.ActiveUIDocument;
            var document = uiDocument.Document;

            var viewCollector = new FilteredElementCollector(document);
            var allViews = viewCollector.OfClass(typeof(View));

            var findForm = new FindAndReplaceWindow(allViews, document.ActiveView);
            findForm.ShowDialog();
            if (findForm.NotCancelled)
            {
                //force show the dockable pane on search
                //var dPid = new DockablePaneId(DockConstants.Id);
                //var pane = commandData.Application.GetDockablePane(dPid);
                //pane.Show();
                
                var textFinder = TextFinderBuilder.BuildTextFinder(findForm.GetFinderSettings(), document);

                var elementCollector = new FilteredElementCollector(document);
                var textCollector = new FilteredElementCollector(document);

                var allFamilies = elementCollector.OfClass(typeof (FamilyInstance));
                var allTextBoxes = textCollector.OfClass(typeof (TextElement));

                var matchingElements = textFinder.FindMatchingElements(allFamilies, allTextBoxes, findForm.GetSearchableViews());

                var resultsForm = new DebugFindResult();
                resultsForm.UpdateElements(matchingElements);
                resultsForm.ShowDialog();
                //Globals.MatchingElementSet = matchingElements;   
  

                //Changing the views and stuff
                View currentView = uiDocument.ActiveView;
                UIView uiview = null;
                IList<UIView> uiviews = uiDocument.GetOpenUIViews(); //this is dumb but is the way thebuildingcoder does it

                foreach (UIView uv in uiviews)
                {
                    if (!uv.ViewId.Equals(currentView.Id)) continue;
                    uiview = uv;
                    break;
                }
                if (Globals.SelectedElement != null)
                {
                    var elem = document.GetElement(Globals.SelectedElement);
                    var boundingbox = elem.get_BoundingBox(uiDocument.ActiveView);
                    uiview.ZoomAndCenterRectangle(boundingbox.get_Bounds(0), boundingbox.get_Bounds(1));
                }

            }
            return Result.Succeeded;
        }
    }
}
