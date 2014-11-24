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
            var findForm = new FindAndReplaceWindow();
            findForm.ShowDialog();
            if (findForm.NotCancelled)
            {
                //force show the dockable pane on search
                //var dPid = new DockablePaneId(DockConstants.Id);
                //var pane = commandData.Application.GetDockablePane(dPid);
                //pane.Show();
                
                var uiDocument = commandData.Application.ActiveUIDocument;
                var document = uiDocument.Document;
                
                var textFinder = TextFinderBuilder.BuildTextFinder(findForm.GetFinderSettings(), document);

                var elementCollector = new FilteredElementCollector(document);
                var textCollector = new FilteredElementCollector(document);

                var allFamilies = elementCollector.OfClass(typeof (FamilyInstance));
                var allTextBoxes = textCollector.OfClass(typeof (TextElement));

                var matchingElements = textFinder.FindMatchingElements(allFamilies, allTextBoxes);

                var resultsForm = new DebugFindResult();
                resultsForm.UpdateElements(matchingElements);
                resultsForm.ShowDialog();
                //Globals.MatchingElementSet = matchingElements;     
                View currentView = uiDocument.ActiveView;
                UIView uiview = null;
                IList<UIView> uiviews = uiDocument.GetOpenUIViews();

                foreach (UIView uv in uiviews)
                {
                    if (!uv.ViewId.Equals(currentView.Id)) continue;
                    uiview = uv;
                    break;
                }
                var elem = document.GetElement(Globals.SelectedElement);
                var boundingbox = elem.get_BoundingBox(uiDocument.ActiveView);
                uiview.ZoomAndCenterRectangle(boundingbox.get_Bounds(0), boundingbox.get_Bounds(1));

            }
            return Result.Succeeded;
        }
    }
}
