using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElectricalToolSuite.FindAndReplace.UI;

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
                var dPid = new DockablePaneId(DockConstants.Id);
                var pane = commandData.Application.GetDockablePane(dPid);
                pane.Show();
                
                var textFinder = TextFinderBuilder.BuildTextFinder(findForm.GetFinderSettings(), document);

                var elementCollector = new FilteredElementCollector(document);
                var textCollector = new FilteredElementCollector(document);

                var allFamilies = elementCollector.OfClass(typeof (FamilyInstance));
                var allTextBoxes = textCollector.OfClass(typeof (TextElement));

                var matchingElements = textFinder.FindMatchingElements(allFamilies, allTextBoxes, findForm.GetSearchableViews());

                Globals.MatchingElementSet = matchingElements;   

            }
            return Result.Succeeded;
        }
    }
}
