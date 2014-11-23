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

                var collector = new FilteredElementCollector(document);

                var allElements = collector.OfClass(typeof (TextElement));
                var debugging = allElements.OfType<AnnotationSymbol>();
                var matchingElements = textFinder.FindMatchingElements(allElements);

                var resultsForm = new DebugFindResult();
                resultsForm.UpdateElements(matchingElements);
                resultsForm.ShowDialog();
                Globals.MatchingElementSet = matchingElements;                
            }
            return Result.Succeeded;
        }
    }
}
