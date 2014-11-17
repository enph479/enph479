using System;
using System.Linq;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
            var textFinder = TextFinderBuilder.BuildTextFinder(findForm.GetFinderSettings());

            var uiDocument = commandData.Application.ActiveUIDocument;
            var document = uiDocument.Document;
            
            var collector = new FilteredElementCollector(document);

            //currently getting ALL elements is really slow.
            var allElements = collector.OfClass(typeof (FamilyInstance));
            var matchingElements = textFinder.FindMatchingElements(allElements);

            var resultsForm = new FindResultsWindow();
            resultsForm.UpdateElements(matchingElements);
            Globals.MatchingElementSet = matchingElements;
            //resultsForm.Show(); //right now passes the data to the idling loop

            return Result.Succeeded;
        }
    }
}
