using System;
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
            try
            {
                var findForm = new FindAndReplaceUi();
                findForm.ShowDialog();
                var textFinder = TextFinderBuilder.BuildTextFinder(findForm.GetFinderSettings());
                
                String info = "the matching elements are: \n";

                var uiDocument = commandData.Application.ActiveUIDocument;
                var document = uiDocument.Document;

                var collector = new FilteredElementCollector(document);

                //currently getting ALL elements is really slow.
                var allElements = collector.WherePasses(new LogicalOrFilter(new ElementIsElementTypeFilter(false),
                    new ElementIsElementTypeFilter(true)));

                var matchingElements = textFinder.FindMatchingElements(allElements);

                foreach (Element elem in matchingElements)
                {
                    info += elem.Name + "\n";
                }
                TaskDialog.Show("Revit", info);
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
