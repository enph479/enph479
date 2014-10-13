using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace FindAndReplace
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ExternalCommand : IExternalCommand
    {
        public ElementSet FindMatchingElements(FilteredElementCollector allElements, String matchingString)
        {
            var matchingElements = new ElementSet();
            try
            {
                foreach (Element elem in allElements)
                {
                    foreach (Parameter param in elem.Parameters)
                    {
                        if (!matchingString.Equals(param.AsString()))
                            continue;
                        matchingElements.Insert(elem);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return matchingElements;
        }

        public Result Execute(ExternalCommandData commandData,
                ref string message, ElementSet elements)
        {
            try
            {
                const string matchingString = " 208 V/3-0 VA";
                String info = "the matching elements are: \n";

                var uiDocument = commandData.Application.ActiveUIDocument;
                var document = uiDocument.Document;

                var collector = new FilteredElementCollector(document);

                //currently getting ALL elements is really slow. But who cares, we can fix this later. (threads? filtering? screw the user, who cares about performance?)
                var allElements = collector.WherePasses(new LogicalOrFilter(new ElementIsElementTypeFilter(false),
                    new ElementIsElementTypeFilter(true)));

                var matchingElements = FindMatchingElements(allElements, matchingString);

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
