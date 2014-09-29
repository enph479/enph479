using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElectricalToolSuite.MECoordination.UI;

namespace ElectricalToolSuite.MECoordination
{
    [Transaction(TransactionMode.Automatic)]
    public class ExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            var familySymbols = new FilteredElementCollector(doc).OfClass(typeof (FamilySymbol)).Cast<FamilySymbol>();
//            var families = new FilteredElementCollector(doc).OfClass(typeof (Family)).Cast<Family>();
            var categories = new ElementCategorizer().GroupByFamilyAndCategoryNames(familySymbols);
            var categoryTreeViewItems = new List<TreeViewItemWithCheckbox>();

            foreach (var categoryGroup in categories)
            {
                var categoryItem = new TreeViewItemWithCheckbox(categoryGroup.Key);
                foreach (var familyGroup in categoryGroup)
                {
                    var familyItem = new TreeViewItemWithCheckbox(familyGroup.Key);
                    foreach (var familySymbol in familyGroup)
                    {
                        var familySymbolItem = new FamilySymbolItem(familySymbol.Name, familySymbol);
                        familyItem.AddChild(familySymbolItem);
                    }
                    categoryItem.AddChild(familyItem);
                }
                categoryTreeViewItems.Add(categoryItem);
            }

            var wnd = new TreeViewMultipleTemplatesSample();
            wnd.Categories.ItemsSource = categoryTreeViewItems;

            wnd.ShowDialog();

            return Result.Succeeded;
        }
    }
}
