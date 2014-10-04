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
            var categories = new ElementCategorizer().GroupByFamilyAndCategoryNames(familySymbols);

            var wnd = new MainWindow();
            wnd.MechanicalTree.ItemsSource = GenerateTreeViewData(categories);
            wnd.ElectricalTree.ItemsSource = GenerateTreeViewData(categories);

            wnd.ShowDialog();

            return Result.Succeeded;
        }

        private static List<TreeViewItemWithCheckbox> GenerateTreeViewData(
            IEnumerable<IGrouping<string, IGrouping<string, FamilySymbol>>> categories)
        {
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
            return categoryTreeViewItems;
        }
    }
}